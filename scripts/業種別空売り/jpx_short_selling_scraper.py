#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JPX 空売り集計（業種別集計）日次データ → Googleスプレッドシート 自動追記スクリプト

取得元 : https://www.jpx.co.jp/markets/statistics-equities/short-selling/index.html
環境変数: GOOGLE_SERVICE_ACCOUNT_JSON（サービスアカウントJSONの中身をそのまま）

日付指定（省略時は最新）:
  コマンドライン : python jpx_short_selling_scraper.py 2026/08/05
  環境変数       : TARGET_DATE=2026/08/05 python jpx_short_selling_scraper.py
  ※ JPXの一覧ページには直近5営業日分しか掲載されないため、それより古い日付は取得不可。
"""

import io
import json
import os
import re
import sys
from typing import Optional

import gspread
import pdfplumber
import requests
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials


# ================================================================
#  定数定義
# ================================================================

# --- JPX サイト ---
JPX_INDEX_URL = "https://www.jpx.co.jp/markets/statistics-equities/short-selling/index.html"
JPX_BASE_URL  = "https://www.jpx.co.jp"
HTTP_TIMEOUT  = 30  # 秒
HTTP_HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}

# --- Googleスプレッドシート ---
SPREADSHEET_ID = "1kWST0CkkIvo3irPSbMgtVtUqqRDXFwvRREYZDRQAFMY"

#   PDFカテゴリ(a/b/c/d) → 書き込み先シートGID
SHEET_MAP = {
    "a": {"gid": 162001397,  "name": "実注文"},
    "b": {"gid": 692441937,  "name": "空売り（価格規制あり）"},
    "c": {"gid": 1306991615, "name": "空売り（価格規制なし）"},
    "d": {"gid": 668166970,  "name": "合計"},
}

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# --- 業種リスト（PDF掲載順）33業種 ＋ その他（33業種外）---
INDUSTRY_ORDER = [
    "水産・農林業",         "鉱業",               "建設業",
    "食料品",               "繊維製品",           "パルプ・紙",
    "化学",                 "医薬品",             "石油・石炭製品",
    "ゴム製品",             "ガラス・土石製品",   "鉄鋼",
    "非鉄金属",             "金属製品",           "機械",
    "電気機器",             "輸送用機器",         "精密機器",
    "その他製品",           "電気・ガス業",       "陸運業",
    "海運業",               "空運業",             "倉庫・運輸関連業",
    "情報・通信業",         "卸売業",             "小売業",
    "銀行業",               "証券、商品先物取引業", "保険業",
    "その他金融業",         "不動産業",           "サービス業",
    "その他（33業種外）",
]

# --- 正規表現 ---
#   業種名に数字が含まれる「その他（33業種外）」対策で
#   INDUSTRY_ORDER の完全一致リスト（長い順）からパターンを生成する
_INDUSTRY_PATTERN = "|".join(
    re.escape(n) for n in sorted(INDUSTRY_ORDER, key=len, reverse=True)
)
# 例: "水産・農林業 2,091 54.3% 1,693 44.0% 67 1.7% 3,852"
ROW_RE = re.compile(
    rf"^(?P<name>{_INDUSTRY_PATTERN})\s+"
    r"(?P<a>[\d,]+)\s+[\d.]+%\s+"
    r"(?P<b>[\d,]+)\s+[\d.]+%\s+"
    r"(?P<c>[\d,]+)\s+[\d.]+%\s+"
    r"(?P<d>[\d,]+)\s*$"
)
DATE_RE = re.compile(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日")


# ================================================================
#  PDF 取得・パース
# ================================================================

def get_pdf_url(target_date: Optional[str] = None) -> str:
    """
    一覧ページから業種別集計PDF URLを取得する。
    target_date が None なら最新日付、"YYYY/MM/DD" 形式で指定すると該当日付を返す。
    ※ 一覧には直近5営業日分のみ掲載。それより古い日付はエラーになる。
    """
    resp = requests.get(JPX_INDEX_URL, headers=HTTP_HEADERS, timeout=HTTP_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding
    soup = BeautifulSoup(resp.text, "html.parser")

    table = soup.find("table")
    if not table:
        raise RuntimeError("一覧テーブルが見つかりません。サイト構造が変更された可能性があります。")

    # テーブルの全行を {日付: URL} でマッピング
    date_url_map = {}
    for row in table.find_all("tr"):
        links = row.find_all("a")
        cells = row.find_all("td")
        if not links or not cells:
            continue
        date_text = cells[0].get_text(strip=True)  # 例: "2026/08/07"
        href = next(
            (a.get("href", "") for a in links if "-g.pdf" in a.get("href", "")),
            links[1].get("href") if len(links) >= 2 else None,
        )
        if date_text and href:
            date_url_map[date_text] = href if href.startswith("http") else JPX_BASE_URL + href

    if not date_url_map:
        raise RuntimeError("PDFリンクを含む行が見つかりませんでした。")

    # 日付未指定 → 先頭行（最新）
    if target_date is None:
        return next(iter(date_url_map.values()))

    if target_date not in date_url_map:
        available = ", ".join(date_url_map.keys())
        raise RuntimeError(
            f"指定日付 '{target_date}' が一覧に見つかりません。\n"
            f"取得可能な日付: {available}"
        )

    return date_url_map[target_date]


def download_pdf(url: str) -> bytes:
    resp = requests.get(url, headers=HTTP_HEADERS, timeout=HTTP_TIMEOUT)
    resp.raise_for_status()
    return resp.content


def parse_pdf(pdf_bytes: bytes) -> tuple[str, dict]:
    """
    PDF から日付と業種別データを抽出する。

    Returns:
        target_date : "YYYY/MM/DD"
        data        : {業種名: {"a": int, "b": int, "c": int, "d": int}}  単位: 百万円
    """
    full_text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            full_text += (page.extract_text(layout=True) or page.extract_text() or "") + "\n"

    # 日付抽出
    date_m = DATE_RE.search(full_text)
    if not date_m:
        raise RuntimeError("PDFから日付を抽出できませんでした。")
    y, m, d = date_m.groups()
    target_date = f"{int(y):04d}/{int(m):02d}/{int(d):02d}"

    # 業種データ抽出
    data = {}
    for line in full_text.splitlines():
        row_m = ROW_RE.match(line.strip())
        if not row_m:
            continue
        name = row_m.group("name")
        data[name] = {
            key: int(row_m.group(key).replace(",", ""))
            for key in ("a", "b", "c", "d")
        }

    missing = [n for n in INDUSTRY_ORDER if n not in data]
    if missing:
        raise RuntimeError(
            f"以下の業種を抽出できませんでした（PDFレイアウト変更の可能性）: {missing}\n"
            f"--- 抽出テキスト（デバッグ用） ---\n{full_text}"
        )

    return target_date, data


# ================================================================
#  Googleスプレッドシート 書き込み
# ================================================================

def get_gspread_client() -> gspread.Client:
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        raise RuntimeError("環境変数 GOOGLE_SERVICE_ACCOUNT_JSON が設定されていません。")
    creds = Credentials.from_service_account_info(json.loads(sa_json), scopes=GOOGLE_SCOPES)
    return gspread.authorize(creds)


def col_letter(col_idx: int) -> str:
    """列番号（1始まり）→ A1形式の列名  例: 1→A, 26→Z, 27→AA"""
    result = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def ensure_industry_column(ws: gspread.Worksheet) -> None:
    """A列の業種名が未設定の場合のみ初期書き込みする。"""
    if len(ws.col_values(1)) >= len(INDUSTRY_ORDER) + 1:
        return  # 設定済みならスキップ
    header = [["業種（単位：百万円）", ""]]
    rows   = [[name, ""] for name in INDUSTRY_ORDER]
    ws.update(range_name="A1", values=header + rows)
    print(f"      → A列に業種名を初期書き込みしました（{len(INDUSTRY_ORDER)}業種）")


def write_date_column(ws: gspread.Worksheet, target_date: str, key: str, data: dict) -> None:
    """C列以降に日付列を書き込む。
    - 同日付が既存 → 上書き（位置はそのまま）
    - 新規日付     → C列に1列挿入して既存データを右にシフト
    """
    header_row = ws.row_values(1)

    # C列（idx=3）以降から同日付の既存列を探す
    target_col: Optional[int] = None
    for idx, val in enumerate(header_row, start=1):
        if idx >= 3 and val == target_date:
            target_col = idx
            break

    if target_col is None:
        # ── 新規日付: C列（0-based index=2）に1列挿入して右シフト ──
        ws.spreadsheet.batch_update({
            "requests": [{
                "insertDimension": {
                    "range": {
                        "sheetId": ws.id,
                        "dimension": "COLUMNS",
                        "startIndex": 2,   # 0-based: A=0, B=1, C=2
                        "endIndex": 3,
                    },
                    "inheritFromBefore": False,
                }
            }]
        })
        target_col = 3  # C列（1-based）

    letter = col_letter(target_col)
    ws.update(range_name=f"{letter}1", values=[[target_date]])
    ws.update(range_name=f"{letter}2", values=[[data[name][key]] for name in INDUSTRY_ORDER])


# ================================================================
#  メイン
# ================================================================

def main() -> None:
    # 日付の優先順位: コマンドライン引数 > 環境変数 TARGET_DATE > None（最新）
    target_date: Optional[str] = None
    if len(sys.argv) >= 2:
        target_date = sys.argv[1]
    elif os.environ.get("TARGET_DATE"):
        target_date = os.environ["TARGET_DATE"]

    print("=" * 55)
    print("  JPX 空売り業種別集計 取得スクリプト")
    print(f"  対象日付: {target_date if target_date else '最新（自動）'}")
    print("=" * 55)

    # Step 1: PDF URL 取得
    pdf_url = get_pdf_url(target_date)
    print(f"[1/3] PDF URL  : {pdf_url}")

    # Step 2: PDF ダウンロード＆パース
    pdf_bytes = download_pdf(pdf_url)
    parsed_date, data = parse_pdf(pdf_bytes)
    print(f"[2/3] 対象日付 : {parsed_date}  /  取得業種数 : {len(data)}")

    # Step 3: スプレッドシート書き込み
    print("[3/3] スプレッドシートに書き込み中...")
    gc = get_gspread_client()
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)

    for key, sheet_info in SHEET_MAP.items():
        ws = spreadsheet.get_worksheet_by_id(sheet_info["gid"])
        print(f"      シート『{sheet_info['name']}』")
        ensure_industry_column(ws)
        write_date_column(ws, parsed_date, key, data)

    print("=" * 55)
    print("  完了しました。")
    print("=" * 55)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[ERROR] {e}", file=sys.stderr)
        sys.exit(1)
