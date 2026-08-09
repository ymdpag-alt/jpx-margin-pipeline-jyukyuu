#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JPX「空売り集計（業種別集計）」日次データを取得し、Googleスプレッドシートに追記するスクリプト。

取得元:
    https://www.jpx.co.jp/markets/statistics-equities/short-selling/index.html
    上記ページに掲載される日次の「業種別集計」PDF（最新日付分）から、
    業種別の売買代金（実注文 / 空売り(価格規制あり) / 空売り(価格規制なし) / 合計）を抽出する。

出力（スプレッドシート）:
    1つのスプレッドシートの中に、以下4枚のシート（タブ）を作成・更新する。
        - 実注文
        - 空売り（価格規制あり）
        - 空売り（価格規制なし）
        - 合計

    各シートのレイアウト:
        A列 : 業種名（初回実行時に自動で書き込み。33業種＋その他（33業種外）＝34行）
        B列 : 空列（予備）
        C列以降 : 実行のたびに新しい列を追加。
                  1行目 = その列が対象とする日付（PDF記載の日付。例: 2026/08/07）
                  2行目以降 = 各業種の売買代金（単位：百万円）
        ※ 同じ日付の列が既に存在する場合は、その列を上書きする（再実行時の重複防止）。

環境変数:
    JPX_SPREADSHEET_ID          書き込み先スプレッドシートのID（必須）
    GOOGLE_SERVICE_ACCOUNT_JSON サービスアカウントの認証情報（JSON文字列そのもの。必須）

必要ライブラリ（requirements.txt参照）:
    requests, beautifulsoup4, pdfplumber, gspread, google-auth
"""

import io
import json
import os
import re
import sys
from typing import Dict, List, Optional, Tuple

import gspread
import pdfplumber
import requests
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials

INDEX_URL = "https://www.jpx.co.jp/markets/statistics-equities/short-selling/index.html"
BASE_URL = "https://www.jpx.co.jp"

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}

# 33業種 ＋ その他（33業種外）。PDF掲載順と完全に一致させる（初回シート生成時の行順に使用）。
INDUSTRY_ORDER: List[str] = [
    "水産・農林業", "鉱業", "建設業", "食料品", "繊維製品", "パルプ・紙",
    "化学", "医薬品", "石油・石炭製品", "ゴム製品", "ガラス・土石製品",
    "鉄鋼", "非鉄金属", "金属製品", "機械", "電気機器", "輸送用機器",
    "精密機器", "その他製品", "電気・ガス業", "陸運業", "海運業",
    "空運業", "倉庫・運輸関連業", "情報・通信業", "卸売業", "小売業",
    "銀行業", "証券、商品先物取引業", "保険業", "その他金融業",
    "不動産業", "サービス業", "その他（33業種外）",
]

# PDF内のカテゴリ(a,b,c,d) と 書き込み先シート名 の対応
SHEET_NAMES: Dict[str, str] = {
    "a": "実注文",
    "b": "空売り（価格規制あり）",
    "c": "空売り（価格規制なし）",
    "d": "合計",
}

# PDFテキストの1行分パターン
#   例: "水産・農林業 2,091 54.3% 1,693 44.0% 67 1.7% 3,852"
#        業種名        a(実注文) 比率  b(規制あり) 比率  c(規制なし) 比率  d(合計)
#
# 「その他（33業種外）」のように業種名自体に数字を含むケースがあるため、
# 業種名は INDUSTRY_ORDER の完全一致リスト（長い順）から判定する。
_INDUSTRY_ALT = "|".join(
    re.escape(name) for name in sorted(INDUSTRY_ORDER, key=len, reverse=True)
)
ROW_RE = re.compile(
    rf"^(?P<name>{_INDUSTRY_ALT})\s+"
    r"(?P<a>[\d,]+)\s+[\d.]+%\s+"
    r"(?P<b>[\d,]+)\s+[\d.]+%\s+"
    r"(?P<c>[\d,]+)\s+[\d.]+%\s+"
    r"(?P<d>[\d,]+)\s*$"
)

DATE_RE = re.compile(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日")


def get_latest_pdf_url() -> str:
    """トップページ（日次集計一覧）から、最新日付の「業種別集計」PDFのURLを取得する。"""
    resp = requests.get(INDEX_URL, headers=REQUEST_HEADERS, timeout=30)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding
    soup = BeautifulSoup(resp.text, "html.parser")

    table = soup.find("table")
    if table is None:
        raise RuntimeError("一覧テーブルが見つかりませんでした。サイト構造が変更された可能性があります。")

    rows = table.find_all("tr")
    # ヘッダー行を除いた最初のデータ行 = 最新日付
    data_rows = [r for r in rows if r.find_all("a")]
    if not data_rows:
        raise RuntimeError("PDFリンクを含む行が見つかりませんでした。")
    first_row = data_rows[0]

    links = first_row.find_all("a")
    if len(links) < 2:
        raise RuntimeError("PDFリンクが2つ（空売り集計・業種別集計）見つかりませんでした。")

    # ファイル名が "-g.pdf" で終わるものが「業種別集計」
    gyoshu_href: Optional[str] = None
    for a in links:
        href = a.get("href", "")
        if "-g.pdf" in href:
            gyoshu_href = href
            break
    if gyoshu_href is None:
        # 想定外のファイル名パターンの場合は2番目のリンクを使う
        gyoshu_href = links[1].get("href")

    if gyoshu_href.startswith("http"):
        return gyoshu_href
    return BASE_URL + gyoshu_href


def download_pdf(url: str) -> bytes:
    resp = requests.get(url, headers=REQUEST_HEADERS, timeout=30)
    resp.raise_for_status()
    return resp.content


def parse_pdf(pdf_bytes: bytes) -> Tuple[str, Dict[str, Dict[str, int]]]:
    """
    PDFバイト列から 対象日付 と 業種別データ を抽出する。

    戻り値:
        target_date: "YYYY/MM/DD" 形式の文字列
        data: { 業種名: {"a": 実注文, "b": 価格規制あり, "c": 価格規制なし, "d": 合計}, ... }
              単位はいずれも百万円。
    """
    full_text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(layout=True) or page.extract_text() or ""
            full_text += text + "\n"

    date_match = DATE_RE.search(full_text)
    if not date_match:
        raise RuntimeError("PDFから対象日付（例: 2026年8月7日）を抽出できませんでした。")
    y, m, d = date_match.groups()
    target_date = f"{int(y):04d}/{int(m):02d}/{int(d):02d}"

    data: Dict[str, Dict[str, int]] = {}
    for raw_line in full_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        m_ = ROW_RE.match(line)
        if not m_:
            continue
        name = m_.group("name")
        if name not in INDUSTRY_ORDER:
            continue
        data[name] = {
            "a": int(m_.group("a").replace(",", "")),
            "b": int(m_.group("b").replace(",", "")),
            "c": int(m_.group("c").replace(",", "")),
            "d": int(m_.group("d").replace(",", "")),
        }

    missing = [n for n in INDUSTRY_ORDER if n not in data]
    if missing:
        raise RuntimeError(
            "以下の業種のデータをPDFから抽出できませんでした。PDFレイアウトが変更された可能性があります。\n"
            f"未抽出の業種: {missing}\n"
            "----- 抽出した全文テキスト（デバッグ用） -----\n"
            f"{full_text}"
        )

    return target_date, data


def get_gspread_client() -> gspread.Client:
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        raise RuntimeError("環境変数 GOOGLE_SERVICE_ACCOUNT_JSON が設定されていません。")
    info = json.loads(sa_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)


def col_index_to_letter(idx: int) -> str:
    """1始まりの列番号をA1形式の列名（A, B, ..., Z, AA, ...）に変換する。"""
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def ensure_sheet(spreadsheet: gspread.Spreadsheet, sheet_name: str) -> gspread.Worksheet:
    """シートが無ければ作成し、A列（業種名）が未設定なら初期化する。"""
    try:
        ws = spreadsheet.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=100, cols=1000)

    a_col = ws.col_values(1)
    if len(a_col) < len(INDUSTRY_ORDER) + 1:
        header = ["業種（単位：百万円）", ""]
        rows = [[name, ""] for name in INDUSTRY_ORDER]
        ws.update(range_name="A1", values=[header] + rows)
    return ws


def append_data_to_sheet(
    ws: gspread.Worksheet, target_date: str, key: str, data: Dict[str, Dict[str, int]]
) -> None:
    """C列以降の該当日付の列にデータを書き込む（同日付が既にあれば上書き）。"""
    header_row = ws.row_values(1)

    target_col: Optional[int] = None
    for idx, val in enumerate(header_row, start=1):
        if idx < 3:
            continue
        if val == target_date:
            target_col = idx
            break
    if target_col is None:
        target_col = max(len(header_row) + 1, 3)

    # 列数が足りない場合は拡張しておく
    if target_col > ws.col_count:
        ws.resize(cols=target_col + 50)

    col_letter = col_index_to_letter(target_col)
    ws.update(range_name=f"{col_letter}1", values=[[target_date]])

    values = [[data[name][key]] for name in INDUSTRY_ORDER]
    ws.update(range_name=f"{col_letter}2", values=values)


def main() -> None:
    print("JPX 空売り業種別集計の取得を開始します...")
    pdf_url = get_latest_pdf_url()
    print(f"対象PDF: {pdf_url}")

    pdf_bytes = download_pdf(pdf_url)
    target_date, data = parse_pdf(pdf_bytes)
    print(f"対象日付: {target_date} / 取得業種数: {len(data)}")

    spreadsheet_id = os.environ.get("JPX_SPREADSHEET_ID")
    if not spreadsheet_id:
        raise RuntimeError("環境変数 JPX_SPREADSHEET_ID が設定されていません。")

    gc = get_gspread_client()
    spreadsheet = gc.open_by_key(spreadsheet_id)

    for key, sheet_name in SHEET_NAMES.items():
        print(f"シート『{sheet_name}』を更新中...")
        ws = ensure_sheet(spreadsheet, sheet_name)
        append_data_to_sheet(ws, target_date, key, data)

    print("完了しました。")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # noqa: BLE001
        print(f"エラーが発生しました: {e}", file=sys.stderr)
        sys.exit(1)
