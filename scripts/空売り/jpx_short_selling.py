#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JPX「空売りの残高に関する情報」を取得し、Google スプレッドシートへ列追加形式で記録する。

対象ページ:
    https://www.jpx.co.jp/markets/public/short-selling/index.html
    （毎営業日 17:00 JST 目処に YYYYMMDD_Short_Positions.xls が掲載される）

出力レイアウト（各シート共通）:
    A列 = コード / B列 = 銘柄名 / 1行目 = 公表日(例: 2026年8月7日(金))
    新しい公表日のデータは C列 に挿入され、既存データは右へずれる。
    既存のコードは同じ行に書き込み、新規コードは最終行に追加する。
    データが無い銘柄のセルは空白のまま。

コマンドライン引数（他スクリプトと同じ呼び出し方式。省略時は環境変数を使用）:
    --date YYYY-MM-DD   取得する公表日（省略時は最新）
    --force             既にその日付の列がある場合でも上書きする
    --dry-run           書き込みせず内容だけ表示する

環境変数:
    GOOGLE_SERVICE_ACCOUNT_JSON  必須  サービスアカウントJSON（文字列そのもの）
                                        ※旧名 GCP_SERVICE_ACCOUNT_JSON も後方互換で使用可
    SPREADSHEET_ID               任意  未設定時はスクリプト内デフォルトを使用
    TARGET_DATE                  任意  --date 未指定時のフォールバック
    RATIO_AS_PERCENT             任意  1(既定)=割合を%表記(0.0051→0.51) / 0=生値のまま
    MAX_DATE_COLS                任意  保持する日付列の最大数（既定 300、0で無制限）
    DRY_RUN                      任意  --dry-run 未指定時のフォールバック
"""

from __future__ import annotations

import argparse
import datetime as dt
import io
import json
import os
import re
import sys
import time
import unicodedata
from typing import Any
from urllib.parse import urljoin

import gspread
import pandas as pd
import requests
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials


# ====================================================================
# 定数定義
# ====================================================================

# --- タイムゾーン / URL -----------------------------------------------------
JST = dt.timezone(dt.timedelta(hours=9))
ORIGIN = "https://www.jpx.co.jp"
INDEX_URL = f"{ORIGIN}/markets/public/short-selling/index.html"

# --- スプレッドシート ---------------------------------------------------
# SPREADSHEET_ID 環境変数が未設定の場合に使うデフォルト
DEFAULT_SPREADSHEET_ID = "1kWST0CkkIvo3irPSbMgtVtUqqRDXFwvRREYZDRQAFMY"
DEFAULT_MAX_DATE_COLS = 300  # MAX_DATE_COLS 環境変数の既定値(保持する日付列の最大数)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# 指標名 -> シートID(gid)
SHEET_GIDS: dict[str, int] = {
    "空売り残高割合": 1029807466,
    "空売り残高数量": 282742022,
    "空売り残高売買単位数": 2012675956,
    "直近計算年月日": 1155532164,
    "直近空売り残高割合": 1391711324,
    "機関数": 1674201498,
    "主要機関": 1621375335,
}

# --- HTTP ---------------------------------------------------------------
REQ_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en;q=0.8",
}
DOWNLOAD_RETRIES = 3
DOWNLOAD_RETRY_BASE_DELAY_SEC = 3  # リトライ毎に (i+1) 倍して待機する基準秒数
API_RATE_LIMIT_DELAY_SEC = 1.5  # シート更新1件ごとのAPIレート制限対策の待機秒数

# 一覧ページ内の Excel ファイルリンクを検出する正規表現(YYYYMMDD_Short_Positions.xls[x])
SHORT_POSITIONS_FILE_RE = re.compile(r"(\d{8})_Short_Positions\.xlsx?$", re.IGNORECASE)

# --- Excel 読み込み -------------------------------------------------------
OLE2_SIGNATURE = b"\xd0\xcf\x11\xe0"  # 旧形式 .xls (OLE2) のファイルシグネチャ
ZIP_SIGNATURE = b"PK"                  # .xlsx (ZIP) のファイルシグネチャ

# Excel シリアル値として date に変換して良いと判断する範囲(概ね1954年〜2119年)
EXCEL_EPOCH = dt.date(1899, 12, 30)
EXCEL_SERIAL_MIN = 20000
EXCEL_SERIAL_MAX = 80000

# --- ヘッダ検出 / パース --------------------------------------------------
HEADER_SCAN_MAX_ROWS = 30  # ヘッダ行を探す際にスキャンする最大行数
REQUIRED_COLUMNS = {"コード", "空売り残高割合", "空売り残高数量"}  # これらが揃わなければヘッダ未検出扱い
CODE_VALIDATION_RE = re.compile(r"[0-9A-Z]{4,5}")  # 見出し・注記行を除外するためのコード形式チェック

# Excel のヘッダ検出用（長い語を優先的に割り当てる）
COLUMN_PATTERNS: dict[str, list[str]] = {
    "コード": ["銘柄コード", "コード"],
    "銘柄名": ["銘柄名", "銘柄"],
    "商号": ["商号・名称・氏名", "商号", "名称又は氏名", "報告義務者"],
    "計算年月日": ["計算年月日"],
    "空売り残高割合": ["空売り残高割合"],
    "空売り残高数量": ["空売り残高数量"],
    "空売り残高売買単位数": ["空売り残高売買単位数"],
    "直近計算年月日": ["直近計算年月日"],
    "直近空売り残高割合": ["直近空売り残高割合"],
}

# --- 日付表示フォーマット用 -------------------------------------------------
# dt.date.weekday(): 月=0, 火=1, 水=2, 木=3, 金=4, 土=5, 日=6
WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]


# ====================================================================
# 日付フォーマット
# ====================================================================

def format_date_jp(d: dt.date) -> str:
    """
    date オブジェクトを "2026年8月7日(金)" 形式の文字列に変換する。
    スプレッドシートの日付列見出し(公表日)として使用する。
    """
    return f"{d.year}年{d.month}月{d.day}日({WEEKDAY_JP[d.weekday()]})"


# ====================================================================
# ユーティリティ
# ====================================================================

def log(msg: str) -> None:
    now = dt.datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S JST")
    print(f"[{now}] {msg}", flush=True)


def norm_text(v: Any) -> str:
    """全角半角・空白・改行を正規化した文字列を返す。"""
    if v is None:
        return ""
    s = unicodedata.normalize("NFKC", str(v))
    return re.sub(r"\s+", "", s)


def col_letter(idx: int) -> str:
    """1 -> A, 27 -> AA"""
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_date_like(v: Any) -> dt.date | None:
    """datetime / Excelシリアル値 / 文字列 から date を取り出す。"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (dt.datetime, pd.Timestamp)):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)):
        n = int(v)
        if EXCEL_SERIAL_MIN < n < EXCEL_SERIAL_MAX:
            return EXCEL_EPOCH + dt.timedelta(days=n)
        return None
    s = norm_text(v)
    if not s:
        return None
    m = re.search(r"(\d{4})[/\-年](\d{1,2})[/\-月](\d{1,2})", s)
    if m:
        try:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", s)
    if m:
        try:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def to_number(v: Any) -> float | None:
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return None if pd.isna(v) else float(v)
    s = norm_text(v).replace(",", "").replace("%", "")
    if s in ("", "-", "－", "―", "N/A"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def norm_code(v: Any) -> str:
    """銘柄コードを 4桁ゼロ埋め文字列に正規化（130A のような英数字コードも許容）。"""
    s = norm_text(v)
    if not s:
        return ""
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    s = s.upper()
    if re.fullmatch(r"\d{1,4}", s):
        s = s.zfill(4)
    return s


# ====================================================================
# JPX からのダウンロード
# ====================================================================

def fetch_file_index() -> dict[dt.date, str]:
    """一覧ページから {公表日: Excel URL} を取得する。"""
    log(f"インデックス取得: {INDEX_URL}")
    res = requests.get(INDEX_URL, headers=REQ_HEADERS, timeout=60)
    res.raise_for_status()
    res.encoding = res.apparent_encoding or "utf-8"

    soup = BeautifulSoup(res.text, "html.parser")

    found: dict[dt.date, str] = {}
    for a in soup.find_all("a", href=True):
        href = a["href"].split("?")[0]
        m = SHORT_POSITIONS_FILE_RE.search(href)
        if not m:
            continue
        d = parse_date_like(m.group(1))
        if d is None:
            continue
        url = urljoin(INDEX_URL, href)
        found[d] = url

    if not found:
        raise RuntimeError("一覧ページから Short_Positions ファイルのリンクを検出できませんでした。")
    log(f"検出: {len(found)}件 (最新 {max(found)})")
    return found


def resolve_target(index: dict[dt.date, str], target: str | None) -> tuple[dt.date, str]:
    if not target:
        d = max(index)
        return d, index[d]

    want = parse_date_like(target)
    if want is None:
        raise ValueError(f"TARGET_DATE を解釈できません: {target!r}")
    if want in index:
        return want, index[want]

    avail = ", ".join(d.strftime("%Y/%m/%d") for d in sorted(index, reverse=True))
    raise RuntimeError(
        f"指定日 {want:%Y/%m/%d} のファイルが一覧ページにありません。"
        f"（一覧ページは当月分のみ掲載されます。掲載中: {avail}）"
    )


def download(url: str, retries: int = DOWNLOAD_RETRIES) -> bytes:
    last: Exception | None = None
    for i in range(retries):
        try:
            log(f"ダウンロード ({i + 1}/{retries}): {url}")
            res = requests.get(url, headers=REQ_HEADERS, timeout=120)
            res.raise_for_status()
            return res.content
        except Exception as e:  # noqa: BLE001
            last = e
            log(f"  失敗: {e}")
            time.sleep(DOWNLOAD_RETRY_BASE_DELAY_SEC * (i + 1))
    raise RuntimeError(f"ダウンロードに失敗しました: {url}") from last


def load_dataframe(data: bytes) -> pd.DataFrame:
    """拡張子に依存せずバイナリ先頭を見て適切なエンジンで読み込む。"""
    head = data[:8]
    if head[:4] == OLE2_SIGNATURE:                # OLE2 = 旧 .xls
        engine = "xlrd"
    elif head[:2] == ZIP_SIGNATURE:                # ZIP = .xlsx
        engine = "openpyxl"
    else:                                          # HTML テーブルの場合
        log("  Excel シグネチャ無し → HTML テーブルとして解析")
        tables = pd.read_html(io.BytesIO(data))
        if not tables:
            raise RuntimeError("表を検出できませんでした。")
        return max(tables, key=len).astype(object)
    log(f"  Excel 読み込み engine={engine}")
    return pd.read_excel(io.BytesIO(data), sheet_name=0, header=None, engine=engine)


# ====================================================================
# パース
# ====================================================================

def detect_header(df: pd.DataFrame) -> tuple[int, dict[str, int]]:
    """ヘッダ行番号と {論理カラム名: 列番号} を返す。"""
    best_row, best_map = -1, {}
    scan = min(len(df), HEADER_SCAN_MAX_ROWS)

    for r in range(scan):
        cells = [norm_text(x) for x in df.iloc[r].tolist()]
        mapping: dict[str, int] = {}
        used_cols: set[int] = set()
        for c, text in enumerate(cells):
            if not text:
                continue
            # その列に当てはまる中で「最も長いパターン」を採用する
            # （例:「直近空売り残高割合」を「空売り残高割合」より優先）
            cands: list[tuple[int, str]] = []
            for key, pats in COLUMN_PATTERNS.items():
                for p in pats:
                    if p in text:
                        cands.append((len(p), key))
            if not cands:
                continue
            cands.sort(reverse=True)
            for _, key in cands:
                if key not in mapping:
                    mapping[key] = c
                    used_cols.add(c)
                    break
        if len(mapping) > len(best_map):
            best_row, best_map = r, mapping

    if not REQUIRED_COLUMNS.issubset(best_map):
        raise RuntimeError(
            f"ヘッダ行を特定できませんでした。検出内容: {best_map}"
        )
    log(f"ヘッダ行={best_row}  列マップ={best_map}")
    return best_row, best_map


def build_records(df: pd.DataFrame) -> pd.DataFrame:
    """銘柄単位に集計した DataFrame を返す。"""
    header_row, cmap = detect_header(df)
    body = df.iloc[header_row + 1 :].reset_index(drop=True)

    def col(key: str) -> pd.Series:
        if key not in cmap:
            return pd.Series([None] * len(body))
        return body.iloc[:, cmap[key]]

    rec = pd.DataFrame(
        {
            "code": col("コード").map(norm_code),
            "name": col("銘柄名").map(lambda v: str(v).strip() if v is not None and not pd.isna(v) else ""),
            "entity": col("商号").map(lambda v: str(v).strip() if v is not None and not pd.isna(v) else ""),
            "ratio": col("空売り残高割合").map(to_number),
            "volume": col("空売り残高数量").map(to_number),
            "units": col("空売り残高売買単位数").map(to_number),
            "last_date": col("直近計算年月日").map(parse_date_like),
            "last_ratio": col("直近空売り残高割合").map(to_number),
            "calc_date": col("計算年月日").map(parse_date_like),
        }
    )

    rec = rec[rec["code"].str.len() > 0].copy()
    # コード列に紛れ込んだ見出し・注記行を除外
    rec = rec[rec["code"].str.fullmatch(CODE_VALIDATION_RE)].copy()
    if rec.empty:
        raise RuntimeError("有効なデータ行が0件でした。")

    # 直近計算年月日が空の場合は計算年月日で補完
    rec["last_date"] = rec["last_date"].fillna(rec["calc_date"])

    log(f"明細行: {len(rec)}件 / 銘柄数: {rec['code'].nunique()}")

    rows = []
    for code, g in rec.groupby("code", sort=True):
        names = [n for n in g["name"].tolist() if n]
        entities = [e for e in g["entity"].tolist() if e]
        dates = [d for d in g["last_date"].tolist() if d is not None]

        # 空売り残高割合が最大の機関を主要機関とする
        main_entity = ""
        if entities:
            gg = g[g["entity"] != ""]
            if not gg["ratio"].isna().all():
                main_entity = str(gg.loc[gg["ratio"].idxmax(), "entity"])
            else:
                main_entity = entities[0]

        rows.append(
            {
                "code": code,
                "name": names[0] if names else "",
                "空売り残高割合": g["ratio"].sum(min_count=1),
                "空売り残高数量": g["volume"].sum(min_count=1),
                "空売り残高売買単位数": g["units"].sum(min_count=1),
                "直近計算年月日": max(dates) if dates else None,
                "直近空売り残高割合": g["last_ratio"].sum(min_count=1),
                "機関数": int(g["entity"].replace("", pd.NA).nunique()) or len(g),
                "主要機関": main_entity,
            }
        )

    out = pd.DataFrame(rows)
    log(f"集計後: {len(out)}銘柄")
    return out


def format_value(metric: str, value: Any, ratio_as_percent: bool) -> Any:
    """
    シートへ書き込む値へ変換する。None は空白。
    「直近計算年月日」は銘柄ごとの実データ値(公表日ヘッダーとは別物)のため、
    従来通り "%Y/%m/%d" 表記のまま(スコープ外の変更は行っていません)。
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if metric in ("空売り残高割合", "直近空売り残高割合"):
        v = float(value) * 100 if ratio_as_percent else float(value)
        return round(v, 6)
    if metric in ("空売り残高数量", "空売り残高売買単位数", "機関数"):
        return int(round(float(value)))
    if metric == "直近計算年月日":
        return value.strftime("%Y/%m/%d") if isinstance(value, dt.date) else str(value)
    return str(value)


# ====================================================================
# Google スプレッドシート書き込み
# ====================================================================

def open_spreadsheet(spreadsheet_id: str) -> gspread.Spreadsheet:
    raw = (
        os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        or os.environ.get("GCP_SERVICE_ACCOUNT_JSON", "").strip()  # 後方互換
    )
    if not raw:
        raise RuntimeError(
            "GOOGLE_SERVICE_ACCOUNT_JSON が設定されていません "
            "(Secrets に GOOGLE_SERVICE_ACCOUNT_JSON を登録してください)。"
        )
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(spreadsheet_id)


def write_metric_sheet(
    ss: gspread.Spreadsheet,
    metric: str,
    gid: int,
    pub_date: dt.date,
    records: pd.DataFrame,
    ratio_as_percent: bool,
    max_date_cols: int,
    dry_run: bool,
    force: bool,
) -> None:
    ws = ss.get_worksheet_by_id(gid)
    date_label = format_date_jp(pub_date)
    log(f"--- [{metric}] シート'{ws.title}' (gid={gid}) 更新開始 ---")

    values = {r["code"]: format_value(metric, r[metric], ratio_as_percent) for _, r in records.iterrows()}
    names = {r["code"]: r["name"] for _, r in records.iterrows()}

    ab = ws.get_values("A1:B") or []
    header_row = ws.row_values(1) or []

    # ---- 初回（空シート）--------------------------------------------------
    if not ab or not any(norm_text(x) for x in (ab[0] if ab else [])):
        codes = sorted(values)
        table = [["コード", "銘柄名", date_label]]
        table += [[c, names.get(c, ""), values.get(c, "")] for c in codes]
        log(f"  初期化書き込み: {len(codes)}銘柄")
        if dry_run:
            return
        _ensure_size(ws, len(table), 3)
        ws.update(values=table, range_name="A1", value_input_option="USER_ENTERED")
        log("  完了（初期化）")
        return

    existing_codes = [norm_code(row[0]) if row else "" for row in ab[1:]]

    # ---- 書き込み先の列を決める ------------------------------------------
    target_col = None  # 1-based
    for i, cell in enumerate(header_row[2:], start=3):
        d = parse_date_like(cell)
        if d == pub_date:
            target_col = i
            break

    if target_col is None:
        log(f"  C列に新規列を挿入 ({date_label})")
        if not dry_run:
            if ws.col_count < 3:
                ws.add_cols(3 - ws.col_count)
            ss.batch_update(
                {
                    "requests": [
                        {
                            "insertDimension": {
                                "range": {
                                    "sheetId": gid,
                                    "dimension": "COLUMNS",
                                    "startIndex": 2,
                                    "endIndex": 3,
                                },
                                "inheritFromBefore": False,
                            }
                        }
                    ]
                }
            )
        target_col = 3
    elif not force:
        log(f"  {date_label} は既に {col_letter(target_col)}列に存在 → スキップ（上書きするには --force / FORCE=1）")
        return
    else:
        log(f"  {date_label} は既に {col_letter(target_col)}列に存在 → --force により上書き")

    # ---- 新規銘柄を最終行へ追加 ------------------------------------------
    known = set(existing_codes)
    new_codes = [c for c in sorted(values) if c not in known]
    if new_codes:
        log(f"  新規銘柄 {len(new_codes)}件を追加")
        start_row = len(existing_codes) + 2
        block = [[c, names.get(c, "")] for c in new_codes]
        if not dry_run:
            _ensure_size(ws, start_row + len(block) - 1, target_col)
            ws.update(
                values=block,
                range_name=f"A{start_row}:B{start_row + len(block) - 1}",
                value_input_option="USER_ENTERED",
            )
        existing_codes += new_codes

    # ---- 対象列を一括書き込み --------------------------------------------
    # ★ existing_codes はシートのA列と同じ行順(新規追加分もその追加順)なので、
    #    ここで各コードをキーに values 辞書を引くことで、行のズレなくコード一致の
    #    セルにだけ値が入る。
    column = [[date_label]] + [[values.get(c, "")] for c in existing_codes]
    letter = col_letter(target_col)
    rng = f"{letter}1:{letter}{len(column)}"
    hit = sum(1 for c in existing_codes if c in values)
    log(f"  {rng} へ書き込み（該当 {hit} / 全 {len(existing_codes)} 行）")
    if not dry_run:
        _ensure_size(ws, len(column), target_col)
        ws.update(values=column, range_name=rng, value_input_option="USER_ENTERED")

    # ---- 古い日付列の削除 -------------------------------------------------
    if max_date_cols > 0:
        total_cols = len(ws.row_values(1))
        keep = 2 + max_date_cols
        if total_cols > keep:
            log(f"  古い列を削除: {col_letter(keep + 1)} 以降 ({total_cols - keep}列)")
            if not dry_run:
                ss.batch_update(
                    {
                        "requests": [
                            {
                                "deleteDimension": {
                                    "range": {
                                        "sheetId": gid,
                                        "dimension": "COLUMNS",
                                        "startIndex": keep,
                                        "endIndex": total_cols,
                                    }
                                }
                            }
                        ]
                    }
                )
    log(f"--- [{metric}] 完了 ---")


def _ensure_size(ws: gspread.Worksheet, rows: int, cols: int) -> None:
    if ws.row_count < rows:
        ws.add_rows(rows - ws.row_count + 100)
    if ws.col_count < cols:
        ws.add_cols(cols - ws.col_count + 5)


# ====================================================================
# main
# ====================================================================

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="JPX 空売り残高をスプレッドシートへ取り込む")
    p.add_argument("--date", dest="date", default=None, help="取得する公表日 (YYYY-MM-DD)。省略時は最新")
    p.add_argument("--force", dest="force", action="store_true", default=None, help="既存の日付列があっても上書きする")
    p.add_argument("--dry-run", dest="dry_run", action="store_true", default=None, help="書き込まずに内容だけ表示する")
    return p.parse_args()


def main() -> int:
    args = parse_args()

    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip() or DEFAULT_SPREADSHEET_ID

    target_date = args.date or os.environ.get("TARGET_DATE", "").strip() or None
    ratio_as_percent = os.environ.get("RATIO_AS_PERCENT", "1") not in ("0", "false", "False")
    max_date_cols = int(os.environ.get("MAX_DATE_COLS", str(DEFAULT_MAX_DATE_COLS)) or 0)

    dry_run = args.dry_run if args.dry_run is not None else os.environ.get("DRY_RUN", "0") in ("1", "true", "True")
    force = args.force if args.force is not None else os.environ.get("FORCE", "0") in ("1", "true", "True")

    log("=== JPX 空売り残高 取り込み開始 ===")
    log(
        f"TARGET_DATE={target_date or '(最新)'} / RATIO_AS_PERCENT={ratio_as_percent} "
        f"/ DRY_RUN={dry_run} / FORCE={force}"
    )

    index = fetch_file_index()
    pub_date, url = resolve_target(index, target_date)
    log(f"対象公表日: {pub_date:%Y/%m/%d}")

    raw = download(url)
    df = load_dataframe(raw)
    records = build_records(df)

    if dry_run:
        log("DRY_RUN: 先頭5件")
        print(records.head().to_string())

    ss = open_spreadsheet(spreadsheet_id)
    for metric, gid in SHEET_GIDS.items():
        try:
            write_metric_sheet(
                ss, metric, gid, pub_date, records, ratio_as_percent, max_date_cols, dry_run, force
            )
        except Exception as e:  # noqa: BLE001
            log(f"!!! [{metric}] でエラー: {e}")
            raise
        time.sleep(API_RATE_LIMIT_DELAY_SEC)  # API レート制限対策

    log("=== 完了 ===")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # noqa: BLE001
        log(f"FATAL: {exc}")
        raise
