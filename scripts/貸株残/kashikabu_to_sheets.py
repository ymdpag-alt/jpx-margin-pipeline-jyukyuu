"""
貸株残：日本証券業協会「銘柄別株券等貸借週末残高」（週次） → Google スプレッドシート

取得元:
    日証協 株券等貸借取引状況（週間）  files/{申込週末日YYYYMMDD}z.xlsx
    https://www.jsda.or.jp/shiryoshitsu/toukei/kabu-taiw/index.html
    報告週の翌週木曜に公表。制度貸借（日証金）は対象外の統計。

書き込み先（KASHIKABU_SPREADSHEET_ID のブック）:
    貸株残 シート

シートの形:
    A列=銘柄コード / B列=銘柄名 / C列以降=申込日（左ほど新しい）
    - 新しい日付は、日付の降順が保たれる位置に1列挿入する（通常はC列）。
    - シートに無い銘柄（新規上場など）は最下部に追加する。
    - 同じ日付の列が既にあれば何もしない（何度実行しても安全）。

使い方:
    python scripts/貸株残/kashikabu_to_sheets.py             # シートを更新
    python scripts/貸株残/kashikabu_to_sheets.py --inspect   # 最新xlsxの先頭と読み取り結果の確認だけ（書き込みなし）

環境変数:
    GOOGLE_SERVICE_ACCOUNT_JSON  サービスアカウントのJSON（必須）
    KASHIKABU_SPREADSHEET_ID     書き込み先ブック（未設定なら下の既定値）
    JSDA_BACKFILL_N              未取得の週を最大何週分入れるか（既定 2）
"""

from __future__ import annotations

import io
import json
import os
import re
import sys
import time
from datetime import datetime

import gspread
import pandas as pd
import requests
from google.oauth2.service_account import Credentials
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# #############################################################################
#
#  設定（変更が必要になるのはこのブロックだけ）
#
# #############################################################################

def _env(name: str, default: str = "") -> str:
    """環境変数を読む。未設定・空文字（Secrets未登録など）のときは default。"""
    return os.environ.get(name, "").strip() or default


# -----------------------------------------------------------------------------
# 1. 実行時の設定（GitHub Actions の env / secrets から受け取る）
# -----------------------------------------------------------------------------
GOOGLE_SERVICE_ACCOUNT_JSON_ENV = "GOOGLE_SERVICE_ACCOUNT_JSON"

SPREADSHEET_ID = _env(                                     # 書き込み先ブック
    "KASHIKABU_SPREADSHEET_ID", "1kWST0CkkIvo3irPSbMgtVtUqqRDXFwvRREYZDRQAFMY"
)
BACKFILL_N = int(_env("JSDA_BACKFILL_N", "2"))             # 未取得の週を最大何週分入れるか

# -----------------------------------------------------------------------------
# 2. シート構成
# -----------------------------------------------------------------------------
SHEET_NAME = "貸株残"
VALUE_COL = "貸株残"

FIXED_COLS = 2               # A列=銘柄コード, B列=銘柄名
FIRST_DATE_COL = 3           # C列（1始まり）
NEW_SHEET_COLS = 50          # シート新規作成時の列数
NEW_SHEET_ROWS = 100         # シート新規作成時の行数

GOOGLE_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# -----------------------------------------------------------------------------
# 3. 取得元（日証協）
# -----------------------------------------------------------------------------
JSDA_INDEX_URL = "https://www.jsda.or.jp/shiryoshitsu/toukei/kabu-taiw/index.html"
JSDA_FILE_URL_TEMPLATE = "https://www.jsda.or.jp/shiryoshitsu/toukei/kabu-taiw/files/{date}z.xlsx"
JSDA_LINK_PATTERN = re.compile(r"files/(\d{8})z\.xlsx")

# -----------------------------------------------------------------------------
# 4. xlsx の読み取りルール（列見出しの候補。上から優先）
# -----------------------------------------------------------------------------
CODE_KEYWORDS = ["銘柄コード", "コード", "code"]
NAME_KEYWORDS = ["銘柄名", "名称", "銘柄"]
BALANCE_KEYWORDS = ["貸付残高", "貸付", "貸株", "週末残高", "残高"]
BALANCE_EXCLUDE = ["借入", "返済", "新規", "成約"]   # 残高列の候補から外す見出し
HEADER_SCAN_ROWS = 25                                 # 見出し行を探す範囲（先頭から）

# -----------------------------------------------------------------------------
# 5. HTTP（日証協は間隔を空けないと 429 になりやすい）
# -----------------------------------------------------------------------------
HTTP_TIMEOUT_PAGE = 30
HTTP_TIMEOUT_FILE = 60
HTTP_RETRY_TOTAL = 5             # 429・5xx のときの再試行回数
HTTP_RETRY_BACKOFF = 10          # 再試行の待ち時間の基準（秒）。10, 20, 40, ... と伸びる
DOWNLOAD_INTERVAL = 15           # ファイル取得ごとの待ち時間（秒）
HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en;q=0.8",
}

WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]


# #############################################################################
#
#  ここから処理
#
# #############################################################################

# =============================================================================
# 共通ユーティリティ
# =============================================================================

def to_japanese_date(yyyymmdd: str) -> str:
    """'20260925' → '2026年9月25日(金)'"""
    dt = datetime.strptime(yyyymmdd, "%Y%m%d")
    return f"{dt.year}年{dt.month}月{dt.day}日({WEEKDAY_JP[dt.weekday()]})"


def from_japanese_date(text: str) -> str:
    """'2026年9月25日(金)' → '20260925'。読めなければ元の文字列を返す。"""
    try:
        dt = datetime.strptime(text.split("(")[0].strip(), "%Y年%m月%d日")
        return dt.strftime("%Y%m%d")
    except (ValueError, IndexError):
        return text


def is_yyyymmdd(text: str) -> bool:
    return bool(re.fullmatch(r"\d{8}", text or ""))


def http_session() -> requests.Session:
    """429・5xx は Retry-After に従って自動で再試行するセッション。"""
    retry = Retry(
        total=HTTP_RETRY_TOTAL,
        backoff_factor=HTTP_RETRY_BACKOFF,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
        respect_retry_after_header=True,
        raise_on_status=False,
    )
    sess = requests.Session()
    sess.headers.update(HTTP_HEADERS)
    sess.mount("https://", HTTPAdapter(max_retries=retry))
    return sess


def fetch_text(sess: requests.Session, url: str) -> str:
    resp = sess.get(url, timeout=HTTP_TIMEOUT_PAGE)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def fetch_file(sess: requests.Session, url: str) -> bytes:
    print(f"  ダウンロード: {url}")
    resp = sess.get(url, timeout=HTTP_TIMEOUT_FILE)
    if resp.status_code == 404:
        raise FileNotFoundError(f"ファイルが見つかりません: {url}")
    resp.raise_for_status()
    return resp.content


# =============================================================================
# 日証協: 掲載一覧と xlsx の読み取り
# =============================================================================

def list_dates(sess: requests.Session) -> list[str]:
    """日証協ページに載っている申込日を新しい順で返す。"""
    dates = sorted(set(JSDA_LINK_PATTERN.findall(fetch_text(sess, JSDA_INDEX_URL))), reverse=True)
    if not dates:
        raise ValueError("日証協のページにファイルリンクがありません。ページ構成が変わった可能性があります。")
    return dates


def normalize_stock_code(value) -> str | None:
    """'1301' / '1301.0' / '13010' / '402A' → 4桁コード（信用残シートと揃える）"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = re.sub(r"\.0$", "", str(value).strip().upper())
    if re.fullmatch(r"[0-9A-Z]{4}0", s):
        return s[:4]
    if re.fullmatch(r"[0-9A-Z]{4}", s):
        return s
    return None


def find_column(columns: list[str], keywords: list[str], exclude: list[str] | None = None) -> str | None:
    """keywords の優先順で最初に見つかった列名を返す。exclude を含む列は除く。"""
    for keyword in keywords:
        for col in columns:
            if exclude and any(ex in col for ex in exclude):
                continue
            if keyword in col:
                return col
    return None


def parse_xlsx(xlsx_bytes: bytes) -> pd.DataFrame:
    """xlsx → DataFrame（列: 銘柄コード, 銘柄名, 貸株残）"""
    raw = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0, header=None)

    # ---- 見出し行を探す（表題行が先頭にあるため）----
    header_row = next(
        (i for i in range(min(HEADER_SCAN_ROWS, len(raw)))
         if any(kw in str(v) for v in raw.iloc[i] for kw in CODE_KEYWORDS)),
        None,
    )
    if header_row is None:
        raise ValueError("見出し行が見つかりません。書式が変わった可能性があります（--inspect で確認）。")

    columns = [re.sub(r"[\s　]", "", str(c)) for c in raw.iloc[header_row]]
    body = raw.iloc[header_row + 1:].copy()
    body.columns = columns

    # ---- 使う列を決める ----
    code_col = find_column(columns, CODE_KEYWORDS)
    name_col = find_column(columns, NAME_KEYWORDS)
    value_col = find_column(columns, BALANCE_KEYWORDS, exclude=BALANCE_EXCLUDE)
    if code_col is None or value_col is None:
        raise ValueError(f"必要な列が見つかりません。検出列: {columns}")
    print(f"  使用列: コード='{code_col}' / 残高='{value_col}' / 銘柄名='{name_col}'")

    # ---- 1行ずつ読む（合計行・空行はコードが読めないので自然に除外される）----
    records = []
    for _, row in body.iterrows():
        code = normalize_stock_code(row[code_col])
        value = pd.to_numeric(str(row[value_col]).replace(",", "").strip(), errors="coerce")
        if code is None or pd.isna(value):
            continue
        name = str(row[name_col]).strip() if name_col else ""
        records.append({
            "銘柄コード": code,
            "銘柄名": "" if name in ("nan", "None") else name,
            VALUE_COL: float(value),
        })

    if not records:
        raise ValueError("貸株残データを読めませんでした。書式が変わった可能性があります。")
    print(f"  {len(records)}銘柄")

    # 同じコードが複数行あれば合算
    df = pd.DataFrame(records).groupby("銘柄コード", as_index=False).agg(
        {"銘柄名": "first", VALUE_COL: "sum"}
    )
    return df.sort_values("銘柄コード").reset_index(drop=True)


def inspect_latest() -> None:
    """--inspect 用。最新xlsxの先頭と読み取り結果を表示する（書き込みなし）。"""
    sess = http_session()
    dates = list_dates(sess)
    print(f"掲載中（新しい順）: {dates[:8]}")
    content = fetch_file(sess, JSDA_FILE_URL_TEMPLATE.format(date=dates[0]))

    raw = pd.read_excel(io.BytesIO(content), sheet_name=0, header=None)
    with pd.option_context("display.max_columns", None, "display.width", 250):
        print(f"\n===== 先頭15行（shape: {raw.shape}）=====")
        print(raw.head(15))
        print("\n===== 読み取り結果（先頭15件）=====")
        print(parse_xlsx(content).head(15))


# =============================================================================
# Google Sheets: 接続とユーティリティ
# =============================================================================

def connect_google_sheets() -> gspread.Client:
    """サービスアカウントで認証する。429（レート制限）は自動リトライ。"""
    raw = os.environ.get(GOOGLE_SERVICE_ACCOUNT_JSON_ENV, "")
    try:
        info = json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(f"{GOOGLE_SERVICE_ACCOUNT_JSON_ENV} をJSONとして読めません（{len(raw)}文字）: {e}") from e

    creds = Credentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)
    try:
        from gspread.http_client import BackOffHTTPClient
        return gspread.authorize(creds, http_client=BackOffHTTPClient)
    except (ImportError, TypeError):   # gspread 5.x
        return gspread.authorize(creds)


def open_worksheet(gc: gspread.Client, spreadsheet_id: str, sheet_name: str) -> gspread.Worksheet:
    """シートを開く。無ければ作る。"""
    book = gc.open_by_key(spreadsheet_id)
    try:
        return book.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  シート '{sheet_name}' を新規作成します")
        return book.add_worksheet(title=sheet_name, rows=NEW_SHEET_ROWS, cols=NEW_SHEET_COLS)


def ensure_size(ws: gspread.Worksheet, rows: int = 0, cols: int = 0) -> None:
    """行数・列数が足りなければ広げる。"""
    if cols and ws.col_count < cols:
        ws.add_cols(cols - ws.col_count)
    if rows and ws.row_count < rows:
        ws.add_rows(rows - ws.row_count)


def to_cell(value):
    """None/NaN → 空欄、整数の float → int"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def dates_in_header(header: list[str]) -> set[str]:
    return {from_japanese_date(c) for c in header[FIXED_COLS:] if c.strip()}


def insert_position(header: list[str], yyyymmdd: str) -> int:
    """日付の降順（左ほど新しい）を保つ挿入列（1始まり）。最新ならC列=3。"""
    for col, cell in enumerate(header[FIXED_COLS:], start=FIRST_DATE_COL):
        d = from_japanese_date(cell)
        if is_yyyymmdd(d) and d < yyyymmdd:
            return col
    last_used = max((i for i, c in enumerate(header, start=1) if c.strip()), default=FIXED_COLS)
    return max(last_used + 1, FIRST_DATE_COL)


def col_letter(col: int) -> str:
    return gspread.utils.rowcol_to_a1(1, col).rstrip("1")


# =============================================================================
# Google Sheets: 1週分の列を書き込む
# =============================================================================

def write_date_column(ws: gspread.Worksheet, df: pd.DataFrame, yyyymmdd: str) -> None:
    existing = ws.get_all_values()
    label = to_japanese_date(yyyymmdd)
    tag = f"  [{ws.title}]"

    # ---- 1. シートが空 → 初回書き込み ----
    if not existing or not any(c.strip() for c in existing[0]):
        print(f"{tag} 初回書き込み")
        rows = [["銘柄コード", "銘柄名", label]]
        rows += [[r["銘柄コード"], r["銘柄名"], to_cell(r[VALUE_COL])] for _, r in df.iterrows()]
        ensure_size(ws, rows=len(rows), cols=FIRST_DATE_COL)
        ws.update(values=rows, range_name="A1", value_input_option="USER_ENTERED")
        return

    header = existing[0]

    # ---- 2. 同じ日付が既にある → 何もしない ----
    if yyyymmdd in dates_in_header(header):
        print(f"{tag} {label} は追加済み。スキップします")
        return

    # ---- 3. 既存銘柄の値を1列にして挿入 ----
    codes = [row[0] for row in existing[1:]]
    col = insert_position(header, yyyymmdd)
    lookup = dict(zip(df["銘柄コード"].astype(str), df[VALUE_COL]))
    column = [label] + [to_cell(lookup.get(code)) for code in codes]

    print(f"{tag} {col_letter(col)}列に {label} を挿入")
    ws.insert_cols([column], col=col, value_input_option="USER_ENTERED")

    # ---- 4. シートに無い銘柄を最下部に追加 ----
    new_stocks = df[~df["銘柄コード"].astype(str).isin(set(codes))]
    if not new_stocks.empty:
        total_cols = max(len(header), col - 1) + 1
        new_rows = []
        for _, r in new_stocks.iterrows():
            row = [""] * total_cols
            row[0], row[1], row[col - 1] = str(r["銘柄コード"]), r["銘柄名"], to_cell(r[VALUE_COL])
            new_rows.append(row)

        start = len(existing) + 1
        end = start + len(new_rows) - 1
        ensure_size(ws, rows=end, cols=total_cols)
        ws.update(
            values=new_rows,
            range_name=f"A{start}:{gspread.utils.rowcol_to_a1(end, total_cols)}",
            value_input_option="USER_ENTERED",
        )
        print(f"{tag} 新規銘柄を追加: {len(new_rows)}件")

    print(f"{tag} 完了")


# =============================================================================
# 更新処理
# =============================================================================

def update(gc: gspread.Client) -> bool:
    """戻り値: 問題なければ True（新しいデータが無いだけなら True）"""
    print("=" * 60)
    print("貸株残（日本証券業協会 週次）")
    print("=" * 60)
    sess = http_session()

    # ---- 対象日を決める ----
    available = list_dates(sess)
    print(f"  掲載中: {available[:5]}{' ...' if len(available) > 5 else ''}")

    ws = open_worksheet(gc, SPREADSHEET_ID, SHEET_NAME)
    done = dates_in_header(ws.row_values(1))
    print(f"  シートの最新: {max((d for d in done if is_yyyymmdd(d)), default='なし')}")

    targets = sorted(sorted((d for d in available if d not in done), reverse=True)[:BACKFILL_N])
    if not targets:
        print("  新しいデータはありません")
        return True
    print(f"  取得対象: {targets}")

    # ---- 1週ずつ取得してシートへ ----
    ok = True
    for i, yyyymmdd in enumerate(targets):
        if i:
            time.sleep(DOWNLOAD_INTERVAL)
        print(f"\n[{yyyymmdd}] {to_japanese_date(yyyymmdd)}")
        try:
            df = parse_xlsx(fetch_file(sess, JSDA_FILE_URL_TEMPLATE.format(date=yyyymmdd)))
        except Exception as e:
            print(f"  ✗ 取得/読み取りに失敗: {e}")
            ok = False
            continue
        write_date_column(ws, df, yyyymmdd)

    return ok


# =============================================================================
# メイン
# =============================================================================

def main() -> None:
    if len(sys.argv) >= 2 and sys.argv[1] == "--inspect":
        inspect_latest()
        return

    if not update(connect_google_sheets()):
        sys.exit("\n貸株残の取得に失敗した週があります。ログを確認してください。")


if __name__ == "__main__":
    main()
