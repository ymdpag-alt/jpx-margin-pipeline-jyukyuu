"""
JPX「銘柄別信用取引週末残高」PDFを取得し、
信用買い残・信用売り残・信用倍率をGoogle Spreadsheetへ書き込む。

添付いただいたyfinanceコードの update_spreadsheet() と同じ設計思想:
  - 初回はヘッダー＋全銘柄を書き込み
  - 2回目以降は新しい「申込日（日本語形式）」列だけ右端に追加
  - 既存銘柄は行を維持したままセルを更新、新規銘柄は末尾に追加

【2026/7/10申込分のPDFで実データ確認済み】
PDFは罫線なしのテキストレイアウトのため、pdfplumberのextract_tables()は使わず、
extract_text()で取った行を正規表現でパースする方式にしている。
また、このPDF自体には「信用倍率」列が存在しないため、
買残高(合計)÷売残高(合計) で自前計算している。
"""

import io
import os
import re
import sys
from datetime import datetime, timedelta

import gspread
import pandas as pd
import pdfplumber
import requests
from google.oauth2.service_account import Credentials

# =============================================================================
# 定数定義
# =============================================================================

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
MARGIN_SHEET_NAME = "信用残データ"

JPX_PDF_URL_TEMPLATE = (
    "https://www.jpx.co.jp/markets/statistics-equities/"
    "margin/tvdivq0000001rnl-att/syumatsu{date}00.pdf"
)

# シート構成: A列=銘柄コード, B列=銘柄名, C列以降=申込日ごとのデータ
# 1日付につき「買い残/売り残/倍率」の3列を使う
FIXED_COLS = 2
COLS_PER_DATE = 3  # 買い残・売り残・倍率

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

_WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]


# =============================================================================
# 日付ユーティリティ
# =============================================================================

def get_target_friday(run_date: datetime | None = None) -> str:
    """
    実行日から直近の「申込日（金曜日）」を YYYYMMDD 形式で返す。
    火曜18:00に実行される想定 → 前週金曜が対象。
    """
    if run_date is None:
        run_date = datetime.now()
    # 月曜=0 ... 金曜=4 ... 日曜=6
    days_since_friday = (run_date.weekday() - 4) % 7
    if days_since_friday == 0 and run_date.hour < 16:
        # 当日金曜かつ16時前は前週分を使う（安全側）
        days_since_friday = 7
    target = run_date - timedelta(days=days_since_friday)
    return target.strftime("%Y%m%d")


def to_japanese_date(date_str_yyyymmdd: str) -> str:
    """'20260710' -> '2026年7月10日(金)'"""
    dt = datetime.strptime(date_str_yyyymmdd, "%Y%m%d")
    weekday = _WEEKDAY_JP[dt.weekday()]
    return f"{dt.year}年{dt.month}月{dt.day}日({weekday})"


def from_japanese_date(date_str: str) -> str:
    """'2026年7月10日(金)' -> '20260710'"""
    try:
        core = date_str.split("(")[0]
        dt = datetime.strptime(core, "%Y年%m月%d日")
        return dt.strftime("%Y%m%d")
    except (ValueError, IndexError):
        return date_str


# =============================================================================
# PDF取得・パース
# =============================================================================

def download_pdf(date_yyyymmdd: str) -> bytes:
    """指定申込日のPDFをダウンロードする。"""
    url = JPX_PDF_URL_TEMPLATE.format(date=date_yyyymmdd)
    print(f"  ダウンロード中: {url}")
    resp = requests.get(url, timeout=30)
    if resp.status_code == 404:
        raise FileNotFoundError(
            f"PDFが見つかりません（まだ公表されていない可能性）: {url}"
        )
    resp.raise_for_status()
    return resp.content


ISIN_PATTERN = re.compile(r"JP[A-Z0-9]{10}")


def parse_margin_pdf(pdf_bytes: bytes) -> pd.DataFrame:
    """
    PDFのテキストを1行ずつ正規表現でパースし、DataFrameで返す。
    列: 銘柄コード, 銘柄名, 買い残, 売り残, 倍率

    1データ行の並び（実データで確認済み）:
      [貸借フラグB] 銘柄名 株式種別 5桁コード ISIN
      売残高(合計) 前週比 買残高(合計) 前週比
      売残高(一般信用) 前週比 売残高(制度信用) 前週比
      買残高(一般信用) 前週比 買残高(制度信用) 前週比

    「倍率」列はPDFに存在しないため、買残高(合計)÷売残高(合計)で算出する。
    """
    records = []
    skipped = 0

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        print(f"  総ページ数: {len(pdf.pages)}")
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                m = ISIN_PATTERN.search(line)
                if not m:
                    continue  # ISINを含まない行（見出し・区切り等）はスキップ

                before = line[: m.start()].strip()
                after = line[m.end():].strip()

                code_match = re.search(r"(\d{5})\s*$", before)
                if not code_match:
                    skipped += 1
                    continue

                code5 = code_match.group(1)
                code = code5[:4]  # 末尾の付番(通常は0)を除いた4桁コード
                name = before[: code_match.start()].strip()
                name = re.sub(r"^B\s+", "", name)  # 先頭の貸借銘柄フラグを除去
                name = re.sub(
                    r"\s*(普通株式|出資証券|投資口|受益証券|優先株式)\s*$", "", name
                )  # 末尾の株式種別を除去

                tokens = after.split()
                values, _ = _parse_signed_tokens(tokens, 12)
                sell_total, _, buy_total = values[0], values[1], values[2]

                if sell_total is None or buy_total is None:
                    skipped += 1
                    continue

                ratio = round(buy_total / sell_total, 2) if sell_total else None

                records.append(
                    {
                        "銘柄コード": code,
                        "銘柄名": name,
                        "買い残": buy_total,
                        "売り残": sell_total,
                        "倍率": ratio,
                    }
                )

    df = pd.DataFrame(records)
    print(f"  抽出件数: {len(df)} 銘柄（パース失敗でスキップ: {skipped} 行）")
    if df.empty:
        raise ValueError("PDFから銘柄データを抽出できませんでした。レイアウトが変わった可能性があります。")
    return df


def _parse_signed_tokens(tokens: list[str], count: int):
    """
    トークン列から count 個の数値を読み取る。
    '▲ 数字' は負の値として扱う（前週比の減少表記）。
    戻り値: (数値リスト, 未使用トークンの残り)
    """
    values = []
    i = 0
    for _ in range(count):
        if i >= len(tokens):
            values.append(None)
            continue
        if tokens[i] in ("▲", "△"):
            i += 1
            if i < len(tokens):
                num = _to_number(tokens[i])
                values.append(-num if num is not None else None)
                i += 1
            else:
                values.append(None)
        else:
            values.append(_to_number(tokens[i]))
            i += 1
    return values, tokens[i:]


def _to_number(s):
    if s is None:
        return None
    s = s.replace(",", "").replace("−", "-").strip()
    if s in ("", "-", "―"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


# =============================================================================
# Google Sheets 認証・書き込み
# =============================================================================

def authenticate_google_sheets() -> gspread.Client:
    """サービスアカウントで認証する（GitHub Actions用）。"""
    creds_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    import json

    info = json.loads(creds_json)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def get_or_create_worksheet(gc: gspread.Client, sheet_name: str, min_cols: int = 50):
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  シート '{sheet_name}' を新規作成します")
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=3000, cols=min_cols)
    if ws.col_count < min_cols:
        ws.add_cols(min_cols - ws.col_count)
        ws = spreadsheet.worksheet(sheet_name)
    return ws


def update_spreadsheet(gc: gspread.Client, df: pd.DataFrame, date_yyyymmdd: str):
    """
    週次で「買い残/売り残/倍率」の3列を右端に追加する。
    既存銘柄は行を維持したままセルを更新、新規銘柄は末尾に追加。
    """
    ws = get_or_create_worksheet(gc, MARGIN_SHEET_NAME)
    existing = ws.get_all_values()
    jp_date = to_japanese_date(date_yyyymmdd)

    date_header_group = [f"{jp_date}_買い残", f"{jp_date}_売り残", f"{jp_date}_倍率"]

    # --- 初回書き込み ---
    if not existing:
        print("  初回書き込み")
        header = ["銘柄コード", "銘柄名"] + date_header_group
        rows = [
            [row["銘柄コード"], row["銘柄名"], row["買い残"], row["売り残"], row["倍率"]]
            for _, row in df.iterrows()
        ]
        ws = get_or_create_worksheet(gc, MARGIN_SHEET_NAME, min_cols=len(header))
        ws.update(values=[header] + _native_rows(rows), range_name="A1", value_input_option="USER_ENTERED")
        return

    header = existing[0]
    existing_codes = [row[0] for row in existing[1:]]

    if any(h.startswith(jp_date) for h in header):
        print(f"  {jp_date} のデータは既に追加済みです。スキップします")
        return

    col_start = len(header) + 1
    col_end = col_start + COLS_PER_DATE - 1
    ws = get_or_create_worksheet(gc, MARGIN_SHEET_NAME, min_cols=col_end)

    header_range = (
        f"{gspread.utils.rowcol_to_a1(1, col_start)}"
        f":{gspread.utils.rowcol_to_a1(1, col_end)}"
    )
    ws.update(values=[date_header_group], range_name=header_range, value_input_option="USER_ENTERED")

    batch_updates = []
    new_rows = []
    for _, row in df.iterrows():
        code = str(row["銘柄コード"])
        values = [row["買い残"], row["売り残"], row["倍率"]]
        if code in existing_codes:
            row_idx = existing_codes.index(code) + 2
            cell_range = (
                f"{gspread.utils.rowcol_to_a1(row_idx, col_start)}"
                f":{gspread.utils.rowcol_to_a1(row_idx, col_end)}"
            )
            batch_updates.append({"range": cell_range, "values": [_native_row(values)]})
        else:
            full_row = [code, row["銘柄名"]] + [""] * (len(header) - 2) + values
            new_rows.append(_native_row(full_row))

    if batch_updates:
        print(f"  既存銘柄を更新: {len(batch_updates)} 件")
        ws.batch_update(batch_updates, value_input_option="USER_ENTERED")

    if new_rows:
        print(f"  新規銘柄を追加: {len(new_rows)} 件")
        start_row = len(existing) + 1
        total_cols = len(header) + COLS_PER_DATE
        ws = get_or_create_worksheet(gc, MARGIN_SHEET_NAME, min_cols=total_cols)
        end_a1 = gspread.utils.rowcol_to_a1(start_row + len(new_rows) - 1, total_cols)
        ws.update(values=new_rows, range_name=f"A{start_row}:{end_a1}", value_input_option="USER_ENTERED")

    print("  書き込み完了！")


def _native_row(values):
    return [("" if v is None or (isinstance(v, float) and pd.isna(v)) else v) for v in values]


def _native_rows(rows):
    return [_native_row(r) for r in rows]


# =============================================================================
# メイン処理
# =============================================================================

def main():
    if not SPREADSHEET_ID:
        print("エラー: 環境変数 SPREADSHEET_ID が設定されていません")
        sys.exit(1)

    override = os.environ.get("TARGET_DATE_OVERRIDE", "").strip()
    target_date = override if override else get_target_friday()
    print(f"対象申込日: {target_date} ({to_japanese_date(target_date)})")

    try:
        pdf_bytes = download_pdf(target_date)
    except FileNotFoundError as e:
        print(f"警告: {e}")
        print("公表が遅れている可能性があります。ワークフローを再実行するか翌日リトライしてください。")
        sys.exit(0)  # 失敗扱いにしない（公表遅延は珍しくないため）

    df = parse_margin_pdf(pdf_bytes)

    gc = authenticate_google_sheets()
    update_spreadsheet(gc, df, target_date)


if __name__ == "__main__":
    main()
