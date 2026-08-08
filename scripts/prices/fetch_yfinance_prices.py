import yfinance as yf
import pandas as pd
import numpy as np
import time
import os
import re
import random
import sys
import gspread
from datetime import datetime
from zoneinfo import ZoneInfo

# =============================================================================
# 定数定義
# =============================================================================

SPREADSHEET_ID = "1QheVVw97DnHjdymEYNFwvgiQhgX8SX-bPxjlHmZpG2I"

# ---- 各データ種別ごとのタブ (gid) --------------------------------------------
OPEN_SHEET_GID   = 1348359438    # 始値
HIGH_SHEET_GID   = 13194030      # 高値
LOW_SHEET_GID    = 876490713     # 安値
CLOSE_SHEET_GID  = 2080765326    # 終値
VOLUME_SHEET_GID = 328764273     # 出来高

# yfinance のフィールド名 → (書き込み先gid, 日本語ラベル, ログ見出し)
FIELD_CONFIG = {
    "Open":   {"gid": OPEN_SHEET_GID,   "label_jp": "始値", "heading": "【始値データ】"},
    "High":   {"gid": HIGH_SHEET_GID,   "label_jp": "高値", "heading": "【高値データ】"},
    "Low":    {"gid": LOW_SHEET_GID,    "label_jp": "安値", "heading": "【安値データ】"},
    "Close":  {"gid": CLOSE_SHEET_GID,  "label_jp": "終値", "heading": "【終値データ】"},
    "Volume": {"gid": VOLUME_SHEET_GID, "label_jp": "出来高", "heading": "【出来高データ】"},
}
# 取得・書き込みを行うフィールド一覧（この順で処理される）
FIELDS = ["Open", "High", "Low", "Close", "Volume"]

SERVICE_ACCOUNT_FILE = os.environ.get("SERVICE_ACCOUNT_FILE", "service_account.json")

CODES_SHEET_GID = 1376419996

CODE_PATTERN = re.compile(r"^\d[0-9A-Za-z]{3}$")

# 取得日付の設定 -----------------------------------------------------------
# None のままにすると実行当日（日本時間）を自動使用
# 日付を指定したい場合は文字列で書き換える（コメントアウトを外す）
TARGET_DATE_OVERRIDE: str | None = None

# TARGET_DATE_OVERRIDE = "2026-07-24"

TARGET_DATES = (
    [TARGET_DATE_OVERRIDE]
    if TARGET_DATE_OVERRIDE
    else [datetime.now(ZoneInfo("Asia/Tokyo")).strftime("%Y-%m-%d")]
)
# -------------------------------------------------------------------------

# ---- API制御（環境変数で上書き可）------------------------------------------
CHUNK_SIZE   = int(os.environ.get("CHUNK_SIZE", "70"))
SLEEP_TIME   = float(os.environ.get("SLEEP_TIME", "0.7"))
MAX_ATTEMPTS = int(os.environ.get("MAX_ATTEMPTS", "3"))
BACKOFF_BASE = float(os.environ.get("BACKOFF_BASE", "6.0"))

MAX_CODES = int(os.environ.get("MAX_CODES", "0"))

MIN_FILL_RATE = float(os.environ.get("MIN_FILL_RATE", "0.50"))

# 連続空振りがこの回数を超えたらスキップして書き込みへ進む（die しない）
ABORT_AFTER_EMPTY_CHUNKS = int(os.environ.get("ABORT_AFTER_EMPTY_CHUNKS", "3"))

FIXED_COLS = 2

_WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]


def log(msg: str = "") -> None:
    print(msg, flush=True)


def die(msg: str) -> None:
    log(f"::error::{msg}")
    sys.exit(1)


# =============================================================================
# 日付フォーマット変換
# =============================================================================

def to_japanese_date(date_str: str) -> str:
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return f"{dt.year}年{dt.month}月{dt.day}日({_WEEKDAY_JP[dt.weekday()]})"
    except ValueError:
        return date_str


def from_japanese_date(date_str: str) -> str:
    try:
        core = date_str.split("(")[0]
        dt = datetime.strptime(core, "%Y年%m月%d日")
        return dt.strftime("%Y-%m-%d")
    except (ValueError, IndexError):
        return date_str


# =============================================================================
# 認証・シート操作
# =============================================================================

def authenticate_google_sheets() -> gspread.Client:
    log("Google認証中...")
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        die(
            f"認証ファイルが見つかりません → {SERVICE_ACCOUNT_FILE}\n"
            "  ワークフローで GOOGLE_SERVICE_ACCOUNT_JSON を書き出すステップが"
            "実行されているか確認してください。"
        )
    return gspread.service_account(filename=SERVICE_ACCOUNT_FILE)


def get_worksheet_by_gid(
    gc: gspread.Client, gid: int, min_cols: int = 50
) -> gspread.Worksheet:
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    ws = spreadsheet.get_worksheet_by_id(gid)
    if ws is None:
        die(f"gid={gid} のタブが見つかりません。")
    if ws.col_count < min_cols:
        ws.add_cols(min_cols - ws.col_count)
        ws = spreadsheet.get_worksheet_by_id(gid)
    return ws


# =============================================================================
# 銘柄コード読み込み
# =============================================================================

def load_stock_codes_from_sheet(gc: gspread.Client, gid: int) -> list[str]:
    ws = get_worksheet_by_gid(gc, gid)
    log(f"銘柄コード読み込み: タブ '{ws.title}' (gid={gid})")

    rows = ws.get_all_values()
    if not rows:
        die(f"タブ '{ws.title}' にデータがありません。")

    max_cols = max(len(r) for r in rows)
    best_col, best_count = -1, 0
    for c in range(max_cols):
        count = sum(
            1 for r in rows if c < len(r) and CODE_PATTERN.match(r[c].strip())
        )
        if count > best_count:
            best_col, best_count = c, count

    if best_col < 0 or best_count == 0:
        die(f"タブ '{ws.title}' から銘柄コードらしい列が見つかりませんでした。")

    codes = [
        r[best_col].strip()
        for r in rows
        if best_col < len(r) and CODE_PATTERN.match(r[best_col].strip())
    ]
    codes = list(dict.fromkeys(codes))

    col_letter = gspread.utils.rowcol_to_a1(1, best_col + 1).rstrip("1")
    log(f"  {col_letter}列から {len(codes)} 銘柄を読み込み")

    if MAX_CODES > 0 and len(codes) > MAX_CODES:
        codes = codes[:MAX_CODES]
        log(f"  MAX_CODES={MAX_CODES} により先頭{MAX_CODES}銘柄に制限")

    return codes


# =============================================================================
# データ取得
# =============================================================================

def _download_with_retry(chunk: list[str], start: str, end: str):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        raw = None
        try:
            raw = yf.download(
                chunk,
                start=start,
                end=end,
                progress=False,
                auto_adjust=False,
                group_by="column",
                threads=False,
                timeout=45,
            )
        except Exception as e:
            log(f"    [attempt {attempt}/{MAX_ATTEMPTS}] {type(e).__name__}: {e}")

        if raw is not None and not raw.empty:
            return raw

        if attempt < MAX_ATTEMPTS:
            wait = BACKOFF_BASE * (2 ** (attempt - 1)) + random.uniform(0, 3.0)
            log(f"    [attempt {attempt}/{MAX_ATTEMPTS}] 空データ。{wait:.1f}秒待機")
            time.sleep(wait)
    return None


def _extract_field(raw: pd.DataFrame, field: str, chunk: list[str]):
    cols = raw.columns
    if isinstance(cols, pd.MultiIndex):
        if field in set(cols.get_level_values(0)):
            sub = raw[field]
        else:
            try:
                sub = raw.xs(field, axis=1, level=1)
            except KeyError:
                return None
    else:
        if field not in cols:
            return None
        sub = raw[[field]].copy()
        sub.columns = [chunk[0]]

    if isinstance(sub, pd.Series):
        sub = sub.to_frame(name=chunk[0])
    return sub


def fetch_prices(codes: list[str], target_dates: list[str]) -> dict:
    """始値・高値・安値・終値・出来高の5項目を取得する。

    yf.download は1回の呼び出しで OHLCV をまとめて返すため、
    追加のAPIコストなしに Open/High/Low を Close/Volume と同時に取得できる。
    """
    tickers = [f"{c}.T" for c in codes]
    start = str(pd.to_datetime(min(target_dates)) - pd.Timedelta(days=7))[:10]
    end = str(pd.to_datetime(max(target_dates)) + pd.Timedelta(days=1))[:10]
    target_date_map = {pd.to_datetime(d).date(): d for d in target_dates}

    total_chunks = (len(tickers) + CHUNK_SIZE - 1) // CHUNK_SIZE
    log(f"\n株価データ取得中（{len(codes)}銘柄 / {CHUNK_SIZE}銘柄ずつ / {total_chunks}チャンク）")
    log(f"  期間: {start} 〜 {end}（対象日: {', '.join(target_dates)}）")
    log(f"  取得項目: {', '.join(FIELDS)}")

    results = {field: {d: {} for d in target_dates} for field in FIELDS}
    filled_codes: set[str] = set()
    consecutive_empty = 0

    for i in range(0, len(tickers), CHUNK_SIZE):
        chunk = tickers[i : i + CHUNK_SIZE]
        chunk_num = i // CHUNK_SIZE + 1

        raw = _download_with_retry(chunk, start, end)

        if raw is None:
            consecutive_empty += 1
            log(f"  [{chunk_num}/{total_chunks}] 取得0件")
            if consecutive_empty >= ABORT_AFTER_EMPTY_CHUNKS:
                # ↓ die() → break に変更
                # 未上場コード等で後半が空になっても取得済みデータを書き込む
                log(
                    f"  ⚠ {consecutive_empty}チャンク連続で取得できませんでした。"
                    f"（未上場コード or レート制限）"
                    f"残り {total_chunks - chunk_num} チャンクをスキップして書き込みへ進みます。"
                )
                break
            time.sleep(SLEEP_TIME * 2)
            continue

        if chunk_num == 1:
            log(f"  DEBUG raw.shape   : {raw.shape}")
            log(f"  DEBUG columns[:6] : {list(raw.columns[:6])}")
            log(f"  DEBUG index[-5:]  : {[str(x)[:10] for x in raw.index[-5:]]}")

        idx = raw.index
        if getattr(idx, "tz", None) is not None:
            idx = idx.tz_localize(None)
        date_to_pos = {ts.date(): pos for pos, ts in enumerate(idx)}

        if chunk_num == 1:
            log(f"  DEBUG 取得できた日付 : {sorted(date_to_pos.keys())}")
            log(f"  DEBUG 探している日付 : {sorted(target_date_map.keys())}")
            missing = [d for d in target_date_map if d not in date_to_pos]
            if missing:
                log(f"  ⚠ 対象日が取得範囲に存在しません: {missing}")

        filled_chunk = 0
        for field in FIELDS:
            sub = _extract_field(raw, field, chunk)
            if sub is None:
                log(f"  [{chunk_num}/{total_chunks}] ⚠ '{field}' 列が見つからず")
                continue
            for ticker in chunk:
                if ticker not in sub.columns:
                    continue
                code = ticker[:-2]
                col = sub[ticker]
                for d_date, d_str in target_date_map.items():
                    pos = date_to_pos.get(d_date)
                    if pos is None:
                        continue
                    val = col.iloc[pos]
                    if pd.notna(val):
                        results[field][d_str][code] = (
                            int(val) if field == "Volume" else float(val)
                        )
                        filled_codes.add(code)
                        filled_chunk += 1

        if filled_chunk == 0:
            consecutive_empty += 1
            log(f"  [{chunk_num}/{total_chunks}] {len(chunk)}銘柄中 0件（未上場コード or レート制限）")
            if consecutive_empty >= ABORT_AFTER_EMPTY_CHUNKS:
                # ↓ die() → break に変更
                # 600A系など Yahoo 未対応コードが後半に固まっていても止まらない
                log(
                    f"  ⚠ {consecutive_empty}チャンク連続で空データ。"
                    f"残り {total_chunks - chunk_num} チャンクをスキップして書き込みへ進みます。"
                )
                break
        else:
            consecutive_empty = 0
            log(f"  [{chunk_num}/{total_chunks}] {len(chunk)}銘柄中 {filled_chunk}件 取得")

        time.sleep(SLEEP_TIME + random.uniform(0, 1.0))

    fill_rate = len(filled_codes) / len(codes) if codes else 0.0
    log(f"\n取得成功: {len(filled_codes)}/{len(codes)} 銘柄 ({fill_rate:.1%})")

    if fill_rate < MIN_FILL_RATE:
        die(
            f"取得率 {fill_rate:.1%} が閾値 {MIN_FILL_RATE:.0%} を下回りました。"
            "空の日付列を作らないため、シートへの書き込みを中止しました。"
        )

    return results


def build_dataframe(codes: list[str], values: dict, target_dates: list[str]) -> pd.DataFrame:
    df = pd.DataFrame({"銘柄コード": codes, "備考": ""})
    for d in target_dates:
        df[d] = df["銘柄コード"].map(values.get(d, {}))
    return df


# =============================================================================
# スプレッドシート書き込み
# =============================================================================

def to_native(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    if isinstance(value, (np.integer, int)):
        return int(value)
    if isinstance(value, (np.floating, float)):
        return float(value)
    return str(value)


def update_spreadsheet(gc: gspread.Client, df: pd.DataFrame, gid: int):
    ws = get_worksheet_by_gid(gc, gid)
    log(f"\n{ws.title} にデータ書き込み中...")
    existing = ws.get_all_values()

    df_date_cols = list(df.columns[FIXED_COLS:])
    iso_to_jp = {d: to_japanese_date(d) for d in df_date_cols}

    non_empty = [d for d in df_date_cols if df[d].notna().any()]
    skipped = [d for d in df_date_cols if d not in non_empty]
    if skipped:
        log(f"  ⚠ 値が0件のためスキップ: {', '.join(skipped)}")
    if not non_empty:
        log("  書き込む列がありません")
        return
    df_date_cols = non_empty

    if not existing:
        log("  初回書き込み")
        header = list(df.columns[:FIXED_COLS]) + [iso_to_jp[c] for c in df_date_cols]
        out = df[list(df.columns[:FIXED_COLS]) + df_date_cols]
        rows = [[to_native(v) for v in row] for row in out.values]
        ws = get_worksheet_by_gid(gc, gid, min_cols=len(header))
        ws.update(values=[header] + rows, range_name="A1",
                  value_input_option="USER_ENTERED")
        return

    header = existing[0]
    code_to_row = {}
    for r, row in enumerate(existing[1:], start=2):
        if row and row[0]:
            code_to_row.setdefault(row[0].strip(), r)

    existing_dates_iso = {from_japanese_date(d) for d in header[FIXED_COLS:]}
    new_cols_iso = [d for d in df_date_cols if d not in existing_dates_iso]

    if not new_cols_iso:
        log("  追加する新しい日付はありません")
        return

    new_cols_jp = [iso_to_jp[d] for d in new_cols_iso]
    log(f"  追加日付: {', '.join(new_cols_jp)}")

    col_start = len(header) + 1
    col_end = col_start + len(new_cols_jp) - 1
    ws = get_worksheet_by_gid(gc, gid, min_cols=col_end)
    header_range = (
        f"{gspread.utils.rowcol_to_a1(1, col_start)}"
        f":{gspread.utils.rowcol_to_a1(1, col_end)}"
    )
    ws.update(values=[new_cols_jp], range_name=header_range,
              value_input_option="USER_ENTERED")

    batch_updates = []
    new_rows = []
    for _, row in df.iterrows():
        code = str(row["銘柄コード"])
        values = [to_native(row[d]) for d in new_cols_iso]
        if all(v == "" for v in values):
            continue
        row_idx = code_to_row.get(code)
        if row_idx:
            cell_range = (
                f"{gspread.utils.rowcol_to_a1(row_idx, col_start)}"
                f":{gspread.utils.rowcol_to_a1(row_idx, col_end)}"
            )
            batch_updates.append({"range": cell_range, "values": [values]})
        else:
            new_rows.append([code] + [""] * (len(header) - 1) + values)

    if batch_updates:
        log(f"  既存銘柄を更新: {len(batch_updates)} 件")
        for i in range(0, len(batch_updates), 500):
            ws.batch_update(batch_updates[i : i + 500],
                            value_input_option="USER_ENTERED")

    if new_rows:
        log(f"  新規銘柄を追加: {len(new_rows)} 件")
        start_row = len(existing) + 1
        total_cols = len(header) + len(new_cols_jp)
        ws = get_worksheet_by_gid(gc, gid, min_cols=total_cols)
        end_a1 = gspread.utils.rowcol_to_a1(start_row + len(new_rows) - 1, total_cols)
        ws.update(values=new_rows, range_name=f"A{start_row}:{end_a1}",
                  value_input_option="USER_ENTERED")

    log("  書き込み完了！")


def sort_date_columns(gc: gspread.Client, gid: int):
    ws = get_worksheet_by_gid(gc, gid)
    log(f"\n{ws.title} の日付列をソート中...")
    all_data = ws.get_all_values()

    if not all_data or len(all_data) < 2 or len(all_data[0]) <= FIXED_COLS:
        log("  ソート対象の日付列がありません。スキップします")
        return

    header = all_data[0]
    fixed_header = header[:FIXED_COLS]
    date_header = header[FIXED_COLS:]

    def parse_date(s: str) -> datetime:
        try:
            return datetime.strptime(from_japanese_date(s), "%Y-%m-%d")
        except ValueError:
            return datetime.min

    order = sorted(range(len(date_header)),
                   key=lambda i: parse_date(date_header[i]), reverse=True)
    sorted_header = [date_header[i] for i in order]

    new_data = [fixed_header + sorted_header]
    for row in all_data[1:]:
        fixed_vals = row[:FIXED_COLS] + [""] * max(0, FIXED_COLS - len(row))
        date_vals = row[FIXED_COLS:]
        date_vals += [""] * max(0, len(date_header) - len(date_vals))
        new_data.append(fixed_vals + [date_vals[i] for i in order])

    total_cols = len(new_data[0])
    ws = get_worksheet_by_gid(gc, gid, min_cols=total_cols)
    end_a1 = gspread.utils.rowcol_to_a1(len(new_data), total_cols)
    ws.update(values=new_data, range_name=f"A1:{end_a1}",
              value_input_option="USER_ENTERED")

    preview = " → ".join(sorted_header[:5])
    suffix = "..." if len(sorted_header) > 5 else ""
    log(f"  ソート完了: {len(sorted_header)} 列（{preview}{suffix}）")


# =============================================================================
# 疎通確認
# =============================================================================

def smoke_test(start: str, end: str) -> None:
    log(f"\n[疎通確認] 7203.T を単独取得... (yfinance {yf.__version__})")
    raw = _download_with_retry(["7203.T"], start, end)
    if raw is None or raw.empty:
        die(
            "疎通確認に失敗しました。考えられる原因:\n"
            "  (1) yfinance が古い（0.2.51以前）→ pip install --upgrade yfinance curl_cffi\n"
            "  (2) Yahoo Finance のIP制限         → 時間をおいて再実行\n"
            f"  現在のバージョン: yfinance {yf.__version__}"
        )
    sub = _extract_field(raw, "Close", ["7203.T"])
    if sub is None or sub.dropna(how="all").empty:
        die("疎通確認: 終値が全てNaNでした。pip install --upgrade yfinance curl_cffi を試してください。")
    log(f"  OK: {len(sub.dropna(how='all'))} 営業日分の終値を取得")


# =============================================================================
# メイン処理
# =============================================================================

def main():
    log("=" * 60)
    log(f"取得対象日（日本時間）: {TARGET_DATES[0]}")
    log(f"yfinance {yf.__version__} / pandas {pd.__version__}")
    log(f"CHUNK_SIZE={CHUNK_SIZE} SLEEP_TIME={SLEEP_TIME} MAX_CODES={MAX_CODES or '無制限'}")
    log("=" * 60)

    gc = authenticate_google_sheets()

    codes = load_stock_codes_from_sheet(gc, CODES_SHEET_GID)
    if not codes:
        die("銘柄コードが0件でした。タブの内容を確認してください。")

    start = str(pd.to_datetime(min(TARGET_DATES)) - pd.Timedelta(days=7))[:10]
    end = str(pd.to_datetime(max(TARGET_DATES)) + pd.Timedelta(days=1))[:10]
    smoke_test(start, end)

    values = fetch_prices(codes, TARGET_DATES)

    # 始値・高値・安値・終値・出来高を、それぞれ対応するタブへ書き込む
    for field in FIELDS:
        cfg = FIELD_CONFIG[field]
        log("\n" + "=" * 50)
        log(cfg["heading"])
        log("=" * 50)
        field_df = build_dataframe(codes, values[field], TARGET_DATES)
        update_spreadsheet(gc, field_df, cfg["gid"])
        sort_date_columns(gc, cfg["gid"])

    log("\n" + "=" * 50)
    log("全処理完了！")
    log(f"  スプレッドシートID : {SPREADSHEET_ID}")
    log(f"  銘柄数             : {len(codes)}")
    log(f"  取得日付           : {', '.join(TARGET_DATES)}")
    log(f"  取得項目           : {', '.join(FIELD_CONFIG[f]['label_jp'] for f in FIELDS)}")
    log("=" * 50)


if __name__ == "__main__":
    main()
