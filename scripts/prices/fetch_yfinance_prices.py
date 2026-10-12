#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetch_yfinance_prices.py  —  JPX全銘柄の OHLCV を Google スプレッドシートへ書き込む

【旧版からの主な変更】
1. 営業日カレンダーを yfinance(7203.T) から取得し、休場日・土日・大引け前の当日は
   最初から対象外にする（空の日付列を作らない / 朝の実行で全滅しない）。
2. 実行のたびに「直近 BACKFILL_DAYS 日」を対象にし、既存の日付列であっても
   "空セルだけ" を補完する。過去に取りこぼした日が次回実行で自動的に埋まる。
   （旧版は見出しに日付があるだけで「追加なし」と判定していたため、
     見出しだけ存在して中身が空の列は永久に埋まらなかった）
3. 見出しとデータを同じ batch_update に入れて書く（見出しだけ残る事故を防ぐ）。
4. 取得率の判定を日付ごとに行う（1日ダメでも他の日は書き込む）。
5. 日付列がすでに新しい順なら、全面書き直し（ソート）をスキップ。
6. 休場日にできた空の日付列を掃除するオプション（CLEAN_HOLIDAY_COLUMNS=1）。
7. Sheets API の 429/5xx は自動リトライ。

【環境変数】
  TARGET_DATE          空=自動 / "2026-10-09" / "2026-10-06,2026-10-09" / "2026-10-01:2026-10-09"
  BACKFILL_DAYS        自動モードで遡るカレンダー日数（既定 14）
  MARKET_CLOSE_HHMM    この時刻より前は当日を対象外にする（既定 16:00）
  FORCE_INTRADAY       1 で大引け前でも当日を対象にする（場中の暫定値が入るので非推奨）
  CLEAN_HOLIDAY_COLUMNS 1 で、休場日の「完全に空の」日付列を削除する
  CHUNK_SIZE / SLEEP_TIME / MAX_ATTEMPTS / BACKOFF_BASE / MAX_CODES / MIN_FILL_RATE
  ABORT_AFTER_EMPTY_CHUNKS / SERVICE_ACCOUNT_FILE
"""

from __future__ import annotations

import os
import random
import re
import sys
import time
from datetime import date, datetime, time as dtime, timedelta
from zoneinfo import ZoneInfo

import gspread
import numpy as np
import pandas as pd
import yfinance as yf

# =============================================================================
# 定数定義
# =============================================================================

SPREADSHEET_ID = "1QheVVw97DnHjdymEYNFwvgiQhgX8SX-bPxjlHmZpG2I"

# ---- 各データ種別ごとのタブ (gid) --------------------------------------------
OPEN_SHEET_GID = 1348359438    # 始値
HIGH_SHEET_GID = 13194030      # 高値
LOW_SHEET_GID = 876490713      # 安値
CLOSE_SHEET_GID = 2080765326   # 終値
VOLUME_SHEET_GID = 328764273   # 出来高

FIELD_CONFIG = {
    "Open":   {"gid": OPEN_SHEET_GID,   "label_jp": "始値",   "heading": "【始値データ】"},
    "High":   {"gid": HIGH_SHEET_GID,   "label_jp": "高値",   "heading": "【高値データ】"},
    "Low":    {"gid": LOW_SHEET_GID,    "label_jp": "安値",   "heading": "【安値データ】"},
    "Close":  {"gid": CLOSE_SHEET_GID,  "label_jp": "終値",   "heading": "【終値データ】"},
    "Volume": {"gid": VOLUME_SHEET_GID, "label_jp": "出来高", "heading": "【出来高データ】"},
}
FIELDS = ["Open", "High", "Low", "Close", "Volume"]

SERVICE_ACCOUNT_FILE = os.environ.get("SERVICE_ACCOUNT_FILE", "service_account.json")

CODES_SHEET_GID = 1376419996
CODE_PATTERN = re.compile(r"^\d[0-9A-Za-z]{3}$")

# 先頭の固定列（A:銘柄コード, B:備考/銘柄名）
FIXED_COLS = 2
FIXED_HEADERS = ["銘柄コード", "備考"]

# ---- 取得日付の設定 -----------------------------------------------------------
JST = ZoneInfo("Asia/Tokyo")
TARGET_DATE_ENV = (os.environ.get("TARGET_DATE") or "").strip()
BACKFILL_DAYS = int(os.environ.get("BACKFILL_DAYS", "14"))
MARKET_CLOSE_HHMM = os.environ.get("MARKET_CLOSE_HHMM", "16:00")
FORCE_INTRADAY = os.environ.get("FORCE_INTRADAY", "0") == "1"
CLEAN_HOLIDAY_COLUMNS = os.environ.get("CLEAN_HOLIDAY_COLUMNS", "0") == "1"
CALENDAR_TICKER = "7203.T"   # 営業日カレンダーの基準にする銘柄

# ---- API制御（環境変数で上書き可）------------------------------------------
CHUNK_SIZE = int(os.environ.get("CHUNK_SIZE", "90"))
SLEEP_TIME = float(os.environ.get("SLEEP_TIME", "0.25"))
MAX_ATTEMPTS = int(os.environ.get("MAX_ATTEMPTS", "2"))
BACKOFF_BASE = float(os.environ.get("BACKOFF_BASE", "4.0"))
MAX_CODES = int(os.environ.get("MAX_CODES", "0"))
MIN_FILL_RATE = float(os.environ.get("MIN_FILL_RATE", "0.30"))
ABORT_AFTER_EMPTY_CHUNKS = int(os.environ.get("ABORT_AFTER_EMPTY_CHUNKS", "3"))

YF_DOWNLOAD_TIMEOUT_SEC = 45
BACKOFF_JITTER_MAX_SEC = 3.0
EMPTY_CHUNK_SLEEP_MULTIPLIER = 2
CHUNK_SLEEP_JITTER_MAX_SEC = 1.0
PRICE_FETCH_LOOKAHEAD_DAYS = 1     # yfinance の end は排他的なため +1日

# ---- Google スプレッドシート ------------------------------------------------
DEFAULT_MIN_COLS = 50
BATCH_UPDATE_CHUNK_SIZE = 500
SORT_PREVIEW_COUNT = 5
SHEETS_MAX_ATTEMPTS = 4
SHEETS_BACKOFF_SEC = 20

_WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]
_DATE_RE = re.compile(r"(\d{4})\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})")


def log(msg: str = "") -> None:
    print(msg, flush=True)


def die(msg: str) -> None:
    log(f"::error::{msg}")
    sys.exit(1)


# =============================================================================
# 日付ユーティリティ
# =============================================================================

def to_japanese_date(iso: str) -> str:
    """'2026-10-09' → '2026年10月9日(金)'"""
    try:
        d = date.fromisoformat(iso)
        return f"{d.year}年{d.month}月{d.day}日({_WEEKDAY_JP[d.weekday()]})"
    except ValueError:
        return iso


def parse_date_text(s) -> str | None:
    """見出しセルの文字列から ISO 日付を取り出す。
    '2026年10月9日(金)' / '2026年10月9日' / '2026/10/09' / '2026-10-09' に対応。"""
    m = _DATE_RE.search(str(s))
    if not m:
        return None
    try:
        return date(int(m[1]), int(m[2]), int(m[3])).isoformat()
    except ValueError:
        return None


def parse_target_dates(raw: str) -> list[str]:
    """TARGET_DATE の書式: 単日 / カンマ区切り / 'start:end' の範囲"""
    raw = raw.strip()
    try:
        if ":" in raw:
            a, b = (x.strip() for x in raw.split(":", 1))
            s, e = date.fromisoformat(a), date.fromisoformat(b)
            if s > e:
                s, e = e, s
            return [(s + timedelta(days=i)).isoformat() for i in range((e - s).days + 1)]
        return sorted({date.fromisoformat(x.strip()).isoformat()
                       for x in raw.split(",") if x.strip()})
    except ValueError:
        die(f"TARGET_DATE の書式が不正です: '{raw}'  "
            "（例: 2026-10-09 / 2026-10-06,2026-10-09 / 2026-10-01:2026-10-09）")
        return []


def last_complete_day(now: datetime) -> date:
    """データが確定している最新の暦日。大引け後(既定16:00以降)なら当日、それ以前は前日。"""
    if FORCE_INTRADAY:
        return now.date()
    h, m = (int(x) for x in MARKET_CLOSE_HHMM.split(":"))
    return now.date() if now.time() >= dtime(h, m) else now.date() - timedelta(days=1)


def build_candidate_dates(now: datetime) -> tuple[list[str], bool, date]:
    """(候補日リスト[ISO昇順], 明示指定か, 確定済み最新日) を返す。"""
    last = last_complete_day(now)
    if TARGET_DATE_ENV:
        wanted = parse_target_dates(TARGET_DATE_ENV)
        too_new = [d for d in wanted if date.fromisoformat(d) > last]
        if too_new:
            log(f"⚠ 大引け前または未来のため対象外にしました: {', '.join(too_new)}"
                f"（確定済みの最新日: {last}）")
        return [d for d in wanted if d not in too_new], True, last
    n = max(BACKFILL_DAYS, 1)
    days = [(last - timedelta(days=i)).isoformat() for i in range(n - 1, -1, -1)]
    return days, False, last


# =============================================================================
# 認証・シート操作
# =============================================================================

def sheets_call(fn, *args, **kwargs):
    """Sheets API 呼び出し。429/5xx は待ってリトライする。"""
    for attempt in range(1, SHEETS_MAX_ATTEMPTS + 1):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = getattr(getattr(e, "response", None), "status_code", None)
            if status not in (429, 500, 502, 503, 504) or attempt == SHEETS_MAX_ATTEMPTS:
                raise
            wait = SHEETS_BACKOFF_SEC * attempt
            log(f"    Sheets API {status}。{wait}秒待機して再試行 ({attempt}/{SHEETS_MAX_ATTEMPTS})")
            time.sleep(wait)


def authenticate_google_sheets() -> gspread.Client:
    log("Google認証中...")
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        die(
            f"認証ファイルが見つかりません → {SERVICE_ACCOUNT_FILE}\n"
            "  ワークフローで GOOGLE_SERVICE_ACCOUNT_JSON を書き出すステップが"
            "実行されているか確認してください。"
        )
    return gspread.service_account(filename=SERVICE_ACCOUNT_FILE)


_SPREADSHEET = None


def get_spreadsheet(gc: gspread.Client):
    global _SPREADSHEET
    if _SPREADSHEET is None:
        _SPREADSHEET = sheets_call(gc.open_by_key, SPREADSHEET_ID)
    return _SPREADSHEET


def get_worksheet_by_gid(gc: gspread.Client, gid: int, min_cols: int = DEFAULT_MIN_COLS):
    sp = get_spreadsheet(gc)
    try:
        ws = sheets_call(sp.get_worksheet_by_id, gid)
    except gspread.exceptions.WorksheetNotFound:
        ws = None
    if ws is None:
        die(f"gid={gid} のタブが見つかりません。")
    if ws.col_count < min_cols:
        sheets_call(ws.add_cols, min_cols - ws.col_count)
        ws = sheets_call(sp.get_worksheet_by_id, gid)
    return ws


def ensure_rows(ws, n_rows: int) -> None:
    if ws.row_count < n_rows:
        sheets_call(ws.add_rows, n_rows - ws.row_count)


def last_nonempty_len(row: list) -> int:
    for i in range(len(row), 0, -1):
        if str(row[i - 1]).strip():
            return i
    return 0


def cell_is_empty(grid: list[list], r: int, c: int) -> bool:
    """grid は get_all_values の結果。r, c は 1 始まり。範囲外は空扱い。"""
    if r - 1 >= len(grid):
        return True
    row = grid[r - 1]
    return c - 1 >= len(row) or not str(row[c - 1]).strip()


# =============================================================================
# 銘柄コード読み込み
# =============================================================================

def load_stock_codes_from_sheet(gc: gspread.Client, gid: int) -> list[str]:
    ws = get_worksheet_by_gid(gc, gid)
    log(f"銘柄コード読み込み: タブ '{ws.title}' (gid={gid})")

    rows = sheets_call(ws.get_all_values)
    if not rows:
        die(f"タブ '{ws.title}' にデータがありません。")

    max_cols = max(len(r) for r in rows)
    best_col, best_count = -1, 0
    for c in range(max_cols):
        count = sum(1 for r in rows if c < len(r) and CODE_PATTERN.match(r[c].strip()))
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
                timeout=YF_DOWNLOAD_TIMEOUT_SEC,
            )
        except Exception as e:
            log(f"    [attempt {attempt}/{MAX_ATTEMPTS}] {type(e).__name__}: {e}")

        if raw is not None and not raw.empty:
            return raw

        if attempt < MAX_ATTEMPTS:
            wait = BACKOFF_BASE * (2 ** (attempt - 1)) + random.uniform(0, BACKOFF_JITTER_MAX_SEC)
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


def get_trading_calendar(days: list[str]) -> set[date]:
    """7203.T の日足が存在する日 = 東証の営業日。休場日・土日・未確定日は含まれない。"""
    start = min(days)
    end = (date.fromisoformat(max(days)) + timedelta(days=PRICE_FETCH_LOOKAHEAD_DAYS)).isoformat()
    log(f"\n[営業日カレンダー取得] {CALENDAR_TICKER}  {start} 〜 {end} (yfinance {yf.__version__})")

    raw = _download_with_retry([CALENDAR_TICKER], start, end)
    if raw is None or raw.empty:
        die(
            "営業日カレンダーを取得できませんでした。考えられる原因:\n"
            "  (1) 指定期間がすべて休場日\n"
            "  (2) yfinance が古い → pip install --upgrade yfinance curl_cffi\n"
            "  (3) Yahoo Finance のIP制限 → 時間をおいて再実行"
        )
    sub = _extract_field(raw, "Close", [CALENDAR_TICKER])
    if sub is None:
        die("営業日カレンダー: Close 列が見つかりませんでした。")
    sub = sub.dropna(how="all")
    if sub.empty:
        die("営業日カレンダー: 終値が全てNaNでした。pip install --upgrade yfinance curl_cffi を試してください。")

    idx = sub.index
    if getattr(idx, "tz", None) is not None:
        idx = idx.tz_localize(None)
    cal = {ts.date() for ts in idx}
    log(f"  取得できた営業日: {', '.join(sorted(d.isoformat() for d in cal))}")
    return cal


def fetch_prices(codes: list[str], trade_dates: list[str]) -> tuple[dict, list[str]]:
    """OHLCV を取得する。戻り値: (results[field][date][code], 取得率を満たした日付リスト)"""
    tickers = [f"{c}.T" for c in codes]
    start = min(trade_dates)
    end = (date.fromisoformat(max(trade_dates)) + timedelta(days=PRICE_FETCH_LOOKAHEAD_DAYS)).isoformat()
    target_date_map = {date.fromisoformat(d): d for d in trade_dates}

    total_chunks = (len(tickers) + CHUNK_SIZE - 1) // CHUNK_SIZE
    log(f"\n株価データ取得中（{len(codes)}銘柄 / {CHUNK_SIZE}銘柄ずつ / {total_chunks}チャンク）")
    log(f"  期間: {start} 〜 {end}（対象日: {', '.join(trade_dates)}）")
    log(f"  取得項目: {', '.join(FIELDS)}")

    results = {field: {d: {} for d in trade_dates} for field in FIELDS}
    filled_by_date: dict[str, set[str]] = {d: set() for d in trade_dates}
    consecutive_empty = 0

    for i in range(0, len(tickers), CHUNK_SIZE):
        chunk = tickers[i: i + CHUNK_SIZE]
        chunk_num = i // CHUNK_SIZE + 1

        raw = _download_with_retry(chunk, start, end)

        if raw is None:
            consecutive_empty += 1
            log(f"  [{chunk_num}/{total_chunks}] 取得0件")
            if consecutive_empty >= ABORT_AFTER_EMPTY_CHUNKS:
                log(
                    f"  ⚠ {consecutive_empty}チャンク連続で取得できませんでした。"
                    f"（未上場コード or レート制限）"
                    f"残り {total_chunks - chunk_num} チャンクをスキップして書き込みへ進みます。"
                )
                break
            time.sleep(SLEEP_TIME * EMPTY_CHUNK_SLEEP_MULTIPLIER)
            continue

        idx = raw.index
        if getattr(idx, "tz", None) is not None:
            idx = idx.tz_localize(None)
        date_to_pos = {ts.date(): pos for pos, ts in enumerate(idx)}

        if chunk_num == 1:
            log(f"  DEBUG raw.shape   : {raw.shape}")
            log(f"  DEBUG 取得できた日付 : {[d.isoformat() for d in sorted(date_to_pos)]}")
            missing = [d.isoformat() for d in target_date_map if d not in date_to_pos]
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
                        results[field][d_str][code] = int(val) if field == "Volume" else float(val)
                        filled_by_date[d_str].add(code)
                        filled_chunk += 1

        if filled_chunk == 0:
            consecutive_empty += 1
            log(f"  [{chunk_num}/{total_chunks}] {len(chunk)}銘柄中 0件（未上場コード or レート制限）")
            if consecutive_empty >= ABORT_AFTER_EMPTY_CHUNKS:
                log(
                    f"  ⚠ {consecutive_empty}チャンク連続で空データ。"
                    f"残り {total_chunks - chunk_num} チャンクをスキップして書き込みへ進みます。"
                )
                break
        else:
            consecutive_empty = 0
            log(f"  [{chunk_num}/{total_chunks}] {len(chunk)}銘柄中 {filled_chunk}件 取得")

        time.sleep(SLEEP_TIME + random.uniform(0, CHUNK_SLEEP_JITTER_MAX_SEC))

    # ---- 日付ごとの取得率判定 ----
    log("\n日付ごとの取得率:")
    ok_dates: list[str] = []
    for d in trade_dates:
        n = len(filled_by_date[d])
        rate = n / len(codes) if codes else 0.0
        ok = rate >= MIN_FILL_RATE
        log(f"  {d}: {n}/{len(codes)} ({rate:.1%})  {'OK' if ok else f'NG（閾値{MIN_FILL_RATE:.0%}未満のため書き込まない）'}")
        if ok:
            ok_dates.append(d)

    if not ok_dates:
        die(
            f"全ての対象日で取得率が閾値 {MIN_FILL_RATE:.0%} を下回りました。"
            "空の日付列を作らないため、シートへの書き込みを中止しました。"
        )
    return results, ok_dates


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


def update_sheet(gc: gspread.Client, gid: int, codes: list[str],
                 field_values: dict[str, dict[str, float]], dates: list[str]) -> None:
    """日付ごとに:
       ・見出しが無ければ新規列を作る
       ・見出しがあっても、空のセルには値を補完する（既存の値は上書きしない）
    """
    ws = get_worksheet_by_gid(gc, gid)
    log(f"\n{ws.title} にデータ書き込み中...")
    existing = sheets_call(ws.get_all_values)

    dates = sorted(d for d in dates if field_values.get(d))
    if not dates:
        log("  書き込める値がありません")
        return

    # ---- 初回書き込み ----
    if not existing:
        log("  初回書き込み")
        desc = sorted(dates, reverse=True)
        header = FIXED_HEADERS + [to_japanese_date(d) for d in desc]
        rows = [[code, ""] + [to_native(field_values[d].get(code)) for d in desc] for code in codes]
        ws = get_worksheet_by_gid(gc, gid, min_cols=len(header))
        ensure_rows(ws, len(rows) + 1)
        sheets_call(ws.update, values=[header] + rows, range_name="A1",
                    value_input_option="USER_ENTERED")
        log("  書き込み完了！")
        return

    # ---- 既存シートの構造を把握 ----
    header = existing[0]
    hdr_len = max(last_nonempty_len(header), FIXED_COLS)

    col_of: dict[str, int] = {}          # ISO日付 → 列番号(1始まり)
    for ci in range(FIXED_COLS, hdr_len):
        iso = parse_date_text(header[ci]) if ci < len(header) else None
        if iso and iso not in col_of:
            col_of[iso] = ci + 1

    code_to_row: dict[str, int] = {}     # 銘柄コード → 行番号(1始まり)
    for r, row in enumerate(existing[1:], start=2):
        if row and str(row[0]).strip():
            code_to_row.setdefault(str(row[0]).strip(), r)

    updates: list[dict] = []

    # ---- 新規日付列（見出しは先頭の更新に入れ、データと同じ batch で送る）----
    new_dates = [d for d in dates if d not in col_of]
    next_col = hdr_len + 1
    for d in new_dates:
        col_of[d] = next_col
        updates.append({"range": gspread.utils.rowcol_to_a1(1, next_col),
                        "values": [[to_japanese_date(d)]]})
        next_col += 1
    total_cols = next_col - 1
    if new_dates:
        log(f"  新規日付列: {', '.join(to_japanese_date(d) for d in new_dates)}")

    # ---- 空セルの補完（連続する行は1つの範囲にまとめる）----
    filled_cells = 0
    healed_dates: list[str] = []
    for d in dates:
        c = col_of[d]
        vals = field_values[d]
        targets: list[tuple[int, float]] = []
        for code in codes:
            if code not in vals:
                continue
            r = code_to_row.get(code)
            if r is None:
                continue
            if cell_is_empty(existing, r, c):
                targets.append((r, vals[code]))
        targets.sort()

        if targets and d not in new_dates:
            healed_dates.append(d)

        i = 0
        while i < len(targets):
            j = i
            while j + 1 < len(targets) and targets[j + 1][0] == targets[j][0] + 1:
                j += 1
            r0, r1 = targets[i][0], targets[j][0]
            rng = f"{gspread.utils.rowcol_to_a1(r0, c)}:{gspread.utils.rowcol_to_a1(r1, c)}"
            updates.append({"range": rng,
                            "values": [[to_native(v)] for _, v in targets[i:j + 1]]})
            filled_cells += j - i + 1
            i = j + 1

    if healed_dates:
        log(f"  既存列の空セルを補完: {', '.join(to_japanese_date(d) for d in healed_dates)}")

    # ---- シートに無い新規銘柄 ----
    new_row_codes = [c for c in codes
                     if c not in code_to_row and any(c in field_values[d] for d in dates)]

    if not updates and not new_row_codes:
        log("  追加・補完が必要なセルはありません")
        return

    ws = get_worksheet_by_gid(gc, gid, min_cols=max(total_cols, hdr_len))

    for i in range(0, len(updates), BATCH_UPDATE_CHUNK_SIZE):
        sheets_call(ws.batch_update, updates[i: i + BATCH_UPDATE_CHUNK_SIZE],
                    value_input_option="USER_ENTERED")
    log(f"  書き込みセル数: {filled_cells}")

    if new_row_codes:
        start_row = len(existing) + 1
        grid = []
        for code in new_row_codes:
            row = [code] + [""] * (total_cols - 1)
            for d in dates:
                v = field_values[d].get(code)
                if v is not None:
                    row[col_of[d] - 1] = to_native(v)
            grid.append(row)
        ensure_rows(ws, start_row + len(grid) - 1)
        end_a1 = gspread.utils.rowcol_to_a1(start_row + len(grid) - 1, total_cols)
        sheets_call(ws.update, values=grid, range_name=f"A{start_row}:{end_a1}",
                    value_input_option="USER_ENTERED")
        log(f"  新規銘柄を追加: {len(new_row_codes)} 件")

    log("  書き込み完了！")


def remove_empty_holiday_columns(gc: gspread.Client, gid: int, holiday_isos: set[str]) -> None:
    """休場日の日付列のうち、中身が完全に空のものだけを削除する。"""
    ws = get_worksheet_by_gid(gc, gid)
    data = sheets_call(ws.get_all_values)
    if not data:
        return
    header = data[0]
    targets = []
    for ci in range(FIXED_COLS, len(header)):
        iso = parse_date_text(header[ci])
        if iso and iso in holiday_isos:
            if all(ci >= len(r) or not str(r[ci]).strip() for r in data[1:]):
                targets.append(ci + 1)
    if not targets:
        return
    deleter = getattr(ws, "delete_columns", None) or ws.delete_cols
    for idx in sorted(targets, reverse=True):
        sheets_call(deleter, idx)
    log(f"  休場日の空列を削除: {len(targets)} 列")


def sort_date_columns(gc: gspread.Client, gid: int) -> None:
    ws = get_worksheet_by_gid(gc, gid)
    log(f"\n{ws.title} の日付列をソート中...")
    all_data = sheets_call(ws.get_all_values)

    if not all_data or len(all_data) < 2:
        log("  ソート対象の日付列がありません。スキップします")
        return
    width = max(last_nonempty_len(all_data[0]), FIXED_COLS)
    if width <= FIXED_COLS:
        log("  ソート対象の日付列がありません。スキップします")
        return

    header = all_data[0][:width]
    header += [""] * (width - len(header))
    fixed_header = header[:FIXED_COLS]
    date_header = header[FIXED_COLS:]
    n = len(date_header)

    order = sorted(range(n), key=lambda i: parse_date_text(date_header[i]) or "", reverse=True)
    if order == list(range(n)):
        log("  すでに新しい順に並んでいます（書き直しなし）")
        return

    sorted_header = [date_header[i] for i in order]
    new_data = [fixed_header + sorted_header]
    for row in all_data[1:]:
        fixed_vals = (list(row[:FIXED_COLS]) + [""] * FIXED_COLS)[:FIXED_COLS]
        date_vals = list(row[FIXED_COLS:width])
        date_vals += [""] * (n - len(date_vals))
        new_data.append(fixed_vals + [date_vals[i] for i in order])

    ws = get_worksheet_by_gid(gc, gid, min_cols=width)
    end_a1 = gspread.utils.rowcol_to_a1(len(new_data), width)
    sheets_call(ws.update, values=new_data, range_name=f"A1:{end_a1}",
                value_input_option="USER_ENTERED")

    preview = " → ".join(sorted_header[:SORT_PREVIEW_COUNT])
    suffix = "..." if len(sorted_header) > SORT_PREVIEW_COUNT else ""
    log(f"  ソート完了: {len(sorted_header)} 列（{preview}{suffix}）")


# =============================================================================
# メイン処理
# =============================================================================

def main():
    now = datetime.now(JST)
    log("=" * 60)
    log(f"実行時刻（日本時間）: {now:%Y-%m-%d %H:%M}")
    log(f"yfinance {yf.__version__} / pandas {pd.__version__}")
    log(f"CHUNK_SIZE={CHUNK_SIZE} SLEEP_TIME={SLEEP_TIME} MAX_CODES={MAX_CODES or '無制限'}")
    mode = f"明示指定 TARGET_DATE={TARGET_DATE_ENV}" if TARGET_DATE_ENV else f"自動（直近{BACKFILL_DAYS}日をバックフィル）"
    log(f"対象日モード: {mode}")
    log("=" * 60)

    candidates, explicit, last_day = build_candidate_dates(now)
    if not candidates:
        die("対象日がありません（指定日がすべて未確定です）。大引け後に実行するか、"
            "FORCE_INTRADAY=1 を指定してください。")
    log(f"確定済みの最新日: {last_day} / 候補日: {candidates[0]} 〜 {candidates[-1]}")

    gc = authenticate_google_sheets()

    codes = load_stock_codes_from_sheet(gc, CODES_SHEET_GID)
    if not codes:
        die("銘柄コードが0件でした。タブの内容を確認してください。")

    calendar = get_trading_calendar(candidates)
    trade_dates = [d for d in candidates if date.fromisoformat(d) in calendar]
    holidays = [d for d in candidates if d not in trade_dates]
    log(f"  対象営業日: {', '.join(trade_dates) or 'なし'}")
    if holidays:
        log(f"  休場日・土日（スキップ）: {', '.join(holidays)}")
    if not trade_dates:
        die("対象期間に営業日がありません。")

    values, ok_dates = fetch_prices(codes, trade_dates)

    for field in FIELDS:
        cfg = FIELD_CONFIG[field]
        log("\n" + "=" * 50)
        log(cfg["heading"])
        log("=" * 50)
        update_sheet(gc, cfg["gid"], codes, values[field], ok_dates)
        if CLEAN_HOLIDAY_COLUMNS and holidays:
            remove_empty_holiday_columns(gc, cfg["gid"], set(holidays))
        sort_date_columns(gc, cfg["gid"])

    log("\n" + "=" * 50)
    log("全処理完了！")
    log(f"  スプレッドシートID : {SPREADSHEET_ID}")
    log(f"  銘柄数             : {len(codes)}")
    log(f"  処理した営業日     : {', '.join(ok_dates)}")
    log(f"  取得項目           : {', '.join(FIELD_CONFIG[f]['label_jp'] for f in FIELDS)}")
    log("=" * 50)


if __name__ == "__main__":
    main()
