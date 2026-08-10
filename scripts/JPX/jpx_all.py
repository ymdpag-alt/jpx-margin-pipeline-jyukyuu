#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JPX 一括スクレイパー  v1.0
─────────────────────────────────────────────────────────────
  1. ToSTNeT 超大口約定情報       → GID 1044963009
  2. 自己株式立会外買付取引情報   → GID 2116229549
  3. 業種別 33 業種指数           → GID 742855575
─────────────────────────────────────────────────────────────

【シート書き込みルール】
  1 & 2: 新データを row=2 に挿入し既存データを下シフト
  3    : 新データを列 C に挿入し既存データを右シフト（4 指標×33 業種）

【重複防止】
  1 & 2: row=2 の「公表日」と新データの公表日を比較
  3    : A35 セルに今日の JST 日付を記録して比較
"""

import json
import logging
import os
import re
import time
from datetime import datetime, timezone, timedelta

import requests
from bs4 import BeautifulSoup
import gspread

# ─────────────────────────────────────────────────────────────
#  定数
# ─────────────────────────────────────────────────────────────
SPREADSHEET_ID = "1kWST0CkkIvo3irPSbMgtVtUqqRDXFwvRREYZDRQAFMY"

GID_TOSTNET    = 1044963009
GID_OWN_SHARES = 2116229549
GID_SECTOR     = 742855575

TOSTNET_URL    = "https://www.jpx.co.jp/markets/equities/tostnet/index.html"
OWN_SHARES_URL = "https://www.jpx.co.jp/markets/equities/off-auction-ownshares/index.html"
SECTOR_URL     = "https://www.jpx.co.jp/markets/indices/realvalues/"

# 業種別指数シート: 各指標の開始行 (1-indexed, 33 業種分)
SECTOR_ROWS = {
    "current": 1,    # 現在値
    "low":     40,   # 安値
    "high":    80,   # 高値
    "open":    120,  # 始値
}
SECTOR_LABELS = {
    "current": "現在値",
    "low":     "安値",
    "high":    "高値",
    "open":    "始値",
}
# row=35 は現在値ブロック(1-33)と安値ブロック(40-)の間の空白行→マーカー用
SECTOR_MARKER = "A35"

JST = timezone(timedelta(hours=9))

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
}

# ─────────────────────────────────────────────────────────────
#  ロギング
# ─────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────
#  Google Sheets ユーティリティ
# ─────────────────────────────────────────────────────────────
def build_gc() -> gspread.Client:
    """GOOGLE_CREDENTIALS_JSON 環境変数からクライアントを生成"""
    info = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
    return gspread.service_account_from_dict(info)


def open_ss(gc: gspread.Client) -> gspread.Spreadsheet:
    return gc.open_by_key(SPREADSHEET_ID)


def get_ws(ss: gspread.Spreadsheet, gid: int) -> gspread.Worksheet:
    for ws in ss.worksheets():
        if ws.id == gid:
            return ws
    raise ValueError(f"GID {gid} not found in spreadsheet")


def prepend_rows(ws: gspread.Worksheet, rows: list) -> None:
    """
    row=2 に新データを挿入し、既存の row=2 以降を下シフトする。
    ヘッダ行 (row=1) は維持される。
    """
    ws.insert_rows(rows, row=2, value_input_option="USER_ENTERED")
    log.info("    ✓ %d 行を row=2 に挿入", len(rows))


def insert_col_c(ss: gspread.Spreadsheet, ws: gspread.Worksheet) -> None:
    """
    列 C (0-indexed startIndex=2) に空列を挿入し、
    既存の C 列以降を右シフトする。
    A 列・B 列は不変。
    """
    ss.batch_update({
        "requests": [{
            "insertDimension": {
                "range": {
                    "sheetId":   ws.id,
                    "dimension": "COLUMNS",
                    "startIndex": 2,   # A=0, B=1, C=2
                    "endIndex":   3,
                },
                "inheritFromBefore": False,
            }
        }]
    })
    log.info("    ✓ 列 C に新列挿入")


# ─────────────────────────────────────────────────────────────
#  HTTP
# ─────────────────────────────────────────────────────────────
def fetch(url: str) -> BeautifulSoup:
    resp = requests.get(url, headers=HTTP_HEADERS, timeout=30)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return BeautifulSoup(resp.text, "html.parser")


# ═════════════════════════════════════════════════════════════
#  1. ToSTNeT 超大口約定情報
#  列: 公表日 | 取引日 | コード | 銘柄 | 価格 | 売買高 | 売買代金
# ═════════════════════════════════════════════════════════════
def scrape_tostnet() -> list:
    """
    コード列 (index=2) が 4 桁数字の行を抽出する。
    戻り値: [[公表日, 取引日, コード, 銘柄, 価格, 売買高, 売買代金], ...]
    """
    soup = fetch(TOSTNET_URL)
    out = []
    for tbl in soup.find_all("table"):
        for tr in tbl.find_all("tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all("td")]
            if len(cells) >= 7 and re.fullmatch(r"\d{4}", cells[2]):
                out.append(cells[:7])

    # フォールバック: コードが別列にある場合 (index=3) も試す
    if not out:
        log.info("  index=2 で取得なし → index=3 でリトライ")
        for tbl in soup.find_all("table"):
            for tr in tbl.find_all("tr"):
                cells = [td.get_text(strip=True) for td in tr.find_all("td")]
                if len(cells) >= 8 and re.fullmatch(r"\d{4}", cells[3]):
                    # [市場, 公表日, 取引日, コード, 銘柄, 価格, 売買高, 売買代金]
                    out.append([cells[1], cells[2], cells[3], cells[4],
                                cells[5], cells[6], cells[7]])
    return out


def run_tostnet(ss: gspread.Spreadsheet) -> None:
    log.info("  スクレイプ: ToSTNeT 超大口...")
    rows = scrape_tostnet()
    log.info("  → %d 件取得", len(rows))
    if not rows:
        log.warning("  データなし → スキップ")
        return

    ws = get_ws(ss, GID_TOSTNET)
    cur = ws.row_values(2)  # 現在の row=2 を確認
    if cur and cur[0] == rows[0][0]:
        log.info("  既挿入済み (%s) → スキップ", rows[0][0])
        return

    prepend_rows(ws, rows)


# ═════════════════════════════════════════════════════════════
#  2. 自己株式立会外買付取引情報
#  列: 公表日 | 実施日 | コード | 銘柄 | 価格 | 買付数量 | 約定数量
# ═════════════════════════════════════════════════════════════
def scrape_own_shares() -> list:
    """
    コード列 (index=2) が 4 桁数字の行を抽出する。
    戻り値: [[公表日, 実施日, コード, 銘柄, 価格, 買付数量, 約定数量], ...]
    """
    soup = fetch(OWN_SHARES_URL)
    out = []
    for tbl in soup.find_all("table"):
        for tr in tbl.find_all("tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all("td")]
            if len(cells) >= 7 and re.fullmatch(r"\d{4}", cells[2]):
                out.append(cells[:7])

    if not out:
        log.info("  index=2 で取得なし → index=3 でリトライ")
        for tbl in soup.find_all("table"):
            for tr in tbl.find_all("tr"):
                cells = [td.get_text(strip=True) for td in tr.find_all("td")]
                if len(cells) >= 8 and re.fullmatch(r"\d{4}", cells[3]):
                    out.append([cells[1], cells[2], cells[3], cells[4],
                                cells[5], cells[6], cells[7]])
    return out


def run_own_shares(ss: gspread.Spreadsheet) -> None:
    log.info("  スクレイプ: 自己株式立会外買付...")
    rows = scrape_own_shares()
    log.info("  → %d 件取得", len(rows))
    if not rows:
        log.warning("  データなし → スキップ")
        return

    ws = get_ws(ss, GID_OWN_SHARES)
    cur = ws.row_values(2)
    if cur and cur[0] == rows[0][0]:
        log.info("  既挿入済み (%s) → スキップ", rows[0][0])
        return

    prepend_rows(ws, rows)


# ═════════════════════════════════════════════════════════════
#  3. 業種別 33 業種指数
#  デフォルト列順: 指数名(0) 現在値(1) 前日比(2) 騰落率%(3) 高値(4) 安値(5) 始値(6) 前日終値(7)
# ═════════════════════════════════════════════════════════════
_DEFAULT_COLS = {"name": 0, "current": 1, "high": 4, "low": 5, "open": 6}


def detect_sector_cols(soup: BeautifulSoup) -> dict:
    """
    テーブルヘッダを動的に解析して列インデックスを返す。
    ヘッダが見つからない場合はデフォルト値を返す。
    """
    for tbl in soup.find_all("table"):
        for tr in tbl.find_all("tr"):
            hdrs = [c.get_text(strip=True) for c in tr.find_all(["th", "td"])]
            if not any("現在値" in h for h in hdrs):
                continue
            col = {}
            for i, h in enumerate(hdrs):
                if ("指数" in h or "銘柄" in h) and "name" not in col:
                    col["name"] = i
                elif "現在値" in h and "current" not in col:
                    col["current"] = i
                elif "高値" in h and "high" not in col:
                    col["high"] = i
                elif "安値" in h and "low" not in col:
                    col["low"] = i
                elif "始値" in h and "open" not in col:
                    col["open"] = i
            if len(col) >= 3:
                log.info("  列マップ検出: %s", col)
                return {**_DEFAULT_COLS, **col}

    log.info("  列マップ: デフォルト使用 %s", _DEFAULT_COLS)
    return _DEFAULT_COLS.copy()


def scrape_sector() -> list:
    """
    33 業種の指数データを取得する。
    戻り値: [{"name":str, "current":str, "high":str, "low":str, "open":str}, ...]

    NOTE: jpx.co.jp/markets/indices/realvalues/ が JS レンダリングの場合は
          データが取得できないことがある。その際はログに "0 業種" と出るので
          ページのソースを確認して SECTOR_URL or スクレイプ方法を調整すること。
    """
    soup = fetch(SECTOR_URL)
    col = detect_sector_cols(soup)
    max_idx = max(col.values())
    out, seen = [], set()

    for tbl in soup.find_all("table"):
        for tr in tbl.find_all("tr"):
            cells = [c.get_text(strip=True) for c in tr.find_all(["th", "td"])]
            if len(cells) <= max_idx:
                continue
            name = cells[col["name"]]
            # 日本語文字(ひらがな/カタカナ/漢字)を含む行のみ対象
            if not name or not re.search(r"[\u3040-\u9fff]", name):
                continue
            # 現在値が数値らしいことを確認
            if not re.search(r"[\d,.]", cells[col["current"]]):
                continue
            if name in seen:
                continue
            seen.add(name)
            out.append({
                "name":    name,
                "current": cells[col["current"]],
                "high":    cells[col["high"]],
                "low":     cells[col["low"]],
                "open":    cells[col["open"]],
            })

    return out


def write_sector_block(
    ws: gspread.Worksheet,
    data: list,
    metric: str,
    start_row: int,
) -> None:
    """
    A 列: 指数名 (毎回上書き、冪等)
    C 列: 今回の指標値 (insert_col_c で挿入済みの空列 C に書き込む)
    """
    n = len(data)
    if n == 0:
        return

    ws.update(
        f"A{start_row}:A{start_row + n - 1}",
        [[d["name"]] for d in data],
        value_input_option="USER_ENTERED",
    )
    ws.update(
        f"C{start_row}:C{start_row + n - 1}",
        [[d[metric]] for d in data],
        value_input_option="USER_ENTERED",
    )
    log.info(
        "    ✓ %s: row%d〜%d (%d 件)",
        SECTOR_LABELS[metric], start_row, start_row + n - 1, n,
    )
    time.sleep(1.2)   # Sheets API レート制限対策


def run_sector(ss: gspread.Spreadsheet) -> None:
    log.info("  スクレイプ: 業種別 33 業種指数...")
    data = scrape_sector()
    log.info("  → %d 業種取得", len(data))
    if not data:
        log.warning("  データなし → スキップ")
        return

    ws = get_ws(ss, GID_SECTOR)
    today = datetime.now(JST).strftime("%Y/%m/%d")

    # 重複チェック: SECTOR_MARKER セル (A35) に今日の日付があれば既実行済み
    marker = (ws.acell(SECTOR_MARKER).value or "").strip()
    if marker == today:
        log.info("  既挿入済み (%s) → スキップ", today)
        return

    # 列 C に空列を挿入 → 既存データが右シフト
    insert_col_c(ss, ws)
    time.sleep(1.5)

    # 4 指標を各ブロックへ書き込み
    for metric, start_row in SECTOR_ROWS.items():
        write_sector_block(ws, data, metric, start_row)

    # 実行済みマーカーを更新
    ws.update(SECTOR_MARKER, [[today]], value_input_option="USER_ENTERED")
    log.info("  ✓ マーカー更新: %s ← %s", SECTOR_MARKER, today)


# ═════════════════════════════════════════════════════════════
#  メイン
# ═════════════════════════════════════════════════════════════
def main() -> None:
    now = datetime.now(JST).strftime("%Y-%m-%d %H:%M JST")
    log.info("━━━ JPX 一括スクレイプ 開始: %s ━━━", now)

    gc = build_gc()
    ss = open_ss(gc)

    tasks = [
        ("ToSTNeT 超大口",     lambda: run_tostnet(ss)),
        ("自己株式立会外買付", lambda: run_own_shares(ss)),
        ("業種別 33 業種指数", lambda: run_sector(ss)),
    ]

    for label, fn in tasks:
        log.info("▶ %s", label)
        try:
            fn()
        except Exception as e:
            log.error("  ✗ エラー: %s", e, exc_info=True)
        time.sleep(2)

    log.info("━━━ 完了 ━━━")


if __name__ == "__main__":
    main()
