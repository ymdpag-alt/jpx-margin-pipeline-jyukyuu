#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JPX 一括スクレイパー  v1.4
  1. ToSTNeT 超大口約定情報       → GID 1044963009
  2. 自己株式立会外買付取引情報   → GID 2116229549
  3. 業種別 33 業種指数           → GID 742855575  ※Playwright 使用

TARGET_DATE 環境変数 (YYYYMMDD):
  - 指定あり → その日付のデータを取得
  - 指定なし → ページ上の最新日を自動取得
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

SPREADSHEET_ID = "1kWST0CkkIvo3irPSbMgtVtUqqRDXFwvRREYZDRQAFMY"
GID_TOSTNET    = 1044963009
GID_OWN_SHARES = 2116229549
GID_SECTOR     = 742855575

TOSTNET_URL    = "https://www.jpx.co.jp/markets/equities/tostnet/index.html"
OWN_SHARES_URL = "https://www.jpx.co.jp/markets/equities/off-auction-ownshares/index.html"
SECTOR_URL     = "https://www.jpx.co.jp/markets/indices/realvalues/"
JPX_BASE       = "https://www.jpx.co.jp"

SECTOR_ROWS   = {"current": 1, "low": 40, "high": 80, "open": 120}
SECTOR_LABELS = {"current": "現在値", "low": "安値", "high": "高値", "open": "始値"}
SECTOR_MARKER = "A35"
SECTOR_COL    = {"name": 0, "current": 1, "open": 3, "high": 4, "low": 5}

# 東証33業種（TOPIX-17ではなく、東証業種別株価指数33業種）
SECTOR33 = [
    "水産・農林業", "鉱業", "建設業", "食料品", "繊維製品", "パルプ・紙",
    "化学", "医薬品", "石油・石炭製品", "ゴム製品", "ガラス・土石製品",
    "鉄鋼", "非鉄金属", "金属製品", "機械", "電気機器", "輸送用機器",
    "精密機器", "その他製品", "電気・ガス業", "陸運業", "海運業", "空運業",
    "倉庫・運輸関連業", "情報・通信業", "卸売業", "小売業", "銀行業",
    "証券・商品先物取引業", "保険業", "その他金融業", "不動産業", "サービス業",
]

def _norm(s: str) -> str:
    """比較用に正規化: 空白・中黒・全角/半角の揺れを吸収"""
    s = s.replace("\u3000", "").replace(" ", "").replace("　", "")
    s = s.replace("・", "").replace("･", "")
    s = s.replace("（", "(").replace("）", ")")
    return s.strip()

SECTOR33_NORM = {_norm(x): x for x in SECTOR33}

SKIP_NAMES = frozenset([
    "指数名", "銘柄", "業種", "指数", "現在値", "前日比", "始値", "高値", "安値",
    "前日終値", "騰落率", "リアルタイムグラフ", "ヒストリカルグラフ",
    "データ更新日時", "", "-", "－", "−",
])

JST      = timezone(timedelta(hours=9))
DATE_PAT = re.compile(r"^\d{4}[/\-年]\d{1,2}[/\-月]\d{1,2}")

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

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────
#  対象日の解決
# ─────────────────────────────────────────────────────────────
def resolve_target_date():
    """
    TARGET_DATE 環境変数 (YYYYMMDD) を読み YYYY/MM/DD 形式で返す。
    未指定・空欄の場合は None（呼び出し側で「最新日」として扱う）。
    """
    raw = os.environ.get("TARGET_DATE", "").strip()
    if raw and re.fullmatch(r"\d{8}", raw):
        d = f"{raw[:4]}/{raw[4:6]}/{raw[6:8]}"
        log.info("対象日: %s（指定）", d)
        return d
    log.info("対象日: 未指定 → 各ページの最新日を自動取得")
    return None


# ─────────────────────────────────────────────────────────────
#  Google Sheets
# ─────────────────────────────────────────────────────────────
def build_gc():
    info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    return gspread.service_account_from_dict(info)

def open_ss(gc):
    return gc.open_by_key(SPREADSHEET_ID)

def get_ws(ss, gid):
    for ws in ss.worksheets():
        if ws.id == gid:
            return ws
    raise ValueError(f"GID {gid} not found")

def prepend_rows(ws, rows):
    ws.insert_rows(rows, row=2, value_input_option="USER_ENTERED")
    log.info("    ✓ %d 行を row=2 に挿入", len(rows))

def insert_col_c(ss, ws):
    ss.batch_update({"requests": [{"insertDimension": {
        "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                  "startIndex": 2, "endIndex": 3},
        "inheritFromBefore": False,
    }}]})
    log.info("    ✓ 列 C に新列挿入")


# ─────────────────────────────────────────────────────────────
#  HTTP
# ─────────────────────────────────────────────────────────────
def fetch(url):
    resp = requests.get(url, headers=HTTP_HEADERS, timeout=30)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return BeautifulSoup(resp.text, "html.parser")


# ═════════════════════════════════════════════════════════════
#  1. ToSTNeT 超大口約定情報
#  ページ構造: 公表日(0) | 取引内容リンク(1)
#  出力列:    公表日 | 取引日 | コード | 銘柄 | 価格 | 売買高 | 売買代金
# ═════════════════════════════════════════════════════════════
def _parse_tostnet_xlsx(url, kouhyo_date):
    """
    ToSTNeT 超大口約定情報の xlsx をダウンロードしてパース。
    出力: [[公表日, 取引日, コード, 銘柄, 価格, 売買高, 売買代金], ...]
    """
    import io
    import openpyxl

    resp = requests.get(url, headers=HTTP_HEADERS, timeout=60)
    resp.raise_for_status()

    wb = openpyxl.load_workbook(io.BytesIO(resp.content), data_only=True)
    out = []

    for ws in wb.worksheets:
        log.info("  [xlsx] シート '%s' (%d行 x %d列)",
                 ws.title, ws.max_row, ws.max_column)

        for row in ws.iter_rows(values_only=True):
            cells = ["" if c is None else str(c).strip() for c in row]
            if not any(cells):
                continue

            # 4桁コードを含む列を探す
            code_idx = None
            for i, c in enumerate(cells):
                # 数値セルは '1234' or '1234.0' の形になり得る
                c2 = c.split(".")[0] if "." in c else c
                if re.fullmatch(r"\d{4}", c2):
                    code_idx = i
                    break

            if code_idx is None:
                continue

            code = cells[code_idx].split(".")[0]
            # コード列の前=取引日、後ろ=銘柄/価格/売買高/売買代金 と想定
            torihiki = cells[code_idx - 1] if code_idx >= 1 else ""
            rest = cells[code_idx + 1: code_idx + 5]
            rest += [""] * (4 - len(rest))

            out.append([kouhyo_date, torihiki, code,
                        rest[0], rest[1], rest[2], rest[3]])

    if not out:
        # デバッグ: xlsx の先頭数行を出力
        ws = wb.worksheets[0]
        log.info("  [DEBUG xlsx] 先頭6行:")
        for j, row in enumerate(ws.iter_rows(max_row=6, values_only=True)):
            log.info("    row[%d]: %s", j,
                     [str(c)[:20] if c is not None else "" for c in row])

    return out


def _parse_tostnet_html(soup, kouhyo_date):
    """リンク先が HTML の場合のフォールバックパーサ"""
    out = []
    for tbl in soup.find_all("table"):
        for tr in tbl.find_all("tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all("td")]
            if not cells:
                continue
            for idx, c in enumerate(cells):
                if re.fullmatch(r"\d{4}", c) and idx >= 1:
                    row = [kouhyo_date] + cells[max(0, idx-1): idx+6]
                    if len(row) >= 7:
                        out.append(row[:7])
                    break
    return out


def scrape_tostnet(target_date=None):
    """
    target_date (YYYY/MM/DD): 指定日のリンクを追う
    None: テーブルの最新行（先頭）のリンクを追う
    """
    soup = fetch(TOSTNET_URL)
    out = []

    for tbl in soup.find_all("table"):
        for tr in tbl.find_all("tr"):
            tds = tr.find_all("td")
            if not tds:
                continue
            kouhyo = tds[0].get_text(strip=True)
            if not DATE_PAT.match(kouhyo):
                continue

            # 日付フィルタ
            if target_date and kouhyo != target_date:
                continue

            link_tag = tds[1].find("a", href=True) if len(tds) > 1 else None
            if not link_tag:
                log.info("  ToSTNeT: %s にリンクなし → セルHTML: %s",
                         kouhyo, str(tds[1])[:200] if len(tds) > 1 else "N/A")
                if not target_date:
                    break   # 最新日のみ試みて終了
                continue

            href = link_tag["href"]
            if not href.startswith("http"):
                href = JPX_BASE + href

            log.info("  ToSTNeT: %s → %s", kouhyo, href)
            try:
                if href.lower().endswith((".xlsx", ".xls")):
                    rows = _parse_tostnet_xlsx(href, kouhyo)
                    log.info("  xlsx: %d 件取得", len(rows))
                else:
                    detail_soup = fetch(href)
                    rows = _parse_tostnet_html(detail_soup, kouhyo)
                    log.info("  HTML: %d 件取得", len(rows))
                out.extend(rows)
            except Exception as e:
                log.warning("  取得失敗: %s", e, exc_info=True)

            if not target_date:
                break   # 最新日のみ

    return out


def run_tostnet(ss, target_date):
    log.info("  スクレイプ: ToSTNeT 超大口...")
    rows = scrape_tostnet(target_date)
    log.info("  → %d 件", len(rows))
    if not rows:
        log.warning("  データなし → スキップ")
        return
    ws = get_ws(ss, GID_TOSTNET)
    # 重複チェック（公表日で比較）
    cur = ws.row_values(2)
    if cur and rows and cur[0] == rows[0][0] and not target_date:
        log.info("  既挿入済み (%s) → スキップ", rows[0][0])
        return
    prepend_rows(ws, rows)


# ═════════════════════════════════════════════════════════════
#  2. 自己株式立会外買付取引情報
#  ページ構造（実測）5列:
#    実施日 | 銘柄名（コード） | 値段 | 買付数量 | 約定数量
#  出力列: 公表日(今日) | 実施日 | コード | 銘柄 | 価格 | 買付数量 | 約定数量
# ═════════════════════════════════════════════════════════════
def _parse_name_code(text):
    """'銘柄名（コード）　株式' → ('銘柄名', 'コード')"""
    m = re.search(r"[（(](\d{4})[）)]", text)
    code = m.group(1) if m else ""
    name = text[:m.start()].strip() if m else text
    name = re.sub(r"[\u3000\s]+(株式|株|全株)\s*$", "", name).strip()
    return name, code


def scrape_own_shares(target_date=None):
    """
    target_date (YYYY/MM/DD): 実施日でフィルタ
    None: 全行取得
    """
    soup = fetch(OWN_SHARES_URL)
    today = datetime.now(JST).strftime("%Y/%m/%d")
    out = []

    for tbl in soup.find_all("table"):
        for tr in tbl.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) < 4:
                continue
            jisshi = tds[0].get_text(strip=True)
            if not DATE_PAT.match(jisshi):
                continue

            # 日付フィルタ
            if target_date and jisshi != target_date:
                continue

            name_code = tds[1].get_text(strip=True)
            price     = tds[2].get_text(strip=True)
            buy_qty   = tds[3].get_text(strip=True)
            exec_qty  = tds[4].get_text(strip=True) if len(tds) > 4 else ""

            name, code = _parse_name_code(name_code)
            # 公表日はページに存在しないため当日 JST を使用
            out.append([today, jisshi, code, name, price, buy_qty, exec_qty])

    return out


def run_own_shares(ss, target_date):
    log.info("  スクレイプ: 自己株式立会外買付...")
    rows = scrape_own_shares(target_date)
    log.info("  → %d 件", len(rows))
    if not rows:
        log.warning("  データなし → スキップ")
        return
    ws = get_ws(ss, GID_OWN_SHARES)
    # 重複チェック（実施日=B列で比較）
    cur = ws.row_values(2)
    if cur and len(cur) > 1 and rows and cur[1] == rows[0][1] and not target_date:
        log.info("  既挿入済み (実施日: %s) → スキップ", rows[0][1])
        return
    prepend_rows(ws, rows)


# ═════════════════════════════════════════════════════════════
#  3. 業種別 33 業種指数  ※JS レンダリング → Playwright 使用
#  ヘッダー（実測）: 指数名(0) 現在値(1) 前日比(2) 始値(3) 高値(4) 安値(5)
# ═════════════════════════════════════════════════════════════
def _parse_sector_html(soup):
    max_idx = max(SECTOR_COL.values())
    out, seen = [], set()
    for tbl in soup.find_all("table"):
        for tr in tbl.find_all("tr"):
            cells = [c.get_text(strip=True) for c in tr.find_all(["th", "td"])]
            if len(cells) <= max_idx:
                continue
            name = cells[SECTOR_COL["name"]]
            if not name or name in SKIP_NAMES:
                continue
            cur = cells[SECTOR_COL["current"]]
            if not cur or cur in SKIP_NAMES:
                continue
            if not re.search(r"[\d０-９,，.]", cur):
                continue
            # ── 33業種ホワイトリストで絞り込み ──
            key = _norm(name)
            if key not in SECTOR33_NORM:
                continue
            canonical = SECTOR33_NORM[key]   # 正式名称に統一

            if canonical in seen:
                continue
            seen.add(canonical)
            out.append({
                "name":    canonical,
                "current": cur,
                "high":    cells[SECTOR_COL["high"]],
                "low":     cells[SECTOR_COL["low"]],
                "open":    cells[SECTOR_COL["open"]],
            })

    # SECTOR33 の定義順にソート（毎回同じ行順を保証）
    order = {n: i for i, n in enumerate(SECTOR33)}
    out.sort(key=lambda d: order.get(d["name"], 999))
    return out


def _fetch_with_playwright():
    from playwright.sync_api import sync_playwright
    log.info("  Playwright 起動中...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        page = browser.new_page()
        page.goto(SECTOR_URL, wait_until="networkidle", timeout=30_000)
        try:
            page.wait_for_function(
                "document.querySelectorAll('table tr').length > 14",
                timeout=10_000,
            )
        except Exception:
            log.info("  待機タイムアウト（現状の HTML を使用）")
        html = page.content()
        browser.close()
    log.info("  Playwright: HTML 取得完了")
    return html


def scrape_sector():
    # まず requests で試す（週次バッチ等で静的 HTML になる場合に対応）
    soup = fetch(SECTOR_URL)
    data = _parse_sector_html(soup)
    if not data:
        log.info("  requests: 0件 → Playwright で再試行...")
        try:
            html = _fetch_with_playwright()
            soup2 = BeautifulSoup(html, "html.parser")
            data = _parse_sector_html(soup2)
            log.info("  Playwright後: %d 業種", len(data))
            if not data:
                tables = soup2.find_all("table")
                log.info("  [DEBUG Playwright後] テーブル数=%d", len(tables))
                for i, tbl in enumerate(tables[:4]):
                    trs = tbl.find_all("tr")
                    log.info("    table[%d]: %d行", i, len(trs))
                    for j, tr in enumerate(trs[:3]):
                        cells = [c.get_text(strip=True) for c in tr.find_all(["th","td"])]
                        log.info("      row[%d] %d列: %s", j, len(cells),
                                 [c[:20] for c in cells])
        except ImportError:
            log.error("  Playwright 未インストール")
        except Exception as e:
            log.error("  Playwright エラー: %s", e, exc_info=True)
    return data


def write_sector_block(ws, data, metric, start_row):
    n = len(data)
    if n == 0:
        return
    ws.update(values=[[d["name"]] for d in data],
              range_name=f"A{start_row}:A{start_row+n-1}",
              value_input_option="USER_ENTERED")
    ws.update(values=[[d[metric]] for d in data],
              range_name=f"C{start_row}:C{start_row+n-1}",
              value_input_option="USER_ENTERED")
    log.info("    ✓ %s: row%d〜%d (%d件)",
             SECTOR_LABELS[metric], start_row, start_row+n-1, n)
    time.sleep(1.2)


def run_sector(ss):
    log.info("  スクレイプ: 業種別 33 業種指数...")
    data = scrape_sector()
    log.info("  → %d 業種取得", len(data))
    if not data:
        log.warning("  データなし → スキップ")
        return
    ws = get_ws(ss, GID_SECTOR)
    today = datetime.now(JST).strftime("%Y/%m/%d")
    marker = (ws.acell(SECTOR_MARKER).value or "").strip()
    if marker == today:
        log.info("  既挿入済み (%s) → スキップ", today)
        return
    insert_col_c(ss, ws)
    time.sleep(1.5)
    for metric, start_row in SECTOR_ROWS.items():
        write_sector_block(ws, data, metric, start_row)
    ws.update(values=[[today]], range_name=SECTOR_MARKER,
              value_input_option="USER_ENTERED")
    log.info("  ✓ マーカー更新: %s ← %s", SECTOR_MARKER, today)


# ═════════════════════════════════════════════════════════════
#  メイン
# ═════════════════════════════════════════════════════════════
def main():
    now = datetime.now(JST).strftime("%Y-%m-%d %H:%M JST")
    log.info("━━━ JPX 一括スクレイプ 開始: %s ━━━", now)

    target_date = resolve_target_date()   # None = 最新日自動

    gc = build_gc()
    ss = open_ss(gc)

    tasks = [
        ("ToSTNeT 超大口",     lambda: run_tostnet(ss, target_date)),
        ("自己株式立会外買付", lambda: run_own_shares(ss, target_date)),
        ("業種別 33 業種指数", lambda: run_sector(ss)),   # 日付指定なし（リアルタイム）
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
