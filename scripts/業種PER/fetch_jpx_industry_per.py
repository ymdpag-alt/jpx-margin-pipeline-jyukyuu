# -*- coding: utf-8 -*-
"""
JPX「規模別・業種別PER・PBR（連結・単体）一覧」（月次Excel）を取得し、
スプレッドシート「００８‐ニュース」の「業種別PER」シートに追記する。

- JPXのページから perpbrYYYYMM.xlsx のリンクを探し、最新月（または FORCE_YM で指定した月）を取得
- Excelの全シートを、行ごとにそのまま転記する（JPXの表の作りが変わっても取りこぼさないため）
- 「データ年月」が既にシートにあれば何もしない（毎月1〜7日に毎日実行しても二重登録しない）
"""
import io
import json
import os
import re
import sys
from datetime import date, datetime, timedelta, timezone
from urllib.parse import urljoin

import gspread
import pandas as pd
import requests
from google.oauth2.service_account import Credentials

# ===== 定数 =====
JST = timezone(timedelta(hours=9))
PAGE_URL = "https://www.jpx.co.jp/markets/statistics-equities/misc/04.html"
LINK_PATTERN = re.compile(r'href="([^"]*?perpbr(\d{6})\.xlsx)"')
SPREADSHEET_ID = "1dZtTjwpeaE8p5QKX1lYfP3mQTKZck7q_K-F2jqVjoNA"  # ００８‐ニュース
SHEET_NAME = "業種別PER"
BASE_HEADER = ["データ年月", "取得日時", "元シート名", "元の行番号"]
MAX_VALUE_COLS = 40            # 元の表から転記する最大列数
APPEND_CHUNK = 500             # 1回の追記で送る行数
CREDENTIALS_ENV = "GOOGLE_SERVICE_ACCOUNT_JSON"
FORCE_YM_ENV = "FORCE_YM"      # 例: '202610'。空なら最新月
TIMEOUT_SEC = 60
USER_AGENT = "Mozilla/5.0 (compatible; jpx-industry-per-fetcher/1.0)"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def now_jst_str():
    return datetime.now(JST).strftime("%Y年%m月%d日 %H:%M")


def ym_label(ym):
    """'202610' → '2026年10月末'"""
    return f"{ym[:4]}年{int(ym[4:]):d}月末"


def find_link(html, force_ym=""):
    """ページ内の perpbrYYYYMM.xlsx から、指定月または最新月の (ym, 絶対URL) を返す"""
    links = {}
    for href, ym in LINK_PATTERN.findall(html):
        links[ym] = urljoin(PAGE_URL, href)
    if not links:
        raise RuntimeError("perpbrYYYYMM.xlsx のリンクが見つかりません（JPXのページ構成が変わった可能性）")
    if force_ym:
        if force_ym not in links:
            raise RuntimeError(f"{force_ym} のファイルがページにありません。掲載月: {sorted(links)}")
        return force_ym, links[force_ym]
    ym = max(links)
    return ym, links[ym]


def cell_to_value(v):
    """Excelのセル値をスプレッドシートに書ける値に変換"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        if pd.isna(v):
            return ""
        return int(v) if v.is_integer() else v
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return v.strftime("%Y年%m月%d日")
    return str(v).strip()


def excel_to_rows(content, ym, fetched_at):
    """全シートを [データ年月, 取得日時, 元シート名, 元の行番号, 値...] の行に変換（空行は除く）"""
    sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None, dtype=object)
    label = ym_label(ym)
    rows = []
    for sheet_name, df in sheets.items():
        for idx, raw in enumerate(df.itertuples(index=False, name=None), start=1):
            values = [cell_to_value(v) for v in raw][:MAX_VALUE_COLS]
            while values and values[-1] == "":
                values.pop()
            if not values:
                continue
            rows.append([label, fetched_at, str(sheet_name), idx] + values)
    if not rows:
        raise RuntimeError("Excelにデータ行がありません")
    return rows


def open_worksheet(gc):
    ss = gc.open_by_key(SPREADSHEET_ID)
    header = BASE_HEADER + [f"列{i}" for i in range(1, MAX_VALUE_COLS + 1)]
    try:
        ws = ss.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=SHEET_NAME, rows=1000, cols=len(header))
        ws.update(range_name="A1", values=[header], value_input_option="RAW")
        ws.freeze(rows=1)
        print(f"シート「{SHEET_NAME}」を作成しました")
    if ws.col_count < len(header):
        ws.resize(cols=len(header))
    return ws


def main():
    force_ym = (os.environ.get(FORCE_YM_ENV) or "").strip()
    creds_json = os.environ.get(CREDENTIALS_ENV)
    if not creds_json:
        sys.exit(f"環境変数 {CREDENTIALS_ENV} が未設定です")

    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT
    page = session.get(PAGE_URL, timeout=TIMEOUT_SEC)
    page.raise_for_status()
    page.encoding = page.apparent_encoding or "utf-8"
    ym, url = find_link(page.text, force_ym)
    label = ym_label(ym)
    print(f"対象: {label} {url}")

    creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
    ws = open_worksheet(gspread.authorize(creds))

    # 二重登録の防止（A列のデータ年月で判定）
    existing = set(ws.col_values(1)[1:])
    if label in existing:
        print(f"{label} は登録済みのため終了します")
        return

    res = session.get(url, timeout=TIMEOUT_SEC)
    res.raise_for_status()
    rows = excel_to_rows(res.content, ym, now_jst_str())

    for i in range(0, len(rows), APPEND_CHUNK):
        ws.append_rows(rows[i:i + APPEND_CHUNK], value_input_option="RAW", table_range="A1")
    print(f"{label}: {len(rows)}行を追記しました")


if __name__ == "__main__":
    main()
