#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JPX 東京証券取引所日報(株式相場表)を毎日取得し、Google スプレッドシートに書き込むスクリプト。

対象データ (銘柄ごと):
    コード, 銘柄名(和文/英文), 業種, 市場区分, 前場 始値/高値/安値/終値, 後場 始値/高値/安値/終値,
    VWAP(売買高加重平均価格), 出来高(千株)

出力先: Google スプレッドシート 10シート (項目ごとに1シート、gidで指定)
    前場始値 / 前場高値 / 前場安値 / 前場終値
    後場始値 / 後場高値 / 後場安値 / 後場終値
    VWAP / 出来高

各シートの列構成(ワイド形式): A=コード, B=銘柄名, C=業種, D=市場, E列以降=日付ごとの値
    - A〜D列は銘柄が初めて登場したときのみ書き込み、以降は変更しない
    - 新しい日付のデータはE列に挿入し、既存の日付列は右にずれていく
    - 当日データが取得できなかった銘柄は、その日の列を空欄にする
    - まだシートに存在しない銘柄(新規コード)は最終行に新しい行として追加する

実行方法:
    python fetch_jpx_daily.py                     # 最新日を自動検出して取得
    python fetch_jpx_daily.py --date 2026-08-06    # 特定日を指定して取得(バックフィル用)
    python fetch_jpx_daily.py --dry-run            # スプレッドシートに書き込まず件数だけ確認
    python fetch_jpx_daily.py --force              # 同じ日付列が既にあっても強制的に挿入し直す

必要な環境変数:
    GOOGLE_SERVICE_ACCOUNT_JSON : サービスアカウントのJSON鍵(文字列そのもの)
    SPREADSHEET_ID              : (任意) 書き込み先スプレッドシートIDを上書きしたい場合のみ指定
"""

import argparse
import io
import json
import os
import re
import sys
import time
from datetime import datetime, date
from typing import Optional

import requests


# ====================================================================
# 定数定義
# ====================================================================

# --- JPX 日報 index ページ -------------------------------------------------
JPX_INDEX_URL = "https://www.jpx.co.jp/markets/statistics-equities/daily/00-archives-01.html"

# --- 日付表示フォーマット用 -------------------------------------------------
# date.weekday(): 月=0, 火=1, 水=2, 木=3, 金=4, 土=5, 日=6
WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]

# --- テキストパース用の正規表現 ---------------------------------------------
CODE_RE_STR = r'[0-9][0-9A-Za-z]{3}'
NUM = r'[\d,]+(?:\.\d+)?'

DATA_LINE_RE = re.compile(
    r'^(?P<code>' + CODE_RE_STR + r')\s+'
    r'(?P<unit>\d+)\s+'
    r'(?P<name_jp>\S+)\s+'
    r'(?P<am_open>' + NUM + r')\s+'
    r'(?P<am_high>' + NUM + r')\s+'
    r'(?P<am_low>' + NUM + r')\s+'
    r'(?P<am_close>' + NUM + r')\s+'
    r'(?P<pm_open>' + NUM + r')\s+'
    r'(?P<pm_high>' + NUM + r')\s+'
    r'(?P<pm_low>' + NUM + r')\s+'
    r'(?P<pm_close>' + NUM + r')\s+'
    # 「最終気配」が"－"(特殊気配なし)の場合、文字幅が非常に狭いため
    # 直後の「前日比」の数値とスペース無しで連結されてしまうことがある(例: "－30.00")。
    # そのためこの境界だけ空白ゼロ許容(\s*)にする。
    r'(?P<final_quote>[－ー―]|' + NUM + r')\s*'
    r'(?P<net_change>-?' + NUM + r')\s+'
    r'(?P<vwap>' + NUM + r')\s+'
    r'(?P<volume>' + NUM + r')\s+'
    r'(?P<value>' + NUM + r')\s*$'
)

# ページ毎に繰り返される定型ヘッダー/フッター行(実際のPDF出力に合わせたもの)
HEADER_JUNK_EXACT = {
    "株 式 相 場 表", "Stock Quotations",
    "立 会 市 場 普 通 取 引", "Auction Trades Regular Way",
    "午前 (The morning trading session) 午後 (The afternoon trading session)",
    "売買 売買高加重",
    "コード 銘柄名 最終気配 前日比 売買高 売買代金",
    "単位 始値 高値 安値 終値 始値 高値 安値 終値 平均価格",
    "Trading",
    "Code Issues Open High Low Close Open High Low Close Final special quote "
    "Net Change VWAP Trading Volume Trading Value",
    "Unit",
    "円[￥] 円[￥] 円[￥] 円[￥] 円[￥] 円[￥] 円[￥] 円[￥] 円[￥] 円[￥] 円[￥] "
    "千株[thous.shs.] 千円[￥thous.]",
    "千口/千個[thous.units.]",
}
DATE_PAGE_RE = re.compile(r'^\d{4}年\d{1,2}月\d{1,2}日')
COPYRIGHT_RE = re.compile(r'^Copyright \(c\) Tokyo Stock Exchange')

# 「市場」列に転記する対象(和文見出しそのまま)
MARKET_HEADERS = {"プライム市場", "スタンダード市場", "グロース市場"}
# 内容としては情報を持つが、業種欄には転記したくない行(英語版の市場見出し等)
NON_INDUSTRY_OTHER_LINES = {
    "Prime Market", "Standard Market", "Growth Market",
    "内国株式", "Domestic Stock", "外国株式", "Foreign Stock",
    "Ｒ Ｅ Ｉ Ｔ", "REIT",
}

# --- Google スプレッドシート関連 --------------------------------------------
# 対象スプレッドシートID (URLの /d/ と /edit の間の文字列)
# 環境変数 SPREADSHEET_ID2 が設定されていればそちらを優先する(別スプレッドシートに切り替えたい場合用)
SPREADSHEET_ID_DEFAULT = "1kWST0CkkIvo3irPSbMgtVtUqqRDXFwvRREYZDRQAFMY"

# 各項目シートのgid (URLの gid= の値)
SHEET_SPECS = [
    # (表示名, gid, レコードのキー)
    ("前場始値", 519462541, "am_open"),
    ("前場高値", 1702721662, "am_high"),
    ("前場安値", 1992276623, "am_low"),
    ("前場終値", 1876491327, "am_close"),
    ("後場始値", 584913773, "pm_open"),
    ("後場高値", 1279938624, "pm_high"),
    ("後場安値", 1826620378, "pm_low"),
    ("後場終値", 1061743785, "pm_close"),
    ("VWAP", 582188033, "vwap"),
    ("出来高", 46406858, "volume"),
]
HEADER_ROW = ["コード", "銘柄名", "業種", "市場"]
DATA_START_COL = 5  # E列(1-indexed)

LOG_SHEET_TITLE = "_ProcessedDates"


# ====================================================================
# 日付フォーマット
# ====================================================================

def format_date_jp(d: date) -> str:
    """
    date オブジェクトを "2026年8月7日(金)" 形式の文字列に変換する。
    スプレッドシートに書き込む日付列の見出し・ログシートの日付として使用する。
    """
    return f"{d.year}年{d.month}月{d.day}日({WEEKDAY_JP[d.weekday()]})"


# ====================================================================
# 1. JPXサイトから最新(または指定日)の「株式相場表」PDFリンクを取得
# ====================================================================

JPX_HOST = "https://www.jpx.co.jp"

# JPXサイトは python-requests の既定UAだと別内容を返す/弾くことがあるため、ブラウザ相当のヘッダーを付ける
HTTP_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
    "Accept-Language": "ja,en;q=0.8",
}

STQ_RE = re.compile(r"""(?P<path>[^\s"'<>()]*stq_(?P<ymd>\d{8})\.pdf)""")

# 一覧がJS描画になったため、PDFの置き場所を直接推測するときの候補(上から順に試す)
PDF_URL_TEMPLATES = [
    JPX_HOST + "/markets/statistics-equities/daily/stq_{ymd}.pdf",
    JPX_HOST + "/automation/markets/statistics-equities/daily/stq_{ymd}.pdf",
    JPX_HOST + "/markets/statistics-equities/daily/pdf/stq_{ymd}.pdf",
]


def _abs_url(path: str, base: str = JPX_INDEX_URL) -> str:
    from urllib.parse import urljoin
    return urljoin(base, path)


def _collect_stq_links(text: str, base: str = JPX_INDEX_URL) -> dict:
    """テキスト(HTML/JS/JSON問わず)から stq_YYYYMMDD.pdf を全部拾い {date: url} で返す。"""
    found = {}
    for m in STQ_RE.finditer(text.replace("\\/", "/")):
        try:
            d = datetime.strptime(m.group("ymd"), "%Y%m%d").date()
        except ValueError:
            continue
        found.setdefault(d, _abs_url(m.group("path"), base))
    return found


def _links_from_static_html() -> dict:
    """方法1: 静的HTML本文 + そこから参照されるJS/JSONファイルを調べる。"""
    resp = requests.get(JPX_INDEX_URL, headers=HTTP_HEADERS, timeout=30)
    resp.raise_for_status()
    html = resp.text

    found = _collect_stq_links(html)
    if found:
        return found

    # 一覧がJSで後から描画されている場合、データ元(js/json)に日付一覧が入っていることが多い
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "html.parser")
    refs = [s["src"] for s in soup.find_all("script", src=True)]
    refs += re.findall(r"""["']([^"']+\.(?:json|js)(?:\?[^"']*)?)["']""", html)
    seen = set()
    for ref in refs:
        url = _abs_url(ref)
        if url in seen or "googletagmanager" in url or "facebook" in url:
            continue
        seen.add(url)
        try:
            r = requests.get(url, headers=HTTP_HEADERS, timeout=30)
            if r.ok:
                found.update(_collect_stq_links(r.text, url))
        except requests.RequestException:
            continue
        if found:
            break
    return found


def _url_exists(url: str) -> bool:
    try:
        r = requests.get(url, headers=HTTP_HEADERS, timeout=30, stream=True)
        ok = r.status_code == 200 and "pdf" in r.headers.get("Content-Type", "").lower()
        r.close()
        return ok
    except requests.RequestException:
        return False


def _guess_pdf_url(d: date) -> Optional[str]:
    """方法2: 日付からPDFのURLを直接組み立てて存在確認する。"""
    ymd = d.strftime("%Y%m%d")
    for tmpl in PDF_URL_TEMPLATES:
        url = tmpl.format(ymd=ymd)
        if _url_exists(url):
            return url
    return None


def _ensure_playwright():
    """
    Playwright が無ければ pip と chromium を自動インストールして sync_playwright を返す。
    (ワークフロー側でインストール手順を足していなくても動くようにするため)
    失敗したら None を返す。
    """
    import subprocess

    try:
        from playwright.sync_api import sync_playwright
        return sync_playwright
    except ImportError:
        pass

    print("  [INFO] playwright が未インストールのため自動インストールします...", file=sys.stderr)
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", "playwright"],
                       check=True)
        # まずOS依存ライブラリ込みで、権限が無ければブラウザ本体のみで試す
        r = subprocess.run([sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"])
        if r.returncode != 0:
            subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
        from playwright.sync_api import sync_playwright
        return sync_playwright
    except Exception as e:
        print(f"  [WARN] playwright の自動インストールに失敗: {e}", file=sys.stderr)
        return None


def _links_from_browser() -> dict:
    """方法3: Playwright(任意)でJS描画後のDOMからリンクを取得する。未インストールならスキップ。"""
    sync_playwright = _ensure_playwright()
    if sync_playwright is None:
        return {}
    found = {}
    seen_urls = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(user_agent=HTTP_HEADERS["User-Agent"], locale="ja-JP")

        def on_response(resp):
            # 一覧データをXHR/JSONで取得している場合に備え、通信内容からもリンクを拾う
            try:
                seen_urls.append(resp.url)
                ctype = resp.headers.get("content-type", "").lower()
                if any(k in ctype for k in ("json", "javascript", "html", "text")):
                    found.update(_collect_stq_links(resp.text(), resp.url))
            except Exception:
                pass

        page.on("response", on_response)
        page.goto(JPX_INDEX_URL, wait_until="networkidle", timeout=60000)
        try:
            page.wait_for_selector("a[href*='stq_']", timeout=15000)
        except Exception:
            pass
        hrefs = page.eval_on_selector_all("a[href]", "els => els.map(e => e.href)")
        browser.close()
    found.update(_collect_stq_links("\n".join(hrefs)))
    if not found:
        print("  [DEBUG] ブラウザ描画でもstqリンクなし。ページが読み込んだjpx.co.jp内URL:", file=sys.stderr)
        for u in seen_urls:
            if "jpx.co.jp" in u:
                print(f"    {u}", file=sys.stderr)
    return found


def find_pdf_url(target_date: Optional[date] = None) -> tuple[str, date]:
    """
    株式相場表(stq_YYYYMMDD.pdf)のURLを取得する。
    日報indexページの一覧がJavaScript描画になり、静的HTMLにリンクが無くなったため、
      1. 静的HTML/参照JS・JSON内のリンク
      2. 日付からのURL直接推測(最新日指定の場合は直近の平日を遡って探索)
      3. Playwrightによるブラウザ描画(インストールされている場合のみ)
    の順に試す。target_date 未指定なら最新日を返す。
    """
    # --- 方法1 ---
    try:
        links = _links_from_static_html()
    except requests.RequestException as e:
        print(f"  [WARN] indexページ取得に失敗: {e}", file=sys.stderr)
        links = {}

    # --- 方法3(方法1が空のときだけ。方法2より先に試して実際の一覧を優先) ---
    if not links:
        try:
            links = _links_from_browser()
        except Exception as e:  # playwrightの起動失敗などは致命的にしない
            print(f"  [WARN] ブラウザ描画での取得に失敗: {e}", file=sys.stderr)
            links = {}

    if links:
        if target_date is None:
            d = max(links)
            return links[d], d
        if target_date in links:
            return links[target_date], target_date
        # 一覧(当月分のみ等)に無い過去日は、見つかったリンクのURLの型から組み立てて試す
        sample_d = max(links)
        sample_ymd = sample_d.strftime("%Y%m%d")
        PDF_URL_TEMPLATES.insert(0, links[sample_d].replace(sample_ymd, "{ymd}"))
        url = _guess_pdf_url(target_date)
        if url:
            return url, target_date
        raise RuntimeError(f"指定日 {target_date} の株式相場表PDFが見つかりませんでした"
                           "(休場日、まだ掲載されていない、または一覧の表示範囲外の可能性があります)。")

    # --- 方法2: URL直接推測 ---
    print("  [INFO] 一覧からリンクを取得できないため、URLを直接推測します", file=sys.stderr)
    if target_date is not None:
        url = _guess_pdf_url(target_date)
        if url:
            return url, target_date
    else:
        from datetime import timedelta
        d = datetime.now().date()
        for _ in range(10):  # 直近10日(土日・祝日を飛ばす分)を遡る
            if d.weekday() < 5:
                url = _guess_pdf_url(d)
                if url:
                    return url, d
            d -= timedelta(days=1)

    raise RuntimeError("株式相場表PDFのURLを取得できませんでした。"
                       "JPXの一覧はJS描画のため、PDF_URL_TEMPLATES を実際のURLに合わせるか、"
                       "Playwright(pip install playwright && playwright install chromium)を導入してください。")


# ====================================================================
# 2. PDFダウンロード & テキスト抽出
# ====================================================================

def download_pdf(url: str) -> bytes:
    resp = requests.get(url, headers=HTTP_HEADERS, timeout=60)
    resp.raise_for_status()
    return resp.content


def _normalize_numeric_spacing(line: str) -> str:
    """
    単語抽出時に、数値の小数点やカンマの前後で不自然にスペースが入ってしまうことがあるため、
    "4,465 .00" -> "4,465.00" のように数値表記を正規化する。
    """
    line = re.sub(r'(\d)\s+\.(\d)', r'\1.\2', line)   # "123 .45" -> "123.45"
    line = re.sub(r'(\d)\s*,\s+(\d{3})', r'\1,\2', line)  # "4 ,465" -> "4,465"
    line = re.sub(r'-\s+(\d)', r'-\1', line)          # "- 51.00" -> "-51.00"
    return line


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """
    PDFからテキストを抽出する。

    pdfplumber の素朴な extract_text() は、密な表組みのPDFで隣接する列同士の間の
    空白を正しく再現できず、数値が連結されてしまう(例: "4,465.004,490.00")ことがある。
    そのため、単語(word)単位で座標(top/x0)を取得し、Y座標が近い単語をまとめて1行とし、
    X座標順に並べて明示的に半角スペースで結合する方式で行を再構成する。
    """
    import pdfplumber

    lines = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=1, y_tolerance=3, keep_blank_chars=False)
            if not words:
                continue
            words.sort(key=lambda w: (w["top"], w["x0"]))

            current_line = []
            current_top = None
            for w in words:
                if current_top is None or abs(w["top"] - current_top) <= 3:
                    current_line.append(w)
                    current_top = w["top"] if current_top is None else (current_top + w["top"]) / 2
                else:
                    current_line.sort(key=lambda x: x["x0"])
                    lines.append(_normalize_numeric_spacing(" ".join(x["text"] for x in current_line)))
                    current_line = [w]
                    current_top = w["top"]
            if current_line:
                current_line.sort(key=lambda x: x["x0"])
                lines.append(_normalize_numeric_spacing(" ".join(x["text"] for x in current_line)))

    return "\n".join(lines)


# ====================================================================
# 3. テキストパース
# ====================================================================

def is_junk(line: str) -> bool:
    s = line.strip()
    if not s:
        return True
    if s in HEADER_JUNK_EXACT:
        return True
    if DATE_PAGE_RE.match(s):
        return True
    if COPYRIGHT_RE.match(s):
        return True
    return False


# 数値 または「－」(その場で売買なし/特殊気配なし)。全角数字(２２５など)を数値と誤認しないよう ASCII のみ。
_NUMD = r'(?:[0-9,]+(?:\.[0-9]+)?|-[0-9,]+(?:\.[0-9]+)?|[－ー―])'
_NUMD_RE = re.compile(r'^' + _NUMD + r'$')

# 「コード [マーク] 売買単位 銘柄名…」で始まる行。
# マークは D / N / W / V / # など1文字で、コードと売買単位の間に入ることがある
# (2026-09-29 のPDFでは約5,000行に付いており、これを想定していなかったため大量に取りこぼしていた)。
_ROW_HEAD_RE = re.compile(r'^(' + CODE_RE_STR + r')\s+(?:([A-Za-z#＃])\s+)?(\d+)\s+\S')

_DATA_KEYS = ["am_open", "am_high", "am_low", "am_close",
              "pm_open", "pm_high", "pm_low", "pm_close",
              "final_quote", "net_change", "vwap", "volume", "value"]

# 内国株式以外のセクション見出し(空白除去後の先頭一致)。INCLUDE_NON_STOCK_SECTIONS=False ならこれ以降の行は対象外
NON_STOCK_SECTION_PREFIXES = (
    "出資証券", "内国投資信託受益証券", "内国投資証券", "外国投信受益証券", "外国投資証券",
)
STOCK_SECTION_HEADS = {"内国株式", "外国株式"}
STOCK_SECTION_HEADS_EN = {"DomesticStock", "ForeignStock"}
# ETF・ETN・REIT・出資証券などもシートに載せたい場合は True にする(業種・市場は空欄になる)
INCLUDE_NON_STOCK_SECTIONS = False


def _parse_data_line(s: str) -> Optional[dict]:
    """
    データ行を右端の13個の数値(または「－」)から逆算して解析する。
        前場OHLC(4) + 後場OHLC(4) + 最終気配 + 前日比 + VWAP + 出来高 + 売買代金
    左側の「コード [マーク] 売買単位」の後ろ、数値より前の部分を和文銘柄名とする。
    こうすると、銘柄名に空白が含まれていても(例: "ＴＳＩ ＨＤ")、片方の場が売買なしで
    「－ － － －」になっていても、同じように解析できる。
    """
    m = _ROW_HEAD_RE.match(s)
    if not m:
        return None
    code, flag, unit = m.group(1), m.group(2), m.group(3)
    toks = re.sub(r'([－ー―])(?=[0-9])', r'\1 ', s).split()  # 「－30.00」のような連結を分離
    head = 3 if flag else 2  # コード, [マーク], 売買単位
    if len(toks) < head + 1 + 13:
        return None

    tail = toks[-13:]
    if all(_NUMD_RE.match(t) for t in tail):
        name_jp, name_en = " ".join(toks[head:-13]), ""
    else:
        # 英文社名が同じ行に混ざった場合の救済: 数値トークンだけ順に拾い、残りを社名とみなす
        rest = toks[head:]
        nums = [t for t in rest if _NUMD_RE.match(t)]
        words = [t for t in rest if not _NUMD_RE.match(t)]
        if len(nums) != 13 or not words:
            return None
        tail, name_jp, name_en = nums, words[0], " ".join(words[1:])

    rec = {"code": code, "unit": unit, "flag": flag or "", "name_jp": name_jp, "name_en": name_en}
    rec.update(dict(zip(_DATA_KEYS, tail)))
    return rec


def parse_text(text: str) -> list[dict]:
    """
    PDFから抽出したテキストを解析し、銘柄ごとのレコード(dict)のリストを返す。

    実際のPDFのレイアウトは、1銘柄につき以下の2行で構成される:
        1行目: コード [マーク] 売買単位 銘柄名(和文) 前場OHLC 後場OHLC 最終気配 前日比 VWAP 出来高 売買代金
        2行目: 銘柄名(英文)

    その手前に出現する「業種」「市場」見出し行を状態として保持し、各レコードに付与する。
    PDFには通常取引のあとにPTS等の別表も続くが、数値が少ないためデータ行とは見なさず読み飛ばす。
    """
    lines = text.splitlines()
    content = [l for l in lines if not is_junk(l)]

    current_market = None
    current_industry = None
    in_stock_section = True  # 内国株式のセクション内か
    pending_record = None  # データ行は解析できたが、まだ英文社名行を待っている状態

    records = []
    unparsed = []  # 通常取引の行のはずなのに解析できなかった行(取りこぼし調査用)
    n_skipped_non_stock = 0

    for raw in content:
        s = raw.strip()

        is_row = _ROW_HEAD_RE.match(s) is not None
        rec = _parse_data_line(s) if is_row else None

        if pending_record is not None:
            if is_row:
                # 英文社名行が欠けて、次の銘柄の行が来た場合(次の銘柄を英文社名として食わない)
                records.append(pending_record)
                pending_record = None
            else:
                pending_record["name_en"] = s
                records.append(pending_record)
                pending_record = None
                continue

        if is_row:
            if rec is None:
                # 数値が多いのに解析できない行だけ警告(PTSなど数値の少ない別表の行は黙って除外)
                # (通常取引の行は数値が13個、PTSの行は6個。コード・売買単位を除いた数で判定)
                if in_stock_section and sum(1 for t in s.split()[2:] if _NUMD_RE.match(t)) >= 10:
                    unparsed.append(s)
                continue
            if not in_stock_section and not INCLUDE_NON_STOCK_SECTIONS:
                n_skipped_non_stock += 1
                continue
            if in_stock_section:
                rec["industry"] = current_industry
                rec["market"] = current_market
            else:
                rec["industry"] = ""
                rec["market"] = ""
            if rec["name_en"]:
                records.append(rec)
            else:
                pending_record = rec
            continue

        # ここに来るのは「業種見出し」「市場見出し」等のセクション行(または英文社名行)
        compact = re.sub(r'\s+', '', s)
        if s in MARKET_HEADERS:
            current_market = s.replace("市場", "")
            in_stock_section = True
        elif compact in STOCK_SECTION_HEADS:
            in_stock_section = True
        elif compact in STOCK_SECTION_HEADS_EN or s in NON_INDUSTRY_OTHER_LINES:
            pass  # 英語表記など、業種欄には転記しない
        elif compact.startswith(NON_STOCK_SECTION_PREFIXES):
            in_stock_section = False
        elif in_stock_section:
            # 業種見出しは "水産・農林業 Fishery,Agriculture & Forestry" のように
            # 和文+英文が空白区切りで並んでいるため、和文部分(先頭トークン)のみを採用する
            current_industry = s.split()[0] if s.split() else s

    if unparsed:
        print(f"  [WARN] 通常取引の行のはずなのに解析できなかった行が {len(unparsed)} 件あります"
              "(先頭20件を表示):", file=sys.stderr)
        for l in unparsed[:20]:
            print(f"    {l!r}", file=sys.stderr)

    if n_skipped_non_stock:
        print(f"  [INFO] 内国株式以外(ETF・REIT・出資証券等)の {n_skipped_non_stock} 行は対象外として除外しました",
              file=sys.stderr)

    if pending_record is not None:
        # 最終行がデータ行で終わっていて英文社名行が取得できなかった場合
        pending_record["name_en"] = ""
        records.append(pending_record)

    return records


# ====================================================================
# 4. Google スプレッドシートへの書き込み(ワイド形式: 銘柄=行, 日付=列)
# ====================================================================

def get_gspread_client():
    import gspread
    from google.oauth2.service_account import Credentials

    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not creds_json:
        raise RuntimeError("環境変数 GOOGLE_SERVICE_ACCOUNT_JSON が設定されていません。")
    info = json.loads(creds_json)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)


def col_letter(n: int) -> str:
    """1-indexed列番号をA1表記の列文字に変換 (5 -> 'E')。"""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def format_value(metric_key: str, raw: str) -> str:
    """
    シートに書き込む直前の値の整形。
    出来高(volume)は千株単位で抽出しているため、実株数(×1000)に変換する。
    """
    if raw in ("－", "ー", "―"):
        return ""  # その場で売買なし(「－」)は空欄にする。全角ダッシュの文字列を数値列に入れない
    if metric_key == "volume" and raw:
        try:
            v = float(raw.replace(",", "")) * 1000
            return f"{v:,.0f}"
        except ValueError:
            return raw
    return raw


def ensure_header(ws, all_values: list[list[str]]) -> list[list[str]]:
    """
    A1:D1 が正しく設定されているか確認し、必要なら補完する。
    シートが完全に空でない場合(例: E列以降に既存データがあるがA〜D列が未設定)でも
    正しくヘッダーを補完できるよう、"シートが空かどうか"ではなく
    "1行目のA〜D列の中身"を直接見て判定する。
    """
    if not all_values:
        ws.update("A1:D1", [HEADER_ROW])
        return [HEADER_ROW]

    header = all_values[0]
    header_ad = header[:4] if len(header) >= 4 else header + [""] * (4 - len(header))
    if header_ad != HEADER_ROW:
        ws.update("A1:D1", [HEADER_ROW])
        header = HEADER_ROW + header[4:]
        all_values = [header] + all_values[1:]
    return all_values


_HEADER_DATE_RE = re.compile(r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日')


def parse_header_date(s: str) -> Optional[date]:
    """ヘッダーセルの「2026年9月29日(火)」形式から date を取り出す(日付でなければ None)。"""
    m = _HEADER_DATE_RE.search(s or "")
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def update_wide_sheet(ws, date_str: str, records: list[dict], metric_key: str, force: bool = False):
    """
    1つの項目シート(ワイド形式)を更新する。
    - A=コード, B=銘柄名, C=業種, D=市場 は既存行がある場合は変更しない
    - 新しい日付は E列に挿入し、既存の日付列は右にずれる
    - 該当銘柄の当日データが無い場合は空欄
    - 未登録の銘柄(新規コード)は最終行に新しい行として追加
    """
    all_values = ws.get_all_values()
    all_values = ensure_header(ws, all_values)

    header = all_values[0]
    data_rows = all_values[1:]

    # 既に同じ日付列が存在する場合は二重挿入を避ける
    target_d = parse_header_date(date_str)
    cur_cols = ws.col_count  # 現在のシート列数(列削除に合わせて自分で更新する)
    # 同じ日付の列(重複していても全部)の列番号(0-indexed)。ヘッダーの曜日表記ゆれは日付で照合する。
    same_date_cols = [i for i in range(DATA_START_COL - 1, len(header))
                      if parse_header_date(header[i]) == target_d]
    if same_date_cols:
        if not force:
            print(f"  [SKIP] 既に {date_str} 列が存在するため挿入をスキップします(--force で作り直し可)")
            return
        # --force: 同じ日付の列(重複含む)をいったん全部削除し、正しい日付順の位置に入れ直す。
        # (従来は --force のたびに E列へ新しい列が挿入され、重複列や日付順の乱れが発生していた)
        ws.spreadsheet.batch_update({
            "requests": [{
                "deleteDimension": {
                    "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                              "startIndex": i, "endIndex": i + 1}
                }
            } for i in sorted(same_date_cols, reverse=True)]  # 右から消してズレを防ぐ
        })
        print(f"  [INFO] 既存の {date_str} 列 {len(same_date_cols)} 本を削除して作り直します")
        header = [h for i, h in enumerate(header) if i not in set(same_date_cols)]
        cur_cols -= len(same_date_cols)

    code_to_rowidx = {}
    for idx, row in enumerate(data_rows):
        if row and row[0]:
            code_to_rowidx[row[0]] = idx

    records_by_code = {r["code"]: r for r in records}
    n_existing_rows = len(data_rows)

    # 新しい列(E列)に入れる値: ヘッダー(日付) + 既存行それぞれの値(無ければ空欄)
    # ★ ここで data_rows を1行ずつ辿り、その行のA列コード(row[0])と同じコードの
    #    レコードを records_by_code から探して積んでいるため、既存シートの行順が
    #    保たれたまま、コードが一致する行にだけ値が入る(ズレは発生しない)。
    col_values = [date_str]
    for row in data_rows:
        code = row[0] if row else ""
        rec = records_by_code.get(code)
        col_values.append(format_value(metric_key, rec[metric_key]) if rec else "")

    # Step1: 日付の新しい順(左ほど新しい)になるよう、挿入位置を決めて空列を1つ入れる。
    # 通常の日次実行では最新日なので E列。過去日のバックフィルは該当する位置に入る。
    target_col = DATA_START_COL
    for i in range(DATA_START_COL - 1, len(header)):
        hd = parse_header_date(header[i])
        if hd is None or (target_d is not None and hd < target_d):
            break
        target_col = i + 2  # 1-indexed で「この列の右隣」
    if target_col <= cur_cols:
        ws.spreadsheet.batch_update({
            "requests": [{
                "insertDimension": {
                    "range": {
                        "sheetId": ws.id,
                        "dimension": "COLUMNS",
                        "startIndex": target_col - 1,
                        "endIndex": target_col,
                    },
                    "inheritFromBefore": False,
                }
            }]
        })
    else:
        ws.add_cols(target_col - cur_cols)  # シートの右端に追加する場合

    # Step2: 対象列に日付+各行の値をまとめて書き込み
    end_row = 1 + n_existing_rows
    col = col_letter(target_col)
    ws.update(f"{col}1:{col}{end_row}", [[v] for v in col_values],
              value_input_option="USER_ENTERED")

    # Step3: 既存シートに無い新規銘柄は末尾に新しい行として追加
    new_rows = []
    for r in records:
        if r["code"] not in code_to_rowidx:
            new_rows.append([
                r["code"], r["name_jp"], r["industry"] or "", r["market"] or "",
            ] + [""] * (target_col - DATA_START_COL) + [
                format_value(metric_key, r[metric_key]),
            ])
    if new_rows:
        ws.append_rows(new_rows, value_input_option="USER_ENTERED")

    print(f"  [OK] 既存{n_existing_rows}行を更新 / 新規{len(new_rows)}行を追加")


def mark_processed(spreadsheet, date_str: str, n_records: int):
    try:
        log_ws = spreadsheet.worksheet(LOG_SHEET_TITLE)
    except Exception:
        log_ws = spreadsheet.add_worksheet(title=LOG_SHEET_TITLE, rows="1000", cols="3")
        log_ws.update("A1:C1", [["日付", "件数", "実行日時"]])
    log_ws.append_row([date_str, n_records, datetime.now().isoformat(timespec="seconds")],
                       value_input_option="USER_ENTERED")


COUNT_DROP_TOLERANCE = 0.97  # 前回件数のこの割合を下回ったら異常とみなす


def _last_processed_count(spreadsheet) -> Optional[int]:
    """_ProcessedDates シートの最終行の件数を返す(無ければ None)。"""
    try:
        rows = spreadsheet.worksheet(LOG_SHEET_TITLE).get_all_values()
    except Exception:
        return None
    for row in reversed(rows[1:]):
        if len(row) >= 2:
            try:
                return int(str(row[1]).replace(",", ""))
            except ValueError:
                continue
    return None


def write_to_sheets(records: list[dict], target_date: date, dry_run: bool = False, force: bool = False):
    date_str = format_date_jp(target_date)

    if dry_run:
        print(f"[DRY-RUN] {date_str} 分、{len(records)}銘柄を各シートのE列に挿入する予定です"
              "(実際には書き込みません)。")
        sample = records[0]
        for name, _gid, key in SHEET_SPECS:
            print(f"  - {name}: 例 {sample['code']} {sample['name_jp']} "
                  f"(業種:{sample['industry']} / 市場:{sample['market']}) = {format_value(key, sample[key])}")
        return

    gc = get_gspread_client()
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", SPREADSHEET_ID_DEFAULT)
    sh = gc.open_by_key(spreadsheet_id)

    # 取りこぼし防止: 直近の処理件数より大きく減っていたら、シートを壊さないよう書き込み前に止める
    prev = _last_processed_count(sh)
    if prev and len(records) < prev * COUNT_DROP_TOLERANCE:
        msg = (f"抽出件数 {len(records)} が前回処理分 {prev} を大きく下回っています"
               f"(許容: {COUNT_DROP_TOLERANCE:.0%} 以上)。PDF解析の取りこぼしの可能性があります。")
        if not force:
            raise RuntimeError(msg + " 内容を確認のうえ、問題なければ --force で再実行してください。")
        print(f"  [WARN] {msg} --force 指定のため続行します。", file=sys.stderr)

    for name, gid, key in SHEET_SPECS:
        print(f"[{name}] 更新中...")
        ws = sh.get_worksheet_by_id(gid)
        if ws is None:
            print(f"  [ERROR] gid={gid} のシートが見つかりません。スキップします。", file=sys.stderr)
            continue
        update_wide_sheet(ws, date_str, records, key, force=force)
        time.sleep(1)  # API レート制限対策

    mark_processed(sh, date_str, len(records))


# ====================================================================
# デバッグ補助
# ====================================================================

def _dump_debug_info(text: str):
    """解析失敗時に、実際に抽出されたテキストの様子をログに出して原因調査を助ける。"""
    lines = text.splitlines()
    non_empty = [l for l in lines if l.strip()]
    print(f"[DEBUG] 抽出テキスト: 総行数={len(lines)}, 非空行数={len(non_empty)}", file=sys.stderr)
    print("[DEBUG] 先頭60行:", file=sys.stderr)
    for l in non_empty[:60]:
        print(f"    {l!r}", file=sys.stderr)
    # コードらしき行(4桁の英数字+空白+日本語)がそもそも存在するか確認
    code_like = [l for l in non_empty if re.match(r'^' + CODE_RE_STR + r'\s', l.strip())][:10]
    print(f"[DEBUG] コード行らしきもの: {len(code_like)}件(先頭10件)", file=sys.stderr)
    for l in code_like:
        print(f"    {l!r}", file=sys.stderr)
    # 数値が並んでいそうな行(データ行候補)を探す
    numeric_like = [l for l in non_empty if len(re.findall(r'[\d,]+\.\d+', l)) >= 4][:10]
    print(f"[DEBUG] 数値が4つ以上並ぶ行(データ行候補): {len(numeric_like)}件(先頭10件)", file=sys.stderr)
    for l in numeric_like:
        print(f"    {l!r}", file=sys.stderr)


# ====================================================================
# main
# ====================================================================

def main():
    parser = argparse.ArgumentParser(description="JPX 東京証券取引所日報 株式相場表 取得スクリプト")
    parser.add_argument("--date", type=str, default=None,
                         help="取得したい日付 (YYYY-MM-DD)。省略時は最新日を自動検出。")
    parser.add_argument("--dry-run", action="store_true",
                         help="スプレッドシートへの書き込みを行わず、解析結果件数のみ表示する。")
    parser.add_argument("--force", action="store_true",
                         help="同日データが処理済みでも強制的に再書き込みする。")
    args = parser.parse_args()

    target_date = None
    if args.date:
        target_date = datetime.strptime(args.date, "%Y-%m-%d").date()

    print("[STEP1] JPXサイトからPDFリンクを検索中...")
    pdf_url, resolved_date = find_pdf_url(target_date)
    print(f"  -> {resolved_date} : {pdf_url}")

    print("[STEP2] PDFをダウンロード中...")
    pdf_bytes = download_pdf(pdf_url)
    print(f"  -> {len(pdf_bytes):,} bytes")

    print("[STEP3] PDFからテキストを抽出中...")
    text = extract_text_from_pdf(pdf_bytes)

    print("[STEP4] テキストを解析中...")
    records = parse_text(text)
    print(f"  -> {len(records)} 銘柄を抽出しました")

    if not records:
        print("[ERROR] 1件も解析できませんでした。PDFフォーマットが変更された可能性があります。", file=sys.stderr)
        _dump_debug_info(text)
        sys.exit(1)

    print("[STEP5] Google スプレッドシートへ書き込み中...")
    write_to_sheets(records, resolved_date, dry_run=args.dry_run, force=args.force)

    print("完了しました。")


if __name__ == "__main__":
    main()
