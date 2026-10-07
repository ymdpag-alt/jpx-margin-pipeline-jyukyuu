"""
JPX 銘柄別信用取引残高（日次）＋ 日証協 貸株残（週次） → Google スプレッドシート

[信用残] JPX「銘柄別信用取引残高」 {申込日}_mtall.pdf
    2026/9/28 から全銘柄・毎営業日16:00頃に前営業日分を公表（掲載ページ: margin/01.html）。
    一般信用買／一般信用売／制度信用買／制度信用売 の4シートへ書き込む。

[貸株残] 日本証券業協会「銘柄別株券等貸借週末残高」 {申込日}z.xlsx
    週次（報告週の翌週木曜公表）。新しいファイルが無い日は何もしない。

シートの形:  A列=銘柄コード / B列=銘柄名 / C列以降=申込日（左ほど新しい）
    - 新しい日付は、日付の降順が保たれる位置に1列挿入する（通常はC列）。
    - シートに無い銘柄（新規上場など）は最下部に追加する。
    - 同じ日付の列が既にあれば何もしない（何度実行しても安全）。

使い方:
    python margin_balance_to_sheets.py                 # 通常実行（シートを更新）
    python margin_balance_to_sheets.py --inspect 20261005
                                                       # PDFの読み取り確認だけ（書き込みなし）
"""

from __future__ import annotations

import io
import json
import os
import re
import sys
import time
from datetime import datetime
from urllib.parse import urljoin

import gspread
import pandas as pd
import pdfplumber
import requests
from google.oauth2.service_account import Credentials
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# #############################################################################
#
#  設定（変更が必要になるのはこのブロックだけ）
#
# #############################################################################

# -----------------------------------------------------------------------------
# 1. 実行時の設定（GitHub Actions の env / secrets から受け取る）
# -----------------------------------------------------------------------------
GOOGLE_SERVICE_ACCOUNT_JSON_ENV = "GOOGLE_SERVICE_ACCOUNT_JSON"   # 認証JSONの環境変数名

def _env(name: str, default: str = "") -> str:
    """環境変数を読む。未設定・空文字（Secrets未登録など）のときは default。"""
    return os.environ.get(name, "").strip() or default


SPREADSHEET_ID = _env("SPREADSHEET_ID")                           # 信用残の書き込み先ブック（必須）
KASHIKABU_SPREADSHEET_ID = _env(                                  # 貸株残の書き込み先ブック
    "KASHIKABU_SPREADSHEET_ID", "1kWST0CkkIvo3irPSbMgtVtUqqRDXFwvRREYZDRQAFMY"
)

MARGIN_BACKFILL_N = int(_env("MARGIN_BACKFILL_N", "5"))           # 掲載分から最大何日分を埋めるか
MARGIN_EXTRA_DATES = _env("MARGIN_EXTRA_DATES")                   # 掲載外の過去日（例: 20260925,20260928）
TARGET_DATE_OVERRIDE = _env("TARGET_DATE_OVERRIDE")               # 1日分だけ強制取得（EXTRAと同じ扱い）
MAX_DATE_COLS = int(_env("MAX_DATE_COLS", "0"))                   # 信用残シートの日付列の上限（0=削除しない）

JSDA_ENABLED = _env("JSDA_ENABLED", "1") != "0"                   # "0" で貸株残をスキップ
JSDA_BACKFILL_N = int(_env("JSDA_BACKFILL_N", "1"))               # 貸株残を最大何週分埋めるか

# -----------------------------------------------------------------------------
# 2. シート構成
# -----------------------------------------------------------------------------
MARGIN_SHEETS = {            # DataFrameの列名 → 書き込み先シート名
    "一般信用買残高": "一般信用買残高",
    "一般信用売残高": "一般信用売残高",
    "制度信用買残高": "制度信用買残高",
    "制度信用売残高": "制度信用売残高",
}
KASHIKABU_SHEET_NAME = "貸株残"
KASHIKABU_VALUE_COL = "貸株残"

FIXED_COLS = 2               # A列=銘柄コード, B列=銘柄名
FIRST_DATE_COL = 3           # C列（1始まり）
NEW_SHEET_COLS = 50          # シート新規作成時の列数
NEW_SHEET_ROWS = 100         # シート新規作成時の行数

GOOGLE_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# -----------------------------------------------------------------------------
# 3. JPX（信用残）
# -----------------------------------------------------------------------------
JPX_LIST_URL = "https://www.jpx.co.jp/markets/statistics-equities/margin/01.html"
JPX_PDF_URL_TEMPLATE = (     # 掲載から外れた日付を直接取りに行くとき用
    "https://www.jpx.co.jp/markets/statistics-equities/margin/"
    "tvdivq0000001rnl-att/{date}_mtall.pdf"
)
JPX_LINK_PATTERN = re.compile(r'href="([^"]*?(\d{8})_mtall\.pdf)"')

# -----------------------------------------------------------------------------
# 4. 日証協（貸株残）
# -----------------------------------------------------------------------------
JSDA_INDEX_URL = "https://www.jsda.or.jp/shiryoshitsu/toukei/kabu-taiw/index.html"
JSDA_FILE_URL_TEMPLATE = "https://www.jsda.or.jp/shiryoshitsu/toukei/kabu-taiw/files/{date}z.xlsx"
JSDA_LINK_PATTERN = re.compile(r"files/(\d{8})z\.xlsx")

JSDA_CODE_KEYWORDS = ["銘柄コード", "コード", "code"]          # 列見出しの候補（優先順）
JSDA_NAME_KEYWORDS = ["銘柄名", "名称", "銘柄"]
JSDA_BALANCE_KEYWORDS = ["貸付残高", "貸付", "貸株", "週末残高", "残高"]
JSDA_BALANCE_EXCLUDE = ["借入", "返済", "新規", "成約"]

# -----------------------------------------------------------------------------
# 5. JPX PDF の読み取りルール（2026/10/5 申込分の実データで確認済み）
# -----------------------------------------------------------------------------
# 1銘柄は「株数」行と「金額」行の2行。株数行だけを読む。
#   B 極洋　普通株式 プライム 貸 13010 JP3257200000 株数 Shs. 9,500 100 0.1% 161,300 1,200 1.3% ...
#   └単位 └銘柄名  └種別  └市場 └貸借 └コード └ISIN
#
# 「株数 Shs.」以降の14個の値の並び:
#    0 売残高   1 前日比   2 上場比
#    3 買残高   4 前日比   5 上場比
#    6 売・一般 7 前日比   8 売・制度 9 前日比
#   10 買・一般 11 前日比 12 買・制度 13 前日比
N_VALUES = 14
IDX = {
    "売残高合計": 0,
    "買残高合計": 3,
    "一般信用売残高": 6,
    "制度信用売残高": 8,
    "一般信用買残高": 10,
    "制度信用買残高": 12,
}

STOCK_LINE_PATTERN = re.compile(          # 方式A: 1行にすべて並んでいる場合
    r"^(?P<prefix>.*?)\s*(?P<code>[0-9A-Z]{5})\s+"
    r"(?P<isin>[A-Z]{2}[A-Z0-9]{9}[0-9])\s+株数\s*(?:Shs\.?)?\s*(?P<body>.*)$"
)
# 方式B: 銘柄名・コード・数値の高さが微妙にずれて別の行に分かれる場合は、
#        「株数」の文字の座標を基準に、同じ銘柄の部品を位置で集め直す。
ISIN_PATTERN = re.compile(r"[A-Z]{2}[A-Z0-9]{9}[0-9]")
CODE_PATTERN = re.compile(r"^[0-9A-Z]{5}$")
WORD_X_TOLERANCE = 1.5                     # 文字を単語にまとめる横方向の許容幅（pt）
DEFAULT_ROW_GAP_RATIO = 2.2                # 「金額」行が見つからないときの行間（文字高さの倍率）
VALUE_TOKEN_PATTERN = re.compile(          # ▲=減少 / 0.1%・*=上場比
    r"([▲△])?\s*(\d[\d,]*\.\d+%|\d[\d,]*%|\d[\d,]*|\*)"
)

NAME_CLEANUP_PATTERNS = [                  # 銘柄名から取り除く部分（上から順に適用）
    re.compile(r"^[AJKBMCTF][\s　]+"),                               # 先頭の売買単位フラグ
    re.compile(r"[\s　]*(貸|制|他)$"),                               # 末尾の貸借区分
    re.compile(r"[\s　]*(プライム|スタンダード|グロース|投信等|ＰＲＯ|PRO|TOKYO PRO Market)$"),
    re.compile(r"[\s　]*(普通株式|受益証券|投資証券|投資口|出資証券|優先株式|優先出資証券|"
               r"新株予約権証券|外国株預託証券|株式)$"),
]

MAX_PARSE_FAILURES = 20          # 失敗行がこれ（または全体の1%）を超えたらレイアウト変更とみなす
MAX_PARSE_FAILURE_RATIO = 0.01

# 小計・総合計行のラベル付け
SEGMENT_KEYWORDS = ["プライム", "スタンダード", "グロース", "投信等"]
CATEGORY_ORDER = ["貸借銘柄", "制度信用銘柄", "その他", "総合計", "全体"]
SEGMENT_ORDER = ["合計", "プライム", "スタンダード", "グロース", "投信等"]
SUBTOTAL_CODE_PREFIX = "SUB_"

# -----------------------------------------------------------------------------
# 6. HTTP
# -----------------------------------------------------------------------------
HTTP_TIMEOUT_PAGE = 30
HTTP_TIMEOUT_FILE = 60
HTTP_RETRY_TOTAL = 5                 # 429・5xx のときの再試行回数
HTTP_RETRY_BACKOFF = 5               # 再試行の待ち時間の基準（秒）。5, 10, 20, ... と伸びる
DOWNLOAD_INTERVAL_JPX = 2            # ファイル取得ごとの待ち時間（秒）
DOWNLOAD_INTERVAL_JSDA = 10          # 日証協は間隔を空けないと 429 になりやすい
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
    """'20261005' → '2026年10月5日(月)'"""
    dt = datetime.strptime(yyyymmdd, "%Y%m%d")
    return f"{dt.year}年{dt.month}月{dt.day}日({WEEKDAY_JP[dt.weekday()]})"


def from_japanese_date(text: str) -> str:
    """'2026年10月5日(月)' → '20261005'。読めなければ元の文字列を返す。"""
    try:
        dt = datetime.strptime(text.split("(")[0].strip(), "%Y年%m月%d日")
        return dt.strftime("%Y%m%d")
    except (ValueError, IndexError):
        return text


def is_yyyymmdd(text: str) -> bool:
    return bool(re.fullmatch(r"\d{8}", text or ""))


def parse_date_list(raw: str) -> list[str]:
    """'20260925, 20260928' → ['20260925', '20260928']"""
    dates = [d for d in re.split(r"[,\s]+", raw or "") if d]
    invalid = [d for d in dates if not is_yyyymmdd(d)]
    if invalid:
        raise ValueError(f"日付は YYYYMMDD で指定してください: {invalid}")
    return dates


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


def print_title(title: str) -> None:
    print("\n" + "=" * 60)
    print(title)
    print("=" * 60)


# =============================================================================
# JPX: 掲載一覧と PDF の取得
# =============================================================================

def list_jpx_files(sess: requests.Session) -> dict[str, str]:
    """01.html に載っている日次PDFを {申込日: URL} で返す。"""
    html = fetch_text(sess, JPX_LIST_URL)
    files = {date: urljoin(JPX_LIST_URL, href) for href, date in JPX_LINK_PATTERN.findall(html)}
    if not files:
        raise ValueError("01.html に *_mtall.pdf のリンクがありません。ページ構成かファイル名が変わった可能性があります。")
    return files


def download_jpx_pdf(sess: requests.Session, url: str) -> bytes:
    content = fetch_file(sess, url)
    if not content.startswith(b"%PDF"):
        raise ValueError(f"PDFではない応答が返りました: {url}")
    return content


# =============================================================================
# JPX: PDF の読み取り
# =============================================================================

def parse_values(body: str) -> list[float | None]:
    """数値を順に取り出す。上場比（% / *）は位置合わせのため None で1個と数える。"""
    values: list[float | None] = []
    for sign, token in VALUE_TOKEN_PATTERN.findall(body):
        if token == "*" or token.endswith("%"):
            values.append(None)
        else:
            number = float(token.replace(",", ""))
            values.append(-number if sign else number)
    return values


def is_consistent(values: list[float | None]) -> bool:
    """売残高合計＝一般＋制度、買残高合計＝一般＋制度 か（列ずれ検出）。"""
    try:
        return (
            values[IDX["売残高合計"]] == values[IDX["一般信用売残高"]] + values[IDX["制度信用売残高"]]
            and values[IDX["買残高合計"]] == values[IDX["一般信用買残高"]] + values[IDX["制度信用買残高"]]
        )
    except TypeError:
        return False


def pick_balances(values: list[float | None]) -> dict[str, float]:
    return {col: values[IDX[col]] for col in MARGIN_SHEETS}


def clean_name(prefix: str) -> str:
    name = prefix.strip()
    for pattern in NAME_CLEANUP_PATTERNS:
        name = pattern.sub("", name)
    return name.strip(" 　")


def parse_stock_line(line: str) -> tuple[dict | None, str | None]:
    """
    銘柄の株数行を読む。
    戻り値: (レコード, None) / (None, 失敗理由) / 銘柄行でなければ (None, None)
    """
    m = STOCK_LINE_PATTERN.match(line.strip())
    if not m:
        return None, None

    values = parse_values(m.group("body"))
    if len(values) < N_VALUES:
        return None, f"数値が{len(values)}個しかありません"
    values = values[:N_VALUES]
    if not is_consistent(values):
        return None, "合計≠一般＋制度（列ずれの可能性）"

    record = {
        "銘柄コード": m.group("code")[:4],   # 5桁目の付番（通常0）を除く
        "銘柄名": clean_name(m.group("prefix")),
        **pick_balances(values),
    }
    return record, None


def _failure_code(text: str) -> str | None:
    """失敗の記録から4桁の銘柄コードを取り出す（ISINの直前の5桁、無ければ最初の5桁）。"""
    m = re.search(r"\b([0-9A-Z]{5})\s+[A-Z]{2}[A-Z0-9]{9}[0-9]\b", text) or re.search(r"\b([0-9A-Z]{5})\b", text)
    return m.group(1)[:4] if m else None


def _mid(word: dict) -> float:
    return (word["top"] + word["bottom"]) / 2


def parse_stock_words(page) -> tuple[list[dict], list[tuple[str, str]]]:
    """
    方式B: 単語の座標から銘柄を組み立てる。

      ┌ 銘柄名・市場・貸借 ┐┌ コード ISIN ┐┌ 株数 Shs. 9,500 100 0.1% ... ┐  ← 株数の行
      └ 英語名 ...        ┘└            ┘└ 金額 Val. ...                ┘  ← 金額の行

    「株数」を起点に、
      - 数値     : 同じ高さで右側にある単語
      - ISIN/コード: 左側で、株数の行〜金額の行の高さにある単語
      - 銘柄名   : コードより左で、株数の行の高さにある単語
    を集める。ISINが見つからない「株数」（＝合計行）は対象外。
    """
    words = page.extract_words(x_tolerance=WORD_X_TOLERANCE, keep_blank_chars=False)
    anchors = [w for w in words if "株数" in w["text"]]
    val_rows = [w for w in words if "金額" in w["text"]]

    records: list[dict] = []
    failures: list[tuple[str, str]] = []

    for a in anchors:
        a_mid = _mid(a)
        height = a["bottom"] - a["top"]

        # 行間 = この「株数」から直下の「金額」までの距離
        below = [v["top"] - a["top"] for v in val_rows if v["top"] > a["top"] + height * 0.5]
        gap = min(below, default=height * DEFAULT_ROW_GAP_RATIO)
        gap = min(gap, height * 4)
        half = gap / 2

        # ---- ISIN（左側・株数の行〜金額の行）----
        isins = [
            w for w in words
            if ISIN_PATTERN.fullmatch(w["text"])
            and w["x1"] <= a["x0"] + 1
            and a_mid - half <= _mid(w) <= a_mid + gap + half * 0.5
        ]
        if not isins:
            continue  # 合計行など
        isin = min(isins, key=lambda w: (abs(_mid(w) - a_mid), a["x0"] - w["x1"]))

        # ---- コード（ISINのすぐ左）----
        codes = [
            w for w in words
            if CODE_PATTERN.match(w["text"])
            and w["x1"] <= isin["x0"] + 1
            and abs(_mid(w) - _mid(isin)) <= half
        ]
        if not codes:
            failures.append(("コードが見つかりません", isin["text"]))
            continue
        code = max(codes, key=lambda w: w["x1"])

        # ---- 数値（同じ高さ・右側）----
        head = re.sub(r"^.*?株数\s*(?:Shs\.?)?", "", a["text"])
        right = sorted(
            (w for w in words if w["x0"] >= a["x1"] - 0.5 and abs(_mid(w) - a_mid) <= half * 0.8),
            key=lambda w: w["x0"],
        )
        body = " ".join([head] + [w["text"] for w in right if not w["text"].startswith("Shs")])
        values = parse_values(body)
        if len(values) < N_VALUES:
            failures.append((f"数値が{len(values)}個しかありません", f"{code['text']} {body[:80]}"))
            continue
        values = values[:N_VALUES]
        if not is_consistent(values):
            failures.append(("合計≠一般＋制度（列ずれの可能性）", f"{code['text']} {body[:80]}"))
            continue

        # ---- 銘柄名（コードより左・株数の行の高さ）----
        name_words = sorted(
            (w for w in words if w["x1"] <= code["x0"] + 0.5 and abs(_mid(w) - a_mid) <= half),
            key=lambda w: (round(w["top"]), w["x0"]),
        )
        records.append({
            "銘柄コード": code["text"][:4],
            "銘柄名": clean_name(" ".join(w["text"] for w in name_words)),
            **pick_balances(values),
        })

    return records, failures


def label_subtotal(prefix: str, state: dict) -> str | None:
    """小計・合計行のラベルを行の文字とカテゴリの流れから決める。"""
    if "総合計" in prefix:
        state["passed_total"] = True
        return "総合計"

    if "貸借銘柄" in prefix:
        state["category"] = "貸借銘柄"
    elif "制度信用銘柄" in prefix:
        state["category"] = "制度信用銘柄"
    elif "その他" in prefix:
        state["category"] = "その他"

    if "小計" not in prefix:
        return f"{state['category']} 合計" if state["category"] else None

    for segment in SEGMENT_KEYWORDS:
        if segment in prefix:
            category = "全体" if state["passed_total"] else state["category"]
            return f"{category} {segment} 小計" if category else None
    return None


def parse_subtotal_line(line: str, state: dict) -> tuple[str, list[float]] | None:
    """ISINを含まず「株数」と「計」を含む行を小計・合計行として読む。"""
    if "株数" not in line or "計" not in line:
        return None

    prefix, _, body = line.partition("株数")
    values = parse_values(re.sub(r"^\s*Shs\.?", "", body))
    if len(values) < N_VALUES or not is_consistent(values[:N_VALUES]):
        return None

    label = label_subtotal(prefix, state) or (re.sub(r"\s+", " ", prefix).strip()[:40] or "不明な合計行")
    return label, values[:N_VALUES]


def subtotal_sort_key(label: str) -> int:
    cat = next((i for i, c in enumerate(CATEGORY_ORDER) if label.startswith(c)), len(CATEGORY_ORDER))
    seg = next((i for i, s in enumerate(SEGMENT_ORDER) if s in label), len(SEGMENT_ORDER))
    return cat * 10 + seg


def parse_margin_pdf(pdf_bytes: bytes) -> pd.DataFrame:
    """
    日次PDF → DataFrame
    列: 銘柄コード, 銘柄名, 一般信用買残高, 一般信用売残高, 制度信用買残高, 制度信用売残高
    小計・合計行は銘柄コード "SUB_<ラベル>" として先頭に置く。
    """
    records: dict[str, dict] = {}                 # 銘柄コード → レコード
    subtotals: dict[str, list[float]] = {}
    failures: list[tuple[str, str]] = []
    state = {"category": None, "passed_total": False}
    n_line, n_word = 0, 0
    sample_lines: list[str] = []                  # 読めなかったときの調査用

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pages = len(pdf.pages)
        for page in pdf.pages:
            # ---- 方式A: 1行単位 ＋ 合計行 ----
            for line in (page.extract_text() or "").split("\n"):
                record, reason = parse_stock_line(line)
                if record:
                    records.setdefault(record["銘柄コード"], record)
                    n_line += 1
                elif reason:
                    failures.append((reason, line))
                elif sub := parse_subtotal_line(line, state):
                    label, values = sub
                    if label in subtotals:
                        print(f"  警告: 合計行 '{label}' が重複。後の値で上書きします")
                    subtotals[label] = values
                elif len(sample_lines) < 8 and ("株数" in line or ISIN_PATTERN.search(line)):
                    sample_lines.append(line)

            # ---- 方式B: 座標で組み立て（方式Aで読めなかった銘柄を補う）----
            word_records, word_failures = parse_stock_words(page)
            for record in word_records:
                if record["銘柄コード"] not in records:
                    records[record["銘柄コード"]] = record
                    n_word += 1
            failures += word_failures

    # どちらかの方式で読めた銘柄の失敗は数えない
    failures = [(reason, text) for reason, text in failures if _failure_code(text) not in records]

    # ---- 結果のチェック ----
    print(
        f"  {n_pages}ページ / {len(records)}銘柄（1行方式 {n_line}・座標方式 {n_word}）"
        f" / 失敗 {len(failures)}件 / 合計行 {len(subtotals)}件"
    )
    for reason, line in failures[:5]:
        print(f"    失敗例（{reason}）: {line[:120]!r}")
    if not records:
        print("  読めなかった行の例:")
        for line in sample_lines:
            print(f"    {line[:160]!r}")
        raise ValueError("銘柄データを1件も読めませんでした。PDFのレイアウトが変わった可能性があります。")
    if len(failures) > max(MAX_PARSE_FAILURES, len(records) * MAX_PARSE_FAILURE_RATIO):
        raise ValueError(f"読み取り失敗が多すぎます（{len(failures)}行）。PDFのレイアウトが変わった可能性があります。")

    # ---- DataFrame にまとめる（合計行 → 銘柄コード順）----
    stocks = pd.DataFrame(list(records.values()))
    stocks["_sort"] = "1_" + stocks["銘柄コード"]

    rows = [
        {
            "銘柄コード": SUBTOTAL_CODE_PREFIX + label.replace(" ", ""),
            "銘柄名": label,
            **pick_balances(values),
            "_sort": f"0_{subtotal_sort_key(label):03d}",
        }
        for label, values in subtotals.items()
    ]
    df = pd.concat([pd.DataFrame(rows), stocks], ignore_index=True) if rows else stocks
    return df.sort_values("_sort").drop(columns="_sort").reset_index(drop=True)


def inspect_margin_pdf(yyyymmdd: str) -> None:
    """--inspect 用。PDFの先頭・末尾ページと読み取り結果を表示する（書き込みなし）。"""
    sess = http_session()
    try:
        url = list_jpx_files(sess).get(yyyymmdd)
    except Exception as e:
        print(f"掲載一覧を取得できませんでした: {e}")
        url = None
    pdf_bytes = download_jpx_pdf(sess, url or JPX_PDF_URL_TEMPLATE.format(date=yyyymmdd))

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for title, page in (("先頭ページ", pdf.pages[0]), ("末尾ページ", pdf.pages[-1])):
            print(f"\n===== {title} =====")
            print((page.extract_text() or "")[:4000])

    df = parse_margin_pdf(pdf_bytes)
    with pd.option_context("display.max_columns", None, "display.width", 200):
        print("\n===== 合計行 =====")
        print(df[df["銘柄コード"].str.startswith(SUBTOTAL_CODE_PREFIX)])
        print("\n===== 銘柄（先頭15件）=====")
        print(df[~df["銘柄コード"].str.startswith(SUBTOTAL_CODE_PREFIX)].head(15))


# =============================================================================
# 日証協: 貸株残の取得と読み取り
# =============================================================================

def list_jsda_dates(sess: requests.Session) -> list[str]:
    """日証協ページに載っている申込日を新しい順で返す。"""
    dates = sorted(set(JSDA_LINK_PATTERN.findall(fetch_text(sess, JSDA_INDEX_URL))), reverse=True)
    if not dates:
        raise ValueError("日証協のページにファイルリンクがありません。ページ構成が変わった可能性があります。")
    return dates


def normalize_stock_code(value) -> str | None:
    """'1301' / '1301.0' / '13010' / '402A' → 4桁コード"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = re.sub(r"\.0$", "", str(value).strip().upper())
    if re.fullmatch(r"[0-9A-Z]{4}0", s):
        return s[:4]
    if re.fullmatch(r"[0-9A-Z]{4}", s):
        return s
    return None


def find_column(columns: list[str], keywords: list[str], exclude: list[str] | None = None) -> str | None:
    for keyword in keywords:
        for col in columns:
            if exclude and any(ex in col for ex in exclude):
                continue
            if keyword in col:
                return col
    return None


def parse_kashikabu_xlsx(xlsx_bytes: bytes) -> pd.DataFrame:
    """xlsx → DataFrame（列: 銘柄コード, 銘柄名, 貸株残）"""
    raw = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0, header=None)

    header_row = next(
        (i for i in range(min(25, len(raw)))
         if any(kw in str(v) for v in raw.iloc[i] for kw in JSDA_CODE_KEYWORDS)),
        None,
    )
    if header_row is None:
        raise ValueError("ヘッダー行が見つかりません。書式が変わった可能性があります。")

    columns = [re.sub(r"[\s　]", "", str(c)) for c in raw.iloc[header_row]]
    body = raw.iloc[header_row + 1:].copy()
    body.columns = columns

    code_col = find_column(columns, JSDA_CODE_KEYWORDS)
    name_col = find_column(columns, JSDA_NAME_KEYWORDS)
    value_col = find_column(columns, JSDA_BALANCE_KEYWORDS, exclude=JSDA_BALANCE_EXCLUDE)
    if code_col is None or value_col is None:
        raise ValueError(f"必要な列が見つかりません。検出列: {columns}")
    print(f"  使用列: コード='{code_col}' / 残高='{value_col}' / 銘柄名='{name_col}'")

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
            KASHIKABU_VALUE_COL: float(value),
        })

    if not records:
        raise ValueError("貸株残データを読めませんでした。書式が変わった可能性があります。")
    print(f"  {len(records)}銘柄")

    df = pd.DataFrame(records).groupby("銘柄コード", as_index=False).agg(
        {"銘柄名": "first", KASHIKABU_VALUE_COL: "sum"}
    )
    return df.sort_values("銘柄コード").reset_index(drop=True)


# =============================================================================
# Google Sheets: 接続とユーティリティ
# =============================================================================

def connect_google_sheets() -> gspread.Client:
    """サービスアカウントで認証する。429（レート制限）は自動リトライ。"""
    raw = os.environ[GOOGLE_SERVICE_ACCOUNT_JSON_ENV]
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
    """行数・列数が足りなければ広げる（gspread 6.x では手元の ws を直接広げる）。"""
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
# Google Sheets: 1日分の列を書き込む
# =============================================================================

def write_date_column(
    gc: gspread.Client,
    spreadsheet_id: str,
    sheet_name: str,
    df: pd.DataFrame,
    value_col: str,
    yyyymmdd: str,
    max_date_cols: int = 0,
) -> None:
    ws = open_worksheet(gc, spreadsheet_id, sheet_name)
    existing = ws.get_all_values()
    label = to_japanese_date(yyyymmdd)
    tag = f"  [{sheet_name}]"

    # ---- 1. シートが空 → 初回書き込み ----
    if not existing or not any(c.strip() for c in existing[0]):
        print(f"{tag} 初回書き込み")
        rows = [["銘柄コード", "銘柄名", label]]
        rows += [[r["銘柄コード"], r["銘柄名"], to_cell(r[value_col])] for _, r in df.iterrows()]
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
    lookup = dict(zip(df["銘柄コード"].astype(str), df[value_col]))
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
            row[0], row[1], row[col - 1] = str(r["銘柄コード"]), r["銘柄名"], to_cell(r[value_col])
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

    # ---- 5. 日付列が上限を超えたら古い列を削除 ----
    n_dates = len([c for c in header[FIXED_COLS:] if c.strip()]) + 1
    if 0 < max_date_cols < n_dates:
        first, last = FIXED_COLS + max_date_cols + 1, FIXED_COLS + n_dates
        print(f"{tag} 古い日付列を削除: {last - first + 1}列（上限 {max_date_cols}）")
        ws.delete_columns(first, last)

    print(f"{tag} 完了")


# =============================================================================
# ジョブ1: 信用残（JPX 日次）
# =============================================================================

def margin_done_dates(gc: gspread.Client) -> set[str]:
    """4シートすべてに入っている日付（途中で失敗した日は未完了として扱う）"""
    per_sheet = [
        dates_in_header(open_worksheet(gc, SPREADSHEET_ID, name).row_values(1))
        for name in MARGIN_SHEETS.values()
    ]
    return set.intersection(*per_sheet)


def update_margin(gc: gspread.Client) -> bool:
    """戻り値: 問題なければ True（False のとき main は終了コード1で終える）"""
    print_title("信用取引残高（JPX 日次）")
    sess = http_session()
    ok = True

    # ---- 対象日を決める ----
    try:
        listed = list_jpx_files(sess)
        print(f"  掲載中: {sorted(listed, reverse=True)}")
    except Exception as e:
        print(f"  ✗ 掲載一覧の取得に失敗: {e}")
        listed, ok = {}, False

    extra = parse_date_list(MARGIN_EXTRA_DATES) + parse_date_list(TARGET_DATE_OVERRIDE)
    done = margin_done_dates(gc)
    print(f"  シートの最新: {max((d for d in done if is_yyyymmdd(d)), default='なし')}")

    targets = sorted((d for d in listed if d not in done), reverse=True)[:MARGIN_BACKFILL_N]
    targets = sorted(set(targets) | {d for d in extra if d not in done})
    if not targets:
        print("  新しいデータはありません")
        return ok
    print(f"  取得対象: {targets}")

    # ---- 1日ずつ取得して4シートへ ----
    for i, yyyymmdd in enumerate(targets):
        if i:
            time.sleep(DOWNLOAD_INTERVAL_JPX)
        print(f"\n[{yyyymmdd}] {to_japanese_date(yyyymmdd)}")
        url = listed.get(yyyymmdd) or JPX_PDF_URL_TEMPLATE.format(date=yyyymmdd)
        try:
            df = parse_margin_pdf(download_jpx_pdf(sess, url))
        except FileNotFoundError as e:
            print(f"  ✗ {e}")
            ok = ok and yyyymmdd not in listed   # 掲載外の過去日が消えているのは想定内
            continue
        except Exception as e:
            print(f"  ✗ 取得/読み取りに失敗: {e}")
            ok = False
            continue

        for value_col, sheet_name in MARGIN_SHEETS.items():
            write_date_column(gc, SPREADSHEET_ID, sheet_name, df, value_col, yyyymmdd, MAX_DATE_COLS)

    return ok


# =============================================================================
# ジョブ2: 貸株残（日証協 週次）
# =============================================================================

def update_kashikabu(gc: gspread.Client) -> None:
    print_title("貸株残（日本証券業協会）")
    sess = http_session()

    available = list_jsda_dates(sess)
    print(f"  掲載中: {available[:5]}{' ...' if len(available) > 5 else ''}")

    ws = open_worksheet(gc, KASHIKABU_SPREADSHEET_ID, KASHIKABU_SHEET_NAME)
    done = dates_in_header(ws.row_values(1))
    targets = sorted(sorted((d for d in available if d not in done), reverse=True)[:JSDA_BACKFILL_N])
    if not targets:
        print("  新しいデータはありません")
        return
    print(f"  取得対象: {targets}")

    for i, yyyymmdd in enumerate(targets):
        if i:
            time.sleep(DOWNLOAD_INTERVAL_JSDA)
        print(f"\n[{yyyymmdd}] {to_japanese_date(yyyymmdd)}")
        try:
            df = parse_kashikabu_xlsx(fetch_file(sess, JSDA_FILE_URL_TEMPLATE.format(date=yyyymmdd)))
        except Exception as e:
            print(f"  ✗ 取得/読み取りに失敗: {e}")
            continue
        write_date_column(gc, KASHIKABU_SPREADSHEET_ID, KASHIKABU_SHEET_NAME, df, KASHIKABU_VALUE_COL, yyyymmdd)


# =============================================================================
# メイン
# =============================================================================

def main() -> None:
    if len(sys.argv) >= 3 and sys.argv[1] == "--inspect":
        inspect_margin_pdf(sys.argv[2])
        return

    if not SPREADSHEET_ID:
        sys.exit("エラー: 環境変数 SPREADSHEET_ID が設定されていません")

    gc = connect_google_sheets()

    try:
        margin_ok = update_margin(gc)
    except Exception as e:
        print(f"✗ 信用残の処理で予期しないエラー: {e}")
        margin_ok = False

    if JSDA_ENABLED:
        try:
            update_kashikabu(gc)          # 信用残の成否に関係なく実行
        except Exception as e:
            print(f"警告: 貸株残の処理に失敗しました: {e}")
    else:
        print("\n貸株残はスキップ（JSDA_ENABLED=0）")

    if not margin_ok:
        sys.exit("\n信用残の取得に失敗した日があります。ログを確認してください。")


if __name__ == "__main__":
    main()
