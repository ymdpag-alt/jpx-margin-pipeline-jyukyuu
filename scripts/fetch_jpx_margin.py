"""
JPX「銘柄別信用取引残高」（2026/9/28〜 日次公表・全銘柄）PDFを取得し、
一般信用買残高・一般信用売残高・制度信用買残高・制度信用売残高を
それぞれ別シートにGoogle Spreadsheetへ書き込む。

あわせて、日本証券業協会「銘柄別株券等貸借週末残高」(xlsx)から
貸株残を取得し、別スプレッドシートの「貸株残」シートへ同じ形式で書き込む。

■ 2026/9/28 の JPX 公表方式変更への対応（2026/10 改修）
  旧: 「銘柄別信用取引週末残高」 syumatsu{金曜日}00.pdf を毎週火曜に公表
      → 2026/9/18申込分（9/25公表）で終了
  新: 「銘柄別信用取引残高」 {申込日YYYYMMDD}_mtall.pdf を毎営業日16:00頃に公表
      （前営業日分。掲載ページ: margin/01.html。直近5営業日分のみ掲載）

  変更点:
    - 対象日は実行日から逆算せず、01.html に掲載されているリンクから
      「まだシートに無い日付」を拾う（祝日・公表遅延・再実行に強い）。
    - PDFの1銘柄が「株数 Shs.」行と「金額 Val.」行の2行構成になったため、
      「株数」行だけをパースする。
    - 数値列に「上場比」が加わったため、列位置を新レイアウトに合わせた。
    - 合計＝一般＋制度 の整合性チェックで、パースずれを行単位で検出する。
    - 公表終了・ファイル未検出・パース失敗は終了コード1で失敗させる
      （旧版は 404 を正常終了扱いにしていたため、停止に気づけなかった）。

■ シート構成
  A列 = 銘柄コード
  B列 = 銘柄名
  C列以降 = 申込日（日本語形式）。左ほど新しい。

■ 書き込みロジック
  新しい日付の列は「日付の降順が保たれる位置」に挿入する。
  通常は最新日なのでC列に入り、過去データは右へシフトする。
  過去日をあとから埋めた場合（バックフィル）も、正しい位置に差し込まれる。
  既存銘柄は該当行に値を入れ、シートに無い銘柄（新規上場など）は最下部に追加する。
  同じ日付の列が既にあればスキップする（二重書き込み防止）。

  MAX_DATE_COLS を指定すると、日付列がその数を超えた分を古い順（右端）から削除する。
  （日次化でセル数が年約430万増えるため。スプレッドシートの上限は1ブック1,000万セル）

■ 貸株残について
  取得元: 日本証券業協会「株券等貸借取引状況（週間）」の銘柄別週末残高
          https://www.jsda.or.jp/shiryoshitsu/toukei/kabu-taiw/index.html
  公表タイミング: 報告週の翌週木曜。週次のまま。
  毎日実行しても、新しいファイルが無い日は自動的にスキップする。

■ 環境変数
  GOOGLE_SERVICE_ACCOUNT_JSON  サービスアカウントJSON（必須）
  SPREADSHEET_ID               信用残の書き込み先（必須）
  KASHIKABU_SPREADSHEET_ID     貸株残の書き込み先
  MARGIN_BACKFILL_N            01.html掲載分のうち、未取得を新しい順に最大何日分入れるか（既定 5）
  MARGIN_EXTRA_DATES           掲載から外れた過去日を直接URLで取りに行く（例: 20260925,20260928）
  TARGET_DATE_OVERRIDE         1日分だけ強制的に取得したい場合（MARGIN_EXTRA_DATES と同じ扱い）
  MAX_DATE_COLS                信用残シートで残す日付列の上限（既定 0 = 削除しない）
  JSDA_BACKFILL_N              貸株残の遡り週数（既定 1）
  JSDA_ENABLED                 "0" で貸株残をスキップ

■ 調査用
  python margin_balance_to_sheets.py --inspect 20261005
    → PDFの先頭・末尾ページのテキストとパース結果の統計を表示（シートには書かない）

■ gspread 6.x 対応
  Spreadsheet.client は HTTPClient を返すため、Worksheet から Client を辿らない。
  行数・列数の拡張は手元の Worksheet に対して直接行う。
  429（レート制限）は BackOffHTTPClient で自動リトライする。
"""

from __future__ import annotations

import io
import json
import os
import re
import sys
from datetime import datetime
from urllib.parse import urljoin

import gspread
import pandas as pd
import pdfplumber
import requests
from google.oauth2.service_account import Credentials

# =============================================================================
# 定数定義
# =============================================================================

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")

# 4種類の内訳データをそれぞれ別シートに書き込む（合計はユーザー側の別シートで算出）
SHEET_NAMES = {
    "一般信用買残高": "一般信用買残高",
    "一般信用売残高": "一般信用売残高",
    "制度信用買残高": "制度信用買残高",
    "制度信用売残高": "制度信用売残高",
}

JPX_LIST_URL = "https://www.jpx.co.jp/markets/statistics-equities/margin/01.html"
JPX_LINK_PATTERN = re.compile(r'href="([^"]*?(\d{8})_mtall\.pdf)"')
# 掲載ページから外れた日付を直接取りに行くときのURL（添付フォルダ名が変わったらここを直す）
JPX_PDF_URL_TEMPLATE = (
    "https://www.jpx.co.jp/markets/statistics-equities/"
    "margin/tvdivq0000001rnl-att/{date}_mtall.pdf"
)

# シート構成: A列=銘柄コード, B列=銘柄名, C列以降=申込日
FIXED_COLS = 2
FIRST_DATE_COL = FIXED_COLS + 1  # = 3 (C列)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

_WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]

MARGIN_BACKFILL_N = int(os.environ.get("MARGIN_BACKFILL_N", "5"))
MAX_DATE_COLS = int(os.environ.get("MAX_DATE_COLS", "0"))

# -----------------------------------------------------------------------------
# PDFパース用の正規表現（2026/10/5申込分のPDFで実データ確認済み）
# -----------------------------------------------------------------------------
# 1銘柄の「株数」行の並び:
#   [単位フラグ] 銘柄名 株式種別 市場 貸借区分 5桁コード ISIN 株数 Shs.
#   売残高 前日比 上場比 買残高 前日比 上場比
#   売・一般 前日比 売・制度 前日比 買・一般 前日比 買・制度 前日比
# 例:
#   B 極洋　普通株式 プライム 貸 13010 JP3257200000 株数 Shs. 9,500 100 0.1% 161,300 1,200 1.3% ...
#   A ＳＰＤＲゴールド・シェア　受益証券投信等 貸 13260 US78463V1070 株数 Shs. 67 ▲ 1 * 7,275 ...
STOCK_LINE_PATTERN = re.compile(
    r"^(?P<prefix>.*?)\s*(?P<code>[0-9A-Z]{5})\s+"
    r"(?P<isin>[A-Z]{2}[A-Z0-9]{9}[0-9])\s+株数\s*(?:Shs\.?)?\s*(?P<body>.*)$"
)
# 数値トークン: 「▲」付き（減少）、上場比の「0.1%」、ETF等の上場比「*」
VALUE_TOKEN_PATTERN = re.compile(r"([▲△])?\s*(\d[\d,]*\.\d+%|\d[\d,]*%|\d[\d,]*|\*)")

N_VALUES = 14
IDX_SELL_TOTAL, IDX_BUY_TOTAL = 0, 3
IDX_GENERAL_SELL, IDX_STANDARD_SELL = 6, 8
IDX_GENERAL_BUY, IDX_STANDARD_BUY = 10, 12

UNIT_FLAG_PATTERN = re.compile(r"^[AJKBMCTF][\s\u3000]+")
LOAN_FLAG_PATTERN = re.compile(r"[\s\u3000]*(貸|制|他)$")
MARKET_PATTERN = re.compile(
    r"[\s\u3000]*(プライム|スタンダード|グロース|投信等|ＰＲＯ|PRO|TOKYO PRO Market)$"
)
SECURITY_TYPE_PATTERN = re.compile(
    r"[\s\u3000]*(普通株式|受益証券|投資証券|投資口|出資証券|優先株式|優先出資証券|"
    r"新株予約権証券|外国株預託証券|株式)$"
)

# 小計・総合計行の判定用キーワード（区分数が増減しても対応できるよう、行の中身から動的に判定）
SEGMENT_KEYWORDS = ["プライム", "スタンダード", "グロース", "投信等"]
CATEGORY_ORDER = ["貸借銘柄", "制度信用銘柄", "その他", "総合計", "全体"]
SEGMENT_ORDER = ["合計", "プライム", "スタンダード", "グロース", "投信等"]

# -----------------------------------------------------------------------------
# 貸株残（日本証券業協会）関連の定数
# -----------------------------------------------------------------------------

KASHIKABU_SPREADSHEET_ID = os.environ.get(
    "KASHIKABU_SPREADSHEET_ID", "1kWST0CkkIvo3irPSbMgtVtUqqRDXFwvRREYZDRQAFMY"
)
KASHIKABU_SHEET_NAME = "貸株残"
KASHIKABU_VALUE_COL = "貸株残"

JSDA_INDEX_URL = "https://www.jsda.or.jp/shiryoshitsu/toukei/kabu-taiw/index.html"
JSDA_FILE_URL_TEMPLATE = (
    "https://www.jsda.or.jp/shiryoshitsu/toukei/kabu-taiw/files/{date}z.xlsx"
)
JSDA_LINK_PATTERN = re.compile(r"files/(\d{8})z\.xlsx")

JSDA_CODE_KEYWORDS = ["銘柄コード", "コード", "code"]
JSDA_NAME_KEYWORDS = ["銘柄名", "名称", "銘柄"]
JSDA_BALANCE_KEYWORDS = ["貸付残高", "貸付", "貸株", "週末残高", "残高"]
JSDA_BALANCE_EXCLUDE = ["借入", "返済", "新規", "成約"]

JSDA_BACKFILL_N = int(os.environ.get("JSDA_BACKFILL_N", "1"))
JSDA_ENABLED = os.environ.get("JSDA_ENABLED", "1") != "0"

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en;q=0.8",
}


def _http_session() -> requests.Session:
    sess = requests.Session()
    sess.headers.update(HTTP_HEADERS)
    return sess


# =============================================================================
# 日付ユーティリティ
# =============================================================================

def to_japanese_date(date_str_yyyymmdd: str) -> str:
    """'20261005' -> '2026年10月5日(月)'"""
    dt = datetime.strptime(date_str_yyyymmdd, "%Y%m%d")
    weekday = _WEEKDAY_JP[dt.weekday()]
    return f"{dt.year}年{dt.month}月{dt.day}日({weekday})"


def from_japanese_date(date_str: str) -> str:
    """'2026年10月5日(月)' -> '20261005'。解釈できなければ元の文字列を返す。"""
    try:
        core = date_str.split("(")[0].strip()
        dt = datetime.strptime(core, "%Y年%m月%d日")
        return dt.strftime("%Y%m%d")
    except (ValueError, IndexError):
        return date_str


def _is_yyyymmdd(s: str) -> bool:
    return bool(re.fullmatch(r"\d{8}", s or ""))


def _parse_date_list(raw: str) -> list[str]:
    dates = [d.strip() for d in re.split(r"[,\s]+", raw or "") if d.strip()]
    bad = [d for d in dates if not _is_yyyymmdd(d)]
    if bad:
        raise ValueError(f"日付は YYYYMMDD で指定してください: {bad}")
    return dates


# =============================================================================
# JPX: 掲載ファイル一覧・PDF取得
# =============================================================================

def list_jpx_margin_files(sess: requests.Session) -> dict[str, str]:
    """
    01.html に掲載されている日次PDFを {申込日YYYYMMDD: URL} で返す。
    1件も見つからない場合はページ構成変更とみなして例外を投げる。
    """
    resp = sess.get(JPX_LIST_URL, timeout=30)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    found = {date: urljoin(JPX_LIST_URL, href) for href, date in JPX_LINK_PATTERN.findall(resp.text)}
    if not found:
        raise ValueError(
            "01.html から *_mtall.pdf のリンクを検出できませんでした。"
            "ページ構成かファイル名の規則が変わった可能性があります。"
        )
    return found


def download_margin_pdf(url: str, sess: requests.Session) -> bytes:
    print(f"  ダウンロード中: {url}")
    resp = sess.get(url, timeout=60)
    if resp.status_code == 404:
        raise FileNotFoundError(f"PDFが見つかりません: {url}")
    resp.raise_for_status()
    if not resp.content.startswith(b"%PDF"):
        raise ValueError(f"PDFではない応答が返りました: {url}")
    return resp.content


# =============================================================================
# PDFパース
# =============================================================================

def _parse_values(body: str) -> list[float | None]:
    """
    「株数 Shs.」以降の文字列から数値を順に取り出す。
    上場比（%付き・*）は位置合わせのため None として1個分数える。
    """
    values: list[float | None] = []
    for sign, token in VALUE_TOKEN_PATTERN.findall(body):
        if token == "*" or token.endswith("%"):
            values.append(None)
            continue
        num = float(token.replace(",", ""))
        values.append(-num if sign else num)
    return values


def _check_consistency(v: list[float | None]) -> bool:
    """売残高合計＝一般＋制度、買残高合計＝一般＋制度 になっているか。"""
    try:
        return (
            v[IDX_SELL_TOTAL] == v[IDX_GENERAL_SELL] + v[IDX_STANDARD_SELL]
            and v[IDX_BUY_TOTAL] == v[IDX_GENERAL_BUY] + v[IDX_STANDARD_BUY]
        )
    except TypeError:
        return False


def _clean_name(prefix: str) -> str:
    name = prefix.strip()
    name = UNIT_FLAG_PATTERN.sub("", name)
    name = LOAN_FLAG_PATTERN.sub("", name)
    name = MARKET_PATTERN.sub("", name)
    name = SECURITY_TYPE_PATTERN.sub("", name)
    return name.strip(" \u3000")


def parse_stock_line(line: str) -> tuple[dict | None, str | None]:
    """
    銘柄の「株数」行を1行パースする。
    戻り値: (レコード or None, 失敗理由 or None)。銘柄行でなければ (None, None)。
    """
    m = STOCK_LINE_PATTERN.match(line.strip())
    if not m:
        return None, None

    values = _parse_values(m.group("body"))
    if len(values) < N_VALUES:
        return None, f"数値が{len(values)}個しかありません"
    values = values[:N_VALUES]
    if not _check_consistency(values):
        return None, "合計≠一般＋制度（列ずれの可能性）"

    code5 = m.group("code")
    return {
        "銘柄コード": code5[:4],  # 末尾の付番(通常は0)を除いた4桁コード
        "銘柄名": _clean_name(m.group("prefix")),
        "一般信用買残高": values[IDX_GENERAL_BUY],
        "一般信用売残高": values[IDX_GENERAL_SELL],
        "制度信用買残高": values[IDX_STANDARD_BUY],
        "制度信用売残高": values[IDX_STANDARD_SELL],
    }, None


def _label_subtotal_line(line: str, state: dict) -> str | None:
    """
    小計・総合計行のラベル（例: "貸借銘柄 プライム 小計"）を、
    行の中身のキーワードとカテゴリの追跡状態から動的に判定する。
    """
    if "総合計" in line:
        state["passed_total"] = True
        return "総合計"
    if "貸借銘柄" in line:
        state["current_category"] = "貸借銘柄"
        if "小計" not in line:
            return "貸借銘柄 合計"
    elif "制度信用銘柄" in line:
        state["current_category"] = "制度信用銘柄"
        if "小計" not in line:
            return "制度信用銘柄 合計"
    elif re.search(r"その他", line) and "小計" not in line:
        state["current_category"] = "その他"
        return "その他 合計"

    for seg in SEGMENT_KEYWORDS:
        if seg in line and "小計" in line:
            category = "全体" if state["passed_total"] else state["current_category"]
            if category is None:
                return None
            return f"{category} {seg} 小計"
    return None


def parse_subtotal_line(line: str, state: dict) -> tuple[str, list[float]] | None:
    """
    小計・合計行（ISINを含まず「株数」と「計」を含む行）をパースする。
    ラベルが判定できない場合は行頭の文字列をそのままラベルにする。
    """
    if "株数" not in line or "計" not in line:
        return None
    prefix, _, body = line.partition("株数")
    body = re.sub(r"^\s*Shs\.?", "", body)
    values = _parse_values(body)
    if len(values) < N_VALUES:
        return None
    values = values[:N_VALUES]
    if not _check_consistency(values):
        return None
    label = _label_subtotal_line(prefix, state)
    if label is None:
        label = re.sub(r"\s+", " ", prefix).strip()[:40] or "不明な合計行"
    return label, values


def _subtotal_sort_key(label: str) -> int:
    cat_idx = next((i for i, c in enumerate(CATEGORY_ORDER) if label.startswith(c)), len(CATEGORY_ORDER))
    seg_idx = next((i for i, s in enumerate(SEGMENT_ORDER) if s in label), len(SEGMENT_ORDER))
    return cat_idx * 10 + seg_idx


def parse_margin_pdf(pdf_bytes: bytes, verbose: bool = True) -> pd.DataFrame:
    """
    日次PDFをパースし、DataFrameで返す。
    列: 銘柄コード, 銘柄名, 一般信用買残高, 一般信用売残高, 制度信用買残高, 制度信用売残高
    小計・総合計行が見つかれば "SUB_<ラベル>" という特別なコードの行として先頭に含める。
    """
    records: list[dict] = []
    subtotal_dict: dict[str, list[float]] = {}
    failures: list[tuple[str, str]] = []
    label_state = {"current_category": None, "passed_total": False}

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pages = len(pdf.pages)
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                rec, reason = parse_stock_line(line)
                if rec:
                    records.append(rec)
                    continue
                if reason:
                    failures.append((reason, line))
                    continue
                sub = parse_subtotal_line(line, label_state)
                if sub:
                    label, values = sub
                    if label in subtotal_dict and verbose:
                        print(f"  警告: 合計行ラベル '{label}' が重複しました。後の値で上書きします。")
                    subtotal_dict[label] = values

    stock_df = pd.DataFrame(records)
    if verbose:
        print(f"  総ページ数: {n_pages} / 抽出: {len(stock_df)} 銘柄 / パース失敗: {len(failures)} 行")
        for reason, line in failures[:5]:
            print(f"    失敗例（{reason}）: {line[:120]!r}")
    if stock_df.empty:
        raise ValueError("PDFから銘柄データを抽出できませんでした。レイアウトが変わった可能性があります。")
    if len(failures) > max(20, len(stock_df) * 0.01):
        raise ValueError(
            f"パース失敗が多すぎます（{len(failures)} 行）。レイアウトが変わった可能性があります。"
        )

    stock_df = stock_df.drop_duplicates(subset="銘柄コード", keep="first")
    # JPXの並び（数字<英字の文字列順）に揃える。集計行より必ず後ろにする。
    stock_df["_sort_key"] = "1_" + stock_df["銘柄コード"]

    if verbose:
        print(f"  合計・小計行の検出数: {len(subtotal_dict)} 件")
    if subtotal_dict:
        summary_df = pd.DataFrame(
            [
                {
                    "銘柄コード": "SUB_" + label.replace(" ", ""),
                    "銘柄名": label,
                    "一般信用買残高": v[IDX_GENERAL_BUY],
                    "一般信用売残高": v[IDX_GENERAL_SELL],
                    "制度信用買残高": v[IDX_STANDARD_BUY],
                    "制度信用売残高": v[IDX_STANDARD_SELL],
                    "_sort_key": f"0_{_subtotal_sort_key(label):03d}",
                }
                for label, v in subtotal_dict.items()
            ]
        )
        df = pd.concat([summary_df, stock_df], ignore_index=True)
    else:
        df = stock_df

    return df.sort_values("_sort_key").drop(columns="_sort_key").reset_index(drop=True)


def inspect_margin_pdf(date_yyyymmdd: str) -> None:
    """レイアウト確認用。先頭・末尾ページのテキストとパース結果を表示する（書き込みなし）。"""
    sess = _http_session()
    try:
        url = list_jpx_margin_files(sess).get(date_yyyymmdd)
    except Exception as e:
        print(f"掲載一覧の取得に失敗: {e}")
        url = None
    url = url or JPX_PDF_URL_TEMPLATE.format(date=date_yyyymmdd)
    pdf_bytes = download_margin_pdf(url, sess)
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for label, page in (("先頭ページ", pdf.pages[0]), ("末尾ページ", pdf.pages[-1])):
            print(f"===== {label} =====")
            print((page.extract_text() or "")[:4000])
    df = parse_margin_pdf(pdf_bytes)
    with pd.option_context("display.max_columns", None, "display.width", 200):
        print(df.head(15))
        print(df[df["銘柄コード"].str.startswith("SUB_")])


# =============================================================================
# 貸株残（日本証券業協会）取得・パース
# =============================================================================

def list_jsda_dates(sess: requests.Session | None = None) -> list[str]:
    """日証協のindexページから銘柄別週末残高ファイルの申込日(YYYYMMDD)を新しい順に返す。"""
    sess = sess or _http_session()
    resp = sess.get(JSDA_INDEX_URL, timeout=30)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    dates = sorted(set(JSDA_LINK_PATTERN.findall(resp.text)), reverse=True)
    if not dates:
        raise ValueError(
            "日証協のページからファイルリンクを検出できませんでした。ページ構成が変わった可能性があります。"
        )
    return dates


def download_jsda_xlsx(date_yyyymmdd: str, sess: requests.Session | None = None) -> bytes:
    sess = sess or _http_session()
    url = JSDA_FILE_URL_TEMPLATE.format(date=date_yyyymmdd)
    print(f"  ダウンロード中: {url}")
    resp = sess.get(url, timeout=60)
    if resp.status_code == 404:
        raise FileNotFoundError(f"貸株残ファイルが見つかりません: {url}")
    resp.raise_for_status()
    return resp.content


def _normalize_stock_code(value) -> str | None:
    """'1301' / '1301.0' / '13010' / '402A' などの表記ゆれを4桁形式に寄せる。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip().upper()
    s = re.sub(r"\.0$", "", s)
    if not s:
        return None
    if len(s) == 5 and re.fullmatch(r"[0-9A-Z]{4}0", s):
        return s[:4]
    if re.fullmatch(r"[0-9A-Z]{4}", s):
        return s
    return None


def _pick_jsda_column(columns: list[str], keywords: list[str],
                      exclude: list[str] | None = None) -> str | None:
    for kw in keywords:
        for col in columns:
            text = str(col)
            if exclude and any(ex in text for ex in exclude):
                continue
            if kw in text:
                return col
    return None


def _find_jsda_header_row(raw: pd.DataFrame, max_scan: int = 25) -> int | None:
    for i in range(min(max_scan, len(raw))):
        row = [str(v) for v in raw.iloc[i].tolist()]
        if any(any(kw in v for kw in JSDA_CODE_KEYWORDS) for v in row):
            return i
    return None


def parse_kashikabu_xlsx(xlsx_bytes: bytes) -> pd.DataFrame:
    """日証協の銘柄別株券等貸借週末残高 xlsx をパースする。列: 銘柄コード, 銘柄名, 貸株残"""
    raw = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0, header=None)
    header_row = _find_jsda_header_row(raw)
    if header_row is None:
        raise ValueError("ヘッダー行を特定できませんでした。inspect_jsda_latest() で中身を確認してください。")

    columns = [
        str(c).replace("\n", "").replace(" ", "").replace("\u3000", "").strip()
        for c in raw.iloc[header_row].tolist()
    ]
    df = raw.iloc[header_row + 1:].copy()
    df.columns = columns

    code_col = _pick_jsda_column(columns, JSDA_CODE_KEYWORDS)
    bal_col = _pick_jsda_column(columns, JSDA_BALANCE_KEYWORDS, exclude=JSDA_BALANCE_EXCLUDE)
    name_col = _pick_jsda_column(columns, JSDA_NAME_KEYWORDS)
    if code_col is None or bal_col is None:
        raise ValueError(f"必要な列を特定できませんでした。検出列: {columns}")
    print(f"  使用列: コード='{code_col}' / 残高='{bal_col}' / 銘柄名='{name_col}'")

    records = []
    skipped = 0
    for _, row in df.iterrows():
        code = _normalize_stock_code(row[code_col])
        if code is None:
            skipped += 1
            continue
        val = pd.to_numeric(str(row[bal_col]).replace(",", "").strip(), errors="coerce")
        if pd.isna(val):
            skipped += 1
            continue
        name = str(row[name_col]).strip() if name_col else ""
        if name in ("nan", "None"):
            name = ""
        records.append({"銘柄コード": code, "銘柄名": name, KASHIKABU_VALUE_COL: float(val)})

    out = pd.DataFrame(records)
    print(f"  抽出件数: {len(out)} 銘柄（スキップ: {skipped} 行）")
    if out.empty:
        raise ValueError("貸株残データを抽出できませんでした。書式が変わった可能性があります。")

    out = out.groupby("銘柄コード", as_index=False).agg({"銘柄名": "first", KASHIKABU_VALUE_COL: "sum"})
    return out.sort_values("銘柄コード").reset_index(drop=True)


def inspect_jsda_latest(n_rows: int = 15) -> None:
    """列名がずれたときの調査用。最新ファイルの先頭数行をそのまま表示する。"""
    sess = _http_session()
    dates = list_jsda_dates(sess)
    print(f"掲載されている申込日(新しい順): {dates}")
    raw = pd.read_excel(io.BytesIO(download_jsda_xlsx(dates[0], sess)), sheet_name=0, header=None)
    print(f"shape: {raw.shape}")
    with pd.option_context("display.max_columns", None, "display.width", 250):
        print(raw.head(n_rows))


# =============================================================================
# Google Sheets 認証・ユーティリティ
# =============================================================================

def authenticate_google_sheets() -> gspread.Client:
    """サービスアカウントで認証する（GitHub Actions用）。429は自動リトライ。"""
    creds_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    try:
        info = json.loads(creds_json)
    except json.JSONDecodeError as e:
        print(f"  エラー: GOOGLE_SERVICE_ACCOUNT_JSON をJSONとして読めません（{len(creds_json)}文字）: {e}")
        raise
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    try:
        from gspread.http_client import BackOffHTTPClient
        return gspread.authorize(creds, http_client=BackOffHTTPClient)
    except (ImportError, TypeError):
        return gspread.authorize(creds)


def _ensure_worksheet_size(ws: gspread.Worksheet, min_cols: int = 0, min_rows: int = 0) -> None:
    if min_cols and ws.col_count < min_cols:
        ws.add_cols(min_cols - ws.col_count)
    if min_rows and ws.row_count < min_rows:
        ws.add_rows(min_rows - ws.row_count)


def get_or_create_worksheet(
    gc: gspread.Client,
    sheet_name: str,
    min_cols: int = 50,
    min_rows: int = 100,
    spreadsheet_id: str | None = None,
) -> gspread.Worksheet:
    spreadsheet = gc.open_by_key(spreadsheet_id or SPREADSHEET_ID)
    try:
        ws = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  シート '{sheet_name}' を新規作成します")
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=max(min_rows, 100), cols=min_cols)
    _ensure_worksheet_size(ws, min_cols=min_cols, min_rows=min_rows)
    return ws


def _native_value(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def _native_row(values: list) -> list:
    return [_native_value(v) for v in values]


def _sheet_dates(header: list[str]) -> set[str]:
    return {from_japanese_date(c) for c in header[FIXED_COLS:] if c.strip()}


def _insert_position(header: list[str], date_yyyymmdd: str) -> int:
    """
    日付の降順（左ほど新しい）が保たれる挿入列（1始まり）を返す。
    最新日なら C列(3)。解釈できない見出しは比較対象から外す。
    """
    for i, cell in enumerate(header[FIXED_COLS:], start=FIRST_DATE_COL):
        d = from_japanese_date(cell)
        if _is_yyyymmdd(d) and d < date_yyyymmdd:
            return i
    # すべて新しい日付 → 最右の日付列の次
    last = FIXED_COLS
    for i, cell in enumerate(header, start=1):
        if cell.strip():
            last = i
    return max(last + 1, FIRST_DATE_COL)


# =============================================================================
# Google Sheets 書き込み
# =============================================================================

def _write_first_time(ws: gspread.Worksheet, df: pd.DataFrame, value_col: str, jp_date: str, sheet_name: str) -> None:
    print(f"  [{sheet_name}] 初回書き込み")
    header = ["銘柄コード", "銘柄名", jp_date]
    rows = [[row["銘柄コード"], row["銘柄名"], row[value_col]] for _, row in df.iterrows()]
    _ensure_worksheet_size(ws, min_cols=FIXED_COLS + 1, min_rows=len(rows) + 1)
    ws.update(values=[header] + [_native_row(r) for r in rows], range_name="A1",
              value_input_option="USER_ENTERED")


def _append_new_stocks(
    ws: gspread.Worksheet,
    df: pd.DataFrame,
    existing_codes: set[str],
    value_col: str,
    total_cols: int,
    insert_col: int,
    existing_row_count: int,
    sheet_name: str,
) -> None:
    """シートに無い銘柄を最下部に追加する（挿入した日付列にだけ値、他は空欄）。"""
    new_df = df[~df["銘柄コード"].astype(str).isin(existing_codes)]
    if new_df.empty:
        return

    new_rows = []
    for _, row in new_df.iterrows():
        full_row = [""] * total_cols
        full_row[0] = str(row["銘柄コード"])
        full_row[1] = row["銘柄名"]
        full_row[insert_col - 1] = row[value_col]
        new_rows.append(_native_row(full_row))

    print(f"  [{sheet_name}] 新規銘柄を追加: {len(new_rows)} 件")
    start_row = existing_row_count + 1
    end_row = start_row + len(new_rows) - 1
    _ensure_worksheet_size(ws, min_cols=total_cols, min_rows=end_row)
    end_a1 = gspread.utils.rowcol_to_a1(end_row, total_cols)
    ws.update(values=new_rows, range_name=f"A{start_row}:{end_a1}", value_input_option="USER_ENTERED")


def _trim_old_columns(ws: gspread.Worksheet, n_date_cols: int, max_date_cols: int, sheet_name: str) -> None:
    if max_date_cols <= 0 or n_date_cols <= max_date_cols:
        return
    start = FIXED_COLS + max_date_cols + 1
    end = FIXED_COLS + n_date_cols
    print(f"  [{sheet_name}] 古い日付列を削除: {end - start + 1} 列（上限 {max_date_cols}）")
    ws.delete_columns(start, end)


def update_spreadsheet_single_column(
    gc: gspread.Client,
    df: pd.DataFrame,
    sheet_name: str,
    value_col: str,
    date_yyyymmdd: str,
    spreadsheet_id: str | None = None,
    max_date_cols: int = 0,
) -> None:
    """
    指定シートに1日分の列を追加する。
    - シートが空 → ヘッダー＋全銘柄を書き込む（初回）
    - 既にデータがある → 日付の降順が保たれる位置に列を挿入（通常はC列）
    - 同じ日付が既にあればスキップ
    """
    ws = get_or_create_worksheet(gc, sheet_name, spreadsheet_id=spreadsheet_id)
    existing = ws.get_all_values()
    has_header = bool(existing) and any(cell.strip() for cell in existing[0])
    jp_date = to_japanese_date(date_yyyymmdd)

    if not has_header:
        _write_first_time(ws, df, value_col, jp_date, sheet_name)
        return

    header = existing[0]
    if date_yyyymmdd in _sheet_dates(header):
        print(f"  [{sheet_name}] {jp_date} のデータは既に追加済みです。スキップします")
        return

    existing_codes = [row[0] for row in existing[1:]]
    insert_col = _insert_position(header, date_yyyymmdd)
    col_letter = gspread.utils.rowcol_to_a1(1, insert_col).rstrip("1")
    print(f"  [{sheet_name}] {col_letter}列に {jp_date} を挿入します")

    lookup = dict(zip(df["銘柄コード"].astype(str), df[value_col]))
    new_column = _native_row([jp_date] + [lookup.get(code, "") for code in existing_codes])
    ws.insert_cols([new_column], col=insert_col, value_input_option="USER_ENTERED")

    total_cols = max(len(header), insert_col - 1) + 1
    _append_new_stocks(
        ws=ws,
        df=df,
        existing_codes=set(existing_codes),
        value_col=value_col,
        total_cols=total_cols,
        insert_col=insert_col,
        existing_row_count=len(existing),
        sheet_name=sheet_name,
    )

    n_date_cols = len([c for c in header[FIXED_COLS:] if c.strip()]) + 1
    _trim_old_columns(ws, n_date_cols, max_date_cols, sheet_name)
    print(f"  [{sheet_name}] 書き込み完了")


# =============================================================================
# 信用残（JPX日次）の更新処理
# =============================================================================

def _existing_margin_dates(gc: gspread.Client) -> set[str]:
    """4シートすべてに入っている日付だけを「取得済み」とみなす（途中失敗の再実行に対応）。"""
    sets = []
    for sheet_name in SHEET_NAMES.values():
        ws = get_or_create_worksheet(gc, sheet_name)
        header = ws.row_values(1)
        sets.append(_sheet_dates(header) if header else set())
    return set.intersection(*sets) if sets else set()


def update_margin(gc: gspread.Client) -> bool:
    """
    JPXの日次PDFを取得して4シートへ書き込む。
    戻り値: 正常なら True。掲載一覧の取得失敗・掲載分の取得/パース失敗があれば False。
    """
    print("=" * 60)
    print("信用取引残高（JPX 日次）")
    print("=" * 60)

    sess = _http_session()
    ok = True

    try:
        listed = list_jpx_margin_files(sess)
        print(f"  掲載されている申込日: {sorted(listed, reverse=True)}")
    except Exception as e:
        print(f"  ✗ 掲載一覧の取得に失敗: {e}")
        listed = {}
        ok = False

    extra = _parse_date_list(os.environ.get("MARGIN_EXTRA_DATES", ""))
    override = os.environ.get("TARGET_DATE_OVERRIDE", "").strip()
    if override:
        extra += _parse_date_list(override)

    done = _existing_margin_dates(gc)
    if done:
        latest_done = max((d for d in done if _is_yyyymmdd(d)), default="なし")
        print(f"  シート済の最新: {latest_done}")

    # 掲載分: 未取得を新しい順に MARGIN_BACKFILL_N 件 / 追加指定分: 未取得すべて
    targets = sorted([d for d in listed if d not in done], reverse=True)[:MARGIN_BACKFILL_N]
    targets += [d for d in extra if d not in done and d not in targets]
    targets = sorted(set(targets))  # 古い順に処理（挿入位置は日付から決まるので順序は崩れない）

    if not targets:
        print("  新しい信用残データはありません")
        return ok

    print(f"  取得対象: {targets}")
    for date_yyyymmdd in targets:
        print(f"\n[{date_yyyymmdd}] ({to_japanese_date(date_yyyymmdd)})")
        url = listed.get(date_yyyymmdd) or JPX_PDF_URL_TEMPLATE.format(date=date_yyyymmdd)
        try:
            df = parse_margin_pdf(download_margin_pdf(url, sess))
        except FileNotFoundError as e:
            # 追加指定の過去日は掲載終了で消えていることがあるので警告のみ
            print(f"  ✗ {e}")
            if date_yyyymmdd in listed:
                ok = False
            continue
        except Exception as e:
            print(f"  ✗ 取得/パース失敗: {e}")
            ok = False
            continue

        for value_col, sheet_name in SHEET_NAMES.items():
            update_spreadsheet_single_column(
                gc, df, sheet_name, value_col, date_yyyymmdd, max_date_cols=MAX_DATE_COLS
            )
    return ok


# =============================================================================
# 貸株残の更新処理
# =============================================================================

def _existing_jsda_dates(gc: gspread.Client) -> set[str]:
    ws = get_or_create_worksheet(gc, KASHIKABU_SHEET_NAME, spreadsheet_id=KASHIKABU_SPREADSHEET_ID)
    header = ws.row_values(1)
    return _sheet_dates(header) if header else set()


def update_kashikabu(gc: gspread.Client, backfill_n: int = 1) -> None:
    """日証協の銘柄別週末残高を取得し、貸株残シートへ書き込む（週次。新ファイルが無ければスキップ）。"""
    print("\n" + "=" * 60)
    print("貸株残（日本証券業協会）")
    print("=" * 60)

    sess = _http_session()
    available = list_jsda_dates(sess)
    print(f"  掲載されている申込日: {available[:5]}{' ...' if len(available) > 5 else ''}")

    done = _existing_jsda_dates(gc)
    targets = sorted([d for d in available if d not in done], reverse=True)[:backfill_n]
    targets.sort()
    if not targets:
        print("  新しい貸株残データはありません。スキップします")
        return

    print(f"  取得対象: {targets}")
    for date_yyyymmdd in targets:
        print(f"\n[{date_yyyymmdd}] ({to_japanese_date(date_yyyymmdd)})")
        try:
            df = parse_kashikabu_xlsx(download_jsda_xlsx(date_yyyymmdd, sess))
        except Exception as e:
            print(f"  ✗ 取得/パース失敗: {e}")
            continue
        update_spreadsheet_single_column(
            gc, df, KASHIKABU_SHEET_NAME, KASHIKABU_VALUE_COL, date_yyyymmdd,
            spreadsheet_id=KASHIKABU_SPREADSHEET_ID,
        )


# =============================================================================
# メイン処理
# =============================================================================

def main() -> None:
    if len(sys.argv) >= 3 and sys.argv[1] == "--inspect":
        inspect_margin_pdf(sys.argv[2])
        return

    if not SPREADSHEET_ID:
        print("エラー: 環境変数 SPREADSHEET_ID が設定されていません")
        sys.exit(1)

    gc = authenticate_google_sheets()

    # ---- 1) JPX 信用取引残高（日次） ----
    try:
        margin_ok = update_margin(gc)
    except Exception as e:
        print(f"✗ 信用残の処理で予期しないエラー: {e}")
        margin_ok = False

    # ---- 2) 日証協 貸株残（信用残の成否とは独立して実行） ----
    if JSDA_ENABLED:
        try:
            update_kashikabu(gc, backfill_n=JSDA_BACKFILL_N)
        except Exception as e:
            print(f"警告: 貸株残の処理に失敗しました: {e}")
            print("inspect_jsda_latest() でファイルの中身を確認してください。")
    else:
        print("\n貸株残の取得はスキップされました（JSDA_ENABLED=0）")

    if not margin_ok:
        print("\n信用残の取得に失敗した日があります。ログを確認してください。")
        sys.exit(1)  # 失敗を Actions 上で赤く表示させる


if __name__ == "__main__":
    main()
