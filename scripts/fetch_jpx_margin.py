"""
JPX「銘柄別信用取引週末残高」PDFを取得し、
一般信用買残高・一般信用売残高・制度信用買残高・制度信用売残高を
それぞれ別シートにGoogle Spreadsheetへ書き込む。

あわせて、日本証券業協会「銘柄別株券等貸借週末残高」(xlsx)から
貸株残を取得し、別スプレッドシートの「貸株残」シートへ同じ形式で書き込む。

■ シート構成
  A列 = 銘柄コード
  B列 = 銘柄名
  C列 = 最新の申込日（日本語形式）
  D列以降 = 過去の申込日（新しい順）

■ 書き込みロジック
  新しい週のデータは「一番右に追記」ではなく「常にC列に挿入」する。
  挿入すると、それまでC列以降にあった過去データは自動的に1列ずつ右（D列以降）へ
  シフトされる（Google Sheets の insertDimension の挙動をそのまま利用）。
  → スプレッドシートを開いたときに常に「一番左＝最新データ」になり、
    週を追うごとに右へ過去データが積み上がっていく形になる。

  既存銘柄（行）はそのまま維持し、該当する行のC列にだけ新しい値を書き込む。
  新規上場銘柄など、まだシートに存在しない銘柄コードは、
  シート最下部に新しい行として追加する（C列に値、D列以降は空欄）。

【2026/7/10申込分のPDFで実データ確認済み】
PDFは罫線なしのテキストレイアウトのため、pdfplumberのextract_tables()は使わず、
extract_text()で取った行を正規表現でパースする方式にしている。
数値中の「▲」はマイナスとしてパースする。
合計値(買い残/売り残/倍率)はここでは計算しない（別シートでユーザー側が算出）。

■ 貸株残について（追加分）
  取得元: 日本証券業協会「株券等貸借取引状況（週間）」の銘柄別週末残高
          https://www.jsda.or.jp/shiryoshitsu/toukei/kabu-taiw/index.html
          ファイル名は  files/{申込週末日YYYYMMDD}z.xlsx  という規則。
  公表タイミング: 報告週の翌週木曜。JPXの信用残（翌週火曜公表）より約1週間遅い。
  → 火曜に本スクリプトを実行した時点では、貸株残は「1週前の金曜分」が最新となる。
    日付ラベルは各データ本来の金曜日付を使うため、シート間で列位置は1週ずれるが
    ラベルを見れば対応関係は一意に判別できる。
  制度貸借取引は対象外の統計であり、JPXの信用残とは別系統のデータである点に注意。
"""

from __future__ import annotations

import io
import json
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

# 4種類の内訳データをそれぞれ別シートに書き込む（合計はユーザー側の別シートで算出）
SHEET_NAMES = {
    "一般信用買残高": "一般信用買残高",
    "一般信用売残高": "一般信用売残高",
    "制度信用買残高": "制度信用買残高",
    "制度信用売残高": "制度信用売残高",
}

JPX_PDF_URL_TEMPLATE = (
    "https://www.jpx.co.jp/markets/statistics-equities/"
    "margin/tvdivq0000001rnl-att/syumatsu{date}00.pdf"
)

# シート構成: A列=銘柄コード, B列=銘柄名, C列=最新申込日, D列以降=過去の申込日
FIXED_COLS = 2
DATE_INSERT_COL = FIXED_COLS + 1  # = 3 (C列)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

_WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]

# 銘柄データ抽出用の正規表現
ISIN_PATTERN = re.compile(r"JP[A-Z0-9]{10}")
NUMBER_PATTERN = re.compile(r"(▲?)\s*(\d{1,3}(?:,\d{3})*)")

# 小計・総合計行の判定用キーワード（JPX側の区分数が増減しても対応できるよう、
# 固定の行数リストではなく行の中身から動的に判定する）
SEGMENT_KEYWORDS = ["プライム", "スタンダード", "グロース", "投信等"]
CATEGORY_ORDER = ["貸借銘柄", "制度信用銘柄", "その他", "総合計", "全体"]
SEGMENT_ORDER = ["合計", "プライム", "スタンダード", "グロース", "投信等"]

# -----------------------------------------------------------------------------
# 貸株残（日本証券業協会）関連の定数
# -----------------------------------------------------------------------------

# 貸株残の書き込み先スプレッドシート（信用残とは別ブック）
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

# xlsxの列見出し候補（優先順）。書式が変わった場合はここだけ直せばよい。
JSDA_CODE_KEYWORDS = ["銘柄コード", "コード", "code"]
JSDA_NAME_KEYWORDS = ["銘柄名", "名称", "銘柄"]
JSDA_BALANCE_KEYWORDS = ["貸付残高", "貸付", "貸株", "週末残高", "残高"]
JSDA_BALANCE_EXCLUDE = ["借入", "返済", "新規", "成約"]

# 1回の実行で最大何週分まで遡って埋めるか（初回バックフィルは 26 などを指定）
JSDA_BACKFILL_N = int(os.environ.get("JSDA_BACKFILL_N", "1"))
# 貸株残の取得をスキップしたい場合は "0" を設定
JSDA_ENABLED = os.environ.get("JSDA_ENABLED", "1") != "0"

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en;q=0.8",
}


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
    days_since_friday = (run_date.weekday() - 4) % 7  # 月=0 ... 金=4 ... 日=6
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
# PDF取得
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


# =============================================================================
# PDFパース：数値ユーティリティ
# =============================================================================

def _to_number(s: str | None) -> float | None:
    if s is None:
        return None
    s = s.replace(",", "").replace("−", "-").strip()
    if s in ("", "-", "―"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_signed_tokens(tokens: list[str], count: int) -> tuple[list[float | None], list[str]]:
    """
    トークン列から count 個の数値を読み取る。
    '▲ 数字' は負の値として扱う（前週比の減少表記）。
    戻り値: (数値リスト, 未使用トークンの残り)
    """
    values: list[float | None] = []
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


# =============================================================================
# PDFパース：小計・総合計行
# =============================================================================

def _parse_subtotal_line(line: str) -> list[float] | None:
    """
    小計・総合計行から12個の数値を抽出する。
    数値同士がスペースなしで連結していることがあるが、
    3桁区切りカンマのパターンで正しく分割できる。
    データ行でなければ None を返す。
    """
    m = re.search(r"(\d+)\s*銘柄", line)
    if not m:
        return None
    tail = line[m.end():]
    numbers = []
    for nm in NUMBER_PATTERN.finditer(tail):
        sign, digits = nm.groups()
        if not digits:
            continue
        val = float(digits.replace(",", ""))
        numbers.append(-val if sign else val)
    if len(numbers) < 12:
        return None
    return numbers[:12]


def _label_subtotal_line(line: str, state: dict) -> str | None:
    """
    小計・総合計行のラベル（例: "貸借銘柄 プライム 小計"）を、
    行の中身のキーワードとカテゴリの追跡状態から動的に判定する。
    state: {"current_category": str|None, "passed_total": bool}
    判定できなければ None を返す。
    """
    if "貸借銘柄" in line:
        state["current_category"] = "貸借銘柄"
        return "貸借銘柄 合計"
    if "制度信用銘柄" in line:
        state["current_category"] = "制度信用銘柄"
        return "制度信用銘柄 合計"
    if re.search(r"その他.*other issues", line):
        state["current_category"] = "その他"
        return "その他 合計"
    if "総合計" in line:
        state["passed_total"] = True
        return "総合計"

    for seg in SEGMENT_KEYWORDS:
        if seg in line and "小計" in line:
            # 総合計より後に出てくる区分別小計は、カテゴリを横断した「全体」の集計
            category = "全体" if state["passed_total"] else state["current_category"]
            if category is None:
                return None
            return f"{category} {seg} 小計"

    return None


def _subtotal_sort_key(label: str) -> int:
    cat_idx = next((i for i, c in enumerate(CATEGORY_ORDER) if label.startswith(c)), len(CATEGORY_ORDER))
    seg_idx = next((i for i, s in enumerate(SEGMENT_ORDER) if s in label), len(SEGMENT_ORDER))
    return cat_idx * 10 + seg_idx


# =============================================================================
# PDFパース：メイン
# =============================================================================

def parse_margin_pdf(pdf_bytes: bytes) -> pd.DataFrame:
    """
    PDFのテキストを1行ずつ正規表現でパースし、DataFrameで返す。
    列: 銘柄コード, 銘柄名, 一般信用買残高, 一般信用売残高, 制度信用買残高, 制度信用売残高

    1データ行の並び（実データで確認済み）:
      [貸借フラグB] 銘柄名 株式種別 5桁コード ISIN
      売残高(合計) 前週比 買残高(合計) 前週比
      売残高(一般信用) 前週比 売残高(制度信用) 前週比
      買残高(一般信用) 前週比 買残高(制度信用) 前週比

    銘柄行の他に、小計・総合計行も "SUB_<ラベル名>" という特別なコードの行として
    先頭に含める。ラベルはキーワードから動的判定するため、区分数の増減に対応できる。
    合計値はここでは計算しない（PDFに実際に記載された小計・総合計値をそのまま使う）。
    """
    records = []
    subtotal_dict: dict[str, list] = {}  # label -> 12個の数値（後勝ちで上書き）
    unlabeled_lines = []  # ラベル判定できなかった小計候補行（調査用）
    skipped = 0
    label_state = {"current_category": None, "passed_total": False}

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        print(f"  総ページ数: {len(pdf.pages)}")
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                m = ISIN_PATTERN.search(line)
                if not m:
                    parsed = _parse_subtotal_line(line)
                    if parsed:
                        label = _label_subtotal_line(line, label_state)
                        if label:
                            if label in subtotal_dict:
                                print(f"  警告: 小計ラベル '{label}' が複数回検出されました。後の値で上書きします。")
                            subtotal_dict[label] = parsed
                        else:
                            unlabeled_lines.append(line)
                    continue  # ISINを含まない行（見出し・区切り等）はスキップ

                before = line[: m.start()].strip()
                after = line[m.end():].strip()

                code_match = re.search(r"([0-9A-Z]{5})\s*$", before)
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
                # values[4]=売残高一般信用, [6]=売残高制度信用, [8]=買残高一般信用, [10]=買残高制度信用
                general_sell = values[4]
                standard_sell = values[6]
                general_buy = values[8]
                standard_buy = values[10]

                if None in (general_sell, standard_sell, general_buy, standard_buy):
                    skipped += 1
                    continue

                records.append(
                    {
                        "銘柄コード": code,
                        "銘柄名": name,
                        "一般信用買残高": general_buy,
                        "一般信用売残高": general_sell,
                        "制度信用買残高": standard_buy,
                        "制度信用売残高": standard_sell,
                    }
                )

    stock_df = pd.DataFrame(records)
    print(f"  抽出件数: {len(stock_df)} 銘柄（パース失敗でスキップ: {skipped} 行）")
    if stock_df.empty:
        raise ValueError("PDFから銘柄データを抽出できませんでした。レイアウトが変わった可能性があります。")

    # 銘柄コード順（数値の昇順）にソート。実銘柄の並び順は集計行より必ず後ろになるよう
    # 十分大きいオフセットを足す。これにより初回書き込み時の行順が揃い、
    # 以後の週で新規上場銘柄は既存コードに一致しないため自動的に末尾へ追加される。
    stock_df["_sort_key"] = 10_000 + pd.to_numeric(stock_df["銘柄コード"], errors="coerce").fillna(99_999)

    print(f"  小計/総合計行の検出数: {len(subtotal_dict)} 件")
    if unlabeled_lines:
        print(f"  警告: ラベルを判定できなかった小計候補行が {len(unlabeled_lines)} 件あります:")
        for l in unlabeled_lines:
            print(f"    {l!r}")

    if subtotal_dict:
        summary_records = []
        for label, values in subtotal_dict.items():
            general_sell, standard_sell = values[4], values[6]
            general_buy, standard_buy = values[8], values[10]
            summary_records.append(
                {
                    "銘柄コード": "SUB_" + label.replace(" ", ""),
                    "銘柄名": label,
                    "一般信用買残高": general_buy,
                    "一般信用売残高": general_sell,
                    "制度信用買残高": standard_buy,
                    "制度信用売残高": standard_sell,
                    "_sort_key": _subtotal_sort_key(label),
                }
            )
        summary_df = pd.DataFrame(summary_records)
        df = pd.concat([summary_df, stock_df], ignore_index=True)
    else:
        print("  警告: 小計/総合計行を1件も検出できませんでした。集計行なしで続行します。")
        df = stock_df

    df = df.sort_values("_sort_key", na_position="last").drop(columns="_sort_key").reset_index(drop=True)
    return df


# =============================================================================
# 貸株残（日本証券業協会）取得・パース  ★追加セクション★
# =============================================================================

def _jsda_session() -> requests.Session:
    sess = requests.Session()
    sess.headers.update(HTTP_HEADERS)
    return sess


def list_jsda_dates(sess: requests.Session | None = None) -> list[str]:
    """
    日証協のindexページから銘柄別週末残高ファイルの申込日(YYYYMMDD)を新しい順に返す。
    日付を実行日から逆算せず、実際に掲載されているリンクから拾うため、
    公表が遅れた週でも取りこぼしや404が発生しない。
    """
    sess = sess or _jsda_session()
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
    sess = sess or _jsda_session()
    url = JSDA_FILE_URL_TEMPLATE.format(date=date_yyyymmdd)
    print(f"  ダウンロード中: {url}")
    resp = sess.get(url, timeout=60)
    if resp.status_code == 404:
        raise FileNotFoundError(f"貸株残ファイルが見つかりません: {url}")
    resp.raise_for_status()
    return resp.content


def _normalize_stock_code(value) -> str | None:
    """
    '1301' / '1301.0' / '13010' / '402A' などの表記ゆれを4桁形式に寄せる。
    JPX側は5桁コードの先頭4桁を使っているため、それに合わせる。
    """
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
    """keywords の優先順で最初に一致した列名を返す。exclude を含む列は除外する。"""
    for kw in keywords:
        for col in columns:
            text = str(col)
            if exclude and any(ex in text for ex in exclude):
                continue
            if kw in text:
                return col
    return None


def _find_jsda_header_row(raw: pd.DataFrame, max_scan: int = 25) -> int | None:
    """銘柄コード列を含むヘッダー行のインデックスを推定する（前置きの表題行対策）。"""
    for i in range(min(max_scan, len(raw))):
        row = [str(v) for v in raw.iloc[i].tolist()]
        if any(any(kw in v for kw in JSDA_CODE_KEYWORDS) for v in row):
            return i
    return None


def parse_kashikabu_xlsx(xlsx_bytes: bytes) -> pd.DataFrame:
    """
    日証協の銘柄別株券等貸借週末残高 xlsx をパースし、DataFrameで返す。
    列: 銘柄コード, 銘柄名, 貸株残
    """
    raw = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0, header=None)
    header_row = _find_jsda_header_row(raw)
    if header_row is None:
        raise ValueError(
            "ヘッダー行を特定できませんでした。inspect_jsda_latest() で中身を確認してください。"
        )

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
            continue  # 合計行・空行・見出し繰り返しなど
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

    # 同一コードが複数行にまたがる場合は合算しておく
    out = (
        out.groupby("銘柄コード", as_index=False)
        .agg({"銘柄名": "first", KASHIKABU_VALUE_COL: "sum"})
    )
    out["_sort_key"] = pd.to_numeric(out["銘柄コード"], errors="coerce").fillna(99_999)
    out = out.sort_values("_sort_key").drop(columns="_sort_key").reset_index(drop=True)
    return out


def inspect_jsda_latest(n_rows: int = 15) -> None:
    """列名がずれたときの調査用。最新ファイルの先頭数行をそのまま表示する。"""
    sess = _jsda_session()
    dates = list_jsda_dates(sess)
    print(f"掲載されている申込日(新しい順): {dates}")
    raw = pd.read_excel(io.BytesIO(download_jsda_xlsx(dates[0], sess)), sheet_name=0, header=None)
    print(f"shape: {raw.shape}")
    with pd.option_context("display.max_columns", None, "display.width", 250):
        print(raw.head(n_rows))


# =============================================================================
# Google Sheets 認証
# =============================================================================

def authenticate_google_sheets() -> gspread.Client:
    """サービスアカウントで認証する（GitHub Actions用）。"""
    creds_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]

    # --- デバッグ: 中身は伏せつつ、長さと先頭・末尾だけ出力して原因を特定する ---
    print(f"  DEBUG: GOOGLE_SERVICE_ACCOUNT_JSON の文字数: {len(creds_json)}")
    print(f"  DEBUG: 先頭10文字: {creds_json[:10]!r}")
    print(f"  DEBUG: 末尾10文字: {creds_json[-10:]!r}")
    print(f"  DEBUG: 改行の数: {creds_json.count(chr(10))}")
    # --- デバッグここまで ---

    try:
        info = json.loads(creds_json)
    except json.JSONDecodeError as e:
        print(f"  DEBUG: JSONパース失敗の詳細: {e}")
        raise

    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def get_or_create_worksheet(
    gc: gspread.Client,
    sheet_name: str,
    min_cols: int = 50,
    min_rows: int = 100,
    spreadsheet_id: str | None = None,
):
    """
    [変更点] spreadsheet_id を引数で受け取れるようにした。
    省略時は従来どおり SPREADSHEET_ID（信用残のブック）を使う。
    """
    spreadsheet = gc.open_by_key(spreadsheet_id or SPREADSHEET_ID)
    try:
        ws = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  シート '{sheet_name}' を新規作成します")
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=max(min_rows, 100), cols=min_cols)
    if ws.col_count < min_cols:
        ws.add_cols(min_cols - ws.col_count)
        ws = spreadsheet.worksheet(sheet_name)
    if ws.row_count < min_rows:
        ws.add_rows(min_rows - ws.row_count)
        ws = spreadsheet.worksheet(sheet_name)
    return ws


def _native_row(values: list) -> list:
    """None / NaN を空文字に変換する（gspreadへの書き込み用）。"""
    return [("" if v is None or (isinstance(v, float) and pd.isna(v)) else v) for v in values]


def _native_rows(rows: list[list]) -> list[list]:
    return [_native_row(r) for r in rows]


# =============================================================================
# Google Sheets 書き込み
# =============================================================================

def _write_first_time(ws: gspread.Worksheet, df: pd.DataFrame, value_col: str, jp_date: str, sheet_name: str) -> None:
    """シートが空の場合の初回書き込み。ヘッダー＋全銘柄をA1から書く。"""
    print(f"  [{sheet_name}] 初回書き込み")
    header = ["銘柄コード", "銘柄名", jp_date]
    rows = [[row["銘柄コード"], row["銘柄名"], row[value_col]] for _, row in df.iterrows()]
    ws.update(values=[header] + _native_rows(rows), range_name="A1", value_input_option="USER_ENTERED")


def _insert_new_date_column(
    ws: gspread.Worksheet,
    existing_codes: list[str],
    df: pd.DataFrame,
    value_col: str,
    jp_date: str,
) -> None:
    """
    C列（DATE_INSERT_COL）に新しい日付の1列を挿入する。
    これにより、それまでC列以降にあった過去データは自動的に1列右（D列以降）へシフトする。
    既存の行順（existing_codes の並び）はそのまま維持し、
    各行に対応する銘柄の値を上から順に並べて書き込む。
    データが見つからない銘柄（当該週にデータがない場合）は空欄にする。
    """
    df_lookup = dict(zip(df["銘柄コード"].astype(str), df[value_col]))

    # 1列分のデータ = [ヘッダー(日付), 行1の値, 行2の値, ...]（既存行の並び順どおり）
    new_column = [jp_date] + [df_lookup.get(code, "") for code in existing_codes]
    new_column = _native_row(new_column)

    # insert_cols の values は「挿入する列ごとのリスト」を渡す仕様のため、
    # 1列だけ挿入する場合は [new_column] という二重リストにする。
    ws.insert_cols([new_column], col=DATE_INSERT_COL, value_input_option="USER_ENTERED")


def _append_new_stocks(
    ws: gspread.Worksheet,
    df: pd.DataFrame,
    existing_codes: set[str],
    value_col: str,
    old_header_len: int,
    existing_row_count: int,
    sheet_name: str,
) -> None:
    """
    シートにまだ存在しない銘柄コード（新規上場など）を、シート最下部に新しい行として追加する。
    列構成は「C列に挿入」後の状態に合わせる: C列=新しい日付の値、D列以降（旧C列以降）=空欄。
    """
    new_codes = set(df["銘柄コード"].astype(str)) - existing_codes
    if not new_codes:
        return

    n_old_date_cols = old_header_len - FIXED_COLS  # 旧D列以降に相当する過去日付列の数
    new_rows = []
    for _, row in df.iterrows():
        code = str(row["銘柄コード"])
        if code in new_codes:
            full_row = [code, row["銘柄名"], row[value_col]] + [""] * n_old_date_cols
            new_rows.append(_native_row(full_row))

    print(f"  [{sheet_name}] 新規銘柄を追加: {len(new_rows)} 件")
    total_cols = old_header_len + 1  # C列挿入後の総列数
    start_row = existing_row_count + 1
    end_row = start_row + len(new_rows) - 1

    # [変更点] 書き込み先ブックを取り違えないよう、対象ワークシート自身のIDを使う
    ws = get_or_create_worksheet(
        gc=ws.spreadsheet.client,
        sheet_name=sheet_name,
        min_cols=total_cols,
        min_rows=end_row,
        spreadsheet_id=ws.spreadsheet.id,
    )
    end_a1 = gspread.utils.rowcol_to_a1(end_row, total_cols)
    ws.update(values=new_rows, range_name=f"A{start_row}:{end_a1}", value_input_option="USER_ENTERED")


def update_spreadsheet_single_column(
    gc: gspread.Client,
    df: pd.DataFrame,
    sheet_name: str,
    value_col: str,
    date_yyyymmdd: str,
    spreadsheet_id: str | None = None,
) -> None:
    """
    指定シートを更新する。

    - シートが空 -> ヘッダー＋全銘柄を書き込む（初回）
    - シートに既にデータがある -> C列に新しい日付列を「挿入」し、
      それまでC列以降にあったデータをD列以降へ押し出す。
      既存銘柄は該当行のC列を更新、新規銘柄はシート末尾に行を追加する。
    - 同じ日付が既にどこかの列にあればスキップ（二重書き込み防止）

    [変更点] spreadsheet_id を引数で受け取れるようにした（貸株残は別ブックのため）。
    """
    ws = get_or_create_worksheet(gc, sheet_name, spreadsheet_id=spreadsheet_id)
    existing = ws.get_all_values()
    # 新規作成直後のシートは get_all_values() が [] ではなく
    # 中身が空の行を返すことがあるため、実際に文字が入っているかで判定する
    has_header = bool(existing) and any(cell.strip() for cell in existing[0])
    jp_date = to_japanese_date(date_yyyymmdd)

    if not has_header:
        ws = get_or_create_worksheet(
            gc, sheet_name, min_cols=FIXED_COLS + 1, spreadsheet_id=spreadsheet_id
        )
        _write_first_time(ws, df, value_col, jp_date, sheet_name)
        return

    header = existing[0]
    existing_codes = [row[0] for row in existing[1:]]

    if jp_date in header[FIXED_COLS:]:
        print(f"  [{sheet_name}] {jp_date} のデータは既に追加済みです。スキップします")
        return

    print(f"  [{sheet_name}] C列に {jp_date} を挿入し、既存データを右へシフトします")
    _insert_new_date_column(ws, existing_codes, df, value_col, jp_date)

    _append_new_stocks(
        ws=ws,
        df=df,
        existing_codes=set(existing_codes),
        value_col=value_col,
        old_header_len=len(header),
        existing_row_count=len(existing),
        sheet_name=sheet_name,
    )

    print(f"  [{sheet_name}] 書き込み完了！")


# =============================================================================
# 貸株残の更新処理  ★追加セクション★
# =============================================================================

def _existing_jsda_dates(gc: gspread.Client) -> set[str]:
    """貸株残シートのヘッダーから、既に書き込み済みの申込日(YYYYMMDD)を集める。"""
    try:
        ws = get_or_create_worksheet(
            gc, KASHIKABU_SHEET_NAME, spreadsheet_id=KASHIKABU_SPREADSHEET_ID
        )
    except gspread.exceptions.APIError as e:
        print(f"  警告: 貸株残シートにアクセスできません: {e}")
        raise
    values = ws.get_all_values()
    if not values or not any(c.strip() for c in values[0]):
        return set()
    return {from_japanese_date(c) for c in values[0][FIXED_COLS:] if c.strip()}


def update_kashikabu(gc: gspread.Client, backfill_n: int = 1) -> None:
    """
    日証協の銘柄別週末残高を取得し、貸株残シートへ書き込む。

    JPXの信用残（翌週火曜公表）に対し、日証協は翌週木曜公表のため、
    火曜実行時点で取得できる最新は「1週前の金曜分」になる。
    実行日から日付を逆算するのではなく、掲載されているファイルの一覧から
    「まだシートに無い日付」を拾うため、公表遅延や祝日週があってもズレない。

    backfill_n を増やすと、未取得の週を古い順にまとめて書き込む（初回投入用）。
    """
    print("\n" + "=" * 60)
    print("貸株残（日本証券業協会）")
    print("=" * 60)

    sess = _jsda_session()
    available = list_jsda_dates(sess)
    print(f"  掲載されている申込日: {available[:5]}{' ...' if len(available) > 5 else ''}")

    done = _existing_jsda_dates(gc)
    if done:
        print(f"  シート済の最新: {max(done)}")

    # 未取得のものを新しい順に backfill_n 件取り、古い順に書き込む。
    # 古い順に挿入することで、常にC列＝最新となる並びが維持される。
    targets = sorted([d for d in available if d not in done], reverse=True)[:backfill_n]
    targets.sort()

    if not targets:
        print("  新しい貸株残データはありません。スキップします")
        return

    print(f"  取得対象: {targets}")
    for date_yyyymmdd in targets:
        print(f"\n[{date_yyyymmdd}] ({to_japanese_date(date_yyyymmdd)})")
        try:
            xlsx_bytes = download_jsda_xlsx(date_yyyymmdd, sess)
            df = parse_kashikabu_xlsx(xlsx_bytes)
        except Exception as e:
            print(f"  ✗ 取得/パース失敗: {e}")
            continue
        update_spreadsheet_single_column(
            gc,
            df,
            KASHIKABU_SHEET_NAME,
            KASHIKABU_VALUE_COL,
            date_yyyymmdd,
            spreadsheet_id=KASHIKABU_SPREADSHEET_ID,
        )


# =============================================================================
# メイン処理
# =============================================================================

def main() -> None:
    if not SPREADSHEET_ID:
        print("エラー: 環境変数 SPREADSHEET_ID が設定されていません")
        sys.exit(1)

    override = os.environ.get("TARGET_DATE_OVERRIDE", "").strip()
    target_date = override if override else get_target_friday()
    print(f"対象申込日: {target_date} ({to_japanese_date(target_date)})")

    gc = authenticate_google_sheets()

    # ---- 1) JPX 信用取引週末残高 ----
    margin_failed = False
    try:
        pdf_bytes = download_pdf(target_date)
    except FileNotFoundError as e:
        print(f"警告: {e}")
        print("公表が遅れている可能性があります。ワークフローを再実行するか翌日リトライしてください。")
        margin_failed = True  # 貸株残は独立して処理したいので、ここでは終了しない
    else:
        df = parse_margin_pdf(pdf_bytes)
        for value_col, sheet_name in SHEET_NAMES.items():
            update_spreadsheet_single_column(gc, df, sheet_name, value_col, target_date)

    # ---- 2) 日証協 貸株残（信用残の成否とは独立して実行） ----
    if JSDA_ENABLED:
        try:
            update_kashikabu(gc, backfill_n=JSDA_BACKFILL_N)
        except Exception as e:
            print(f"警告: 貸株残の処理に失敗しました: {e}")
            print("inspect_jsda_latest() でファイルの中身を確認してください。")
    else:
        print("\n貸株残の取得はスキップされました（JSDA_ENABLED=0）")

    if margin_failed:
        sys.exit(0)  # 公表遅延は失敗扱いにしない


if __name__ == "__main__":
    main()
