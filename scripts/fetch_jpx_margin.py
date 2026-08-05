"""
JPX「銘柄別信用取引週末残高」PDFを取得し、
一般信用買残高・一般信用売残高・制度信用買残高・制度信用売残高を
それぞれ別シートにGoogle Spreadsheetへ書き込む。

■ シート構成
  A列 = 銘柄コード
  B列 = 銘柄名
  C列 = 最新の申込日（日本語形式）
  D列以降 = 過去の申込日（新しい順）

■ 書き込みロジック（今回の修正の核心）
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


def get_or_create_worksheet(gc: gspread.Client, sheet_name: str, min_cols: int = 50, min_rows: int = 100):
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
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
# Google Sheets 書き込み（★ 今回の修正の中心 ★）
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

    ws = get_or_create_worksheet(gc=ws.spreadsheet.client, sheet_name=sheet_name, min_cols=total_cols, min_rows=end_row)
    end_a1 = gspread.utils.rowcol_to_a1(end_row, total_cols)
    ws.update(values=new_rows, range_name=f"A{start_row}:{end_a1}", value_input_option="USER_ENTERED")


def update_spreadsheet_single_column(
    gc: gspread.Client, df: pd.DataFrame, sheet_name: str, value_col: str, date_yyyymmdd: str
) -> None:
    """
    指定シートを更新する。

    - シートが空 -> ヘッダー＋全銘柄を書き込む（初回）
    - シートに既にデータがある -> C列に新しい日付列を「挿入」し、
      それまでC列以降にあったデータをD列以降へ押し出す。
      既存銘柄は該当行のC列を更新、新規銘柄はシート末尾に行を追加する。
    - 同じ日付が既にどこかの列にあればスキップ（二重書き込み防止）
    """
    ws = get_or_create_worksheet(gc, sheet_name)
    existing = ws.get_all_values()
    # 新規作成直後のシートは get_all_values() が [] ではなく
    # 中身が空の行を返すことがあるため、実際に文字が入っているかで判定する
    has_header = bool(existing) and any(cell.strip() for cell in existing[0])
    jp_date = to_japanese_date(date_yyyymmdd)

    if not has_header:
        ws = get_or_create_worksheet(gc, sheet_name, min_cols=FIXED_COLS + 1)
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
# メイン処理
# =============================================================================

def main() -> None:
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

    for value_col, sheet_name in SHEET_NAMES.items():
        update_spreadsheet_single_column(gc, df, sheet_name, value_col, target_date)


if __name__ == "__main__":
    main()
