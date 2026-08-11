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

JPX_INDEX_URL = "https://www.jpx.co.jp/markets/statistics-equities/daily/index.html"

# ------------------------------------------------------------------
# 1. JPXサイトから最新(または指定日)の「株式相場表」PDFリンクを取得
# ------------------------------------------------------------------

def find_pdf_url(target_date: Optional[date] = None) -> tuple[str, date]:
    """
    JPX日報indexページをスクレイピングし、株式相場表(stq_YYYYMMDD.pdf)へのリンクを取得する。
    target_date を指定した場合はその日付の行を探す。指定しない場合は一覧の最初(最新)を返す。
    """
    from bs4 import BeautifulSoup

    resp = requests.get(JPX_INDEX_URL, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    ymd_re = re.compile(r"stq_(\d{8})\.pdf")
    candidates = []
    for a in soup.find_all("a", href=True):
        m = ymd_re.search(a["href"])
        if m:
            d = datetime.strptime(m.group(1), "%Y%m%d").date()
            url = a["href"] if a["href"].startswith("http") else "https://www.jpx.co.jp" + a["href"]
            candidates.append((d, url))

    if not candidates:
        raise RuntimeError("JPX日報indexページから株式相場表PDFリンクを検出できませんでした。"
                            "サイト構造が変更された可能性があります。")

    candidates.sort(key=lambda x: x[0], reverse=True)  # 新しい日付順

    if target_date is None:
        return candidates[0][1], candidates[0][0]

    for d, url in candidates:
        if d == target_date:
            return url, d

    raise RuntimeError(f"指定日 {target_date} の株式相場表PDFが見つかりませんでした"
                        "(休場日、まだ掲載されていない、または一覧の表示範囲外の可能性があります)。")


# ------------------------------------------------------------------
# 2. PDFダウンロード & テキスト抽出
# ------------------------------------------------------------------

def download_pdf(url: str) -> bytes:
    resp = requests.get(url, timeout=60)
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


# ------------------------------------------------------------------
# 3. テキストパース
# ------------------------------------------------------------------

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


def parse_text(text: str) -> list[dict]:
    """
    PDFから抽出したテキストを解析し、銘柄ごとのレコード(dict)のリストを返す。

    実際のPDFのレイアウトは、1銘柄につき以下の2行で構成される:
        1行目: コード 売買単位 銘柄名(和文) 前場OHLC 後場OHLC 最終気配 前日比 VWAP 出来高 売買代金
        2行目: 銘柄名(英文)

    その手前に出現する「業種」「市場」見出し行を状態として保持し、各レコードに付与する。
    """
    lines = text.splitlines()
    content = [l for l in lines if not is_junk(l)]

    current_market = None
    current_industry = None
    pending_record = None  # データ行はマッチしたが、まだ英文社名行を待っている状態

    records = []

    for raw in content:
        s = raw.strip()

        if pending_record is not None:
            # 直前のデータ行の次に来る行は英文社名(そのまま採用)
            pending_record["name_en"] = s
            records.append(pending_record)
            pending_record = None
            continue

        dm = DATA_LINE_RE.match(s)
        if dm:
            rec = dm.groupdict()
            rec["industry"] = current_industry
            rec["market"] = current_market
            pending_record = rec
            continue

        # ここに来るのは「業種見出し」または「市場見出し」等のセクション行
        if s in MARKET_HEADERS:
            current_market = s.replace("市場", "")
        elif s in NON_INDUSTRY_OTHER_LINES:
            pass  # 市場の英語表記など、業種欄には転記しない
        else:
            # 業種見出しは "水産・農林業 Fishery,Agriculture & Forestry" のように
            # 和文+英文が空白区切りで並んでいるため、和文部分(先頭トークン)のみを採用する
            current_industry = s.split()[0] if s.split() else s

    if pending_record is not None:
        # 最終行がデータ行で終わっていて英文社名行が取得できなかった場合
        pending_record["name_en"] = ""
        records.append(pending_record)

    return records


# ------------------------------------------------------------------
# 4. Google スプレッドシートへの書き込み(ワイド形式: 銘柄=行, 日付=列)
# ------------------------------------------------------------------

# 対象スプレッドシートID (URLの /d/ と /edit の間の文字列)
# 環境変数 SPREADSHEET_ID が設定されていればそちらを優先する(別スプレッドシートに切り替えたい場合用)
SPREADSHEET_ID_DEFAULT = "1kWST0CkkIvo3irPSbMgtVtUqqRDXFwvRREYZDRQAFMY"

# 各項目シートのgid (URLの gid= の値)
SHEET_SPECS = [
    # (表示名, gid, レコードのキー)
    ("前場始値", 383067295, "am_open"),
    ("前場高値", 2103325248, "am_high"),
    ("前場安値", 820854515, "am_low"),
    ("前場終値", 599538789, "am_close"),
    ("後場始値", 266318248, "pm_open"),
    ("後場高値", 1004322238, "pm_high"),
    ("後場安値", 146259449, "pm_low"),
    ("後場終値", 1352449304, "pm_close"),
    ("VWAP", 519095635, "vwap"),
    ("出来高", 941877137, "volume"),
]
HEADER_ROW = ["コード", "銘柄名", "業種", "市場"]
DATA_START_COL = 5  # E列(1-indexed)


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
    if not force and date_str in header:
        print(f"  [SKIP] 既に {date_str} 列が存在するため挿入をスキップします(--force で強制上書き可)")
        return

    code_to_rowidx = {}
    for idx, row in enumerate(data_rows):
        if row and row[0]:
            code_to_rowidx[row[0]] = idx

    records_by_code = {r["code"]: r for r in records}
    n_existing_rows = len(data_rows)

    # 新しい列(E列)に入れる値: ヘッダー(日付) + 既存行それぞれの値(無ければ空欄)
    col_values = [date_str]
    for row in data_rows:
        code = row[0] if row else ""
        rec = records_by_code.get(code)
        col_values.append(format_value(metric_key, rec[metric_key]) if rec else "")

    # Step1: E列の前に空列を1つ挿入(既存の日付列は右にずれる)
    ws.spreadsheet.batch_update({
        "requests": [{
            "insertDimension": {
                "range": {
                    "sheetId": ws.id,
                    "dimension": "COLUMNS",
                    "startIndex": DATA_START_COL - 1,
                    "endIndex": DATA_START_COL,
                },
                "inheritFromBefore": False,
            }
        }]
    })

    # Step2: 挿入した列(E列)に日付+各行の値をまとめて書き込み
    end_row = 1 + n_existing_rows
    col = col_letter(DATA_START_COL)
    ws.update(f"{col}1:{col}{end_row}", [[v] for v in col_values],
              value_input_option="USER_ENTERED")

    # Step3: 既存シートに無い新規銘柄は末尾に新しい行として追加
    new_rows = []
    for r in records:
        if r["code"] not in code_to_rowidx:
            new_rows.append([
                r["code"], r["name_jp"], r["industry"] or "", r["market"] or "",
                format_value(metric_key, r[metric_key]),
            ])
    if new_rows:
        ws.append_rows(new_rows, value_input_option="USER_ENTERED")

    print(f"  [OK] 既存{n_existing_rows}行を更新 / 新規{len(new_rows)}行を追加")


LOG_SHEET_TITLE = "_ProcessedDates"


def mark_processed(spreadsheet, date_str: str, n_records: int):
    try:
        log_ws = spreadsheet.worksheet(LOG_SHEET_TITLE)
    except Exception:
        log_ws = spreadsheet.add_worksheet(title=LOG_SHEET_TITLE, rows="1000", cols="3")
        log_ws.update("A1:C1", [["日付", "件数", "実行日時"]])
    log_ws.append_row([date_str, n_records, datetime.now().isoformat(timespec="seconds")],
                       value_input_option="USER_ENTERED")


def write_to_sheets(records: list[dict], target_date: date, dry_run: bool = False, force: bool = False):
    date_str = target_date.strftime("%Y-%m-%d")

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

    for name, gid, key in SHEET_SPECS:
        print(f"[{name}] 更新中...")
        ws = sh.get_worksheet_by_id(gid)
        if ws is None:
            print(f"  [ERROR] gid={gid} のシートが見つかりません。スキップします。", file=sys.stderr)
            continue
        update_wide_sheet(ws, date_str, records, key, force=force)
        time.sleep(1)  # API レート制限対策

    mark_processed(sh, date_str, len(records))


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


# ------------------------------------------------------------------
# main
# ------------------------------------------------------------------

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
