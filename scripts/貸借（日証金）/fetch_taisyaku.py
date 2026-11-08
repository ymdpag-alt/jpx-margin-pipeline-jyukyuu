"""
日証金「銘柄別融資・貸株残高一覧表」(zandaka.csv) を毎営業日取得し、
Googleスプレッドシートへ「銘柄コード｜銘柄名｜日付…（新しい日付が左）」形式で書き込む。

- 公表：毎営業日 11:30 頃（前営業日の申込日＝約定日の値）
- 同じ日付の列が既にあれば上書き（何回実行しても二重にならない）
- 生CSVは data/taisyaku/ に保存（ワークフロー側でコミットし、作り直しに使う）

必要な環境変数
  GCP_SA_KEY         サービスアカウントのJSON（文字列）
  TAISYAKU_SHEET_ID  書き込み先スプレッドシートID
"""
import csv
import io
import json
import os
import re
import sys
import time
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from urllib.parse import urljoin

import requests

# ===== 定数 =====
DOWNLOAD_PAGE = "https://www.taisyaku.jp/download/"
# ページからリンクが取れないときの予備（推測のURL。初回に実URLを確認して直すこと）
FALLBACK_URLS = ["https://www.taisyaku.jp/data/zandaka.csv"]
USER_AGENT = "Mozilla/5.0 (compatible; market-data-collector/1.0)"

# CSVの項目名 → シート名
SHEETS = {
    "融資新規": "証金融資新規",
    "融資返済": "証金融資返済",
    "融資残高": "証金融資残高",
    "貸株新規": "証金貸株新規",
    "貸株返済": "証金貸株返済",
    "貸株残高": "証金貸株残高",
}
MARKET_KEEP = "東証"          # 取引所区分名がこれを含む行だけ使う（実CSVは「東証およびＰＴＳ」）
MAX_DATE_COLS = 260           # 残す日付列数（約1年）。セル上限1000万に注意
NAME_MISMATCH_STOP = 0.05     # 既存行と銘柄名の不一致率がこれを超えたら行ずれとみなし中止
STALE_DAYS = 7                # 申込日がこれより古ければ何もしない
RAW_DIR = "data/taisyaku"
WEEKDAY = "月火水木金土日"
JST = timezone(timedelta(hours=9))


# ===== 取得 =====
def http_get(url, tries=3):
    last = None
    for i in range(tries):
        try:
            r = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=30)
            r.raise_for_status()
            return r
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(5 * (i + 1))
    raise RuntimeError(f"取得失敗: {url}: {last}")


def find_csv_url():
    try:
        html = http_get(DOWNLOAD_PAGE).text
        m = re.search(r'href="([^"]*zandaka[^"]*\.csv[^"]*)"', html, re.I)
        if m:
            return urljoin(DOWNLOAD_PAGE, m.group(1))
        print("警告: ダウンロードページに zandaka.csv のリンクが見つからない。予備URLを試す")
    except Exception as e:  # noqa: BLE001
        print(f"警告: ダウンロードページ取得失敗 {e}")
    return FALLBACK_URLS[0]


def decode(raw: bytes) -> str:
    for enc in ("utf-8-sig", "cp932"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise RuntimeError("CSVの文字コードを判定できない")


# ===== 解析 =====
def norm_code(s: str) -> str:
    s = s.strip().upper()
    if len(s) == 5 and s.endswith("0"):   # 5桁表記（末尾0）なら4桁に
        s = s[:4]
    return s


def to_num(s: str):
    s = (s or "").replace(",", "").strip()
    if s in ("", "-", "－"):
        return ""
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return ""


def parse_date(s: str):
    d = re.sub(r"\D", "", s or "")
    if len(d) == 8:
        return date(int(d[:4]), int(d[4:6]), int(d[6:]))
    return None


def parse_csv(text: str):
    rows = list(csv.reader(io.StringIO(text)))
    hi = next((i for i, r in enumerate(rows) if any("申込日" in c for c in r) and any("コード" in c for c in r)), None)
    if hi is None:
        raise RuntimeError("見出し行（申込日・コード）が見つからない。CSVの形式が変わった可能性")
    head = [c.strip() for c in rows[hi]]

    def col(key):
        # 実CSVは「融資新規株数」「融資新規金額」「融資・新規回転日数」が並ぶので株数を優先する
        for cand in (key + "株数", key):
            if cand in head:
                return head.index(cand)
        for i, h in enumerate(head):
            if key in h and "金額" not in h and "回転" not in h:
                return i
        raise RuntimeError(f"列が見つからない: {key}（見出し: {head}）")

    i_d, i_c, i_n = col("申込日"), col("コード"), col("銘柄名")
    # 市場の列は「取引所区分名」（例：東証およびＰＴＳ／名証／福証／札証）
    i_m = next((i for i, h in enumerate(head) if "取引所" in h or "市場" in h), None)
    i_s = next((i for i, h in enumerate(head) if "速報" in h or "確報" in h), None)
    i_f = {k: col(k) for k in SHEETS}

    recs, dates, kinds = {}, Counter(), Counter()
    for r in rows[hi + 1:]:
        if len(r) <= max(i_f.values()):
            continue
        if i_m is not None and MARKET_KEEP not in r[i_m]:
            continue
        code = norm_code(r[i_c])
        if not re.fullmatch(r"[0-9]{3}[0-9A-Z][0-9A-Z]?", code):
            continue
        if i_s is not None:
            kinds[r[i_s].strip()] += 1
        d = parse_date(r[i_d])
        if d:
            dates[d] += 1
        if code in recs:
            continue
        rec = {"name": r[i_n].strip()}
        for k, i in i_f.items():
            rec[k] = to_num(r[i])
        recs[code] = rec
    if not recs or not dates:
        raise RuntimeError("有効な行が0件")
    d, _ = dates.most_common(1)[0]
    if kinds:
        print(f"速報／確報: {dict(kinds)}（同じ申込日の確報が後で来れば上書きされる）")
    return d, recs


def label_of(d: date) -> str:
    return f"{d.year}年{d.month}月{d.day}日({WEEKDAY[d.weekday()]})"


# ===== 書き込み =====
def update_sheet(ss, title, field, label, recs):
    import gspread
    from gspread.utils import rowcol_to_a1

    try:
        ws = ss.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=title, rows=10, cols=3)
        ws.update(range_name="A1:B1", values=[["銘柄コード", "銘柄名"]])

    header = ws.row_values(1)
    ab = ws.get_values("A2:B")
    codes = [str(r[0]).strip().upper() for r in ab if r and r[0]]
    names = {str(r[0]).strip().upper(): (r[1] if len(r) > 1 else "") for r in ab if r and r[0]}

    # 行ずれ検査：既存行の銘柄名がCSVと大きく食い違えば止める
    both = [c for c in codes if c in recs and names.get(c)]
    if len(both) > 100:
        bad = sum(1 for c in both if names[c] != recs[c]["name"])
        if bad / len(both) > NAME_MISMATCH_STOP:
            raise RuntimeError(f"{title}: 銘柄名の不一致 {bad}/{len(both)}。行ずれの疑いで中止")

    new = [c for c in recs if c not in names]
    if new:
        ws.append_rows([[c, recs[c]["name"]] for c in new], value_input_option="RAW")
        codes += new

    colvals = [label] + [recs[c][field] if c in recs else "" for c in codes]
    if label in header:
        ci = header.index(label) + 1
        rng = f"{rowcol_to_a1(1, ci)}:{rowcol_to_a1(len(colvals), ci)}"
        ws.update(range_name=rng, values=[[v] for v in colvals], value_input_option="RAW")
        act = "上書き"
    else:
        ws.insert_cols([colvals], col=3, value_input_option="RAW")
        act = "追加"

    if ws.col_count > 2 + MAX_DATE_COLS:
        ws.delete_columns(3 + MAX_DATE_COLS, ws.col_count)
    print(f"{title}: {label} を{act}（{len(recs)}銘柄、新規行 {len(new)}）")


def main():
    url = find_csv_url()
    print(f"取得: {url}")
    raw = http_get(url).content
    d, recs = parse_csv(decode(raw))
    today = datetime.now(JST).date()
    print(f"申込日 {d}（{len(recs)}銘柄）")
    if (today - d).days > STALE_DAYS:
        print(f"申込日が古い（{d}）。未更新とみなして終了")
        return

    os.makedirs(RAW_DIR, exist_ok=True)
    with open(os.path.join(RAW_DIR, f"zandaka_{d:%Y%m%d}.csv"), "wb") as f:
        f.write(raw)

    if os.environ.get("DRY_RUN") == "1":
        print("DRY_RUN: シートには書かない")
        return

    import gspread
    gc = gspread.service_account_from_dict(json.loads(os.environ["GCP_SA_KEY"]))
    ss = gc.open_by_key(os.environ["TAISYAKU_SHEET_ID"])
    label = label_of(d)
    for field, title in SHEETS.items():
        update_sheet(ss, title, field, label, recs)
        time.sleep(1)   # 書き込みAPIの毎分上限対策


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # noqa: BLE001
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)
