# -*- coding: utf-8 -*-
"""
JPX「規模別・業種別PER・PBR（連結・単体）一覧」（月次Excel）を取得し、
スプレッドシート「00000cloudeデータ」に保存する。

1. 「業種別PER」シート（縦長の蓄積データ）
   年月 / 市場区分名 / Section / 種別 / Industry / 会社数 /
   単純＿PER・PBR・1株当たり当期純利益・1株当たり純資産 /
   加重＿PER・PBR・当期純利益合計・純資産合計
   を毎月追記する（同じ年月が既にあれば追記しない）。
2. 「業種別PER推移」シート（毎回作り直す）
   市場区分名×種別×指標ごとに、月を横に並べ、前月差・前月比・3か月差・12か月差を数式で表示する。

手動実行で FORCE_YM に '202601,202602' のようにカンマ区切りで年月を渡すと、過去分をまとめて取り込める。
"""
import io
import json
import os
import re
import sys
import unicodedata
from datetime import datetime, timedelta, timezone
from urllib.parse import urljoin

import gspread
import pandas as pd
import requests
from google.oauth2.service_account import Credentials

# ===== 定数 =====
JST = timezone(timedelta(hours=9))
PAGE_URL = "https://www.jpx.co.jp/markets/statistics-equities/misc/04.html"
LINK_PATTERN = re.compile(r'href="([^"]*?perpbr(\d{6})\.xlsx)"')
SPREADSHEET_ID = "1QheVVw97DnHjdymEYNFwvgiQhgX8SX-bPxjlHmZpG2I"  # 00000cloudeデータ
DATA_SHEET = "業種別PER"
TREND_SHEET = "業種別PER推移"
TARGET_SHEET_KEYWORD = "連結"   # Excel内でこの文字を含むシートを優先（無ければ見出しが見つかった最初のシート）
MIN_ROWS = 100                   # 取り込み行数がこれ未満なら、表の構造が変わったとみなして中止
CREDENTIALS_ENV = "GOOGLE_SERVICE_ACCOUNT_JSON"
FORCE_YM_ENV = "FORCE_YM"        # 例: '202610' または '202601,202602'。空なら最新月
TIMEOUT_SEC = 60
USER_AGENT = "Mozilla/5.0 (compatible; jpx-industry-per-fetcher/1.1)"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 「業種別PER」シートの見出し（この順で書き込む）
HEADER = [
    "年月", "市場区分名", "Section", "種別", "Industry", "会社数",
    "単純＿PER（倍）", "単純＿PBR（倍）", "単純＿1株当たり\n当期純利益（円）", "単純＿1株当たり\n純資産（円）",
    "加重＿PER（倍）", "加重＿PBR（倍）", "加重＿親会社株主に帰属する\n当期純利益合計（円）", "加重＿純資産合計（円）",
]
# Excelの列を見つけるための手がかり（見出しの文字を正規化して部分一致）
BASE_KEYS = {"市場区分名": ["市場区分名", "市場区分"], "Section": ["section"], "種別": ["種別"],
             "Industry": ["industry"], "会社数": ["会社数"]}
METRIC_KEYS = [("PER", "PER"), ("PBR", "PBR"), ("NI", "当期純利益"), ("NA", "純資産")]
# 指標を見つける手がかり（正規化した見出しに部分一致。NAは当期純利益の列を除く）
METRIC_WORDS = {"PER": ["per", "p/e"], "PBR": ["pbr", "p/b"],
                "NI": ["当期純利益", "純利益", "eps", "netincome"],
                "NA": ["純資産", "bps", "netassets", "bookvalue"]}
GROUP_WORDS = {"単純": ["単純", "simple"], "加重": ["加重", "weighted"]}
HEADER_SCAN_UP = 4   # 「種別」の行から上に何行まで見出しとみなすか

# 推移シートで並べる指標と表示名
TREND_METRICS = [("単純＿PER（倍）", "単純PER"), ("加重＿PER（倍）", "加重PER"),
                 ("単純＿PBR（倍）", "単純PBR"), ("加重＿PBR（倍）", "加重PBR")]
TREND_FIXED_COLS = ["市場区分名", "種別", "指標", "前月差", "前月比", "3か月差", "12か月差"]


# ===== 共通 =====
def norm(s):
    """見出し照合用：全角半角をそろえ、空白・改行を除いて小文字に"""
    s = unicodedata.normalize("NFKC", str(s if s is not None else ""))
    return re.sub(r"\s+", "", s).lower()


def ym_label(ym):
    """'202610' → '2026/10'（既存シートの表記に合わせる）"""
    return f"{ym[:4]}/{ym[4:]}"


def to_cell(v):
    """Excelの値をシートに書ける値へ。数値は数値、'－'などの文字はそのまま"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() else float(v)
    s = str(v).strip()
    try:
        f = float(s.replace(",", ""))
        return int(f) if f.is_integer() else f
    except ValueError:
        return s


def to_num(v):
    """推移計算用：数値に変換できなければ None"""
    try:
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return None


# ===== JPXから取得 =====
def find_links(html):
    links = {ym: urljoin(PAGE_URL, href) for href, ym in LINK_PATTERN.findall(html)}
    if not links:
        raise RuntimeError("perpbrYYYYMM.xlsx のリンクが見つかりません（JPXのページ構成が変わった可能性）")
    return links


def locate_table(df):
    """見出し行を探し、出力列ごとの列番号と、データ開始行を返す。
    見つからない列があれば、見出し部分の中身をログに出して止める（次の修正の手がかり）"""
    nrows, ncol = df.shape
    for r in range(min(nrows, 60)):
        row = [norm(x) for x in df.iloc[r].tolist()]
        if "種別" not in row or "industry" not in row:
            continue
        # 基本列（「種別」の行の前後数行から探す）
        near = range(max(0, r - HEADER_SCAN_UP), min(nrows, r + HEADER_SCAN_UP + 1))
        cols = {}
        for name, keys in BASE_KEYS.items():
            for c in range(ncol):
                if any(norm(df.iat[h, c]) in keys for h in near):
                    cols[name] = c
                    break
        if "会社数" not in cols:
            dump_header(df, near)
            raise RuntimeError("「会社数」の列が見つかりません")
        # データ開始行：会社数が数値または「－」の最初の行
        start = r + 1
        while start < nrows:
            v = to_cell(df.iat[start, cols["会社数"]])
            if isinstance(v, (int, float)) or v in ("－", "-"):
                break
            start += 1
        hdr_rows = list(range(max(0, r - HEADER_SCAN_UP), start))
        base_cols = set(cols.values())

        # 単純／加重のグループ：各行で見つけた語を右へ引き継ぐ（結合セル対策）。上の行を優先
        group = [""] * ncol
        for h in hdr_rows:
            cur = ""
            for c in range(ncol):
                v = norm(df.iat[h, c])
                for g, words in GROUP_WORDS.items():
                    if any(w in v for w in words):
                        cur = g
                if cur and not group[c]:
                    group[c] = cur

        texts = ["".join(norm(df.iat[h, c]) for h in hdr_rows) for c in range(ncol)]
        for key, _ in METRIC_KEYS:
            cands = []
            for c in range(ncol):
                if c in base_cols:
                    continue
                t = texts[c]
                if key == "NA" and any(w in t for w in METRIC_WORDS["NI"]):
                    continue
                if key == "PER" and "pbr" in t:
                    continue
                if any(w in t for w in METRIC_WORDS[key]):
                    cands.append(c)
            for g in ("単純", "加重"):
                hit = [c for c in cands if group[c] == g]
                if hit:
                    cols[(key, g)] = hit[0]
            # グループ名が無い表：左から1つ目を単純、2つ目を加重とみなす
            if (key, "単純") not in cols and (key, "加重") not in cols and len(cands) >= 2:
                cols[(key, "単純")], cols[(key, "加重")] = cands[0], cands[1]
                print(f"注意: {key} は単純/加重の見出しが無いため、左から順に割り当てました（列{cands[0] + 1}, 列{cands[1] + 1}）")

        need = list(BASE_KEYS) + [(k, g) for g in ("単純", "加重") for k, _ in METRIC_KEYS]
        missing = [n for n in need if n not in cols]
        if missing:
            dump_header(df, hdr_rows)
            raise RuntimeError(f"見出しの列が見つかりません: {missing}")
        print("列の対応: " + ", ".join(f"{k}=列{v + 1}" for k, v in cols.items()))
        return cols, start
    return None, None


def dump_header(df, rows):
    """見出し部分をログに出す（列番号: 値）"""
    print("---- 見出し部分の中身（原因調査用）----")
    for h in rows:
        cells = [f"{c + 1}:{str(df.iat[h, c]).strip()!r}" for c in range(df.shape[1])
                 if str(df.iat[h, c]).strip() not in ("", "nan", "None")]
        print(f"行{h + 1}: " + " | ".join(cells))
    print("--------------------------------------")


def parse_excel(content, ym):
    sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None, dtype=object)
    names = list(sheets)
    print(f"Excelのシート: {names}")
    order = [n for n in names if TARGET_SHEET_KEYWORD in str(n)] + [n for n in names if TARGET_SHEET_KEYWORD not in str(n)]
    for name in order:
        df = sheets[name]
        cols, start = locate_table(df)
        if cols is None:
            continue
        print(f"使用するシート: {name}（データ開始 {start + 1}行目）")
        out = []
        for r in range(start, len(df)):
            kind = to_cell(df.iat[r, cols["種別"]])
            count = to_cell(df.iat[r, cols["会社数"]])
            if kind == "" or count == "":
                continue
            if norm(kind) == "種別":   # 途中に見出しが繰り返されている場合
                continue
            row = [ym_label(ym)]
            row += [to_cell(df.iat[r, cols[n]]) for n in ("市場区分名", "Section", "種別", "Industry", "会社数")]
            for g in ("単純", "加重"):
                row += [to_cell(df.iat[r, cols[(k, g)]]) for k, _ in METRIC_KEYS]
            out.append(row)
        if len(out) < MIN_ROWS:
            raise RuntimeError(f"取り込み行数が{len(out)}行で少なすぎます（表の構造が変わった可能性）")
        return out
    raise RuntimeError("「種別」「Industry」の見出しがあるシートが見つかりません")


# ===== スプレッドシート =====
def get_or_create(ss, title, rows, cols):
    try:
        return ss.worksheet(title)
    except gspread.WorksheetNotFound:
        return ss.add_worksheet(title=title, rows=rows, cols=cols)


def append_month(ws, rows):
    if not ws.get_values("A1:A1"):
        ws.update(range_name="A1", values=[HEADER], value_input_option="RAW")
    ws.append_rows(rows, value_input_option="RAW", table_range="A1")


def build_trend(ss, data_ws):
    """蓄積データから推移シートを作り直す"""
    values = data_ws.get_all_values()
    if len(values) < 2:
        return
    head = values[0]
    idx = {h: i for i, h in enumerate(head)}
    have = sorted({r[idx["年月"]] for r in values[1:] if r and re.fullmatch(r"\d{4}/\d{2}", r[idx["年月"]])}, reverse=True)
    if not have:
        return
    latest = have[0]
    # 抜けている月も空欄の列として並べる（3か月差・12か月差が暦どおりの月を指すように）
    months = []
    y, m = int(latest[:4]), int(latest[5:])
    oldest = have[-1]
    while True:
        label = f"{y}/{m:02d}"
        months.append(label)
        if label <= oldest:
            break
        y, m = (y, m - 1) if m > 1 else (y - 1, 12)
    # 行の並びは最新月の並び順に合わせる
    keys, seen = [], set()
    for r in values[1:]:
        if r[idx["年月"]] == latest:
            k = (r[idx["市場区分名"]], r[idx["種別"]])
            if k not in seen:
                seen.add(k)
                keys.append(k)
    table = {}
    for r in values[1:]:
        table[(r[idx["年月"]], r[idx["市場区分名"]], r[idx["種別"]])] = r

    month_cols = [f"{m[:4]}年{int(m[5:])}月" for m in months]
    header = TREND_FIXED_COLS + month_cols
    first_month_col = len(TREND_FIXED_COLS) + 1  # 1始まりの列番号（H列）
    out = [header]
    for market, kind in keys:
        for src, label in TREND_METRICS:
            row = [market, kind, label, "", "", "", ""]
            for m in months:
                rec = table.get((m, market, kind))
                v = to_num(rec[idx[src]]) if rec else None
                row.append(v if v is not None else "")
            out.append(row)

    ws = get_or_create(ss, TREND_SHEET, rows=len(out) + 10, cols=len(header) + 2)
    ws.clear()
    ws.resize(rows=max(len(out) + 10, 50), cols=max(len(header), len(TREND_FIXED_COLS) + 13))
    ws.update(range_name="A1", values=out, value_input_option="RAW")

    # 前月差・前月比・3か月差・12か月差は数式（最新月列との差）
    def a1(col):
        return gspread.utils.rowcol_to_a1(1, col).rstrip("1")
    cur = a1(first_month_col)
    prev = a1(first_month_col + 1)
    m3 = a1(first_month_col + 3)
    m12 = a1(first_month_col + 12)
    formulas = []
    for r in range(2, len(out) + 1):
        formulas.append([
            f'=IF(OR({cur}{r}="",{prev}{r}=""),"",{cur}{r}-{prev}{r})',
            f'=IF(OR({cur}{r}="",{prev}{r}="",{prev}{r}=0),"",{cur}{r}/{prev}{r}-1)',
            f'=IF(OR({cur}{r}="",{m3}{r}=""),"",{cur}{r}-{m3}{r})',
            f'=IF(OR({cur}{r}="",{m12}{r}=""),"",{cur}{r}-{m12}{r})',
        ])
    ws.update(range_name=f"D2:G{len(out)}", values=formulas, value_input_option="USER_ENTERED")
    format_trend(ss, ws, len(out), len(header))
    print(f"推移シートを更新: {len(out) - 1}行 × {len(months)}か月")


def format_trend(ss, ws, nrows, ncols):
    sid = ws.id
    meta = ss.fetch_sheet_metadata({"fields": "sheets(properties(sheetId),conditionalFormats)"})
    has_cf = any(s["properties"]["sheetId"] == sid and s.get("conditionalFormats") for s in meta["sheets"])

    def rng(c0, c1, r0=1, r1=None):
        g = {"sheetId": sid, "startRowIndex": r0, "startColumnIndex": c0, "endColumnIndex": c1}
        if r1 is not None:
            g["endRowIndex"] = r1
        return g

    reqs = [
        {"updateSheetProperties": {"properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1, "frozenColumnCount": 3}},
                                   "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        {"repeatCell": {"range": rng(0, ncols, 0, 1), "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                        "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}}},
                        "fields": "userEnteredFormat.textFormat.bold,userEnteredFormat.backgroundColor"}},
        {"repeatCell": {"range": rng(3, 4), "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "+0.0;-0.0;0.0"}}},
                        "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {"range": rng(4, 5), "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "+0.0%;-0.0%;0.0%"}}},
                        "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {"range": rng(5, 7), "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "+0.0;-0.0;0.0"}}},
                        "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {"range": rng(7, ncols), "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0.0"}}},
                        "fields": "userEnteredFormat.numberFormat"}},
    ]
    if not has_cf:
        # 変化の方向が一目で分かるように：上昇は赤、低下は青（D〜G列）
        ranges = [rng(3, 7)]
        reqs += [
            {"addConditionalFormatRule": {"rule": {"ranges": ranges, "booleanRule": {
                "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0"}]},
                "format": {"textFormat": {"foregroundColor": {"red": 0.8, "green": 0, "blue": 0}}}}}, "index": 0}},
            {"addConditionalFormatRule": {"rule": {"ranges": ranges, "booleanRule": {
                "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0.8}}}}}, "index": 1}},
        ]
    ss.batch_update({"requests": reqs})


# ===== メイン =====
def main():
    creds_json = os.environ.get(CREDENTIALS_ENV)
    if not creds_json:
        sys.exit(f"環境変数 {CREDENTIALS_ENV} が未設定です")
    force = [x.strip() for x in (os.environ.get(FORCE_YM_ENV) or "").split(",") if x.strip()]

    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT
    page = session.get(PAGE_URL, timeout=TIMEOUT_SEC)
    page.raise_for_status()
    page.encoding = page.apparent_encoding or "utf-8"
    links = find_links(page.text)
    targets = force or [max(links)]
    missing = [ym for ym in targets if ym not in links]
    if missing:
        sys.exit(f"ページに無い年月: {missing}。掲載月: {sorted(links)}")

    creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
    ss = gspread.authorize(creds).open_by_key(SPREADSHEET_ID)
    data_ws = get_or_create(ss, DATA_SHEET, rows=1000, cols=len(HEADER))
    existing = set(data_ws.col_values(1)[1:])

    added = 0
    for ym in sorted(targets):
        label = ym_label(ym)
        if label in existing:
            print(f"{label} は登録済みのためスキップ")
            continue
        res = session.get(links[ym], timeout=TIMEOUT_SEC)
        res.raise_for_status()
        rows = parse_excel(res.content, ym)
        append_month(data_ws, rows)
        existing.add(label)
        added += 1
        print(f"{label}: {len(rows)}行を追記")

    build_trend(ss, data_ws)
    print(f"完了（追記 {added}か月）")


if __name__ == "__main__":
    main()
