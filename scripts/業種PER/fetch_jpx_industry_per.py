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
TARGET_SHEET_KEYWORD = "連結"
MARKET_SHEET_KEYWORD = "単体"   # 「市場別（単体）」シート：市場ごとの行（種別「－」）を追加で取り込む
JPX_SYMBOLS = ("－", "＊", "-", "*")
MARKET_NAMES = ("プライム市場", "スタンダード市場", "グロース市場")  # JPXの記号（－：該当なし・マイナス、＊：PER1000倍以上）   # Excel内でこの文字を含むシートを優先（無ければ見出しが見つかった最初のシート）
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
NOTE_MAX_LEN = 40    # これより長いセルは注記とみなして見出しの判定に使わない

# 推移シートで並べる指標と表示名
TREND_METRICS = [("単純＿PER（倍）", "単純PER"), ("加重＿PER（倍）", "加重PER"),
                 ("単純＿PBR（倍）", "単純PBR"), ("加重＿PBR（倍）", "加重PBR")]
TREND_FIXED_COLS = ["市場区分名", "種別", "指標", "前月差", "前月比", "3か月差", "12か月差"]
# JPXが「－」（マイナス）「＊」（1000倍以上）で公表しない値を、同じ行の他の数値から計算するための式
#   株価（単純平均）＝PBR×1株純資産＝PER×1株利益、時価総額（加重）＝PBR×純資産合計＝PER×純利益合計
ESTIMATE_RULES = {
    "単純＿PER（倍）": ("単純＿PBR（倍）", "単純＿1株当たり\n純資産（円）", "単純＿1株当たり\n当期純利益（円）"),
    "加重＿PER（倍）": ("加重＿PBR（倍）", "加重＿純資産合計（円）", "加重＿親会社株主に帰属する\n当期純利益合計（円）"),
    "単純＿PBR（倍）": ("単純＿PER（倍）", "単純＿1株当たり\n当期純利益（円）", "単純＿1株当たり\n純資産（円）"),
    "加重＿PBR（倍）": ("加重＿PER（倍）", "加重＿親会社株主に帰属する\n当期純利益合計（円）", "加重＿純資産合計（円）"),
}
NO_COMPANY_TEXT = "対象なし"   # 会社数が「－」（その市場・業種に会社が無い）
ESTIMATE_NOTE = ("灰色の斜体＝推計値。JPXが「－」（PER・PBRがマイナス）や「＊」（PER1000倍以上）で公表していない月を、"
                 "同じ月の数値から計算（PER＝PBR×1株純資産÷1株利益、加重はPBR×純資産合計÷純利益合計。PBRも同様）。"
                 "マイナスは赤字（または債務超過）。元のPBR・PERが小数1桁のため誤差あり。「対象なし」＝その区分に会社が無い。")


def estimate(rec, idx, src):
    """公表されていない値を計算する。計算できなければ None"""
    rule = ESTIMATE_RULES.get(src)
    if not rule or any(k not in idx for k in rule):
        return None
    a, b, c = (to_num(rec[idx[k]]) for k in rule)
    if a is None or b is None or c is None or c == 0:
        return None
    return round(a * b / c, 1)


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
    # 同じ年月が2回出る場合は最初のリンクを使う（ページ下部に「参考のサンプルファイル」perpbr202111.xlsx があるため）
    links = {}
    for href, ym in LINK_PATTERN.findall(html):
        links.setdefault(ym, urljoin(PAGE_URL, href))
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
        # まず、見出しが保存用の見出し（HEADER）と同じ文字かで探す（今のJPXの表はこれで全列そろう）
        exact = exact_columns(row)
        if exact:
            start = r + 1
            while start < nrows:
                v = to_cell(df.iat[start, exact["会社数"]])
                if isinstance(v, (int, float)) or v in ("－", "-"):
                    break
                start += 1
            print("列の対応（見出し一致）: " + ", ".join(f"{k}=列{v + 1}" for k, v in exact.items()))
            return exact, start
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
                if len(v) > NOTE_MAX_LEN:
                    continue
                for g, words in GROUP_WORDS.items():
                    if any(w in v for w in words):
                        cur = g
                if cur and not group[c]:
                    group[c] = cur

        # 注記のような長い文章（「PER・PBRがマイナス…」など）は見出しとみなさない
        texts = ["".join(norm(df.iat[h, c]) for h in hdr_rows if len(norm(df.iat[h, c])) <= NOTE_MAX_LEN)
                 for c in range(ncol)]
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


def exact_columns(row):
    """見出し行の文字がHEADERと一致する列を返す（全部そろわなければ None）"""
    pos = {}
    for c, v in enumerate(row):
        if v and v not in pos:
            pos[v] = c
    names = {"市場区分名": "市場区分名", "Section": "Section", "種別": "種別", "Industry": "Industry", "会社数": "会社数"}
    cols = {}
    for name, label in names.items():
        c = pos.get(norm(label))
        if c is None:
            return None
        cols[name] = c
    metric_heads = HEADER[6:]   # 単純4項目、加重4項目の順
    keys = [(k, g) for g in ("単純", "加重") for k, _ in METRIC_KEYS]
    for key, label in zip(keys, metric_heads):
        c = pos.get(norm(label))
        if c is None:
            return None
        cols[key] = c
    return cols


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
        have = {(row[1], row[3]) for row in out}
        out += [row for row in parse_market_sheet(sheets, ym, name) if (row[1], row[3]) not in have]
        return out
    raise RuntimeError("「種別」「Industry」の見出しがあるシートが見つかりません")


def metric_col(row, key):
    """見出し行から「単純＿PER」「加重＿当期純利益」のような列を探す。
    単体シートは見出しの文字が少し違う（例：加重＿当期純利益（円）、加重＿純資産（円））ため、
    先頭の「単純／加重」と指標の言葉で判定する"""
    metric, group = key
    for c, v in enumerate(row):
        if not v.startswith(norm(group)):
            continue
        if metric == "NA":
            if "純資産" in v and "当期純利益" not in v:
                return c
        elif metric == "NI":
            if "当期純利益" in v:
                return c
        elif norm(metric) in v:
            return c
    return None


def parse_market_sheet(sheets, ym, used_name):
    """「市場別（単体）」シートから市場ごとの行を取り込む（種別・Industryは「－」）。
    うまく読めなくても全体は止めず、警告だけ出す（連結の業種データを優先）"""
    names = [n for n in sheets if MARKET_SHEET_KEYWORD in str(n) and n != used_name]
    if not names:
        print(f"注意: 「{MARKET_SHEET_KEYWORD}」のシートが無いため、市場別の行は取り込みません")
        return []
    df = sheets[names[0]]
    nrows, ncol = df.shape
    for r in range(min(nrows, 60)):
        row = [norm(x) for x in df.iloc[r].tolist()]
        if "市場区分名" not in row or "会社数" not in row:
            continue
        pos = {}
        for c, v in enumerate(row):
            if v and v not in pos:
                pos[v] = c
        cols = {name: pos.get(norm(name)) for name in ("市場区分名", "Section", "種別", "Industry", "会社数")}
        keys = [(k, g) for g in ("単純", "加重") for k, _ in METRIC_KEYS]
        for key in keys:
            cols[key] = metric_col(row, key)
        lack = [k for k, v in cols.items() if v is None and k not in ("種別", "Industry", "Section")]
        if lack:
            dump_header(df, range(max(0, r - 2), min(nrows, r + 2)))
            print(f"注意: 「{names[0]}」の見出しが見つからない列があるため、市場別の行は取り込みません: {lack}")
            return []
        out = []
        for i in range(r + 1, nrows):
            count = to_cell(df.iat[i, cols["会社数"]])
            market = to_cell(df.iat[i, cols["市場区分名"]])
            if not (isinstance(count, (int, float)) or count in JPX_SYMBOLS) or market == "":
                continue
            kind = to_cell(df.iat[i, cols["種別"]]) if cols["種別"] is not None else "－"
            if kind not in ("", "－", "-"):
                continue   # 連結シートの業種の行と重ならないよう、市場全体の行だけにする
            out.append([ym_label(ym), market,
                        to_cell(df.iat[i, cols["Section"]]) if cols["Section"] is not None else "",
                        "－",
                        to_cell(df.iat[i, cols["Industry"]]) if cols["Industry"] is not None else "－",
                        count]
                       + [to_cell(df.iat[i, cols[k]]) for k in keys])
        print(f"「{names[0]}」から市場別の行を{len(out)}行取り込み")
        return out
    print(f"注意: 「{names[0]}」に見出し行（市場区分名・会社数）が見つからないため、市場別の行は取り込みません")
    return []


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
    # 最新月に無い行（過去の月だけにある行）も消さずに、後ろに並べる
    keys, seen = [], set()
    for want_latest in (True, False):
        for r in values[1:]:
            if (r[idx["年月"]] == latest) != want_latest or not re.fullmatch(r"\d{4}/\d{2}", r[idx["年月"]]):
                continue
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
    est_cells = []   # 推計値のセル（行・列は0始まり）
    for market, kind in keys:
        for src, label in TREND_METRICS:
            row = [market, kind, label, "", "", "", ""]
            for m in months:
                rec = table.get((m, market, kind))
                if not rec:
                    row.append("")            # その月のデータ自体が無い
                    continue
                raw = str(rec[idx[src]]).strip()
                v = to_num(raw)
                if v is not None:
                    row.append(v)
                    continue
                if str(rec[idx["会社数"]]).strip() in JPX_SYMBOLS:
                    row.append(NO_COMPANY_TEXT)
                    continue
                e = estimate(rec, idx, src)
                if e is not None:
                    est_cells.append((len(out), len(row)))
                    row.append(e)
                else:
                    row.append(raw if raw in JPX_SYMBOLS else "")
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
            f'=IF(AND(ISNUMBER({cur}{r}),ISNUMBER({prev}{r})),{cur}{r}-{prev}{r},"")',
            f'=IF(AND(ISNUMBER({cur}{r}),ISNUMBER({prev}{r}),N({prev}{r})<>0),{cur}{r}/{prev}{r}-1,"")',
            f'=IF(AND(ISNUMBER({cur}{r}),ISNUMBER({m3}{r})),{cur}{r}-{m3}{r},"")',
            f'=IF(AND(ISNUMBER({cur}{r}),ISNUMBER({m12}{r})),{cur}{r}-{m12}{r},"")',
        ])
    ws.update(range_name=f"D2:G{len(out)}", values=formulas, value_input_option="USER_ENTERED")
    format_trend(ss, ws, len(out), len(header), est_cells)
    print(f"推計値で補ったセル: {len(est_cells)}")
    print(f"推移シートを更新: {len(out) - 1}行 × {len(months)}か月")


def format_trend(ss, ws, nrows, ncols, est_cells=()):
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
        # 月の列：書式を毎回いったん標準に戻す（前回の推計値の色が別のセルに残らないように）
        {"repeatCell": {"range": rng(7, ncols), "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "NUMBER", "pattern": "0.0"},
            "textFormat": {"italic": False, "foregroundColor": {"red": 0, "green": 0, "blue": 0}}}},
                        "fields": "userEnteredFormat.numberFormat,userEnteredFormat.textFormat.italic,userEnteredFormat.textFormat.foregroundColor"}},
        # 見出し「指標」のセルに凡例のメモ
        {"updateCells": {"range": rng(2, 3, 0, 1), "rows": [{"values": [{"note": ESTIMATE_NOTE}]}], "fields": "note"}},
    ]
    for r, c in est_cells:
        reqs.append({"repeatCell": {"range": rng(c, c + 1, r, r + 1), "cell": {"userEnteredFormat": {
            "textFormat": {"italic": True, "foregroundColor": {"red": 0.45, "green": 0.45, "blue": 0.45}}}},
            "fields": "userEnteredFormat.textFormat.italic,userEnteredFormat.textFormat.foregroundColor"}})
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
    # 登録済みの行（年月・市場区分名・種別）。登録済みの月でも、足りない行（市場別の行など）があれば追記する
    all_rows = data_ws.get_all_values()
    head = all_rows[0] if all_rows else HEADER
    ci = {h: i for i, h in enumerate(head)}
    existing = {(r[ci["年月"]], r[ci["市場区分名"]], r[ci["種別"]]) for r in all_rows[1:] if len(r) > ci["種別"]}
    months_have = {k[0] for k in existing}
    markets_have = {k[0] for k in existing if k[2] in ("－", "-") and k[1] in MARKET_NAMES}

    # 登録済みなのに市場別の行が無い月は、JPXのページにあれば自動で補う
    auto = sorted({k[0].replace("/", "") for k in existing if re.fullmatch(r"\d{4}/\d{2}", k[0])}
                  - {m.replace("/", "") for m in markets_have} - set(targets))
    auto = [ym for ym in auto if ym in links]
    if auto:
        print(f"市場別の行が無い月を補います: {auto}")
        targets = list(targets) + auto

    added = 0
    for ym in sorted(targets):
        label = ym_label(ym)
        if label in months_have and label in markets_have:
            print(f"{label} は登録済み（市場別の行もあり）のためスキップ")
            continue
        res = session.get(links[ym], timeout=TIMEOUT_SEC)
        res.raise_for_status()
        rows = parse_excel(res.content, ym)
        new_rows = [r for r in rows if (r[0], r[1], r[3]) not in existing]
        if not new_rows:
            print(f"{label}: 追記する行なし")
            continue
        append_month(data_ws, new_rows)
        existing |= {(r[0], r[1], r[3]) for r in new_rows}
        added += 1
        print(f"{label}: {len(new_rows)}行を追記" + ("（登録済みの月に不足分を追加）" if label in months_have else ""))
        months_have.add(label)

    build_trend(ss, data_ws)
    print(f"完了（追記 {added}か月）")


if __name__ == "__main__":
    main()
