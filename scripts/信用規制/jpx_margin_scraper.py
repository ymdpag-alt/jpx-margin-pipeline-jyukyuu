#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JPX 信用取引規制データ 一括取得スクリプト

以下2つのJPXページを1回の実行でまとめて取得し、それぞれ対応するシートへ
差分反映する。

    1. 信用取引に関する規制等（増担保規制）
       https://www.jpx.co.jp/markets/equities/margin-reg/index.html
       → 増担保規制シート (gid=1552537213)

    2. 信用取引に関する日々公表（日々公表銘柄）
       https://www.jpx.co.jp/markets/equities/margin-daily/index.html
       → 日々公表銘柄フラグシート (gid=173915474)

シート列構成（両シート共通）:
    A: 新規記載日   B: コード   C: 銘柄名   D: 実施日（＝指定日）
    E: 解除日       F: 規制の内容   G: 該当基準

差分反映ロジック（両フィード共通）:
    - (コード, 実施日/指定日) の組み合わせが既存シートに存在する行はそのまま
    - 存在しない行（＝新規）は2行目にまとめて挿入し、既存データは1行ずつ下にずれる
      （E列＝解除日は空欄のまま）
    - 「解除された銘柄」一覧は、コードが一致しE列が空欄の行を探して解除日を記入
    - 対応する未解除行が見つからない場合は、情報を失わないよう新規行として追加

HTML解析の注意点:
    JPXのテーブルはヘッダー行を <th> ではなく <td> で組んでいるため、
    pandas.read_html が列名を 0,1,2... と振り、ヘッダーが本文1行目に
    入った状態で返ってくる。_match_table() でヘッダー行を検出して
    列名に昇格させることで、<td>・<th> どちらの構成でも解析できるようにしている。

日付の扱い:
    JPXのページは常に「最新の状態」のスナップショットであり、過去日を指定して
    別のデータを取得することはできない。--date は「新規記載日（A列）」に書き込む
    基準日を指定するためのオプションであり、省略時は実行日（JST）を使用する。

非営業日のスキップについて:
    JPXのページは非営業日（土日・祝日・年末年始）には更新されない。無駄な
    実行を避けるため、対象日（--date指定日、または実行日）が非営業日の場合は
    デフォルトで処理をスキップする。--force を付けると営業日判定を無視して
    強制的に取得・反映する（手動バックフィル等で使用）。

使い方:
    python jpx_margin_scraper.py                    # 増担保規制・日々公表銘柄を一括実行（既定）
    python jpx_margin_scraper.py --date 2026-08-15   # 新規記載日を指定して一括実行
    python jpx_margin_scraper.py --only reg          # 増担保規制のみ
    python jpx_margin_scraper.py --only daily        # 日々公表銘柄のみ
    python jpx_margin_scraper.py --force             # 非営業日でも強制実行
"""

import argparse
import json
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from io import StringIO
from typing import Dict, List, Optional

import gspread
import jpholiday
import numpy as np
import pandas as pd
import requests
from google.oauth2.service_account import Credentials
from zoneinfo import ZoneInfo


# =============================================================================
# 共通設定
# =============================================================================
def _env_str(name: str, default: str) -> str:
    # 環境変数が未設定、または空文字列("")の場合の両方でデフォルト値を使う。
    # GitHub Actionsでは未設定のsecretsを参照すると空文字列になるため、
    # os.environ.get(name, default) だけでは意図通りにフォールバックしない。
    val = os.environ.get(name)
    return val if val else default


def _env_int(name: str, default: int) -> int:
    val = os.environ.get(name)
    return int(val) if val else default


SPREADSHEET_ID = _env_str("SPREADSHEET_ID", "1kWST0CkkIvo3irPSbMgtVtUqqRDXFwvRREYZDRQAFMY")
MARGIN_REG_GID = _env_int("MARGIN_REG_GID", 1552537213)
MARGIN_DAILY_GID = _env_int("MARGIN_DAILY_GID", 173915474)

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}

SHEET_HEADER = ["新規記載日", "コード", "銘柄名", "実施日", "解除日", "規制の内容", "該当基準"]


@dataclass
class NewRecord:
    code: str
    start_date: str
    name: str
    content: str = ""
    criteria: str = ""


@dataclass
class ReleasedRecord:
    code: str
    end_date: str
    name: str
    content: str = ""


# =============================================================================
# 日付ユーティリティ
# =============================================================================
def get_target_datetime(date_str: Optional[str]) -> datetime:
    if date_str:
        return datetime.strptime(date_str, "%Y-%m-%d")
    return datetime.now(ZoneInfo("Asia/Tokyo"))


def is_business_day(dt: datetime) -> bool:
    """土日・祝日・年末年始（12/31〜1/3）を非営業日として判定する"""
    if dt.weekday() >= 5:  # 土(5)・日(6)
        return False
    if (dt.month == 1 and dt.day <= 3) or (dt.month == 12 and dt.day == 31):
        return False  # 取引所の年末年始休業日
    if jpholiday.is_holiday(dt.date()):
        return False
    return True


# =============================================================================
# ページ取得
# =============================================================================
def fetch_html(url: str) -> str:
    print(f"取得中: {url}")
    resp = requests.get(url, headers=REQUEST_HEADERS, timeout=30)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding
    return resp.text


# =============================================================================
# HTML解析 共通ユーティリティ
# =============================================================================
def _to_text(v) -> str:
    """セル値を安全に文字列化する。

    ヘッダーが<th>で組まれている場合、コード列が数値型として読み込まれ、
    そのまま str() すると '3907.0' のようになってシート側の文字列と
    一致しなくなる。float表現の末尾 .0 を除去して防ぐ。
    """
    if v is None:
        return ""
    if isinstance(v, float):
        if np.isnan(v):
            return ""
        if v.is_integer():
            return str(int(v))
        return str(v)
    return str(v).strip()


def _flatten_columns(columns) -> List[str]:
    """pandas.read_htmlがMultiIndex列を返した場合でも単純な文字列リストに正規化する"""
    flat = []
    for c in columns:
        if isinstance(c, tuple):
            parts = [
                str(x).strip()
                for x in c
                if str(x).strip() and not str(x).startswith("Unnamed")
            ]
            flat.append("".join(parts) if parts else str(c[-1]).strip())
        else:
            flat.append(str(c).strip())
    return flat


def _match_table(df: pd.DataFrame, required: List[str]) -> Optional[pd.DataFrame]:
    """required の列を全て含むテーブルなら、列名を正規化して返す。含まなければ None。

    JPXのテーブルはヘッダーを <th> ではなく <td> で組んでいるため、
    pandas が列名を 0,1,2... と振り、ヘッダーが本文1行目に入った状態で
    返ってくることがある。その場合はヘッダー行を検出して列名に昇格させる。
    """
    # ケース1: pandasがヘッダーを正しく認識している（<th>構成）
    cols = _flatten_columns(df.columns)
    if all(c in cols for c in required):
        out = df.copy()
        out.columns = cols
        out = out.loc[:, ~pd.Index(cols).duplicated()]
        return out[required].copy().reset_index(drop=True)

    # ケース2: ヘッダーが本文の先頭行に入っている（<td>構成）
    for i in range(min(3, len(df))):
        row = [_to_text(v) for v in df.iloc[i].tolist()]
        if all(c in row for c in required):
            out = df.iloc[i + 1:].copy()
            out.columns = row
            out = out.loc[:, ~pd.Index(row).duplicated()]
            return out[required].copy().reset_index(drop=True)

    return None


def _find_tables(html: str, label: str, specs: Dict[str, List[str]]) -> Dict[str, pd.DataFrame]:
    """HTML内の全テーブルから、specsで指定した列構成に合致するものを探す。

    specs: {"regulated": [必要な列...], "released": [必要な列...]} の形式。
    同じテーブルが複数のspecに重複マッチしないよう、一度使ったテーブルは除外する。
    """
    tables = pd.read_html(StringIO(html), flavor="lxml")
    found: Dict[str, pd.DataFrame] = {}
    used = set()

    for key, required in specs.items():
        for idx, t in enumerate(tables):
            if idx in used:
                continue
            matched = _match_table(t, required)
            if matched is not None:
                found[key] = matched
                used.add(idx)
                break

    missing = [k for k in specs if k not in found]
    if missing:
        # 失敗時に原因を追いやすいよう、取得できたテーブルの概要を出力する
        print(f"[{label}] 取得できたtable数: {len(tables)}")
        for i, t in enumerate(tables):
            print(f"  table[{i}] shape={t.shape} columns={_flatten_columns(t.columns)}")
            if len(t) > 0:
                print(f"    1行目: {[_to_text(v) for v in t.iloc[0].tolist()]}")
        raise RuntimeError(
            f"{label}: 必要なテーブル {missing} が見つかりません。"
            "JPXのページ構成が変更された可能性があります。"
        )
    return found


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """全セルを文字列化し、コード列が空のダミー行を除去する"""
    for col in df.columns:
        df[col] = df[col].map(_to_text)
    df = df[~df["コード"].isin(["", "nan", "－", "-", "―"])]
    return df.reset_index(drop=True)


# =============================================================================
# Google Sheets 共通処理
# =============================================================================
def get_gspread_client() -> gspread.Client:
    key_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not key_json:
        raise RuntimeError("環境変数 GOOGLE_SERVICE_ACCOUNT_JSON が設定されていません。")
    info = json.loads(key_json)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)


def open_worksheet(gc: gspread.Client, gid: int):
    sh = gc.open_by_key(SPREADSHEET_ID)
    return sh.get_worksheet_by_id(gid)


def ensure_header(ws) -> None:
    first_row = ws.row_values(1)
    if first_row[: len(SHEET_HEADER)] != SHEET_HEADER:
        ws.update(range_name="A1:G1", values=[SHEET_HEADER])


def load_existing(ws) -> List[Dict]:
    """シート内容を読み込み、行番号付きの辞書リストとして返す（ヘッダー除く）"""
    values = ws.get_all_values()
    rows = []
    for i, row in enumerate(values[1:], start=2):
        row = row + [""] * (7 - len(row))
        rows.append(
            {
                "sheet_row": i,
                "new_date": row[0].strip(),
                "code": row[1].strip(),
                "name": row[2].strip(),
                "start_date": row[3].strip(),
                "end_date": row[4].strip(),
                "content": row[5].strip(),
                "criteria": row[6].strip(),
            }
        )
    return rows


def process_new_entries(ws, records: List[NewRecord], record_date: str, label: str) -> None:
    """(コード, 実施日/指定日) が未登録のレコードを2行目にまとめて挿入する"""
    existing = load_existing(ws)
    existing_keys = {(r["code"], r["start_date"]) for r in existing}

    new_rows = []
    for rec in records:
        key = (rec.code, rec.start_date)
        if key in existing_keys:
            continue  # 既存データのまま
        new_rows.append(
            [record_date, rec.code, rec.name, rec.start_date, "", rec.content, rec.criteria]
        )

    if new_rows:
        # JPXの一覧は日付が新しい順に並んでいるため、そのままの順序でまとめて
        # 2行目から挿入すれば「新しいものほど上」の並びが保たれる
        ws.insert_rows(new_rows, row=2, value_input_option="USER_ENTERED")
        print(f"[{label}] 新規追加: {len(new_rows)} 件")
        for r in new_rows:
            print(f"  + {r[1]} {r[2]}  実施日:{r[3]}")
    else:
        print(f"[{label}] 新規データなし（更新済み）")


def process_released_entries(
    ws, records: List[ReleasedRecord], record_date: str, label: str
) -> None:
    """コードが一致しE列(解除日)が空欄の行に解除日を記入する"""
    existing = load_existing(ws)

    open_rows_by_code = {}
    for r in existing:
        if r["end_date"] == "" and r["code"]:
            cur = open_rows_by_code.get(r["code"])
            if cur is None or r["start_date"] > cur["start_date"]:
                open_rows_by_code[r["code"]] = r

    # JPXは解除銘柄を解除後1週間程度そのまま掲載し続けるため、同じ解除情報が
    # 何日も連続で取得される。既に解除日が記入済みの (コード, 解除日) を控えて
    # おき、フォールバック行が毎日重複追加されるのを防ぐ。
    released_keys = {(r["code"], r["end_date"]) for r in existing if r["end_date"]}

    cell_updates = []
    append_rows = []
    for rec in records:
        if (rec.code, rec.end_date) in released_keys:
            continue  # 反映済み

        target = open_rows_by_code.get(rec.code)
        if target:
            cell_updates.append({"range": f"E{target['sheet_row']}", "values": [[rec.end_date]]})
            released_keys.add((rec.code, rec.end_date))
            print(
                f"[{label}解除] {rec.code} {rec.name}  解除日:{rec.end_date} "
                f"→ {target['sheet_row']}行目に記入"
            )
        else:
            append_rows.append(
                [record_date, rec.code, rec.name, "", rec.end_date, rec.content, ""]
            )
            released_keys.add((rec.code, rec.end_date))
            print(f"[{label}解除] 対応する未解除行なし: {rec.code} {rec.name} → 新規行として追加")

    if cell_updates:
        ws.batch_update(cell_updates, value_input_option="USER_ENTERED")
    if append_rows:
        ws.insert_rows(append_rows, row=2, value_input_option="USER_ENTERED")
    if not cell_updates and not append_rows:
        print(f"[{label}解除] 更新対象なし")


# =============================================================================
# 1. 信用取引に関する規制等（増担保規制）
# =============================================================================
MARGIN_REG_URL = "https://www.jpx.co.jp/markets/equities/margin-reg/index.html"
REGULATED_COLS = ["銘柄名", "コード", "実施日", "規制の内容", "該当基準"]
RELEASED_COLS_REG = ["銘柄名", "コード", "解除日", "規制の内容"]


def parse_margin_reg(html: str):
    found = _find_tables(
        html,
        "信用取引に関する規制等",
        {"regulated": REGULATED_COLS, "released": RELEASED_COLS_REG},
    )
    regulated_df = _clean_dataframe(found["regulated"])
    released_df = _clean_dataframe(found["released"])

    new_records = [
        NewRecord(
            code=row["コード"],
            start_date=row["実施日"],
            name=row["銘柄名"],
            content=row["規制の内容"],
            criteria=row["該当基準"],
        )
        for _, row in regulated_df.iterrows()
    ]
    released_records = [
        ReleasedRecord(
            code=row["コード"],
            end_date=row["解除日"],
            name=row["銘柄名"],
            content=row["規制の内容"],
        )
        for _, row in released_df.iterrows()
    ]
    return new_records, released_records


def run_margin_reg(gc, record_date, html):
    print("--- 信用取引に関する規制等（増担保規制） ---")
    new_records, released_records = parse_margin_reg(html)
    print(f"取得件数: 規制中 {len(new_records)} 件 / 解除 {len(released_records)} 件")
    ws = open_worksheet(gc, MARGIN_REG_GID)
    ensure_header(ws)
    process_new_entries(ws, new_records, record_date, "規制銘柄")
    process_released_entries(ws, released_records, record_date, "規制銘柄")


# =============================================================================
# 2. 信用取引に関する日々公表（日々公表銘柄）
# =============================================================================
MARGIN_DAILY_URL = "https://www.jpx.co.jp/markets/equities/margin-daily/index.html"
DESIGNATED_COLS = ["銘柄名", "コード", "指定日"]
RELEASED_COLS_DAILY = ["銘柄名", "コード", "解除日"]
DAILY_CONTENT_LABEL = "日々公表銘柄"
MARK_RE = re.compile(r"^[※◆]+")


def split_marks(name: str):
    """銘柄名先頭の ※ ◆ マークを分離し、(素の銘柄名, マーク文字列) を返す"""
    m = MARK_RE.match(name)
    marks = m.group(0) if m else ""
    clean_name = MARK_RE.sub("", name)
    return clean_name, marks


def marks_to_criteria(marks: str) -> str:
    parts = []
    if "※" in marks:
        parts.append("信用規制中(※)")
    if "◆" in marks:
        parts.append("特別周知(◆)")
    return " / ".join(parts)


def parse_margin_daily(html: str):
    found = _find_tables(
        html,
        "信用取引に関する日々公表",
        {"designated": DESIGNATED_COLS, "released": RELEASED_COLS_DAILY},
    )
    designated_df = _clean_dataframe(found["designated"])
    released_df = _clean_dataframe(found["released"])

    new_records = []
    for _, row in designated_df.iterrows():
        clean_name, marks = split_marks(row["銘柄名"])
        new_records.append(
            NewRecord(
                code=row["コード"],
                start_date=row["指定日"],
                name=clean_name,
                content=DAILY_CONTENT_LABEL,
                criteria=marks_to_criteria(marks),
            )
        )

    released_records = []
    for _, row in released_df.iterrows():
        clean_name, _ = split_marks(row["銘柄名"])
        released_records.append(
            ReleasedRecord(
                code=row["コード"],
                end_date=row["解除日"],
                name=clean_name,
                content=DAILY_CONTENT_LABEL,
            )
        )
    return new_records, released_records


def run_margin_daily(gc, record_date, html):
    print("--- 信用取引に関する日々公表（日々公表銘柄） ---")
    new_records, released_records = parse_margin_daily(html)
    print(f"取得件数: 日々公表 {len(new_records)} 件 / 解除 {len(released_records)} 件")
    ws = open_worksheet(gc, MARGIN_DAILY_GID)
    ensure_header(ws)
    process_new_entries(ws, new_records, record_date, "日々公表銘柄")
    process_released_entries(ws, released_records, record_date, "日々公表銘柄")


# =============================================================================
# メイン
# =============================================================================
def parse_args():
    parser = argparse.ArgumentParser(description="JPX 信用取引規制データ 一括取得スクリプト")
    parser.add_argument(
        "--date",
        default=None,
        help="新規記載日(A列)として使う日付 YYYY-MM-DD。省略時は実行日(JST)を使用",
    )
    parser.add_argument(
        "--only",
        choices=["reg", "daily", "all"],
        default="all",
        help="reg=増担保規制のみ / daily=日々公表銘柄のみ / all=両方まとめて実行（既定）",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="非営業日（土日祝・年末年始）でも営業日判定を無視して強制実行する",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    target_dt = get_target_datetime(args.date)
    record_date = target_dt.strftime("%Y/%m/%d")

    if not args.force and not is_business_day(target_dt):
        print(
            f"{record_date} は非営業日（土日祝・年末年始）のためスキップします。"
            "強制実行する場合は --force を指定してください。"
        )
        return

    print(f"=== JPX 信用取引規制データ 一括取得開始 (基準日: {record_date}) ===")

    gc = get_gspread_client()
    errors = []

    if args.only in ("reg", "all"):
        try:
            run_margin_reg(gc, record_date, fetch_html(MARGIN_REG_URL))
        except Exception as exc:  # noqa: BLE001
            print(f"エラー（増担保規制）: {exc}", file=sys.stderr)
            errors.append(("増担保規制", exc))

    if args.only in ("daily", "all"):
        try:
            run_margin_daily(gc, record_date, fetch_html(MARGIN_DAILY_URL))
        except Exception as exc:  # noqa: BLE001
            print(f"エラー（日々公表銘柄）: {exc}", file=sys.stderr)
            errors.append(("日々公表銘柄", exc))

    if errors:
        print(f"=== {len(errors)} 件のフィードでエラーが発生しました ===", file=sys.stderr)
        sys.exit(1)

    print("=== 完了 ===")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"致命的エラー: {exc}", file=sys.stderr)
        sys.exit(1)
