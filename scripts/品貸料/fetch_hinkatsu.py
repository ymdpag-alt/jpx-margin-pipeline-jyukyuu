#!/usr/bin/env python3
"""
JPX 品貸料データ取得・スプレッドシート更新スクリプト

使用方法:
  python fetch_hinkatsu.py               # 最新データを取得・書き込み
  python fetch_hinkatsu.py --date 2026-08-07  # 指定日のデータを取得・書き込み
  python fetch_hinkatsu.py --dry-run     # 書き込まずに確認のみ
"""

import argparse
import io
import json
import os
import re
import time
from datetime import datetime

import gspread
import pandas as pd
import requests
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials

# ============================================================
# 定数
# ============================================================
JPX_BASE = "https://www.jpx.co.jp"
JPX_PAGE_URL = f"{JPX_BASE}/markets/statistics-equities/margin/01.html"
SPREADSHEET_ID = "1kWST0CkkIvo3irPSbMgtVtUqqRDXFwvRREYZDRQAFMY"

# gspread で取得するシート設定 (シート名 → {gid, Excelの列キーワード})
SHEET_CONFIG = {
    "貸株超過株数": {"gid": 1882345841, "excel_col": "貸株超過株数"},
    "最高料率":    {"gid": 1962219247, "excel_col": "最高料率"},
    "品貸料率":    {"gid": 1316793327, "excel_col": "品貸料率"},
}

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja-JP,ja;q=0.9,en-US;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
}

# ============================================================
# Google Sheets 認証
# ============================================================
def get_gspread_client():
    """環境変数 or credentials.json から認証"""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if creds_json:
        creds_info = json.loads(creds_json)
    else:
        creds_path = os.path.join(os.path.dirname(__file__), "..", "credentials.json")
        with open(creds_path) as f:
            creds_info = json.load(f)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)


# ============================================================
# JPX ページスクレイピング
# ============================================================
def fetch_page_info():
    """
    JPX 品貸料ページから (申込日, ExcelURL) を取得する
    Returns: {"date_raw": "20260807", "date_display": "2026/08/07", "excel_url": "https://..."}
    """
    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)

    resp = session.get(JPX_PAGE_URL, timeout=30)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.content, "html.parser")

    excel_url = None
    date_raw = None
    date_display = None

    # ─── Excel リンクを探す ───
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if re.search(r"Premium_Charges\.xlsx?$", href, re.IGNORECASE):
            excel_url = href if href.startswith("http") else f"{JPX_BASE}{href}"

            # 同じ行の日付テキストを探す
            tr = a.find_parent("tr")
            if tr:
                row_text = tr.get_text(" ", strip=True)
                m = re.search(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", row_text)
                if m:
                    y, mo, d = m.group(1), m.group(2).zfill(2), m.group(3).zfill(2)
                    date_raw = f"{y}{mo}{d}"
                    date_display = f"{y}/{mo}/{d}"
            break

    if not excel_url:
        raise ValueError("JPX ページから品貸料 Excel リンクが見つかりませんでした。")

    # 日付が取れなかった場合は URL から推測
    if not date_raw:
        m = re.search(r"(\d{8})", excel_url)
        if m:
            raw = m.group(1)
            date_raw = raw
            date_display = f"{raw[:4]}/{raw[4:6]}/{raw[6:]}"
        else:
            date_display = datetime.now().strftime("%Y/%m/%d")
            date_raw = datetime.now().strftime("%Y%m%d")

    return {
        "date_raw": date_raw,
        "date_display": date_display,
        "excel_url": excel_url,
    }


# ============================================================
# Excel 解析
# ============================================================
def download_excel(url: str) -> bytes:
    headers = dict(REQUEST_HEADERS)
    headers["Referer"] = JPX_PAGE_URL
    resp = requests.get(url, headers=headers, timeout=60)
    resp.raise_for_status()
    return resp.content


def parse_excel(content: bytes) -> pd.DataFrame:
    """
    品貸料 Excel を解析し、
    [コード, 銘柄名, 貸株超過株数, 最高料率, 品貸料率] の DataFrame を返す
    """
    xl = pd.ExcelFile(io.BytesIO(content))

    for sheet_name in xl.sheet_names:
        df_raw = xl.parse(sheet_name, header=None, dtype=str)

        # ヘッダー行（「コード」「銘柄」を含む行）を探す
        header_row_idx = None
        for i in df_raw.index:
            row_vals = " ".join(str(v) for v in df_raw.loc[i] if pd.notna(v) and str(v) != "nan")
            if "コード" in row_vals or "銘柄" in row_vals:
                header_row_idx = i
                break

        if header_row_idx is None:
            continue

        # ヘッダー行を列名に
        header = [
            str(v).strip().replace("\n", "").replace(" ", "") if pd.notna(v) else f"_col{i}"
            for i, v in enumerate(df_raw.loc[header_row_idx])
        ]
        df = df_raw.loc[header_row_idx + 1:].copy()
        df.columns = header
        df = df.dropna(how="all").reset_index(drop=True)

        # 必要列を見つける（部分一致）
        col_map = {}
        for col in df.columns:
            c = col.replace("（", "").replace("）", "").replace("(", "").replace(")", "")
            if "コード" in c and "コード" not in col_map:
                col_map["コード"] = col
            elif ("銘柄名" in c or "銘柄" in c) and "銘柄名" not in col_map:
                col_map["銘柄名"] = col
            elif "貸株超過株数" in c and "貸株超過株数" not in col_map:
                col_map["貸株超過株数"] = col
            elif "最高料率" in c and "最高料率" not in col_map:
                col_map["最高料率"] = col
            elif "品貸料率" in c and "品貸料率" not in col_map:
                col_map["品貸料率"] = col

        if "コード" not in col_map or "銘柄名" not in col_map:
            continue

        # 列をリネーム
        df = df.rename(columns={v: k for k, v in col_map.items()})

        # 必要な列だけ抽出
        needed = ["コード", "銘柄名", "貸株超過株数", "最高料率", "品貸料率"]
        existing_cols = [c for c in needed if c in df.columns]
        df = df[existing_cols].copy()

        # コードを文字列整形（"1234.0" → "1234"）
        def clean_code(x):
            s = str(x).strip()
            if s in ("nan", "", "None"):
                return ""
            # 末尾 .0 を除去
            s = re.sub(r"\.0+$", "", s)
            return s

        df["コード"] = df["コード"].apply(clean_code)
        df["銘柄名"] = df["銘柄名"].apply(lambda x: str(x).strip() if pd.notna(x) and str(x) != "nan" else "")

        # 数値列を文字列に（NaN → ""）
        for col in ["貸株超過株数", "最高料率", "品貸料率"]:
            if col in df.columns:
                df[col] = df[col].apply(
                    lambda x: "" if (pd.isna(x) or str(x) in ("nan", "None", "")) else str(x).strip()
                )

        # コードまたは銘柄名が空の行を除外
        df = df[(df["コード"] != "") | (df["銘柄名"] != "")].reset_index(drop=True)

        print(f"  Excel シート「{sheet_name}」から {len(df)} 件解析")
        return df

    raise ValueError("品貸料データの解析に失敗しました。Excel の構造を確認してください。")


# ============================================================
# スプレッドシート更新
# ============================================================
def get_worksheet_by_gid(spreadsheet, gid: int):
    """gid でワークシートを取得"""
    for ws in spreadsheet.worksheets():
        if ws.id == gid:
            return ws
    raise ValueError(f"gid={gid} のシートが見つかりません")


def insert_column_at_c(spreadsheet, worksheet):
    """
    C列（0-indexed: index=2）の前に空列を挿入する
    Sheets API の insertDimension を直接呼ぶ
    """
    body = {
        "requests": [{
            "insertDimension": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 2,   # 0-indexed → column C
                    "endIndex": 3,
                },
                "inheritFromBefore": False,
            }
        }]
    }
    spreadsheet.batch_update(body)


def update_worksheet(spreadsheet, ws, date_display: str, data_map: dict, dry_run: bool = False):
    """
    シートに新しい日付列を追加する。

    シート構造:
      Row 1:  [コード, 銘柄名, <新日付>, <旧日付>, ...]
      Row 2+: [code,   name,  <値>,     <旧値>, ...]

    新データは常に C列 に挿入（既存列は右にシフト）。
    既存 (コード, 銘柄名) ペアに対応する行に値を書き込み、
    新銘柄は末尾に行追加する。

    Args:
        spreadsheet: gspread Spreadsheet
        ws: gspread Worksheet
        date_display: "YYYY/MM/DD"
        data_map: {(code, name): value_str}
        dry_run: True なら書き込みしない
    """
    existing = ws.get_all_values()

    # ── 空シートの初期化 ──────────────────────────────
    is_empty = not existing or all(
        all(cell == "" for cell in row) for row in existing
    )
    if is_empty:
        print(f"    空シートを初期化します")
        rows = [["コード", "銘柄名", date_display]]
        for (code, name), value in sorted(data_map.items(), key=lambda x: x[0][0]):
            rows.append([code, name, value])
        if not dry_run:
            ws.clear()
            ws.update(rows, "A1")
        print(f"    → {len(data_map)} 件書き込み（新規）")
        return

    # ── 日付重複チェック ──────────────────────────────
    header_row = existing[0] if existing else []
    existing_dates = [str(v).strip() for v in header_row[2:]]  # C列以降

    if date_display in existing_dates:
        print(f"    → {date_display} は既存。スキップ。")
        return

    # ── 既存 (code, name) → 行番号（1-indexed）マップ ──
    existing_map: dict[tuple, int] = {}
    for i, row in enumerate(existing[1:], start=2):
        code = str(row[0]).strip() if len(row) > 0 else ""
        name = str(row[1]).strip() if len(row) > 1 else ""
        if code or name:
            existing_map[(code, name)] = i

    print(f"    既存行数: {len(existing_map)} 件")

    # ── C列に空列を挿入 ──────────────────────────────
    if not dry_run:
        insert_column_at_c(spreadsheet, ws)
        time.sleep(1.5)

    # ── C1 に日付をセット ──────────────────────────────
    if not dry_run:
        ws.update_cell(1, 3, date_display)
        time.sleep(0.5)

    # ── 既存行への値セット（バッチ更新）──────────────
    batch_updates = []
    new_entries: list[tuple] = []

    for (code, name), value in data_map.items():
        if (code, name) in existing_map:
            row_num = existing_map[(code, name)]
            batch_updates.append({
                "range": gspread.utils.rowcol_to_a1(row_num, 3),
                "values": [[value]],
            })
        else:
            new_entries.append((code, name, value))

    if batch_updates:
        print(f"    既存行に値をセット: {len(batch_updates)} 件")
        if not dry_run:
            # 100件ずつバッチ更新（API制限対策）
            for i in range(0, len(batch_updates), 100):
                chunk = batch_updates[i:i + 100]
                ws.batch_update(chunk)
                if i + 100 < len(batch_updates):
                    time.sleep(1)

    # ── 新規銘柄を末尾に行追加 ──────────────────────────
    if new_entries:
        print(f"    新規銘柄を追加: {len(new_entries)} 件")
        if not dry_run:
            # 現在のシート行数（挿入後）= 元の行数 + 1（ヘッダー分含む）
            last_row = len(existing) + 1
            # 新規行: A=code, B=name, C=value（D列以降は空のまま）
            new_rows = [[code, name, value] for code, name, value in new_entries]
            ws.update(new_rows, f"A{last_row}")

    if dry_run:
        print(f"    [DRY-RUN] 更新予定: 既存 {len(batch_updates)} 件, 新規 {len(new_entries)} 件")
    else:
        print(f"    ✅ 完了: 既存 {len(batch_updates)} 件更新, 新規 {len(new_entries)} 件追加")


# ============================================================
# メイン
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="JPX 品貸料データ取得・スプレッドシート更新")
    parser.add_argument(
        "--date",
        type=str,
        default=None,
        help="対象日付 YYYY-MM-DD（省略時はページの最新データを取得）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=False,
        help="スプレッドシートへの書き込みをせず、取得内容だけ確認する",
    )
    args = parser.parse_args()

    # ── 1. JPX ページから最新 Excel 情報を取得 ──────────
    print("=" * 60)
    print("JPX 品貸料ページをスクレイピング中...")
    page_info = fetch_page_info()
    print(f"  申込日: {page_info['date_display']}")
    print(f"  Excel URL: {page_info['excel_url']}")

    # ── 2. 指定日との照合 ──────────────────────────────
    if args.date:
        target_raw = args.date.replace("-", "")
        if page_info["date_raw"] != target_raw:
            print(
                f"\n⚠️  指定日 {args.date} とページの最新日 {page_info['date_display']} が一致しません。"
            )
            print("   JPX は最新データのみ公開しています。スクリプトを終了します。")
            print("   ※ 指定日のデータを取得するには、その日に実行してください。")
            return

    date_display = page_info["date_display"]

    # ── 3. Excel ダウンロード & 解析 ──────────────────
    print("\nExcel をダウンロード中...")
    content = download_excel(page_info["excel_url"])
    print(f"  ダウンロード完了: {len(content):,} bytes")

    print("Excel を解析中...")
    df = parse_excel(content)
    print(f"\n--- データ先頭5件 ---")
    print(df.head().to_string())
    print(f"--- 全 {len(df)} 件 ---\n")

    # ── 4. (code, name) → value のマップ作成 ──────────
    def build_data_map(col_name: str) -> dict:
        result = {}
        if col_name not in df.columns:
            return result
        for _, row in df.iterrows():
            code = str(row.get("コード", "")).strip()
            name = str(row.get("銘柄名", "")).strip()
            value = str(row.get(col_name, "")).strip()
            if code or name:
                result[(code, name)] = value
        return result

    # ── 5. Google Sheets に接続 ───────────────────────
    if not args.dry_run:
        print("Google Sheets に接続中...")
        gc = get_gspread_client()
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    else:
        gc = None
        spreadsheet = None
        print("[DRY-RUN モード] スプレッドシートへの接続をスキップ")

    # ── 6. 各シートを更新 ─────────────────────────────
    for sheet_name, config in SHEET_CONFIG.items():
        print(f"\n[{sheet_name}] を処理中...")

        excel_col = config["excel_col"]
        if excel_col not in df.columns:
            print(f"  ⚠️  列 '{excel_col}' がデータにありません。スキップ。")
            continue

        data_map = build_data_map(excel_col)
        print(f"  データ件数: {len(data_map)} 件")

        if args.dry_run:
            print(f"  [DRY-RUN] サンプル:")
            for i, ((code, name), val) in enumerate(list(data_map.items())[:3]):
                print(f"    {code} {name}: {val}")
            continue

        try:
            ws = get_worksheet_by_gid(spreadsheet, config["gid"])
            update_worksheet(spreadsheet, ws, date_display, data_map, dry_run=False)
        except Exception as e:
            print(f"  ❌ エラー: {e}")
            import traceback
            traceback.print_exc()

        time.sleep(2)  # シート間の API 制限対策

    print("\n" + "=" * 60)
    print("✅ 処理完了")


if __name__ == "__main__":
    main()
