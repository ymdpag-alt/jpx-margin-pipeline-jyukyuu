import yfinance as yf
import pandas as pd
import numpy as np
import time
import os
import re
import gspread
from datetime import datetime
from zoneinfo import ZoneInfo

# =============================================================================
# 定数定義
# =============================================================================

# kabutan/nikkei スクレイパーと同じスプレッドシート
SPREADSHEET_ID = "1QheVVw97DnHjdymEYNFwvgiQhgX8SX-bPxjlHmZpG2I"

# 書き込み先タブを gid（URLの gid=... の数値）で指定する。
# タブ名ではなく gid で狙うので、タブ名が何であっても確実に同じタブへ書き込む。
VOLUME_SHEET_GID = 328764273     # 出来高
CLOSE_SHEET_GID  = 2080765326    # 終値

# サービスアカウントのJSONキー（このスクリプトと同じフォルダに置く想定）
# kabutan/nikkei スクレイパーと同じファイルを流用可。
SERVICE_ACCOUNT_FILE = "service_account.json"

# 銘柄コード一覧を読み込むタブ（同じスプレッドシート内、gid で指定）
CODES_SHEET_GID = 1376419996

# 銘柄コードの形式判定用パターン
# 数字始まりの4文字（4桁数字 例:7203 / 英文字入りの新コード 例:130A）にマッチ。
# このパターンで、ヘッダー行・空欄・社名などを自動的に除外する。
CODE_PATTERN = re.compile(r"^\d[0-9A-Za-z]{3}$")

# 取得したい日付（YYYY-MM-DD形式）
# → 実行した「当日（日本時間）」の日付を自動で設定
TARGET_DATES = [datetime.now(ZoneInfo("Asia/Tokyo")).strftime("%Y-%m-%d")]
# 手動で日付を指定したい場合は↓のように書き換える
#TARGET_DATES = ["2026-06-10"]

# API制御
CHUNK_SIZE = 100   # 一度に取得する銘柄数
SLEEP_TIME = 1     # チャンク間のスリープ秒数

# シート構成: A列=銘柄コード, B列=（予備）, C列以降=日付データ
FIXED_COLS = 2     # ソート時に固定する左端の列数（A・B列）

# 曜日（月〜日）
_WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]


# =============================================================================
# 日付フォーマット変換
# =============================================================================

def to_japanese_date(date_str: str) -> str:
    """
    'YYYY-MM-DD' 形式を '2026年2月20日(金)' 形式に変換する。
    変換に失敗した場合は元の文字列をそのまま返す。
    """
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        weekday = _WEEKDAY_JP[dt.weekday()]
        return f"{dt.year}年{dt.month}月{dt.day}日({weekday})"
    except ValueError:
        return date_str


def from_japanese_date(date_str: str) -> str:
    """
    '2026年2月20日(金)' 形式を 'YYYY-MM-DD' 形式に逆変換する。
    変換に失敗した場合は元の文字列をそのまま返す。
    """
    try:
        # 曜日部分を除去して数値だけ抽出
        core = date_str.split("(")[0]  # "2026年2月20日"
        dt = datetime.strptime(core, "%Y年%m月%d日")
        return dt.strftime("%Y-%m-%d")
    except (ValueError, IndexError):
        return date_str


# =============================================================================
# 認証・シート操作
# =============================================================================

def authenticate_google_sheets() -> gspread.Client:
    """サービスアカウントで Google Sheets API の認証を行い、クライアントを返す。"""
    print("Google認証中...")
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(
            f"認証ファイルが見つかりません → {SERVICE_ACCOUNT_FILE}\n"
            "  サービスアカウントのJSONキーをこのスクリプトと同じ場所に置き、\n"
            "  対象スプレッドシートをそのサービスアカウントのメールアドレス\n"
            "  （JSON内の client_email）と共有（編集者）してください。"
        )
    return gspread.service_account(filename=SERVICE_ACCOUNT_FILE)


def get_worksheet_by_gid(
    gc: gspread.Client, gid: int, min_cols: int = 50
) -> gspread.Worksheet:
    """gid でワークシートを取得し、必要なら列を拡張して返す。"""
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = spreadsheet.get_worksheet_by_id(gid)
    except gspread.exceptions.WorksheetNotFound:
        raise gspread.exceptions.WorksheetNotFound(
            f"gid={gid} のタブが見つかりません。"
            "スプレッドシート内に該当タブが存在するか確認してください。"
        )

    # 列数が足りなければ拡張
    if ws.col_count < min_cols:
        ws.add_cols(min_cols - ws.col_count)
        ws = spreadsheet.get_worksheet_by_id(gid)
    return ws


# =============================================================================
# データ取得
# =============================================================================

def load_stock_codes_from_sheet(gc: gspread.Client, gid: int) -> list[str]:
    """
    指定タブから銘柄コード一覧を読み込む。
    タブ内で「銘柄コードらしい値」が最も多い列を自動判定し、その列から
    コード形式の値だけを抽出する（ヘッダー・空欄・社名などは自動的に除外）。
    コードが社名と同居していても、列の位置に関係なく拾える。
    """
    ws = get_worksheet_by_gid(gc, gid)
    sheet_name = ws.title
    print(f"銘柄コード読み込み: タブ '{sheet_name}' (gid={gid})")

    rows = ws.get_all_values()
    if not rows:
        raise ValueError(f"タブ '{sheet_name}' にデータがありません。")

    # 各列について、コード形式の値がいくつ含まれるかを数え、最多の列を採用
    max_cols = max(len(r) for r in rows)
    best_col, best_count = -1, 0
    for c in range(max_cols):
        count = sum(
            1 for r in rows
            if c < len(r) and CODE_PATTERN.match(r[c].strip())
        )
        if count > best_count:
            best_col, best_count = c, count

    if best_col < 0 or best_count == 0:
        raise ValueError(
            f"タブ '{sheet_name}' から銘柄コードらしい列が見つかりませんでした。\n"
            "  4桁の証券コード（例: 7203）が並んだ列があるか確認してください。"
        )

    # 採用した列からコード形式の値だけを抽出し、重複を除去（順序維持）
    codes = [
        r[best_col].strip()
        for r in rows
        if best_col < len(r) and CODE_PATTERN.match(r[best_col].strip())
    ]
    codes = list(dict.fromkeys(codes))

    col_letter = gspread.utils.rowcol_to_a1(1, best_col + 1).rstrip("1")
    print(f"  {col_letter}列から {len(codes)} 銘柄を読み込み")
    return codes


def fetch_stock_data(
    codes: list[str],
    target_dates: list[str],
    data_type: str = "Volume",
) -> pd.DataFrame:
    """
    Yahoo Finance から出来高 or 終値を取得し、DataFrameで返す。
    内部の列名は 'YYYY-MM-DD' 形式で保持する。

    Args:
        codes: 銘柄コードのリスト（例: ["7203", "9984"]）
        target_dates: 取得対象日のリスト（YYYY-MM-DD形式）
        data_type: "Volume"（出来高）または "Close"（終値）
    """
    tickers = [f"{c}.T" for c in codes]
    # yfinanceは start〜end の範囲が狭すぎる（特に1日だけ）と、
    # 実際にはデータがある日でも0件を返すことがある既知の癖があるため、
    # 取得範囲は前後に余裕を持たせ、対象日の行だけを後で拾う。
    start = str(pd.to_datetime(min(target_dates)) - pd.Timedelta(days=7))[:10]
    end = str(pd.to_datetime(max(target_dates)) + pd.Timedelta(days=1))[:10]
    label = "出来高" if data_type == "Volume" else "終値"
    total_chunks = (len(tickers) + CHUNK_SIZE - 1) // CHUNK_SIZE

    # B列（備考）を追加して日付をC列以降に揃える
    result = pd.DataFrame({"銘柄コード": codes, "備考": ""})
    for d in target_dates:
        result[d] = None

    # 対象日を date で保持する。
    # yfinance の index に tz や時刻成分が付いても日付一致が外れないよう、date ベースで突き合わせる。
    target_date_map = {pd.to_datetime(d).date(): d for d in target_dates}

    print(f"\n{label}データ取得中（{total_chunks} チャンク）...")
    filled_total = 0

    for i in range(0, len(tickers), CHUNK_SIZE):
        chunk = tickers[i : i + CHUNK_SIZE]
        chunk_num = i // CHUNK_SIZE + 1

        # auto_adjust=False で「実際の終値」を取得する
        # （yfinance 1.x の既定は auto_adjust=True で、終値が調整後終値になってしまうため）
        try:
            raw = yf.download(
                chunk, start=start, end=end,
                progress=False, auto_adjust=False, group_by="column",
            )
        except Exception as e:
            print(f"  [{chunk_num}/{total_chunks}] ⚠ ダウンロード失敗: {e}")
            if chunk_num == 1:
                import traceback
                print("  DEBUG: 例外の詳細:")
                traceback.print_exc()
            time.sleep(SLEEP_TIME)
            continue

        if chunk_num == 1:
            print(f"  DEBUG: raw is None: {raw is None}")
            if raw is not None:
                print(f"  DEBUG: raw.shape: {raw.shape}")
                print(f"  DEBUG: raw.empty: {raw.empty}")
                print(f"  DEBUG: raw.columns[:10]: {list(raw.columns[:10])}")
                print(f"  DEBUG: raw.index[-5:]: {list(raw.index[-5:])}")
                print(f"  DEBUG: start={start} end={end}")

        if raw is None or raw.empty:
            print(f"  [{chunk_num}/{total_chunks}] データ0件（休場・未確定・取得制限の可能性）")
            time.sleep(SLEEP_TIME)
            continue

        # data_type の列を DataFrame（index=日付, columns=ティッカー）として取り出す
        if isinstance(raw.columns, pd.MultiIndex):
            if data_type not in raw.columns.get_level_values(0):
                print(f"  [{chunk_num}/{total_chunks}] ⚠ '{data_type}' 列が見つからず")
                time.sleep(SLEEP_TIME)
                continue
            sub = raw[data_type]
        else:
            if data_type not in raw.columns:
                print(f"  [{chunk_num}/{total_chunks}] ⚠ '{data_type}' 列が見つからず")
                time.sleep(SLEEP_TIME)
                continue
            sub = raw[[data_type]]
            sub.columns = [chunk[0]]
        if isinstance(sub, pd.Series):
            sub = sub.to_frame(name=chunk[0])

        # index の tz を外して date 化（日付一致を date ベースに）
        idx = sub.index
        if getattr(idx, "tz", None) is not None:
            idx = idx.tz_localize(None)
        date_to_pos = {ts.date(): pos for pos, ts in enumerate(idx)}

        if chunk_num == 1:
            print(f"  DEBUG: date_to_posのキー(取得できた日付一覧): {sorted(date_to_pos.keys())}")
            print(f"  DEBUG: target_date_mapのキー(探している日付): {sorted(target_date_map.keys())}")

        filled_chunk = 0
        for ticker in chunk:
            if ticker not in sub.columns:
                continue
            code = ticker.replace(".T", "")
            mask = result["銘柄コード"] == code
            col = sub[ticker]
            for d_date, d_str in target_date_map.items():
                pos = date_to_pos.get(d_date)
                if pos is None:
                    continue
                val = col.iloc[pos]
                if pd.notna(val):
                    result.loc[mask, d_str] = (
                        int(val) if data_type == "Volume" else float(val)
                    )
                    filled_chunk += 1

        filled_total += filled_chunk
        print(f"  [{chunk_num}/{total_chunks}] {len(chunk)}銘柄中 {filled_chunk}件 取得")
        time.sleep(SLEEP_TIME)

    print(f"  → {label}: 合計 {filled_total} 件取得")
    if filled_total == 0:
        print("  ⚠ 値が1件も取得できませんでした。考えられる原因:")
        print("     ・対象日のデータがまだ未反映（大引け前）/ 休場日")
        print("     ・yfinance の取得制限（時間をおく・CHUNK_SIZE を下げる）")
        print("     ・銘柄コードの形式（.T 付与後に Yahoo に存在するか）")
    return result


# =============================================================================
# スプレッドシート書き込み
# =============================================================================

def to_native(value):
    """pandas/numpy の値を Google Sheets 用のネイティブ型に変換する。"""
    if pd.isna(value) or value is None:
        return ""
    if isinstance(value, (np.integer, int)):
        return int(value)
    if isinstance(value, (np.floating, float)):
        return float(value)
    return str(value)


def update_spreadsheet(gc: gspread.Client, df: pd.DataFrame, gid: int):
    """
    スプレッドシートに日付列データを追記する。
    初回はヘッダー＋全データ、2回目以降は新しい日付列のみ右側に追加。
    日付列ヘッダーはシートに書き込む際に日本語形式（例: 2026年2月20日(金)）に変換する。
    """
    ws = get_worksheet_by_gid(gc, gid)
    sheet_name = ws.title
    print(f"\n{sheet_name} にデータ書き込み中...")
    existing = ws.get_all_values()

    # df の日付列（YYYY-MM-DD）→ 日本語形式の列名に変換したヘッダーを準備
    df_date_cols = list(df.columns[FIXED_COLS:])                          # ["2026-02-20", ...]
    jp_date_cols = [to_japanese_date(d) for d in df_date_cols]            # ["2026年2月20日(金)", ...]
    iso_to_jp   = dict(zip(df_date_cols, jp_date_cols))                   # 変換マッピング

    # --- 初回書き込み ---
    if not existing:
        print("  初回書き込み")
        # ヘッダーの日付部分のみ日本語形式に置換
        header = (
            list(df.columns[:FIXED_COLS])
            + [iso_to_jp.get(c, c) for c in df_date_cols]
        )
        rows = [[to_native(v) for v in row] for row in df.values]
        ws = get_worksheet_by_gid(gc, gid, min_cols=len(header))
        ws.update(values=[header] + rows, range_name="A1", value_input_option="USER_ENTERED")
        return

    # --- 追記モード: 新しい日付列のみ追加 ---
    header = existing[0]
    existing_codes = [row[0] for row in existing[1:]]

    # シート上の既存日付列は日本語形式なので、比較用に YYYY-MM-DD に逆変換してチェック
    existing_dates_jp  = set(header[FIXED_COLS:])
    existing_dates_iso = {from_japanese_date(d) for d in existing_dates_jp}
    new_date_cols_iso  = [d for d in df_date_cols if d not in existing_dates_iso]

    if not new_date_cols_iso:
        print("  追加する新しい日付はありません")
        return

    new_date_cols_jp = [iso_to_jp[d] for d in new_date_cols_iso]
    print(f"  追加日付: {', '.join(new_date_cols_jp)}")

    # ヘッダー行の右端に新日付（日本語形式）を追加
    col_start = len(header) + 1
    col_end   = col_start + len(new_date_cols_jp) - 1
    ws = get_worksheet_by_gid(gc, gid, min_cols=col_end)
    header_range = (
        f"{gspread.utils.rowcol_to_a1(1, col_start)}"
        f":{gspread.utils.rowcol_to_a1(1, col_end)}"
    )
    ws.update(values=[new_date_cols_jp], range_name=header_range, value_input_option="USER_ENTERED")

    # 既存銘柄→バッチ更新、新規銘柄→末尾に追加
    batch_updates = []
    new_rows      = []

    for _, row in df.iterrows():
        code   = str(row["銘柄コード"])
        values = [to_native(row[d]) for d in new_date_cols_iso]

        if code in existing_codes:
            row_idx    = existing_codes.index(code) + 2
            cell_range = (
                f"{gspread.utils.rowcol_to_a1(row_idx, col_start)}"
                f":{gspread.utils.rowcol_to_a1(row_idx, col_end)}"
            )
            batch_updates.append({"range": cell_range, "values": [values]})
        else:
            full_row = [code] + [""] * (len(header) - 1) + values
            new_rows.append(full_row)

    if batch_updates:
        print(f"  既存銘柄を更新: {len(batch_updates)} 件")
        ws.batch_update(batch_updates, value_input_option="USER_ENTERED")

    if new_rows:
        print(f"  新規銘柄を追加: {len(new_rows)} 件")
        start_row  = len(existing) + 1
        total_cols = len(header) + len(new_date_cols_jp)
        ws         = get_worksheet_by_gid(gc, gid, min_cols=total_cols)
        end_a1     = gspread.utils.rowcol_to_a1(start_row + len(new_rows) - 1, total_cols)
        ws.update(
            values=new_rows,
            range_name=f"A{start_row}:{end_a1}",
            value_input_option="USER_ENTERED",
        )

    print("  書き込み完了！")


def sort_date_columns(gc: gspread.Client, gid: int):
    """
    C列以降の日付列を新しい順（降順）にソートする。
    日付列ヘッダーは日本語形式（例: 2026年2月20日(金)）を想定。
    A列・B列（FIXED_COLS=2）は固定のまま維持。
    """
    ws       = get_worksheet_by_gid(gc, gid)
    sheet_name = ws.title
    print(f"\n{sheet_name} の日付列をソート中...")
    all_data = ws.get_all_values()

    # データ不足またはC列以降に日付がなければスキップ
    if not all_data or len(all_data) < 2 or len(all_data[0]) <= FIXED_COLS:
        print("  ソート対象の日付列がありません。スキップします")
        return

    header       = all_data[0]
    fixed_header = header[:FIXED_COLS]    # A列・B列（固定）
    date_header  = header[FIXED_COLS:]    # C列以降（日付・日本語形式）

    def parse_date(s: str) -> datetime:
        """
        日本語形式 '2026年2月20日(金)' をパース。
        失敗時は datetime.min を返す。
        """
        iso = from_japanese_date(s)
        try:
            return datetime.strptime(iso, "%Y-%m-%d")
        except ValueError:
            return datetime.min

    # 降順（新しい日付が左）にソート
    sorted_indices = sorted(
        range(len(date_header)),
        key=lambda i: parse_date(date_header[i]),
        reverse=True,
    )
    sorted_header = [date_header[i] for i in sorted_indices]

    # 全行を同じ順序で並べ替え
    new_data = [fixed_header + sorted_header]
    for row in all_data[1:]:
        fixed_vals = row[:FIXED_COLS]
        # 行が短い場合に備えて空文字でパディング
        date_vals  = row[FIXED_COLS:] + [""] * max(0, len(date_header) - len(row[FIXED_COLS:]))
        sorted_vals = [date_vals[i] for i in sorted_indices]
        new_data.append(fixed_vals + sorted_vals)

    # シートに書き戻し
    total_cols = len(new_data[0])
    ws         = get_worksheet_by_gid(gc, gid, min_cols=total_cols)
    end_a1     = gspread.utils.rowcol_to_a1(len(new_data), total_cols)
    ws.update(
        values=new_data,
        range_name=f"A1:{end_a1}",
        value_input_option="USER_ENTERED",
    )

    preview = " → ".join(sorted_header[:5])
    suffix  = "..." if len(sorted_header) > 5 else ""
    print(f"  ソート完了: {len(sorted_header)} 列（{preview}{suffix}）")


# =============================================================================
# メイン処理
# =============================================================================

def main():
    """メイン実行関数"""
    print(f"取得対象日（日本時間の当日）: {TARGET_DATES[0]}")
    gc = authenticate_google_sheets()

    print()
    try:
        codes = load_stock_codes_from_sheet(gc, CODES_SHEET_GID)
    except Exception as e:
        print(f"エラー: 銘柄コードの読み込みに失敗 → {e}")
        return

    if not codes:
        print("エラー: 銘柄コードが0件でした。タブの内容を確認してください。")
        return

    # 出来高データ
    print("\n" + "=" * 50)
    print("【出来高データ】")
    print("=" * 50)
    vol_df = fetch_stock_data(codes, TARGET_DATES, "Volume")
    update_spreadsheet(gc, vol_df, VOLUME_SHEET_GID)
    sort_date_columns(gc, VOLUME_SHEET_GID)

    # 終値データ
    print("\n" + "=" * 50)
    print("【終値データ】")
    print("=" * 50)
    close_df = fetch_stock_data(codes, TARGET_DATES, "Close")
    update_spreadsheet(gc, close_df, CLOSE_SHEET_GID)
    sort_date_columns(gc, CLOSE_SHEET_GID)

    # 完了サマリー
    print("\n" + "=" * 50)
    print("全処理完了！")
    print(f"  スプレッドシートID : {SPREADSHEET_ID}")
    print(f"  出来高タブ gid     : {VOLUME_SHEET_GID}")
    print(f"  終値タブ gid       : {CLOSE_SHEET_GID}")
    print(f"  銘柄数             : {len(codes)}")
    print(f"  取得日付           : {', '.join(TARGET_DATES)}")
    print("=" * 50)


if __name__ == "__main__":
    main()
