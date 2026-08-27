# ============================================================
#  株探ランキング・スクレイパー：GitHub Actions版
# ------------------------------------------------------------
#  ローカルPC版からの変更点：
#    (1) 認証：ローカルの service_account.json
#        → GitHub Secrets（環境変数 GOOGLE_SERVICE_ACCOUNT_JSON）から読込
#        ※ ローカルでの動作確認用に、Secretsが無い場合は
#          従来通り同フォルダの service_account.json にフォールバックします
#    (2) 実行環境：
#        GitHub Actionsのランナーには画面(ディスプレイ)が無いため、
#        ワークフロー側で Xvfb（仮想ディスプレイ）を起動した上で
#        headless=False のまま実Chromeを動かします
#        （--headless フラグを使うより、Bot検知を受けにくいためです）
#    (3) Chromeのパス・プロキシ・headless切替を環境変数で指定可能に
#    (4) CAPTCHA発生時や認証失敗時は exit code 1 で終了し、
#        GitHub Actions側で失敗として検知できるようにしました
#
#  必要なリポジトリ設定：
#    - Secrets: GOOGLE_SERVICE_ACCOUNT_JSON
#        （サービスアカウントのJSONキーの中身をそのまま貼り付け）
#    - Variables（任意）: SPREADSHEET_ID
#        （未設定の場合は下記のデフォルト値を使用）
#    - 対象スプレッドシートを、サービスアカウントのメールアドレスに
#      「編集者」権限で共有しておくこと
# ============================================================

import os
import sys
import json
import time
import asyncio
import random
from io import StringIO
from importlib.metadata import version, PackageNotFoundError

import pandas as pd
from bs4 import BeautifulSoup
import nodriver as uc
import gspread


# ============================================================
#  設定
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# ローカル実行時のフォールバック用（GitHub Actionsでは使いません）
SERVICE_ACCOUNT_KEY_PATH = os.path.join(BASE_DIR, "service_account.json")

# データを書き込むスプレッドシートのID（Variablesで上書き可能）
# 注意: GitHub Actionsでは未設定の vars.SPREADSHEET_ID が空文字列として渡ってくるため、
# os.environ.get(key, default) の default は効きません（キー自体は存在するため）。
# 「or」で空文字列も弾いてデフォルトにフォールバックさせています。
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID") or "1QheVVw97DnHjdymEYNFwvgiQhgX8SX-bPxjlHmZpG2I"

# ブラウザ関連の環境変数
HEADLESS = os.environ.get("HEADLESS", "false").lower() == "true"
CHROME_PATH = os.environ.get("CHROME_PATH") or None
PROXY_SERVER = os.environ.get("PROXY_SERVER")  # 例: "http://user:pass@host:port"

TOP_URL = "https://kabutan.jp/"
DEFAULT_PAGES = [1, 2, 3, 4, 5, 6]
PAGE_MAX_ATTEMPTS = 2
WRITE_CHUNK = 300
FIRST_WAIT = 30   # その実行で最初のページだけ長め
RETRY_WAIT = 10    # 2ページ目以降
WRONG_TABLE_KEYWORDS = ["日経平均", "米ドル", "ＮＹダウ", "NYダウ", "ＴＯＰＩＸ"]

KABUTAN_TARGETS = [
    {"sheet": "東証騰落レシオ",   "url_tpl": "https://kabutan.jp/warning/?mode=9_1&market=0&capitalization=-1&stc=zenhiritsu&stm=1&page={page}", "pages": [1, 2, 3]},
    {"sheet": "約定回数",         "url_tpl": "https://kabutan.jp/warning/?mode=2_9&market=0&capitalization=-1&stc=&stm=0&page={page}"},
    {"sheet": "取引高",           "url_tpl": "https://kabutan.jp/warning/trading_value_ranking?market=0&capitalization=-1&stc=&stm=0&page={page}"},
    {"sheet": "出来高",           "url_tpl": "https://kabutan.jp/warning/volume_ranking?market=0&capitalization=-1&stc=&stm=0&page={page}"},
    {"sheet": "株価25日越え",     "url_tpl": "https://kabutan.jp/warning/?mode=6_3&market=0&capitalization=-1&stc=&stm=0&page={page}"},
    {"sheet": "5日25日GC",        "url_tpl": "https://kabutan.jp/warning/?mode=6_1&market=0&capitalization=-1&stc=&stm=0&page={page}"},
    {"sheet": "200日上昇",        "url_tpl": "https://kabutan.jp/tansaku/?mode=2_0262&market=0&capitalization=-1&stc=v3&stm=0&page={page}"},
    {"sheet": "一目均衡表好転",   "url_tpl": "https://kabutan.jp/tansaku/?mode=2_0427&market=0&capitalization=-1&stc=&stm=0&page={page}"},
    {"sheet": "パラボリック陽転", "url_tpl": "https://kabutan.jp/tansaku/?mode=2_0476&market=0&capitalization=-1&stc=&stm=0&page={page}"},
    {"sheet": "RSI20以下",        "url_tpl": "https://kabutan.jp/tansaku/?mode=2_0460&market=0&capitalization=-1&stc=&stm=0&page={page}"},
    {"sheet": "MACD買い",         "url_tpl": "https://kabutan.jp/tansaku/?mode=2_0440&market=0&capitalization=-1&stc=&stm=0&page={page}"},
    {"sheet": "3陽転新値",        "url_tpl": "https://kabutan.jp/tansaku/?mode=2_0490&market=0&capitalization=-1&stc=&stm=0&page={page}"},
    {"sheet": "年初来高値",        "url_tpl": "https://kabutan.jp/tansaku/?mode=3_3&market=0&capitalization=-1&stc=&stm=0&page={page}"},

    {"sheet": "信用高値期日",         "url_tpl": "https://kabutan.jp/warning/?mode=7_5&market=0&capitalization=-1&stc=&stm=0&page={page}"},
    {"sheet": "信用安値期日",        "url_tpl": "https://kabutan.jp/tansaku/?mode=7_6&market=0&capitalization=-1&stc=&stm=0&page={page}"},

    {"sheet": "上昇率",           "url_tpl": "https://kabutan.jp/warning/?mode=2_1&market=0&capitalization=-1&stc=zenhiritsu&stm=1&page={page}"},
    {"sheet": "下落率",           "url_tpl": "https://kabutan.jp/warning/?mode=2_2&market=0&capitalization=-1&stc=zenhiritsu&stm=0&page={page}"},
    {"sheet": "過去1ｗ上昇率",    "url_tpl": "https://kabutan.jp/warning/?mode=11_13&market=0&capitalization=-1&stc=zenhiritsu&stm=1&page={page}"},
    {"sheet": "過去1ｗ下落率",    "url_tpl": "https://kabutan.jp/warning/?mode=11_14&market=0&capitalization=-1&stc=zenhiritsu&stm=0&page={page}"},

    {"sheet": "動意",             "url_tpl": "https://kabutan.jp/warning/?mode=3_5&market=0&capitalization=-1&stc=zenhiritsu&stm=1&page={page}", "pages": [1, 2, 3, 4, 5, 6, 7, 8, 9]},
    {"sheet": "株価が動意づいた材料株",   "url_tpl": "https://kabutan.jp/warning/?mode=3_5&market=0&capitalization=-1&stc=zenhiritsu&stm=1&page={page}", "pages": [1, 2, 3, 4, 5, 6, 7, 8, 9]},
    {"sheet": "インパクトがある開示情報", "url_tpl": "https://kabutan.jp/warning/?mode=4_4&market=0&capitalization=-1&stc=zenhiritsu&stm=1&page={page}", "pages": [1, 2, 3, 4, 5, 6, 7, 8, 9]},
    {"sheet": "朝刊」ニュース銘柄",       "url_tpl": "https://kabutan.jp/warning/?mode=4_1&market=0&capitalization=-1&stc=zenhiritsu&stm=1&page={page}", "pages": [1, 2, 3, 4, 5, 6, 7, 8, 9]},
]


# ============================================================
#  ページ種別の判定
# ============================================================
def classify_page(html: str) -> str:
    if not html:
        return "empty"
    low = html.lower()
    if "stock_table" in low:
        return "ok"
    if "captcha.awswaf.com" in low or "/captcha.js" in low or "awswafcaptcha" in low:
        return "captcha"
    if ("gokuprops" in low or "token.awswaf.com" in low
            or "challenge.js" in low or "awswafcookiedomainlist" in low):
        return "challenge"
    if "<table" in low:
        return "table_other"
    return "unknown"


async def get_waf_token(browser):
    try:
        for c in await browser.cookies.get_all():
            if getattr(c, "name", "") == "aws-waf-token":
                return getattr(c, "value", None)
    except Exception:
        return None
    return None


def parse_ranking(html: str):
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table", class_="stock_table")
    if not tables:
        all_tables = soup.find_all("table")
        if not all_tables:
            return None
        tables = [max(all_tables, key=lambda x: len(x.find_all("tr")))]
    biggest = max(tables, key=lambda x: len(x.find_all("tr")))
    try:
        df = pd.read_html(StringIO(str(biggest)))[0]
    except Exception:
        return None
    if df is None or df.empty or len(df) <= 2:
        return None
    head = "".join(map(str, df.columns)) + "".join(map(str, df.iloc[0].values))
    if any(kw in head for kw in WRONG_TABLE_KEYWORDS):
        return None
    return df


async def fetch_page_robust(browser, url, is_first=False):
    page = await browser.get(url)
    deadline = time.time() + (FIRST_WAIT if is_first else RETRY_WAIT)

    last_tr, stable = -1, 0
    reloaded = False
    token_since = None

    while time.time() < deadline:
        html = await page.get_content()
        kind = classify_page(html)

        if kind == "captcha":
            return None, "captcha", page

        if kind == "ok":
            tr = html.count("<tr")
            if tr == last_tr and tr > 3:
                stable += 1
                if stable >= 2:
                    df = parse_ranking(html)
                    if df is not None:
                        return df, "ok", page
            else:
                stable = 0
            last_tr = tr
        else:
            token = await get_waf_token(browser)
            if token and token_since is None:
                token_since = time.time()
            if token and (not reloaded) and token_since and (time.time() - token_since) > 6:
                try:
                    await page.reload()
                except Exception:
                    pass
                reloaded = True

        await asyncio.sleep(1.5)

    html = await page.get_content()
    if classify_page(html) == "ok":
        df = parse_ranking(html)
        if df is not None:
            return df, "ok", page
    return None, classify_page(html), page


async def warm_up(browser, max_wait=30):
    print("  [warm-up] 接続確認 ...", flush=True)
    page = await browser.get(TOP_URL)
    deadline = time.time() + max_wait
    while time.time() < deadline:
        html = await page.get_content()
        kind = classify_page(html)
        token = await get_waf_token(browser)

        if kind == "captcha":
            print("  [warm-up] ✗ 画像CAPTCHA（自動突破不可）", flush=True)
            return False
        if token:
            print("  [warm-up] ○ aws-waf-token 取得（チャレンジ通過）", flush=True)
            await asyncio.sleep(random.uniform(1, 2))
            return True
        if kind in ("ok", "table_other") and len(html) > 50000:
            print("  [warm-up] ○ チャレンジ無しでページ表示（通過不要）", flush=True)
            await asyncio.sleep(random.uniform(1, 2))
            return True
        await asyncio.sleep(1.5)

    print("  [warm-up] △ 確認できず。各ページ側で対応します", flush=True)
    return False


# ============================================================
#  メイン巡回処理（GitHub Actions：Xvfb上でheadless=Falseとして動作）
# ============================================================
async def scrape_all(targets):
    browser_args = [
        "--disable-blink-features=AutomationControlled",
        "--lang=ja-JP",
        "--accept-lang=ja-JP,ja",
        "--window-size=1920,1080",
        # GitHub Actionsのコンテナ的な実行環境向け（root実行・共有メモリ対策）
        "--no-sandbox",
        "--disable-dev-shm-usage",
    ]
    if PROXY_SERVER:
        browser_args.append(f"--proxy-server={PROXY_SERVER}")

    browser = await uc.start(
        headless=HEADLESS,              # Actionsでは既定でFalse（Xvfb配下で実Chrome相当として動かす）
        browser_executable_path=CHROME_PATH,
        browser_args=browser_args,
    )

    results = {}
    captcha_hit = False

    try:
        await warm_up(browser)
        is_first = True

        for target in targets:
            sheet_name = target["sheet"]
            target_pages = target.get("pages", DEFAULT_PAGES)
            print(f"  {sheet_name:<16}", end=" ", flush=True)

            page_dfs = {}
            for idx, p in enumerate(target_pages):
                url = target["url_tpl"].format(page=p)

                df, kind = None, "unknown"
                for attempt in range(PAGE_MAX_ATTEMPTS):
                    d, kind, _ = await fetch_page_robust(
                        browser, url, is_first=(is_first and attempt == 0)
                    )
                    is_first = False
                    if d is not None:
                        df = d
                        break
                    if kind == "captcha":
                        captcha_hit = True
                        break
                    await asyncio.sleep(3)

                page_dfs[p] = df
                print(f"p{p}={len(df) if df is not None else '×'}", end=" ", flush=True)
                await asyncio.sleep(random.uniform(2, 4))

            valid = [page_dfs[p] for p in target_pages
                     if page_dfs.get(p) is not None and not page_dfs[p].empty]
            if valid:
                combined = pd.concat(valid, ignore_index=True).drop_duplicates().reset_index(drop=True)
                results[sheet_name] = combined
                print(f"→ 計{len(combined)}行")
            else:
                results[sheet_name] = pd.DataFrame()
                print("→ 空")

    except Exception as e:
        print(f"スクレイピング中にエラー: {e}")

    finally:
        try:
            stop = browser.stop()
            if stop:
                await stop
        except Exception as e:
            print(f"ブラウザ終了エラー: {e}")
        print("ブラウザを終了しました。")

    if captcha_hit:
        print("\n⚠ 途中で画像CAPTCHAが発生しました。"
              "\n   GitHub ActionsのIPはデータセンター系のため、"
              "自宅回線より判定が厳しくなる傾向があります。"
              "\n   頻発する場合は、実行時間帯を変える・住宅用プロキシ（PROXY_SERVER）"
              "・J-Quants移行を検討してください。")

    return results, captcha_hit


# ============================================================
#  スプレッドシート書き込み
# ============================================================
def write_df_to_sheet(spreadsheet, sheet_name, df):
    if df is None or df.empty:
        print(f"  [{sheet_name}] データが空のため書き込みをスキップ")
        return

    n_rows, n_cols = df.shape
    try:
        ws = spreadsheet.worksheet(sheet_name)
        ws.clear()
        time.sleep(1.2)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=n_rows + 10, cols=n_cols + 5)
        time.sleep(1.2)

    ws.resize(rows=n_rows + 1, cols=n_cols)
    safe = df.fillna("").astype(str)

    ws.update(values=[safe.columns.tolist()], range_name="A1", value_input_option="USER_ENTERED")
    time.sleep(1.2)

    values = safe.values.tolist()
    for start in range(0, len(values), WRITE_CHUNK):
        chunk = values[start:start + WRITE_CHUNK]
        ws.update(values=chunk, range_name=f"A{start + 2}", value_input_option="USER_ENTERED")
        time.sleep(1.2)

    print(f"  [{sheet_name}] 書き込み完了: {n_rows}行 × {n_cols}列")


# ============================================================
#  認証
#    優先順位: GOOGLE_SERVICE_ACCOUNT_JSON（GitHub Secrets）
#            → ローカルの service_account.json（手元での動作確認用）
# ============================================================
def get_gspread_client():
    env_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if env_json:
        try:
            info = json.loads(env_json)
            gc = gspread.service_account_from_dict(info)
            print("サービスアカウントで認証しました（GitHub Secretsから読込）")
            return gc
        except Exception as e:
            print(f"認証エラー（GOOGLE_SERVICE_ACCOUNT_JSONの読込に失敗）: {e}")
            return None

    if not os.path.exists(SERVICE_ACCOUNT_KEY_PATH):
        print(f"★ 認証ファイルが見つかりません: {SERVICE_ACCOUNT_KEY_PATH}")
        print("  → GitHub Secrets『GOOGLE_SERVICE_ACCOUNT_JSON』を設定するか、")
        print("    ローカル実行時は service_account.json を同じフォルダに置いてください。")
        return None
    try:
        gc = gspread.service_account(filename=SERVICE_ACCOUNT_KEY_PATH)
        print(f"サービスアカウントで認証しました: {os.path.basename(SERVICE_ACCOUNT_KEY_PATH)}")
        return gc
    except Exception as e:
        print(f"認証エラー: {e}")
        return None


# ============================================================
#  メイン
# ============================================================
async def main():
    try:
        nodriver_ver = version("nodriver")
    except PackageNotFoundError:
        nodriver_ver = "不明"
    print(f"===== 環境情報 =====\nnodriver : {nodriver_ver}\nheadless : {HEADLESS}\n"
          f"SPREADSHEET_ID : {SPREADSHEET_ID}\n====================\n")

    gc = get_gspread_client()
    spreadsheet = None
    if gc is not None:
        try:
            spreadsheet = gc.open_by_key(SPREADSHEET_ID)
            print("スプレッドシートを開きました。")
        except gspread.exceptions.SpreadsheetNotFound:
            print(f"エラー: ID '{SPREADSHEET_ID}' のスプレッドシートが見つかりません。"
                  "（サービスアカウントに共有されているか確認してください）")
        except Exception as e:
            print(f"スプレッドシートを開く際にエラー: {e}")
    else:
        print("認証に失敗したため、書き込みはスキップします。")

    if spreadsheet is None:
        print("\nエラー: スプレッドシート未準備のため処理を中断します。")
        sys.exit(1)

    print("\n■ 株探ランキング情報の取得を開始します...")
    results, captcha_hit = await scrape_all(KABUTAN_TARGETS)

    print("\n■ スプレッドシートへの書き込みを開始します...")
    for sheet_name, df in results.items():
        write_df_to_sheet(spreadsheet, sheet_name, df)

    if captcha_hit:
        # 一部でもCAPTCHAが出た実行は、GitHub Actions上で「失敗」として
        # 通知させるため、あえて非ゼロ終了にしています。
        print("\n⚠ CAPTCHA発生のため、このジョブを失敗扱いで終了します（通知目的）。")
        sys.exit(1)


if __name__ == "__main__":
    t0 = time.time()
    uc.loop().run_until_complete(main())
    print(f"\n完了  所要時間: {time.time() - t0:.1f} 秒")
