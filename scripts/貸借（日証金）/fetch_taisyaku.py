name: 日証金 貸借取引残高（日次）

on:
  schedule:
    - cron: '45 3 * * 1-5'   # 12:45 JST（公表 11:30 頃の後）
    - cron: '30 7 * * 1-5'   # 16:30 JST 再実行（遅延・取りこぼし対策。同じ日付は上書きなので二重にならない）
  workflow_dispatch:
    inputs:
      dry_run:
        description: 'シートに書かず取得だけ試す'
        type: boolean
        default: false

permissions:
  contents: write

concurrency:
  group: taisyaku-daily
  cancel-in-progress: false

jobs:
  fetch:
    runs-on: ubuntu-latest
    timeout-minutes: 15
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'
      - run: pip install requests gspread
      - name: 取得してシートへ書き込み
        env:
          GCP_SA_KEY: ${{ secrets.GCP_SA_KEY }}
          TAISYAKU_SHEET_ID: ${{ secrets.TAISYAKU_SHEET_ID }}
          DRY_RUN: ${{ inputs.dry_run && '1' || '0' }}
        run: python scripts/fetch_taisyaku.py
      - name: 生CSVをコミット
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "41898283+github-actions[bot]@users.noreply.github.com"
          git add data/taisyaku || true
          git diff --cached --quiet || (git commit -m "日証金 zandaka $(TZ=Asia/Tokyo date +%F)" && git push)
