JPX信用取引週末残高 自動取得パイプライン
JPX（日本取引所グループ）が毎週火曜16:30頃に公表する「銘柄別信用取引週末残高」PDFを
GitHub Actionsで自動取得し、Googleスプレッドシートに週次で列追加していく。
事前準備
1. Googleサービスアカウントの作成
Google Cloud Consoleでプロジェクトを作成（既存のものでも可）
「APIとサービス」→「有効なAPI」で Google Sheets API を有効化
「認証情報」→「サービスアカウントを作成」
作成したサービスアカウントの「キー」タブから JSON形式のキーを新規作成・ダウンロード
ダウンロードしたJSONファイルの中の `client_email` の値（例: `xxx@yyy.iam.gserviceaccount.com`）を確認
2. スプレッドシートへの共有設定
対象のスプレッドシート（`SPREADSHEET_ID`）を開き、共有設定で上記の `client_email` を
編集者として追加する。これをしないと書き込み権限エラーになる。
3. GitHub Secretsの登録
リポジトリの Settings → Secrets and variables → Actions で以下を登録:
Secret名	内容
`GOOGLE_SERVICE_ACCOUNT_JSON`	ダウンロードしたJSONキーファイルの中身をそのまま貼り付け
`SPREADSHEET_ID`	対象スプレッドシートのID
ファイル構成
```
.github/workflows/jpx_margin.yml   # 毎週火曜18:00 JSTに自動実行
scripts/fetch_jpx_margin.py        # PDF取得〜パース〜書き込み本体
scripts/requirements.txt           # 依存パッケージ
```
動作確認（重要・最初に必ず実施）
PDFの実際の列レイアウトを確認せずに `parse_margin_pdf()` の列インデックスを仮置きしています。
本番投入前に、ローカルで一度dry-runして抽出結果を目視確認してください。
```bash
pip install -r scripts/requirements.txt

export GOOGLE_SERVICE_ACCOUNT_JSON="$(cat your-service-account-key.json)"
export SPREADSHEET_ID="1FP7DAwwGm8Pq5ZIP0kZ_ME7eGoJFvEWhvUsjpiwC2Bw"
export TARGET_DATE_OVERRIDE="20260710"   # 過去の確定分で試す

python scripts/fetch_jpx_margin.py
```
`parse_margin_pdf()` 内で `print(rows[:5])` などを一時的に挟んで実際の列構成を確認し、
`_to_number` を当てる列番号（買い残・売り残・倍率がそれぞれ何番目の列か）を実データに合わせて調整してください。
手動実行（GitHub Actions側）
Actionsタブ →「JPX信用取引週末残高 取得」→「Run workflow」から、
`target_date` に `YYYYMMDD` 形式で過去分を指定して手動実行できる。
既知の注意点
2026年9月28日にJPX側の集計システムが移行予定。現在のPDF URLパターン
（`syumatsu{YYYYMMDD}00.pdf`）は使えなくなる可能性が高い。移行後は
`JPX_PDF_URL_TEMPLATE` を新形式に合わせて更新すること。
祝日等で公表が1日ずれることがある。PDFが404の場合はワークフローは失敗扱いにせず
正常終了するよう設計している（`sys.exit(0)`）ので、翌日の手動再実行で拾える。
現状は「買い残・売り残・倍率」の3列固定。制度信用／一般信用の内訳が必要な場合は
`parse_margin_pdf()` と `update_spreadsheet()` の両方を拡張する必要がある。
