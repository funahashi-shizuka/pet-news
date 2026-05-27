# ペット業界ニュース

Googleアラートのメールをスプレッドシートに入れて、HTMLでニュース一覧を見るだけのシンプルな構成です。

## ファイル

- `index.html`: ニュース一覧ページ
- `data/news-template.csv`: スプレッドシートの列見本
- `apps-script/Code.gs`: Googleアラートメールをスプレッドシートへ入れる最小スクリプト

## スプレッドシート

シート名は `ニュース` にします。スプレッドシートはデータベース扱いなので、列はこの6つだけです。

| 受信日 | キーワード | タイトル | 媒体 | URL | 概要 |
|---|---|---|---|---|---|

## Gmailから取り込む

1. Googleスプレッドシートを作る
2. `拡張機能` から `Apps Script` を開く
3. `apps-script/Code.gs` を貼り付ける
4. `setup` を実行する
5. `previewGoogleAlerts` を実行して、分解結果をログで確認する
6. 問題なければ `importGoogleAlerts` を実行する
7. 自動化する場合だけ `createHourlyTrigger` を一度実行する

## HTMLで見る

スプレッドシートをCSV公開して、そのURLを `index.html` の `SHEET_CSV_URL` に入れます。

ローカルではサンプルCSVを表示します。

`?csv=公開CSVのURL` を付けても読み込めます。
