# 業界情報チェック MVP

GoogleアラートのメールをGoogleスプレッドシートに取り込み、チームで確認するためのスターターキットです。

## 構成

- `index.html` / `styles.css` / `app.js`: 閲覧ページ
- `apps-script/Code.gs`: GmailのGoogleアラートをスプレッドシートへ取り込むApps Script
- `apps-script/Index.html`: チーム共有用のApps Script Webアプリ画面
- `data/sample-alerts.csv`: CSV連携確認用のサンプル

## スプレッドシート列

Apps Scriptは次の列で `alerts` シートを作ります。

```text
id, receivedAt, alertKeyword, title, source, url, snippet, status, priority, owner, category, tags, memo, publishedAt, importedAt
```

## Googleアラート取り込み

1. Googleスプレッドシートを新規作成する
2. `拡張機能` から `Apps Script` を開く
3. `apps-script/Code.gs` の内容を貼り付ける
4. Apps ScriptでHTMLファイル `Index` を作り、`apps-script/Index.html` の内容を貼り付ける
5. `setup` を実行して権限を許可する
6. `importGoogleAlerts` を実行する
7. 自動取り込みにする場合は `createHourlyTrigger` を一度だけ実行する
8. `デプロイ` からWebアプリとして公開する

対象メールは `CONFIG.gmailQuery` で調整できます。最初はGoogleアラートのメールだけを拾う設定です。

## 閲覧ページの使い方

ローカル確認:

```powershell
python -m http.server 8080
```

ブラウザで `http://localhost:8080/industry-info-hub/` を開きます。

スプレッドシート連携:

1. `alerts` シートをCSV形式で公開する
2. 公開CSVのURLを画面右上の `CSV URL` に入れる
3. `読み込み` を押す

ローカル版で編集したステータス、重要度、担当、タグ、メモはブラウザのローカル保存です。チームで同じ編集内容を共有する場合は、Apps Script版のWebアプリを使います。

## 次に固めるところ

- Googleアラートの検索キーワード一覧
- カテゴリの標準値
- 重要度の判断基準
- 誰が週次で確認するか
- スプレッドシートだけで運用するか、別DBに移すか
