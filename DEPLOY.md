# Streamlit Community Cloud 公開手順

## 1. GitHub
このフォルダの中身をGitHubリポジトリ直下にアップロードしてください。
`images/` `data/` `templates/` もアップロードします。

## 2. Streamlit Community Cloud
1. https://share.streamlit.io/ を開く
2. GitHubでログイン
3. Create app / New app
4. Repository: 作成したリポジトリ
5. Branch: main
6. Main file path: app.py
7. Deploy

## 3. Google Drive（任意）
Google Driveを使わなくても、内蔵写真・利用者の写真アップロード・Excel/CSV/テキスト入力で利用できます。
Google Driveを使う場合は、アプリの「Google Drive連携」タブに表示される設定手順に従い、Streamlit CloudのSecretsへ認証情報を登録してください。

## 4. 写真
標準写真は `images/` に45枚入っています。
利用者は大学ごとにPCから写真をアップロードして差し替えることもできます。

## 5. 大学情報
- 画面に直接テキストを貼り付け
- Excel / CSVをアップロード
- Google Sheets URLから取得（公開設定等が必要）
が使えます。
