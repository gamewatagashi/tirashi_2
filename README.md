# 🎓 東進 OC チラシ 自動生成ツール

オープンキャンパスのチラシ（Word/PDF）を自動生成するStreamlitアプリです。

## 機能

- **① 単体作成**：都道府県を選択 → 大学名・日程を確認・編集 → 表紙写真を選択
  （ライブラリ／Driveに複数候補があれば選べます、または自分の端末から直接アップロード）
  → Word・PDFを生成
- **② まとめて作成**：複数の都道府県を選んで一括生成し、ZIPでまとめてダウンロード
  （写真はライブラリ／Google Driveから自動選択・候補が複数あれば最初の1枚）
- **⚙️ Google Drive連携（任意）**：設定した人だけが使える機能。写真ライブラリや
  Wordテンプレートを、GitHubに含めずに共有Google Driveフォルダから直接読み込めます
- 掲載大学数（6 / 8 / 10 / 12）に応じてテンプレートを自動選択
- 裏面PDF（参加チェックリスト）と自動結合

公開（デプロイ）手順は [DEPLOY.md](./DEPLOY.md) を参照してください。

## 写真の扱いについて（気密性）

- **共有ライブラリ／Drive経由**：管理者が用意した写真フォルダから、大学名で自動的に
  候補が絞り込まれます。候補が複数ある場合は画面上で選べます。
- **個人アップロード**：「📤 自分の端末からアップロード」を選ぶと、その場でアップロード
  した写真だけが使われます。共有ライブラリやDriveには一切保存されません（生成中のみ
  一時的にサーバー上のセッション用フォルダに置かれ、他の利用者からは見えません）。
  まだ公開されていない写真や個人が用意した写真を使いたいときはこちらを使ってください。

## セットアップ

### 必要なファイル（gitignoreに含める）

以下のファイルは容量のため Git 管理外です。各自でフォルダに配置してください。

```
oc_app/
├── templates/
│   ├── template_6.docx     ← 6大学数フォーマット
│   ├── template_8.docx     ← 8大学数フォーマット
│   ├── template_10.docx    ← 10大学数フォーマット
│   ├── template_12.docx    ← 12大学数フォーマット
│   └── back_page.pdf       ← 裏面（参加チェックリスト）
├── images/
│   └── *.jpg               ← 大学写真（pixta素材）
└── data/
    └── university_map.xlsx  ← 都道府県×大学マスタ
```

### ローカル起動

```bash
pip install -r requirements.txt
# LibreOffice も必要（PDF変換用）
# macOS: brew install libreoffice
# Ubuntu: sudo apt install libreoffice

cd oc_app
streamlit run app.py
```

### Streamlit Community Cloud にデプロイ

詳しい手順は [DEPLOY.md](./DEPLOY.md) にまとめてあります。概要：

1. このリポジトリを GitHub にプッシュ（`templates/` `images/` `data/` は含まれません）
2. [share.streamlit.io](https://share.streamlit.io) でリポジトリを選択
3. Main file: `oc_app/app.py` を指定
4. テンプレート・写真をコードと分けて管理したい場合は「⚙️ Google Drive連携」を設定
   （DEPLOY.md 手順4）。分けなくてよい場合は `.gitignore` からその行を削除してリポジトリに含める

> ⚠️ LibreOffice が必要なPDF変換はCloud環境では追加設定が必要です（`packages.txt` に `libreoffice` を追加、設定済み）

## OC日程データの入力方法

サイドバーの入力欄に、Googleスプレッドシートからコピーしたデータを貼り付けます。

**形式：** `大学名[TAB]日程テキスト`（1行1大学）

```
大阪大学	6/13(人科)・6/27(外語)・8/4-19 各学部
京都大学	8/6・8/7 吉田キャンパス（来場型）
神戸大学	8/7・8/8・8/10 各学部
```

スプレッドシートの列をそのまま選択してコピーすると自動的にTAB区切りになります。

## ファイル構成

```
oc_app/
├── app.py          ← Streamlit メインアプリ（①単体作成 ②まとめて作成 ③Drive連携）
├── generator.py    ← チラシ生成コアロジック（単体・一括ZIP生成）
├── drive_helper.py ← Google Drive 連携（任意機能）
├── requirements.txt
├── README.md
├── DEPLOY.md       ← 公開手順・Drive連携の設定手順
├── packages.txt    ← Streamlit Cloud 用（LibreOffice）
├── templates/      ← Word テンプレート・裏面PDF（gitignore）
├── images/         ← 大学写真（gitignore）
└── data/           ← マスタデータ（gitignore）
```

## .gitignore の設定

```gitignore
oc_app/templates/
oc_app/images/
oc_app/data/
oc_app/output/
__pycache__/
*.pyc
```
