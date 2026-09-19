# OCチラシ生成アプリ v3.1 編集内容

## 変更
- 「関関同立」を関西大学・関西学院大学・同志社大学・立命館大学の4大学へ独立展開
- Word内の大学名に先頭1文字分の空白を統一して付与
- Gemini + Google検索によるオープンキャンパス情報調査
- Gemini調査結果の出典URLをアプリ上で表示
- 調査結果をチラシの日程欄へ優先反映
- 作成日時・都道府県・年度・大学数を履歴表示
- 出力ファイル名の年度可変は既存仕様を維持

## Gemini設定
Streamlit Cloud の Secrets に以下を設定してください。

```toml
GEMINI_API_KEY = "YOUR_API_KEY"
```

または環境変数 `GEMINI_API_KEY` でも可。

## 注意
作成履歴は `logs/creation_history.json` に保存します。
Streamlit Cloud の再起動・再デプロイをまたいで永久保存したい場合は、次段階でGoogle Drive等への保存機能を追加するのが安全です。

GeminiのWeb検索はGoogle Search groundingを使用します。検索回数・料金はGoogleのGemini API料金体系に従います。
