# Gemini Google Search 対応パッチ

## 変更点
- `oc_research.py` は Gemini Google Search grounding を利用。
- `app.py` に任意のAI検索UIを追加。
- Gemini検索結果は既存の日程データへ反映できる。
- デフォルトでは既存の日程を上書きしない。
- AI検索後も既存のテキスト貼り付け、Excel/CSV、Google Sheets、画面上の直接編集を利用可能。
- AI検索で参照したURLを画面で確認可能。
- `GEMINI_API_KEY` がない場合でも、AI検索以外の機能は使用可能。

## Streamlit Secrets
```
GEMINI_API_KEY = "あなたのGemini APIキー"
```

Google AI Studioで取得したAPIキーを設定してください。
