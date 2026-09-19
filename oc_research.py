"""Gemini + Google Search による大学OC情報調査。"""
from __future__ import annotations
import json
from datetime import datetime, timezone

MODEL = "gemini-2.5-flash"

SYSTEM_PROMPT = """
あなたは大学オープンキャンパス情報の調査・整理担当です。

目的：指定された日本の大学について、指定年度のオープンキャンパス情報をWeb検索し、
チラシ掲載用に正確に整理してください。

必須ルール：
1. Google検索を使って必ず最新情報を確認すること。
2. 大学公式サイトを最優先すること。大学公式サイトで確認できない場合のみ、信頼できる第三者情報を補助的に使うこと。
3. 情報を推測・創作しないこと。確認できない項目は「要確認」とすること。
4. 日付・時間・会場・対象・予約方法などは、検索結果の内容を忠実に整理すること。
5. チラシ用の「日程」は短く整理してよいが、意味を変えないこと。
6. 各情報の根拠となるURLを、検索結果の引用情報から可能な限り取得すること。
7. 「関関同立」は大学名ではなく、関西大学・関西学院大学・同志社大学・立命館大学の4大学を意味する。
8. 出力はJSONのみとする。
"""

SCHEMA = {
    "type": "object",
    "properties": {
        "university": {"type": "string"},
        "schedule": {"type": "string"},
        "details": {"type": "string"},
        "status": {"type": "string", "enum": ["確認済み", "要確認", "情報なし"]}
    },
    "required": ["university", "schedule", "details", "status"]
}


def _get(obj, key, default=None):
    if isinstance(obj, dict):
        return obj.get(key, default)
    return getattr(obj, key, default)


def _extract_sources(response):
    sources = []
    candidates = _get(response, "candidates", []) or []
    for candidate in candidates:
        gm = _get(candidate, "grounding_metadata")
        chunks = _get(gm, "grounding_chunks", []) or []
        for chunk in chunks:
            web = _get(chunk, "web")
            if web:
                url = _get(web, "uri")
                title = _get(web, "title") or url
                if url and url not in {s["url"] for s in sources}:
                    sources.append({"title": title, "url": url})
    return sources


def research_university(university: str, year: int, api_key: str) -> dict:
    """1大学を検索し、構造化結果とGroundingの出典URLを返す。"""
    if not api_key:
        raise ValueError("GEMINI_API_KEY が設定されていません。")

    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    prompt = f"""
大学名：{university}
対象年度：{year}年度

この大学の{year}年度オープンキャンパスについて調査してください。
特に、実際に参加者向けチラシへ掲載できる開催日・時間・会場・予約要否などを確認してください。
複数日程がある場合は、主要な日程をすべてまとめてください。
大学公式サイトの情報を優先してください。
"""

    response = client.models.generate_content(
        model=MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT,
            response_mime_type="application/json",
            response_json_schema=SCHEMA,
            tools=[types.Tool(google_search=types.GoogleSearch())],
        ),
    )
    data = json.loads(response.text)
    data["sources"] = _extract_sources(response)
    data["checked_at"] = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
    return data
