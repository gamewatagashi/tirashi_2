"""
sheet_parser.py
スプレッドシート（Excel / CSV / Google Sheets URL）から
OC日程データを読み込むユーティリティ。

対応フォーマット:
  A) 大学情報スプシ（シート1）: 大学名列 + 日程列 の2列以上
  B) 分担スプシ（シート2相当）: 都道府県番号 / 都道府県 / 掲載大学 の3列
  C) Googleスプレッドシート共有URL（/edit → /export?format=xlsx に変換してDL）

戻り値は常に以下の2つのdict（どちらかがNoneの場合もある）:
  oc_schedule : {大学名: 日程テキスト}
  pref_map    : {都道府県名: [大学名, ...]}  (分担スプシから読んだ場合のみ)
"""

from __future__ import annotations
import io
import re
import unicodedata
import pandas as pd
from univ_utils import canonical_university_name


# ── 内部ユーティリティ ─────────────────────────────────────────

def _normalize(text: str) -> str:
    """全角→半角、余白除去などの正規化。"""
    if not text:
        return ""
    text = str(text).strip()
    text = unicodedata.normalize("NFKC", text)
    return text


def _is_univ_name(val) -> bool:
    """セルが大学名として扱えるか（空・数値・日付は除外）。"""
    if val is None:
        return False
    s = str(val).strip()
    if not s or s in ("None", "nan"):
        return False
    # 数値のみは大学名ではない
    try:
        float(s)
        return False
    except ValueError:
        pass
    return True


# ── フォーマット A: 大学OC情報スプシ ────────────────────────────

def _parse_oc_sheet(df: pd.DataFrame) -> dict[str, str]:
    """
    シート1形式:
      列0: カテゴリ（東京一工 など、Noneの行は前のカテゴリが続く）
      列1: 大学名
      列4: 日程テキスト
    → {大学名: 日程} を返す
    """
    schedule: dict[str, str] = {}

    # ヘッダー行をスキップ（大学名らしい列を探す）
    univ_col = None
    date_col  = None

    for col_idx in range(min(len(df.columns), 10)):
        sample = [str(v) for v in df.iloc[:, col_idx].dropna().head(10)]
        # 「大学」を含む値が多い列 → 大学名列
        if sum("大学" in s or s.endswith("大") or "大、" in s or "大 " in s for s in sample) >= 2:
            univ_col = col_idx
            break

    if univ_col is None:
        return schedule

    # 日程列: 日付パターンだけでなく「今日」「明日」「未定」なども認識。
    # それでも見つからない場合は、大学名列の直後にある非空列を候補にする。
    date_patterns = re.compile(r'\d+[/月]\d+|オンライン|来場|事前|今日|明日|未定|随時|開催')
    for col_idx in range(univ_col + 1, min(len(df.columns), 12)):
        sample = [str(v).strip() for v in df.iloc[:, col_idx].dropna().head(30)]
        if sum(bool(date_patterns.search(s)) for s in sample) >= 1:
            date_col = col_idx
            break

    if date_col is None:
        for col_idx in range(univ_col + 1, min(len(df.columns), 12)):
            sample = [str(v).strip() for v in df.iloc[:, col_idx].dropna().head(30)]
            nonempty = [s for s in sample if s and s.lower() not in ('nan','none')]
            if len(nonempty) >= 2:
                date_col = col_idx
                break

    if date_col is None:
        for _, row in df.iterrows():
            val = row.iloc[univ_col] if univ_col < len(row) else None
            if _is_univ_name(val):
                name = canonical_university_name(_normalize(str(val)))
                if name:
                    schedule[name] = ""
        return schedule

    for _, row in df.iterrows():
        univ_val = row.iloc[univ_col] if univ_col < len(row) else None
        date_val = row.iloc[date_col] if date_col < len(row) else None

        if not _is_univ_name(univ_val):
            continue

        name     = canonical_university_name(_normalize(str(univ_val)))
        date_str = _normalize(str(date_val)) if date_val not in (None, "", "nan", "None") else ""

        # 改行を「／」に変換して短くする
        date_str = re.sub(r'\s*\n\s*', ' / ', date_str)
        # 長すぎる日程は先頭200文字だけ
        if len(date_str) > 200:
            date_str = date_str[:197] + "…"

        if name:
            schedule[name] = date_str

    return schedule


# ── フォーマット B: 分担スプシ（都道府県×大学） ───────────────────

def _parse_pref_sheet(df: pd.DataFrame) -> dict[str, list[str]]:
    """
    シート2 / シート4形式:
      都道府県 | 掲載大学（カンマ・読点区切り）
    → {都道府県名: [大学名, ...]}
    """
    pref_map: dict[str, list[str]] = {}

    pref_col  = None
    univ_col  = None

    prefs_47 = {
        "北海道","青森","岩手","宮城","秋田","山形","福島","茨城","栃木","群馬",
        "埼玉","千葉","東京","神奈川","新潟","富山","石川","福井","山梨","長野",
        "岐阜","静岡","愛知","三重","滋賀","京都","大阪","兵庫","奈良","和歌山",
        "鳥取","島根","岡山","広島","山口","徳島","香川","愛媛","高知","福岡",
        "佐賀","長崎","熊本","大分","宮崎","鹿児島","沖縄",
    }

    # 都道府県列を探す
    for col_idx in range(min(len(df.columns), 5)):
        sample = [str(v).strip() for v in df.iloc[:, col_idx].dropna().head(15)]
        hit = sum(any(p in s for p in prefs_47) for s in sample)
        if hit >= 3:
            pref_col = col_idx
            break

    if pref_col is None:
        return pref_map

    # 大学列: 都道府県列の右隣で「大学」を多く含む列
    for col_idx in range(pref_col + 1, min(len(df.columns), pref_col + 5)):
        sample = [str(v) for v in df.iloc[:, col_idx].dropna().head(15)]
        if sum("大学" in s or s.endswith("大") or "大、" in s or "大 " in s for s in sample) >= 3:
            univ_col = col_idx
            break

    if univ_col is None:
        return pref_map

    for _, row in df.iterrows():
        pref_val = row.iloc[pref_col] if pref_col < len(row) else None
        univ_val = row.iloc[univ_col] if univ_col < len(row) else None

        if not _is_univ_name(pref_val) or not _is_univ_name(univ_val):
            continue

        pref = _normalize(str(pref_val))
        raw  = _normalize(str(univ_val))

        # Skip header rows
        if pref in ('都道府県', '都道府県名', 'prefecture', '県', '府') or            raw in ('掲載大学', '大学名', '大学一覧', '例年チラシに載せる大学（変更する可能性あり）') or            raw.startswith('例年'):
            continue

        # 「（）」内の補足を除去して分割
        raw = re.sub(r'[（(][^）)]*[）)]', '', raw)
        univs = [canonical_university_name(u.strip()) for u in re.split(r'[、,，]+', raw) if u.strip() and len(u.strip()) >= 2]

        if pref and univs:
            pref_map[pref] = univs

    return pref_map


# ── 公開API ──────────────────────────────────────────────────────

def parse_spreadsheet(
    file_bytes: bytes,
    filename: str = "upload.xlsx",
) -> tuple[dict[str, str], dict[str, list[str]]]:
    """
    アップロードされたExcel / CSVバイト列を解析して返す。

    Returns:
        oc_schedule : {大学名: 日程テキスト}
        pref_map    : {都道府県名: [大学名, ...]}
    """
    oc_schedule: dict[str, str] = {}
    pref_map:    dict[str, list[str]] = {}

    ext = filename.rsplit(".", 1)[-1].lower()

    if ext == "csv":
        # CSV: 1シートだけ、まずOC情報として解析
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=str)
            oc_schedule = _parse_oc_sheet(df)
            pm = _parse_pref_sheet(df)
            if pm:
                pref_map = pm
        except Exception as e:
            raise ValueError(f"CSV読み込みエラー: {e}") from e
        return oc_schedule, pref_map

    # Excel: 全シートを走査
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception as e:
        raise ValueError(f"Excelファイルの読み込みに失敗しました: {e}") from e

    for sheet_name in xls.sheet_names:
        try:
            df = xls.parse(sheet_name, header=None, dtype=str)
        except Exception:
            continue

        # OC日程シート（日付パターンが多い）
        date_count = sum(
            bool(re.search(r'\d+[/月]\d+|オンライン|来場', str(v)))
            for v in df.values.flatten() if v and str(v) != "nan"
        )
        if date_count >= 5:
            sched = _parse_oc_sheet(df)
            oc_schedule.update(sched)

        # 分担シート（都道府県×大学）
        pm = _parse_pref_sheet(df)
        if pm:
            pref_map.update(pm)

    return oc_schedule, pref_map


def parse_google_sheets_url(url: str) -> tuple[dict[str, str], dict[str, list[str]]]:
    """
    Google Sheets の共有URL（/edit や /pub）を受け取り、
    xlsx形式でダウンロードして parse_spreadsheet() に渡す。
    """
    import urllib.request

    # https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit...
    # → https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/export?format=xlsx
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if not m:
        raise ValueError("Google Sheets のURLではありません。\n"
                         "「共有」→「リンクをコピー」で取得したURLを貼り付けてください。")

    sheet_id  = m.group(1)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    try:
        req = urllib.request.Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            file_bytes = resp.read()
    except Exception as e:
        raise ValueError(
            f"スプレッドシートのダウンロードに失敗しました。\n"
            f"「リンクを知っている全員が閲覧可能」に設定されているか確認してください。\n"
            f"詳細: {e}"
        ) from e

    return parse_spreadsheet(file_bytes, "google_sheets.xlsx")
