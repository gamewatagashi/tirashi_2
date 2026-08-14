"""
東進 オープンキャンパス チラシ 自動生成アプリ  v3.0
"""
import os, sys, tempfile, io
import streamlit as st
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from generator import (
    generate_chirashi, generate_batch_zip, docx_to_pdf, merge_pdfs,
    get_available_images, find_best_image, find_all_images, get_template_path,
)
from sheet_parser import parse_spreadsheet, parse_google_sheets_url
import drive_helper as dh

# ── Page config ───────────────────────────────────────────────
st.set_page_config(page_title="OC チラシ生成", page_icon="🎓", layout="wide")
st.title("🎓 オープンキャンパス チラシ 自動生成")
st.caption("東進ハイスクール / 東進衛星予備校")

# ── Constants ─────────────────────────────────────────────────
DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "university_map.xlsx")
BACK_PDF  = os.path.join(os.path.dirname(__file__), "templates", "back_page.pdf")

# ── Session state init ────────────────────────────────────────
if "session_tmp_dir" not in st.session_state:
    st.session_state.session_tmp_dir = tempfile.mkdtemp(prefix="oc_session_")
if "oc_schedule" not in st.session_state:
    st.session_state.oc_schedule = {}      # {大学名: 日程}
if "pref_map_imported" not in st.session_state:
    st.session_state.pref_map_imported = {}  # {都道府県名: [大学名,...]}  スプシから

SESSION_TMP = st.session_state.session_tmp_dir

# ── Load master data ──────────────────────────────────────────
@st.cache_data
def load_master():
    df = pd.read_excel(DATA_PATH)
    df.columns = ['番号', '都道府県', '掲載大学']
    df = df.dropna(subset=['番号', '都道府県'])
    df['番号'] = df['番号'].astype(int)
    return df

@st.cache_data
def load_images():
    return get_available_images()

df       = load_master()
img_dict = load_images()
DRIVE_OK = dh.drive_enabled()

# ── Helpers ───────────────────────────────────────────────────
def pref_suffix(name: str) -> str:
    if name in ['大阪', '京都']: return '府'
    if name == '東京':           return '都'
    if name == '北海道':         return ''
    return '県'

def resolve_campus_image(univ_name: str, uploaded_file, tmp_dir: str, use_drive: bool):
    if uploaded_file is not None:
        ext  = uploaded_file.name.rsplit('.', 1)[-1]
        path = os.path.join(tmp_dir, f'_upload.{ext}')
        with open(path, 'wb') as f:
            f.write(uploaded_file.getvalue())
        return path
    local = find_best_image(univ_name)
    if local:
        return local
    if use_drive and DRIVE_OK:
        folder_id = st.secrets.get("drive_images_folder_id")
        match = dh.find_drive_image(univ_name, folder_id)
        if match:
            ext  = match['name'].rsplit('.', 1)[-1]
            path = os.path.join(tmp_dir, f'_drive_{univ_name}.{ext}')
            try:
                dh.download_drive_file(match['id'], path)
                return path
            except Exception:
                return None
    return None

def resolve_template_path(tmpl_key: int, tmp_dir: str, use_drive_template: bool):
    local_path, key = get_template_path(tmpl_key)
    if use_drive_template and dh.templates_from_drive_enabled():
        folder_id = st.secrets.get("drive_templates_folder_id")
        match = dh.find_drive_template(key, folder_id)
        if match:
            dest = os.path.join(tmp_dir, f'_drive_template_{key}.docx')
            try:
                dh.download_drive_file(match['id'], dest)
                return dest, key
            except Exception:
                pass
    return local_path, key

# ════════════════════════════════════════════════════════════════
# SIDEBAR  ── OC日程データの読み込み（全タブ共通）
# ════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("📊 OC日程データの読み込み")

    import_method = st.radio(
        "読み込み方法",
        ["📋 テキストで貼り付け", "📁 ファイルをアップロード", "🔗 Google SheetsのURLを入力"],
        label_visibility="collapsed",
    )

    # ── A: テキスト貼り付け ──────────────────────────────────
    if import_method == "📋 テキストで貼り付け":
        oc_raw = st.text_area(
            "大学名 [TAB] 日程（1行1大学）",
            height=240,
            placeholder=(
                "例（スプレッドシートからコピー）:\n"
                "大阪大学\t6/13(人科)・6/27(外語)・8/4-19 各学部\n"
                "京都大学\t8/6・8/7 吉田キャンパス\n"
                "神戸大学\t8/7・8/8・8/10 各学部"
            ),
            help="Googleスプレッドシートの列をそのまま選択→コピー→貼り付けできます",
        )
        if st.button("反映する", key="paste_apply"):
            sched = {}
            for line in oc_raw.strip().splitlines():
                parts = line.split('\t')
                if len(parts) >= 2:
                    sched[parts[0].strip()] = '\t'.join(parts[1:]).strip()
                elif parts[0].strip():
                    sched[parts[0].strip()] = ''
            st.session_state.oc_schedule.update(sched)
            st.success(f"{len(sched)}件を反映しました")

    # ── B: ファイルアップロード ──────────────────────────────
    elif import_method == "📁 ファイルをアップロード":
        st.caption("Excel（.xlsx）またはCSV（.csv）をアップロード")
        uploaded_sheet = st.file_uploader(
            "スプレッドシートをアップロード",
            type=["xlsx", "xls", "csv"],
            label_visibility="collapsed",
        )
        if uploaded_sheet and st.button("読み込む", key="file_import"):
            with st.spinner("読み込み中..."):
                try:
                    sched, pref_map = parse_spreadsheet(
                        uploaded_sheet.getvalue(), uploaded_sheet.name
                    )
                    st.session_state.oc_schedule.update(sched)
                    if pref_map:
                        st.session_state.pref_map_imported.update(pref_map)
                    msg = f"日程データ {len(sched)}件を読み込みました"
                    if pref_map:
                        msg += f"、都道府県×大学マッピング {len(pref_map)}件も読み込みました"
                    st.success(msg)
                except Exception as e:
                    st.error(f"読み込みエラー: {e}")

    # ── C: Google Sheets URL ─────────────────────────────────
    else:
        st.caption("「共有」→「リンクをコピー」で取得したURLを貼り付け")
        gs_url = st.text_input(
            "Google Sheets URL",
            placeholder="https://docs.google.com/spreadsheets/d/...",
            label_visibility="collapsed",
        )
        st.caption("⚠️ スプレッドシートを「リンクを知っている全員が閲覧可能」に設定してください")
        if gs_url and st.button("読み込む", key="gs_import"):
            with st.spinner("Google Sheetsからダウンロード中..."):
                try:
                    sched, pref_map = parse_google_sheets_url(gs_url)
                    st.session_state.oc_schedule.update(sched)
                    if pref_map:
                        st.session_state.pref_map_imported.update(pref_map)
                    msg = f"日程データ {len(sched)}件を読み込みました"
                    if pref_map:
                        msg += f"、都道府県×大学 {len(pref_map)}件も読み込みました"
                    st.success(msg)
                except Exception as e:
                    st.error(str(e))

    # ── 読み込み済みデータの確認 ─────────────────────────────
    st.divider()
    oc_schedule = st.session_state.oc_schedule
    if oc_schedule:
        st.caption(f"✅ 日程データ読み込み済み: **{len(oc_schedule)}大学**")
        with st.expander("読み込み済み大学一覧"):
            for name, sched in list(oc_schedule.items())[:30]:
                st.caption(f"**{name}**: {sched[:40]}{'…' if len(sched)>40 else ''}")
            if len(oc_schedule) > 30:
                st.caption(f"…他 {len(oc_schedule)-30} 件")
        if st.button("🗑 データをクリア", key="clear_schedule"):
            st.session_state.oc_schedule = {}
            st.session_state.pref_map_imported = {}
            st.rerun()
    else:
        st.caption("日程データ未読み込み（上で読み込んでください）")

    st.divider()
    if DRIVE_OK:
        st.success("✅ Google Drive 連携: 有効")
    else:
        st.info("ℹ️ Google Drive 連携: 未設定")

# ════════════════════════════════════════════════════════════════
# TABS
# ════════════════════════════════════════════════════════════════
tab_single, tab_batch, tab_drive = st.tabs(
    ["① 単体作成", "② まとめて作成", "⚙️ Google Drive連携"]
)

# ════════════════════════════════════════════════════════════════
# TAB 1: 単体作成
# ════════════════════════════════════════════════════════════════
with tab_single:
    st.subheader("① 都道府県を選択")
    pref_options = df.apply(lambda r: f"{int(r['番号']):02d}  {r['都道府県']}", axis=1).tolist()
    default_idx  = next((i for i, s in enumerate(pref_options) if '27' in s), 0)
    selected     = st.selectbox(
        "都道府県", pref_options, index=default_idx,
        label_visibility="collapsed", key="single_pref",
    )
    pref_num  = int(selected.split()[0])
    pref_name = selected.split()[1]

    row = df[df['番号'] == pref_num].iloc[0]
    default_univs_str = str(row['掲載大学'])

    # スプシからインポートした都道府県マッピングがあれば優先
    if pref_name in st.session_state.pref_map_imported:
        default_list = st.session_state.pref_map_imported[pref_name]
    else:
        raw_list     = [u.strip().strip('（）()') for u in default_univs_str.replace('、', ',').split(',')]
        default_list = [u for u in raw_list if u]

    st.subheader("② 掲載大学数・テンプレート")
    col_a, col_b, col_c = st.columns([1, 2, 2])
    with col_a:
        num_univ = st.select_slider("掲載大学数", options=[6, 8, 10, 12], value=12, key="single_num")
    with col_b:
        _, tmpl_key = get_template_path(num_univ)
        st.info(f"テンプレート: **{tmpl_key}大学フォーマット**")
    with col_c:
        use_drive_template = False
        if dh.templates_from_drive_enabled():
            use_drive_template = st.checkbox("テンプレートをDriveから", key="single_tmpl_drive")

    st.subheader("③ 大学名・日程を確認・編集")
    rows = []
    for i in range(num_univ):
        name     = default_list[i] if i < len(default_list) else ''
        schedule = oc_schedule.get(name, '')
        img_ok   = '✅' if find_best_image(name) else '❓'
        rows.append({'大学名': name, '日程': schedule, '写真': img_ok})

    edited = st.data_editor(
        pd.DataFrame(rows), num_rows="fixed", use_container_width=True,
        column_config={
            '大学名': st.column_config.TextColumn('大学名', width='medium'),
            '日程':   st.column_config.TextColumn('日程・場所', width='large'),
            '写真':   st.column_config.TextColumn('写真', width='small', disabled=True),
        },
        key="single_univ_table",
    )
    for i, r in edited.iterrows():
        edited.at[i, '写真'] = '✅' if find_best_image(r['大学名']) else '❓'

    main_univ = edited['大学名'].iloc[0] if len(edited) > 0 else ''

    # ── 写真選択 ─────────────────────────────────────────────
    st.subheader("④ 表紙写真を選択")
    photo_mode = st.radio(
        "写真の選び方",
        ["📚 ライブラリから選ぶ", "📤 端末からアップロード"],
        horizontal=True, key="single_photo_mode",
    )

    selected_campus_path = None

    if photo_mode.startswith("📤"):
        custom_img_file = st.file_uploader(
            "写真をアップロード（この生成にのみ使用）",
            type=['jpg', 'jpeg', 'png'], key="single_upload",
        )
        if custom_img_file:
            ext = custom_img_file.name.rsplit('.', 1)[-1]
            selected_campus_path = os.path.join(SESSION_TMP, f"_upload_{main_univ}.{ext}")
            with open(selected_campus_path, 'wb') as f:
                f.write(custom_img_file.getvalue())
            st.image(selected_campus_path, caption="アップロードした写真", width=360)
    else:
        candidates = [{'label': os.path.basename(p), 'path': p}
                      for p in find_all_images(main_univ)]
        if DRIVE_OK and st.checkbox("Google Drive の写真も候補に含める", value=True,
                                    key="single_use_drive_photo"):
            folder_id = st.secrets.get("drive_images_folder_id")
            try:
                for f in dh.find_all_drive_images(main_univ, folder_id):
                    candidates.append({'label': f"☁️ {f['name']}", 'drive_id': f['id']})
            except Exception as e:
                st.warning(f"Drive検索エラー: {e}")

        if not candidates:
            st.warning(f"「{main_univ}」の写真が見つかりません。「端末からアップロード」を使ってください。")
        else:
            # Drive候補をローカルにキャッシュ
            for cand in candidates:
                if 'path' not in cand and 'drive_id' in cand:
                    cache = os.path.join(SESSION_TMP, f"_drivecache_{cand['drive_id']}.jpg")
                    if not os.path.exists(cache):
                        try:
                            with open(cache, 'wb') as f:
                                f.write(dh.download_drive_file_bytes(cand['drive_id']))
                        except Exception:
                            continue
                    cand['path'] = cache

            # サムネイル表示
            cols = st.columns(min(4, len(candidates)))
            for i, cand in enumerate(candidates):
                with cols[i % len(cols)]:
                    if cand.get('path'):
                        st.image(cand['path'], use_container_width=True, caption=cand['label'])

            choice_idx = st.radio(
                f"「{main_univ}」に使う写真",
                list(range(len(candidates))),
                format_func=lambda i: candidates[i]['label'],
                horizontal=True, key=f"single_photo_choice_{main_univ}",
            )
            selected_campus_path = candidates[choice_idx].get('path')

    # ── 出力設定 & 生成ボタン ────────────────────────────────
    st.divider()
    col_o1, col_o2 = st.columns(2)
    with col_o1:
        do_pdf  = st.checkbox("PDF版も生成する", value=True, key="single_do_pdf")
    with col_o2:
        do_back = st.checkbox("裏面PDFと結合する", value=True, key="single_do_back")

    if st.button("🚀 チラシを生成する", type="primary", use_container_width=True, key="single_generate"):
        univs = [
            {'name': r['大学名'].strip(), 'schedule': r['日程'].strip()}
            for _, r in edited.iterrows() if r['大学名'].strip()
        ]
        if not univs:
            st.error("大学名を入力してください")
            st.stop()

        suffix    = pref_suffix(pref_name)
        base_name = f"{pref_num:02d}_{pref_name}{suffix}_oc2026"

        with st.spinner("生成中..."):
            with tempfile.TemporaryDirectory() as tmp:
                campus_path    = selected_campus_path or find_best_image(univs[0]['name'])
                tmpl_path, _   = resolve_template_path(num_univ, tmp, use_drive_template)
                docx_out       = os.path.join(tmp, f'{base_name}.docx')
                try:
                    generate_chirashi(
                        pref_num, pref_name, univs, campus_path, docx_out,
                        template_path_override=tmpl_path,
                    )
                    docx_bytes = open(docx_out, 'rb').read()
                    st.success("✅ 生成完了！")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.download_button(
                            "📄 Word ダウンロード", data=docx_bytes,
                            file_name=f'{base_name}.docx',
                            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                            use_container_width=True,
                        )
                    with c2:
                        if do_pdf:
                            pdf_front = os.path.join(tmp, f'{base_name}_front.pdf')
                            docx_to_pdf(docx_out, pdf_front)
                            if do_back and os.path.exists(BACK_PDF):
                                pdf_out = os.path.join(tmp, f'{base_name}.pdf')
                                merge_pdfs(pdf_front, BACK_PDF, pdf_out)
                                label = "📄 PDF（表裏合体）"
                            else:
                                pdf_out, label = pdf_front, "📄 PDF（表面のみ）"
                            st.download_button(
                                label, data=open(pdf_out,'rb').read(),
                                file_name=f'{base_name}.pdf', mime='application/pdf',
                                use_container_width=True,
                            )
                except Exception as e:
                    st.error(f"生成エラー: {e}")
                    import traceback; st.code(traceback.format_exc())

    with st.expander("📚 写真ライブラリ（利用可能な大学一覧）"):
        names = sorted(img_dict.keys())
        cols  = st.columns(4)
        for i, name in enumerate(names):
            cols[i % 4].caption(f"• {name}（{len(img_dict[name])}枚）")


# ════════════════════════════════════════════════════════════════
# TAB 2: まとめて作成
# ════════════════════════════════════════════════════════════════
with tab_batch:
    st.subheader("複数の都道府県を一度に生成する")
    st.caption(
        "日程はサイドバーで読み込んだデータから大学名で自動引き当てます。"
        "一括生成では写真はライブラリ（または Drive）から自動選択されます。"
    )

    # ── 都道府県選択 ─────────────────────────────────────────
    pref_options_b = df.apply(lambda r: f"{int(r['番号']):02d}  {r['都道府県']}", axis=1).tolist()
    selected_prefs = st.multiselect("生成する都道府県（複数選択可）", pref_options_b, key="batch_prefs")

    # スプシからインポートした分担データを使って一括選択できるボタン
    if st.session_state.pref_map_imported:
        imported_prefs = sorted(st.session_state.pref_map_imported.keys())
        if st.button(f"📥 読み込み済みスプシの都道府県を全選択（{len(imported_prefs)}件）",
                     key="batch_select_imported"):
            # session_stateに保存してmultiselect側に反映（Streamlitの制約上、
            # multiselect の default は再実行時に反映される）
            st.session_state["batch_prefs"] = [
                s for s in pref_options_b
                if any(p in s for p in imported_prefs)
            ]
            st.rerun()

    # ── 出力設定 ─────────────────────────────────────────────
    col1, col2, col3 = st.columns(3)
    with col1:
        batch_do_pdf  = st.checkbox("PDF版も生成", value=True, key="batch_do_pdf")
    with col2:
        batch_do_back = st.checkbox("裏面と結合",  value=True, key="batch_do_back")
    with col3:
        batch_use_drive_photo = False
        if DRIVE_OK:
            batch_use_drive_photo = st.checkbox("写真にDriveも使う", value=True,
                                                key="batch_use_drive_photo")
    batch_use_drive_template = False
    if dh.templates_from_drive_enabled():
        batch_use_drive_template = st.checkbox("テンプレートをDriveから", key="batch_tmpl_drive")

    # ── プレビューテーブル ────────────────────────────────────
    if selected_prefs:
        preview_rows = []
        for sel in selected_prefs:
            p_num  = int(sel.split()[0])
            p_name = sel.split()[1]
            row    = df[df['番号'] == p_num].iloc[0]

            # 大学リスト: スプシ優先 → マスタ
            if p_name in st.session_state.pref_map_imported:
                names = st.session_state.pref_map_imported[p_name]
            else:
                raw   = [u.strip().strip('（）()') for u in str(row['掲載大学']).replace('、',',').split(',')]
                names = [u for u in raw if u]

            hit_count  = sum(1 for n in names if oc_schedule.get(n))
            img_count  = sum(1 for n in names if find_best_image(n))
            _, tkey    = get_template_path(len(names))
            preview_rows.append({
                '都道府県': f"{p_num:02d} {p_name}",
                '大学数': len(names),
                'テンプレ': f"{tkey}大学",
                '日程あり': f"{hit_count}/{len(names)}",
                '写真あり': f"{img_count}/{len(names)}",
            })

        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

    run_batch = st.button(
        "🚀 まとめて生成する", type="primary", use_container_width=True,
        key="batch_generate", disabled=not selected_prefs,
    )

    if run_batch:
        jobs = []
        with tempfile.TemporaryDirectory() as tmp:
            for sel in selected_prefs:
                p_num  = int(sel.split()[0])
                p_name = sel.split()[1]
                row    = df[df['番号'] == p_num].iloc[0]

                if p_name in st.session_state.pref_map_imported:
                    names = st.session_state.pref_map_imported[p_name]
                else:
                    raw   = [u.strip().strip('（）()') for u in str(row['掲載大学']).replace('、',',').split(',')]
                    names = [u for u in raw if u]

                if not names:
                    continue

                univs         = [{'name': n, 'schedule': oc_schedule.get(n, '')} for n in names]
                suffix        = pref_suffix(p_name)
                base_name     = f"{p_num:02d}_{p_name}{suffix}_oc2026"
                campus_path   = resolve_campus_image(names[0], None, tmp, batch_use_drive_photo)
                tmpl_path, _  = resolve_template_path(len(univs), tmp, batch_use_drive_template)

                jobs.append({
                    'base_name':            base_name,
                    'pref_num':             p_num,
                    'pref_name':            p_name,
                    'universities':         univs,
                    'campus_image_path':    campus_path,
                    'template_path_override': tmpl_path,
                })

            with st.spinner(f"{len(jobs)}件を生成中..."):
                back_pdf_path = BACK_PDF if (batch_do_back and os.path.exists(BACK_PDF)) else None
                zip_bytes, warnings = generate_batch_zip(
                    jobs, batch_do_pdf, batch_do_back, back_pdf_path
                )

        st.success(f"✅ {len(jobs)}件を生成しました！")
        st.download_button(
            "📦 まとめてダウンロード (.zip)", data=zip_bytes,
            file_name="oc_chirashi_batch.zip", mime="application/zip",
            use_container_width=True,
        )
        if warnings:
            with st.expander(f"⚠️ 警告（{len(warnings)}件）", expanded=True):
                for w in warnings:
                    st.write(f"- {w}")


# ════════════════════════════════════════════════════════════════
# TAB 3: Google Drive 連携
# ════════════════════════════════════════════════════════════════
with tab_drive:
    st.subheader("Google Drive 連携について")
    st.write(
        "この機能を設定すると、写真ライブラリやテンプレートをGitHubに含めずに"
        "共有のGoogle Driveフォルダから直接読み込めるようになります。"
        "**設定は任意で、設定していない場合はローカルの `images/` `templates/` フォルダが使われます。**"
    )

    if DRIVE_OK:
        st.success("✅ Google Drive 連携が有効です。")
        images_folder    = st.secrets.get("drive_images_folder_id")
        templates_folder = st.secrets.get("drive_templates_folder_id")
        st.write(f"📁 写真フォルダ ID: `{images_folder}`")
        if templates_folder:
            st.write(f"📁 テンプレートフォルダ ID: `{templates_folder}`")
        if st.button("接続テスト"):
            with st.spinner("Driveに接続中..."):
                try:
                    names = dh.list_drive_image_names(images_folder)
                    st.write(f"写真フォルダから **{len(names)}件** を検出：")
                    st.write("、".join(names[:30]) + ("…" if len(names) > 30 else ""))
                except Exception as e:
                    st.error(f"接続に失敗しました: {e}")
    else:
        st.info("未設定です。以下の手順で設定できます。")
        st.markdown("""
**設定手順（1回だけでOK）**

1. Google Cloud Console で新しいプロジェクトを作成し「Google Drive API」を有効化。
2. 「APIとサービス」→「認証情報」→「サービスアカウントを作成」→鍵タブから **JSONキーをダウンロード**。
3. 写真を入れたGoogle Driveフォルダをサービスアカウントのメールアドレスに **閲覧者として共有**。
4. `.streamlit/secrets.toml`（または Streamlit Cloud の「Settings → Secrets」）に以下を記入：

```toml
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
client_email = "xxxx@xxxx.iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"

drive_images_folder_id    = "フォルダID"
drive_templates_folder_id = "フォルダID"   # テンプレートも共有するなら
```
""")

    st.divider()
    st.caption("東進ハイスクール / 東進衛星予備校 | OC チラシ自動生成ツール v3.0")
