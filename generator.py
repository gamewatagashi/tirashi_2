"""
OC Chirashi Generator - Core document generation logic
"""
import re, os, zipfile, tempfile, io
from PIL import Image
from univ_utils import canonical_university_name, cover_university

BASE_DIR     = os.path.dirname(__file__)
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
IMAGE_DIR    = os.path.join(BASE_DIR, "images")

TEMPLATE_MAP = {6: "template_6.docx", 8: "template_8.docx",
                10: "template_10.docx", 12: "template_12.docx"}

# The large campus photo filename(s) per template (may be duplicated for textbox copy)
CAMPUS_IMAGE_MAP = {
    6:  ["image1.png"],
    8:  ["image3.jpeg", "image7.jpeg"],
    10: ["image1.png"],
    12: ["image3.png",  "image7.png"],
}

# Placeholder text in title paragraphs per template
# Structure: ['PREF', '県の'] or ['PREF', '大学'] as split runs
TITLE_PREF_PLACEHOLDER = {6: "〇〇", 8: "徳島", 10: "○○", 12: "〇〇"}
TITLE_UNIV_PLACEHOLDER = {6: "〇〇", 8: "徳島", 10: "○○", 12: "〇〇"}

# ── Image helpers ─────────────────────────────────────────────
def get_available_images() -> dict:
    images = {}
    if not os.path.exists(IMAGE_DIR):
        return images
    for fname in os.listdir(IMAGE_DIR):
        if not fname.lower().endswith(('.jpg', '.jpeg')):
            continue
        name = canonical_university_name(fname.split('pixta_')[0].split('_pixta')[0].strip().rstrip('_').strip())
        if not name or name.startswith('pixta'):
            continue
        images.setdefault(name, []).append(os.path.join(IMAGE_DIR, fname))
    return images

def find_best_image(university_name: str) -> str | None:
    if not university_name:
        return None
    university_name = canonical_university_name(university_name)
    images = get_available_images()
    def pick(paths):
        for p in paths:
            if '_M.' in p or ' M.' in p:
                return p
        return paths[0]
    if university_name in images:
        return pick(images[university_name])
    for key, paths in images.items():
        if university_name in key or key in university_name:
            return pick(paths)
    return None

def find_all_images(university_name: str) -> list[str]:
    """Return every local photo matching a university name, for a picker UI
    (as opposed to find_best_image, which returns a single best guess)."""
    if not university_name:
        return []
    images = get_available_images()
    if university_name in images:
        return images[university_name]
    matches: list[str] = []
    for key, paths in images.items():
        if university_name in key or key in university_name:
            matches.extend(paths)
    return matches

def get_template_path(num_universities: int):
    key = 6 if num_universities <= 6 else 8 if num_universities <= 8 else 10 if num_universities <= 10 else 12
    return os.path.join(TEMPLATE_DIR, TEMPLATE_MAP[key]), key

# ── Title replacement ─────────────────────────────────────────
def pref_suffix(pref_name: str) -> str:
    if pref_name in ['大阪', '京都']:   return '府'
    if pref_name == '東京':             return '都'
    if pref_name == '北海道':           return ''
    return '県'

def _xml_escape(text: str) -> str:
    """Escape characters that are not valid as-is inside a <w:t> text node.
    Without this, any '&', '<', or '>' in a university name or OC schedule
    string (e.g. 'AO&推薦', 'A<B学部') produces malformed XML and Word refuses
    to open the resulting .docx."""
    if text is None:
        return ''
    return (str(text)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;'))

def replace_title(xml: str, tmpl_key: int, pref_name: str, main_univ: str) -> str:
    """Replace title text only in the document header area (before the first table).
    This avoids confusing the title with university-table cells and works across
    the 6/8/10/12 templates, whose placeholder spellings differ.
    """
    table_pos = xml.find('<w:tbl')
    if table_pos < 0:
        table_pos = len(xml)
    head, tail = xml[:table_pos], xml[table_pos:]
    p_pattern = re.compile(r'(<w:p\b[^>]*>)(.*?)(</w:p>)', re.DOTALL)
    wt_pattern = re.compile(r'<w:t\b[^>]*>.*?</w:t>', re.DOTALL)
    pref_done=0
    univ_done=0

    def process(m):
        nonlocal pref_done, univ_done
        body=m.group(2)
        vals=re.findall(r'<w:t\b[^>]*>(.*?)</w:t>',body,re.DOTALL)
        vals=[re.sub(r'<[^>]+>','',v) for v in vals]
        full=''.join(vals).strip()
        if not full:
            return m.group(0)

        # Prefecture title: e.g. 「神奈川県の」「〇〇県の」
        if full.endswith(('県の','府の','都の')) and pref_done < 2:
            matches=list(wt_pattern.finditer(body))
            if matches:
                first=matches[0]
                body=body[:first.start()]+f'<w:t xml:space="preserve">{_xml_escape(pref_name)}</w:t>'+body[first.end():]
                matches=list(wt_pattern.finditer(body))
                if len(matches)>1:
                    second=matches[1]
                    body=body[:second.start()]+f'<w:t xml:space="preserve">{_xml_escape(pref_suffix(pref_name)+"の")}</w:t>'+body[second.end():]
                    matches=list(wt_pattern.finditer(body))
                    for mm in reversed(matches[2:]):
                        body=body[:mm.start()]+'<w:t></w:t>'+body[mm.end():]
                pref_done += 1
                return m.group(1)+body+m.group(3)

        # Featured university title: short paragraph ending in 「大学」
        # and located before the first table.
        if full.endswith('大学') and len(full) <= 12 and univ_done < 2:
            matches=list(wt_pattern.finditer(body))
            if matches:
                first=matches[0]
                body=body[:first.start()]+f'<w:t xml:space="preserve">{_xml_escape(main_univ)}</w:t>'+body[first.end():]
                matches=list(wt_pattern.finditer(body))
                for mm in reversed(matches[1:]):
                    body=body[:mm.start()]+'<w:t></w:t>'+body[mm.end():]
                univ_done += 1
                return m.group(1)+body+m.group(3)
        return m.group(0)

    head=p_pattern.sub(process,head)
    return head+tail

# ── Table replacement ─────────────────────────────────────────
def replace_cell_content(tc_xml: str, univ_name: str, schedule: str) -> str:
    p_pattern  = re.compile(r'<w:p\b[^>]*>.*?</w:p>', re.DOTALL)
    wt_pattern = re.compile(r'<w:t[^>]*>[^<]*</w:t>')

    paragraphs = list(p_pattern.finditer(tc_xml))
    if not paragraphs:
        return tc_xml

    new_tc = tc_xml
    for pidx, text in enumerate([univ_name, schedule]):
        paragraphs = list(p_pattern.finditer(new_tc))
        if pidx >= len(paragraphs):
            break
        p = paragraphs[pidx].group(0)
        wts = list(wt_pattern.finditer(p))
        if not wts:
            continue
        new_p = p[:wts[0].start()] + f'<w:t xml:space="preserve">{_xml_escape(text)}</w:t>'
        rest   = wt_pattern.sub('<w:t></w:t>', p[wts[0].end():])
        new_p += rest
        new_tc = new_tc[:paragraphs[pidx].start()] + new_p + new_tc[paragraphs[pidx].end():]

    # Empty remaining paragraphs (3rd onward)
    paragraphs = list(p_pattern.finditer(new_tc))
    for p_match in paragraphs[2:]:
        new_p = wt_pattern.sub('<w:t></w:t>', p_match.group(0))
        new_tc = new_tc[:p_match.start()] + new_p + new_tc[p_match.end():]
        paragraphs = list(p_pattern.finditer(new_tc))

    return new_tc

def replace_table_universities(xml: str, universities: list[dict]) -> str:
    """Replace each university cell's two paragraphs, clearing all old text nodes."""
    tc_pattern = re.compile(r'<w:tc\b[^>]*>.*?</w:tc>', re.DOTALL)
    table_pattern = re.compile(r'<w:tbl\b[^>]*>.*?</w:tbl>', re.DOTALL)
    p_pattern = re.compile(r'<w:p\b[^>]*>.*?</w:p>', re.DOTALL)
    wt_pattern = re.compile(r'<w:t\b[^>]*>.*?</w:t>', re.DOTALL)
    table_count = 0
    univ_index = 0

    def replace_one_table(tm):
        nonlocal table_count, univ_index
        table_count += 1
        if table_count > 2:
            return tm.group(0)
        table_xml = tm.group(0)

        def replace_one_cell(cm):
            nonlocal univ_index
            cell = cm.group(0)
            name = universities[univ_index]['name'] if univ_index < len(universities) else ''
            schedule = universities[univ_index]['schedule'] if univ_index < len(universities) else ''
            univ_index += 1
            paragraphs=list(p_pattern.finditer(cell))
            for pidx, text in enumerate([name, schedule]):
                paragraphs=list(p_pattern.finditer(cell))
                if pidx >= len(paragraphs):
                    break
                ps,pe=paragraphs[pidx].span()
                para=paragraphs[pidx].group(0)
                wts=list(wt_pattern.finditer(para))
                if not wts:
                    continue
                # Put all replacement text into the first run and clear the rest.
                first=wts[0]
                new_t=f'<w:t xml:space="preserve">{_xml_escape(text)}</w:t>'
                para=para[:first.start()]+new_t+para[first.end():]
                wts2=list(wt_pattern.finditer(para))
                for mm in reversed(wts2[1:]):
                    para=para[:mm.start()]+'<w:t></w:t>'+para[mm.end():]
                cell=cell[:ps]+para+cell[pe:]
            # Clear unused paragraphs
            paragraphs=list(p_pattern.finditer(cell))
            for p_match in reversed(paragraphs[2:]):
                para=wt_pattern.sub('<w:t></w:t>', p_match.group(0))
                ps,pe=p_match.span()
                cell=cell[:ps]+para+cell[pe:]
            return cell
        return tc_pattern.sub(replace_one_cell, table_xml)
    return table_pattern.sub(replace_one_table, xml)

# ── Image replacement ─────────────────────────────────────────
def replace_campus_image(work_dir: str, tmpl_key: int, new_image_path: str):
    media_dir = os.path.join(work_dir, 'word', 'media')
    img = Image.open(new_image_path).convert('RGB')
    for fname in CAMPUS_IMAGE_MAP[tmpl_key]:
        dst = os.path.join(media_dir, fname)
        ext = fname.split('.')[-1].lower()
        fmt = 'JPEG' if ext in ('jpg', 'jpeg') else 'PNG'
        img.save(dst, fmt, quality=90)

# ── Annual settings / font normalization ─────────────────────
def _replace_paragraph_text(xml: str, predicate, new_text: str) -> str:
    p_pattern = re.compile(r'(<w:p\b[^>]*>)(.*?)(</w:p>)', re.DOTALL)
    wt_pattern = re.compile(r'<w:t\b[^>]*>.*?</w:t>', re.DOTALL)
    def repl(m):
        body=m.group(2)
        texts=[re.sub(r'<[^>]+>','',x) for x in re.findall(r'<w:t\b[^>]*>(.*?)</w:t>', body, re.DOTALL)]
        full=''.join(texts)
        if not predicate(full):
            return m.group(0)
        wts=list(wt_pattern.finditer(body))
        if not wts:
            return m.group(0)
        first=wts[0]
        new_t=f'<w:t xml:space="preserve">{_xml_escape(new_text)}</w:t>'
        body=body[:first.start()]+new_t+body[first.end():]
        # Clear all subsequent text nodes, but preserve their runs/styles.
        matches=list(wt_pattern.finditer(body))
        if len(matches)>1:
            for mm in reversed(matches[1:]):
                body=body[:mm.start()]+'<w:t></w:t>'+body[mm.end():]
        return m.group(1)+body+m.group(3)
    return p_pattern.sub(repl, xml)

def apply_annual_settings(xml: str, year: int, video_count: int) -> str:
    # Only touch the two dedicated paragraphs, leaving the rest of the template intact.
    xml = _replace_paragraph_text(
        xml, lambda s: 'オープンキャンパス' in s and ('2026' in s or '2025' in s or '2027' in s),
        f'オープンキャンパス {year}'
    )
    xml = _replace_paragraph_text(
        xml, lambda s: '現在、約' in s and '大学紹介動画' in s,
        f'現在、約 {video_count} の大学紹介動画を掲載中！'
    )
    return xml

def force_document_font(xml: str, font_name: str = 'UD デジタル 教科書体 NP-B') -> str:
    def repl(m):
        tag=m.group(0)
        tag=re.sub(r'\s+(?:w:ascii|w:hAnsi|w:eastAsia|w:cs)="[^"]*"', '', tag)
        tag=tag[:-2] if tag.endswith('/>') else tag[:-1]
        return tag + f' w:ascii="{font_name}" w:eastAsia="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>'
    return re.sub(r'<w:rFonts\b[^>]*/>', repl, xml)

# ── Main entry point ──────────────────────────────────────────
def generate_chirashi(pref_num, pref_name, universities, campus_image_path, output_path,
                       template_path_override: str | None = None, year: int = 2027, video_count: int = 500,
                       cover_univ: str | None = None) -> str:
    num  = len(universities)
    _, tmpl_key = get_template_path(num)
    template_path = template_path_override if template_path_override else _

    # Normalize university names so abbreviations in the master sheet match
    # full-name photo files and yearly schedule data.
    univs = [
        {'name': canonical_university_name(u.get('name','')), 'schedule': u.get('schedule','')}
        for u in universities
    ]
    # Pad to fill template slots
    while len(univs) < tmpl_key:
        univs.append({'name': '', 'schedule': ''})

    if not campus_image_path and univs:
        target_cover = cover_univ or cover_university(pref_name) or univs[0]['name']
        campus_image_path = find_best_image(target_cover)

    with tempfile.TemporaryDirectory() as work_dir:
        with zipfile.ZipFile(template_path, 'r') as z:
            z.extractall(work_dir)

        doc_path = os.path.join(work_dir, 'word', 'document.xml')
        with open(doc_path, 'r', encoding='utf-8') as f:
            xml = f.read()

        main_univ = canonical_university_name(cover_univ or cover_university(pref_name) or univs[0]['name']) if univs and (cover_univ or cover_university(pref_name) or univs[0]['name']) else pref_name + '大学'
        xml = replace_title(xml, tmpl_key, pref_name, main_univ)
        xml = replace_table_universities(xml, univs)
        xml = apply_annual_settings(xml, year, video_count)
        xml = force_document_font(xml)

        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(xml)

        if campus_image_path and os.path.exists(campus_image_path):
            replace_campus_image(work_dir, tmpl_key, campus_image_path)

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(work_dir):
                for file in files:
                    fp = os.path.join(root, file)
                    zout.write(fp, os.path.relpath(fp, work_dir))

    return output_path

# ── PDF helpers ───────────────────────────────────────────────
def docx_to_pdf(docx_path: str, pdf_path: str) -> str:
    import subprocess, uuid

    out_dir = os.path.dirname(os.path.abspath(pdf_path))
    os.makedirs(out_dir, exist_ok=True)

    # Each call gets its own LibreOffice user profile. Reusing the default
    # profile across back-to-back headless conversions (as happens in batch
    # generation) can leave a stale lock file behind and cause the next
    # conversion to silently fail to produce an output file.
    profile_dir = os.path.join(tempfile.gettempdir(), f'lo_profile_{uuid.uuid4().hex}')

    cmd = [
        'soffice', '--headless', '--norestore',
        f'-env:UserInstallation=file://{profile_dir}',
        '--convert-to', 'pdf', '--outdir', out_dir, docx_path,
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"LibreOfficeの変換がタイムアウトしました（{docx_path}）") from e
    finally:
        import shutil
        shutil.rmtree(profile_dir, ignore_errors=True)

    base     = os.path.splitext(os.path.basename(docx_path))[0]
    expected = os.path.join(out_dir, base + '.pdf')

    if not os.path.exists(expected):
        stderr = (result.stderr or '').strip()[-800:]
        raise RuntimeError(
            f"PDF変換に失敗しました（{os.path.basename(docx_path)}）。"
            f"LibreOfficeがインストールされているか確認してください。詳細: {stderr}"
        )

    if os.path.abspath(expected) != os.path.abspath(pdf_path):
        os.replace(expected, pdf_path)
    return pdf_path

def merge_pdfs(front_pdf: str, back_pdf: str, output_pdf: str) -> str:
    from pypdf import PdfWriter, PdfReader
    for path in [front_pdf, back_pdf]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"PDFファイルが見つかりません: {path}")
    writer = PdfWriter()
    for path in [front_pdf, back_pdf]:
        for page in PdfReader(path).pages:
            writer.add_page(page)
    with open(output_pdf, 'wb') as f:
        writer.write(f)
    return output_pdf

# ── Batch generation ────────────────────────────────────────
def generate_batch_zip(jobs: list[dict], do_pdf: bool, do_back: bool,
                        back_pdf_path: str | None) -> tuple[bytes, list[str]]:
    """
    jobs: list of dicts, each with:
        base_name, pref_num, pref_name, universities, campus_image_path,
        template_path_override (optional)
    Returns (zip_bytes, warnings)
    """
    warnings = []
    buf = io.BytesIO()
    with tempfile.TemporaryDirectory() as tmp:
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for job in jobs:
                base_name = job['base_name']
                docx_out  = os.path.join(tmp, f'{base_name}.docx')
                try:
                    generate_chirashi(
                        job['pref_num'], job['pref_name'], job['universities'],
                        job.get('campus_image_path'), docx_out,
                        template_path_override=job.get('template_path_override'),
                        year=job.get('year', 2027),
                        video_count=job.get('video_count', 500),
                        cover_univ=job.get('cover_univ'),
                    )
                except Exception as e:
                    warnings.append(f"{base_name}: 生成エラー ({e})")
                    continue

                zf.write(docx_out, f'{base_name}.docx')

                if not job.get('campus_image_path'):
                    warnings.append(f"{base_name}: 写真が見つからず既定の画像のままです")

                if do_pdf:
                    try:
                        pdf_front = os.path.join(tmp, f'{base_name}_front.pdf')
                        docx_to_pdf(docx_out, pdf_front)
                        if do_back and back_pdf_path and os.path.exists(back_pdf_path):
                            pdf_final = os.path.join(tmp, f'{base_name}.pdf')
                            merge_pdfs(pdf_front, back_pdf_path, pdf_final)
                        else:
                            pdf_final = pdf_front
                        zf.write(pdf_final, f'{base_name}.pdf')
                    except Exception as e:
                        warnings.append(f"{base_name}: PDF変換エラー ({e})。Wordファイルのみ同梱しました")

    buf.seek(0)
    return buf.getvalue(), warnings
