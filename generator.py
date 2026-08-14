"""
OC Chirashi Generator - Core document generation logic
"""
import re, os, zipfile, tempfile, io
from PIL import Image
from lxml import etree

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

BASE_DIR     = os.path.dirname(__file__)
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
IMAGE_DIR    = os.path.join(BASE_DIR, "images")
FONT_NAME = "UD デジタル 教科書体 NP-B"
DEFAULT_FLYER_YEAR = 2026
DEFAULT_VIDEO_COUNT = 400

TEMPLATE_MAP = {6: "template_6.docx", 8: "template_8.docx",
                10: "template_10.docx", 12: "template_12.docx"}

# The large campus photo filename(s) per template (may be duplicated for textbox copy)
CAMPUS_IMAGE_MAP = {
    6:  ["image1.jpeg"],
    8:  ["image3.jpeg", "image7.jpeg"],
    10: ["image1.jpeg"],
    12: ["image3.png",  "image7.png"],
}

# Placeholder text in title paragraphs per template
# Structure: ['PREF', '県の'] or ['PREF', '大学'] as split runs
TITLE_PREF_PLACEHOLDER = {6: "山口", 8: "徳島", 10: "滋賀", 12: "〇〇"}
TITLE_UNIV_PLACEHOLDER = {6: "山口", 8: "徳島", 10: "滋賀", 12: "〇〇"}

# ── Image helpers ─────────────────────────────────────────────
def get_available_images() -> dict:
    images = {}
    if not os.path.exists(IMAGE_DIR):
        return images
    for fname in os.listdir(IMAGE_DIR):
        if not fname.lower().endswith(('.jpg', '.jpeg')):
            continue
        name = fname.split('pixta_')[0].split('_pixta')[0].strip().rstrip('_').strip()
        if not name or name.startswith('pixta'):
            continue
        images.setdefault(name, []).append(os.path.join(IMAGE_DIR, fname))
    return images

def find_best_image(university_name: str) -> str | None:
    if not university_name:
        return None
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
    """Replace the duplicated title textboxes without leaving old suffix text."""
    root = etree.fromstring(xml.encode('utf-8'))
    old_ph = TITLE_PREF_PLACEHOLDER[tmpl_key]
    suffix = pref_suffix(pref_name)
    univ_replaced = 0
    pref_replaced = 0
    for p in root.xpath('.//w:p', namespaces=NS):
        texts = p.xpath('./w:r/w:t/text()', namespaces=NS)
        if not texts:
            continue
        joined = ''.join(texts).strip()
        # University title: placeholder + "大学" in the original template.
        if univ_replaced < 2 and old_ph in texts and '大学' in texts:
            _set_paragraph_text(p, main_univ)
            univ_replaced += 1
            continue
        # Prefecture title: placeholder + "県の/府の/都の/の".
        if pref_replaced < 2 and old_ph in texts and any('の' in t for t in texts):
            runs = p.xpath('./w:r/w:t', namespaces=NS)
            if runs:
                runs[0].text = pref_name + suffix
                runs[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                if len(runs) >= 2:
                    runs[1].text = 'の'
                    runs[1].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                for node in runs[2:]:
                    node.text = ''
            pref_replaced += 1
    return etree.tostring(root, encoding='unicode', xml_declaration=False)

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

def _set_paragraph_text(paragraph, text: str):
    """Replace all visible text in one paragraph while preserving its formatting."""
    text = '' if text is None else str(text)
    runs = paragraph.xpath('./w:r/w:t', namespaces=NS)
    if not runs:
        return
    # Put all text in the first text node and clear the rest. This is important
    # because the supplied templates split strings such as "〇〇大学" into
    # multiple runs ("〇〇" + "大学"). Replacing only the first run leaves
    # stale text such as "関関同立東京大学" behind.
    runs[0].text = text
    for node in runs[1:]:
        node.text = ''
    if text:
        runs[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


def replace_table_universities(xml: str, universities: list[dict]) -> str:
    """Replace the university cells in both duplicated table copies.

    Each supplied template has two identical copies of the university table
    (one is used for the textbox/print layout). A cell's first paragraph is
    the university name and the second paragraph is its OC schedule. The
    original template contains last year's sample text, often split across
    many runs/paragraphs. We replace the *entire* first two paragraphs and
    blank all remaining paragraphs, so no old university or date can leak
    into the new flyer.
    """
    root = etree.fromstring(xml.encode('utf-8'))
    tables = root.xpath('.//w:tbl', namespaces=NS)
    univ_index = 0
    for table in tables[:2]:
        cells = table.xpath('.//w:tr/w:tc', namespaces=NS)
        for cell in cells:
            name = universities[univ_index]['name'] if univ_index < len(universities) else ''
            schedule = universities[univ_index]['schedule'] if univ_index < len(universities) else ''
            univ_index += 1
            paragraphs = cell.xpath('./w:p', namespaces=NS)
            if not paragraphs:
                continue
            _set_paragraph_text(paragraphs[0], name)
            if len(paragraphs) >= 2:
                _set_paragraph_text(paragraphs[1], schedule)
            for p in paragraphs[2:]:
                _set_paragraph_text(p, '')
    return etree.tostring(root, encoding='unicode', xml_declaration=False)

# ── Image replacement ─────────────────────────────────────────
def replace_campus_image(work_dir: str, tmpl_key: int, new_image_path: str):
    media_dir = os.path.join(work_dir, 'word', 'media')
    img = Image.open(new_image_path).convert('RGB')
    for fname in CAMPUS_IMAGE_MAP[tmpl_key]:
        dst = os.path.join(media_dir, fname)
        ext = fname.split('.')[-1].lower()
        fmt = 'JPEG' if ext in ('jpg', 'jpeg') else 'PNG'
        img.save(dst, fmt, quality=90)

def apply_document_customizations(xml: str, flyer_year: int, video_count: int) -> str:
    """Apply yearly values and normalize font names in all visible text runs."""
    root = etree.fromstring(xml.encode('utf-8'))
    year = int(flyer_year)
    count = int(video_count)

    # Text in the templates is sometimes split across multiple <w:t> runs
    # (e.g. "現在、約" + "40" + "0の大学紹介動画..."). Work at paragraph
    # level so the yearly number can never be left half-old/half-new.
    for p in root.xpath('.//w:p', namespaces=NS):
        texts = p.xpath('./w:r/w:t/text()', namespaces=NS)
        if not texts:
            continue
        joined = ''.join(texts)
        if '現在、約' in joined and '大学紹介動画を掲載中！' in joined:
            new_text = re.sub(
                r'現在、約[0-9０-９,，]+の大学紹介動画を掲載中！',
                f'現在、約{count:,}の大学紹介動画を掲載中！',
                joined,
            )
            _set_paragraph_text(p, new_text)
        elif 'オープンキャンパス' in joined and re.search(r'20\d{2}', joined):
            new_text = re.sub(r'オープンキャンパス20\d{2}', f'オープンキャンパス{year}', joined)
            _set_paragraph_text(p, new_text)

    # Normalize every run-level font declaration in the generated front page.
    # This removes the mixed NP-B/NK-B/UD Digi/MS P Gothic definitions that
    # caused different fallback fonts after LibreOffice PDF conversion.
    font_qname = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    for rfonts in root.xpath('.//w:rFonts', namespaces=NS):
        for attr in ('ascii', 'eastAsia', 'hAnsi', 'cs'):
            q = font_qname + attr
            if q in rfonts.attrib:
                rfonts.set(q, FONT_NAME)
    return etree.tostring(root, encoding='unicode', xml_declaration=False)


# ── Main entry point ──────────────────────────────────────────
def generate_chirashi(pref_num, pref_name, universities, campus_image_path, output_path,
                       template_path_override: str | None = None,
                       flyer_year: int = DEFAULT_FLYER_YEAR,
                       video_count: int = DEFAULT_VIDEO_COUNT) -> str:
    num  = len(universities)
    _, tmpl_key = get_template_path(num)
    template_path = template_path_override if template_path_override else _

    # Pad to fill template slots
    univs = list(universities)
    while len(univs) < tmpl_key:
        univs.append({'name': '', 'schedule': ''})

    if not campus_image_path and univs:
        campus_image_path = find_best_image(univs[0]['name'])

    with tempfile.TemporaryDirectory() as work_dir:
        with zipfile.ZipFile(template_path, 'r') as z:
            z.extractall(work_dir)

        doc_path = os.path.join(work_dir, 'word', 'document.xml')
        with open(doc_path, 'r', encoding='utf-8') as f:
            xml = f.read()

        main_univ = univs[0]['name'] if univs[0]['name'] else pref_name + '大学'
        xml = replace_title(xml, tmpl_key, pref_name, main_univ)
        xml = replace_table_universities(xml, univs)
        xml = apply_document_customizations(xml, flyer_year, video_count)

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
                        flyer_year=job.get('flyer_year', DEFAULT_FLYER_YEAR),
                        video_count=job.get('video_count', DEFAULT_VIDEO_COUNT),
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
