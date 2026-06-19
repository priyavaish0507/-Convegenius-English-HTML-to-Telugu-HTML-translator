"""
Core extract/apply logic — ported from the CLI scripts.
All functions operate on strings/BytesIO, no file system access.
"""
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from bs4 import BeautifulSoup, NavigableString, Tag

SKIP_TAGS = {'script', 'style'}
PURE_NUMBER_RE = re.compile(r'^\s*-?\d+(\.\d+)?\s*$')

EMOJI_RE = re.compile(
    r'[\U0001F000-\U0001FFFF'
    r'\U00002600-\U000027FF'
    r'←-⇿'
    r'⌀-⏿'
    r'■-◿'
    r'✀-➿'
    r']+'
)
_LEAD_PAT = re.compile(
    r'^([\s\U0001F000-\U0001FFFF\U00002600-\U000027FF←-⇿⌀-⏿■-◿✀-➿]*)'
)
_TRAIL_PAT = re.compile(
    r'([\s\U0001F000-\U0001FFFF\U00002600-\U000027FF←-⇿⌀-⏿■-◿✀-➿]*)$'
)


def strip_emoji(text):
    return EMOJI_RE.sub('', text).strip()


def merge_emoji(original, translated):
    lead_m = _LEAD_PAT.match(original)
    trail_m = _TRAIL_PAT.search(original)
    leading = lead_m.group().strip() if lead_m and EMOJI_RE.search(lead_m.group()) else ''
    trailing = trail_m.group().strip() if trail_m and EMOJI_RE.search(trail_m.group()) else ''
    result = translated.strip()
    if leading:
        result = leading + ' ' + result
    if trailing and trailing != leading:
        result = result + ' ' + trailing
    return result


def should_extract(text):
    text = text.strip()
    if not text:
        return False
    if PURE_NUMBER_RE.match(text):
        return False
    if len(re.findall(r'[a-zA-Z]', text)) < 3:
        return False
    return True


def walk_dom(soup):
    index = 0
    seen_attr_ids = set()
    for node in soup.descendants:
        if isinstance(node, Tag):
            if node.has_attr('data-script'):
                node_id = node.get('id', '')
                if node_id not in seen_attr_ids:
                    text = node['data-script'].strip()
                    if should_extract(text):
                        seen_attr_ids.add(node_id)
                        yield (index, 'ATTR', node_id, text)
                        index += 1
        elif isinstance(node, NavigableString) and type(node) is NavigableString:
            if any(p.name in SKIP_TAGS for p in node.parents if isinstance(p, Tag)):
                continue
            text = str(node).strip()
            if should_extract(text):
                yield (index, 'TEXT', '', text)
                index += 1


def extract_to_excel(html_content: str) -> bytes:
    """Parse HTML, extract English text, return xlsx bytes."""
    soup = BeautifulSoup(html_content, 'html.parser')
    records = list(walk_dom(soup))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Translate Here'
    ws.sheet_properties.tabColor = '1A5276'

    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    green_border = Border(
        left=Side(style='medium', color='27AE60'),
        right=thin, top=thin, bottom=thin
    )

    ws.row_dimensions[1].height = 36
    grey_font = Font(italic=True, color='555555', size=10)
    inst = Alignment(wrap_text=True, vertical='center')
    ws['E1'] = 'English text to translate (emoji removed — they will be re-added automatically).'
    ws['F1'] = '✏ Type the Telugu translation here. Write only the words — no need to add emoji.'
    for col in ['E', 'F']:
        ws[f'{col}1'].font = grey_font
        ws[f'{col}1'].alignment = inst

    ws.row_dimensions[2].height = 24
    header_fill = PatternFill(start_color='1A5276', end_color='1A5276', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    center = Alignment(horizontal='center', vertical='center')
    for col, label in [('E', 'Text to Translate'), ('F', 'Telugu Translation'), ('G', 'Row #')]:
        c = ws[f'{col}2']
        c.value = label
        c.font = header_font
        c.fill = header_fill
        c.alignment = center

    wrap_top = Alignment(wrap_text=True, vertical='top')
    for idx, typ, ref, text in records:
        row = idx + 3
        text_only = strip_emoji(text)
        ws.cell(row=row, column=1, value=idx)
        ws.cell(row=row, column=2, value=typ)
        ws.cell(row=row, column=3, value=ref)
        ws.cell(row=row, column=4, value=text)

        e = ws.cell(row=row, column=5, value=text_only)
        e.fill = PatternFill(start_color='EBF5FB', end_color='EBF5FB', fill_type='solid')
        e.font = Font(color='1A252F', size=10)
        e.alignment = wrap_top
        e.border = border

        f = ws.cell(row=row, column=6, value='')
        f.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        f.font = Font(size=11)
        f.alignment = wrap_top
        f.border = green_border

        g = ws.cell(row=row, column=7, value=idx + 1)
        g.font = Font(color='AAAAAA', size=9)
        g.alignment = Alignment(horizontal='center', vertical='top')

        ws.row_dimensions[row].height = max(18, min(80, 15 + (len(text) // 50) * 14))

    for col in ['A', 'B', 'C', 'D']:
        ws.column_dimensions[col].hidden = True
    ws.column_dimensions['E'].width = 50
    ws.column_dimensions['F'].width = 50
    ws.column_dimensions['G'].width = 7
    ws.freeze_panes = 'E3'

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def load_translations_from_bytes(xlsx_bytes: bytes):
    """Read filled Excel, return (translations dict, english_check dict)."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active
    translations = {}
    english_check = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[0] is None:
            continue
        idx = int(row[0])
        english_orig = str(row[3] or '').strip()
        telugu = row[5]
        english_check[idx] = english_orig
        if telugu and str(telugu).strip():
            translations[idx] = str(telugu).strip()
    return translations, english_check


def apply_translations(html_content: str, translations: dict, english_check: dict) -> tuple[str, int, int]:
    """Apply translations to HTML, return (new_html, applied_count, total_count)."""
    soup = BeautifulSoup(html_content, 'html.parser')

    # First pass: collect replacements
    replacements = []
    index = 0
    seen_attr_ids = set()

    for node in soup.descendants:
        if isinstance(node, Tag):
            if node.has_attr('data-script'):
                node_id = node.get('id', '')
                if node_id not in seen_attr_ids:
                    text = node['data-script'].strip()
                    if should_extract(text):
                        seen_attr_ids.add(node_id)
                        if index in translations:
                            final = merge_emoji(english_check.get(index, text), translations[index])
                            replacements.append((node, final, 'ATTR'))
                        index += 1
        elif isinstance(node, NavigableString) and type(node) is NavigableString:
            if any(p.name in SKIP_TAGS for p in node.parents if isinstance(p, Tag)):
                continue
            text = str(node).strip()
            if should_extract(text):
                if index in translations:
                    final = merge_emoji(english_check.get(index, text), translations[index])
                    replacements.append((node, final, 'TEXT'))
                index += 1

    total = index

    # Second pass: apply
    for node, new_text, node_type in replacements:
        if node_type == 'ATTR':
            node['data-script'] = new_text
        else:
            node.replace_with(NavigableString(new_text))

    return str(soup), len(replacements), total
