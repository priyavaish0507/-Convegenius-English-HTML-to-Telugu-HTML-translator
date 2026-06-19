import sys
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from bs4 import BeautifulSoup, NavigableString, Tag

SKIP_TAGS = {'script', 'style'}
PURE_NUMBER_RE = re.compile(r'^\s*-?\d+(\.\d+)?\s*$')

# Matches emoji and special symbols (arrows, checkmarks, speakers, etc.)
EMOJI_RE = re.compile(
    r'[\U0001F000-\U0001FFFF'
    r'\U00002600-\U000027FF'
    r'←-⇿'
    r'⌀-⏿'
    r'■-◿'
    r'✀-➿'
    r']+'
)


def strip_emoji(text):
    return EMOJI_RE.sub('', text).strip()


def should_extract(text):
    text = text.strip()
    if not text:
        return False
    if PURE_NUMBER_RE.match(text):
        return False
    latin_letters = re.findall(r'[a-zA-Z]', text)
    if len(latin_letters) < 3:
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


def write_excel(records, output_path):
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

    # --- Row 1: Instructions ---
    ws.row_dimensions[1].height = 36
    inst = Alignment(wrap_text=True, vertical='center')
    grey_font = Font(italic=True, color='555555', size=10)

    ws['E1'] = 'English text to translate (emoji removed — they will be re-added automatically).'
    ws['F1'] = '✏ Type the Telugu translation here. Write only the words — no need to add emoji.'
    for col in ['E', 'F']:
        ws[f'{col}1'].font = grey_font
        ws[f'{col}1'].alignment = inst

    # --- Row 2: Headers ---
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

    # --- Data rows (row 3 onwards) ---
    wrap_top = Alignment(wrap_text=True, vertical='top')

    for idx, typ, ref, text in records:
        row = idx + 3   # row 1=instructions, row 2=headers, row 3=first data
        text_only = strip_emoji(text)

        # Hidden metadata: A=Index, B=Type, C=Ref, D=Original English with emoji
        ws.cell(row=row, column=1, value=idx)
        ws.cell(row=row, column=2, value=typ)
        ws.cell(row=row, column=3, value=ref)
        ws.cell(row=row, column=4, value=text)   # kept hidden; used by apply script

        # E: Text to translate (emoji stripped) — light blue
        e = ws.cell(row=row, column=5, value=text_only)
        e.fill = PatternFill(start_color='EBF5FB', end_color='EBF5FB', fill_type='solid')
        e.font = Font(color='1A252F', size=10)
        e.alignment = wrap_top
        e.border = border

        # F: Telugu Translation — white, green left border
        f = ws.cell(row=row, column=6, value='')
        f.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        f.font = Font(size=11)
        f.alignment = wrap_top
        f.border = green_border

        # G: Row number
        g = ws.cell(row=row, column=7, value=idx + 1)
        g.font = Font(color='AAAAAA', size=9)
        g.alignment = Alignment(horizontal='center', vertical='top')

        # Row height based on text length
        ws.row_dimensions[row].height = max(18, min(80, 15 + (len(text) // 50) * 14))

    # Hide metadata columns A, B, C, D
    for col in ['A', 'B', 'C', 'D']:
        ws.column_dimensions[col].hidden = True

    ws.column_dimensions['E'].width = 50
    ws.column_dimensions['F'].width = 50
    ws.column_dimensions['G'].width = 7

    ws.freeze_panes = 'E3'

    wb.save(output_path)


def main():
    if len(sys.argv) != 3:
        print('Usage: python extract_text.py "<input.html>" <output.xlsx>')
        sys.exit(1)

    html_path = sys.argv[1]
    xlsx_path = sys.argv[2]

    with open(html_path, encoding='utf-8') as f:
        html = f.read()

    soup = BeautifulSoup(html, 'html.parser')
    records = list(walk_dom(soup))
    write_excel(records, xlsx_path)
    print(f'Extracted {len(records)} text items to {xlsx_path}')


if __name__ == '__main__':
    main()
