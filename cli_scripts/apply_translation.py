import sys
import re
import openpyxl
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

# Matches leading emoji+whitespace at the start of a string
_LEAD_PAT = re.compile(
    r'^([\s\U0001F000-\U0001FFFF\U00002600-\U000027FF←-⇿⌀-⏿■-◿✀-➿]*)'
)
# Matches trailing emoji+whitespace at the end of a string
_TRAIL_PAT = re.compile(
    r'([\s\U0001F000-\U0001FFFF\U00002600-\U000027FF←-⇿⌀-⏿■-◿✀-➿]*)$'
)


def merge_emoji(original, translated):
    """Re-insert leading/trailing emoji from the original into the translated text."""
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
    latin_letters = re.findall(r'[a-zA-Z]', text)
    if len(latin_letters) < 3:
        return False
    return True


def load_translations(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    translations = {}
    english_check = {}

    # Row 1 = instructions, Row 2 = headers, Row 3+ = data
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[0] is None:
            continue
        idx = int(row[0])
        english_orig = str(row[3] or '').strip()   # col D — original with emoji
        telugu = row[5]                             # col F — translation
        english_check[idx] = english_orig
        if telugu and str(telugu).strip():
            translations[idx] = str(telugu).strip()

    return translations, english_check


def collect_replacements(soup, translations, english_check):
    replacements = []
    index = 0
    seen_attr_ids = set()
    warnings = 0

    for node in soup.descendants:
        if isinstance(node, Tag):
            if node.has_attr('data-script'):
                node_id = node.get('id', '')
                if node_id not in seen_attr_ids:
                    text = node['data-script'].strip()
                    if should_extract(text):
                        seen_attr_ids.add(node_id)
                        if index in translations:
                            expected = english_check.get(index, '')
                            if expected and expected != text:
                                print(f'WARNING row {index + 3}: expected "{expected[:60]}" but found "{text[:60]}" — skipping')
                                warnings += 1
                            else:
                                final = merge_emoji(expected, translations[index])
                                replacements.append((node, final, 'ATTR'))
                        index += 1

        elif isinstance(node, NavigableString) and type(node) is NavigableString:
            if any(p.name in SKIP_TAGS for p in node.parents if isinstance(p, Tag)):
                continue
            text = str(node).strip()
            if should_extract(text):
                if index in translations:
                    expected = english_check.get(index, '')
                    if expected and expected != text:
                        print(f'WARNING row {index + 3}: expected "{expected[:60]}" but found "{text[:60]}" — skipping')
                        warnings += 1
                    else:
                        final = merge_emoji(expected, translations[index])
                        replacements.append((node, final, 'TEXT'))
                index += 1

    return replacements, warnings


def apply_replacements(replacements):
    for node, new_text, node_type in replacements:
        if node_type == 'ATTR':
            node['data-script'] = new_text
        else:
            node.replace_with(NavigableString(new_text))


def main():
    if len(sys.argv) != 4:
        print('Usage: python apply_translation.py "<input.html>" <translations.xlsx> "<output.html>"')
        sys.exit(1)

    html_path = sys.argv[1]
    xlsx_path = sys.argv[2]
    output_path = sys.argv[3]

    translations, english_check = load_translations(xlsx_path)
    if not translations:
        print('No translations found in Excel file. Nothing to apply.')
        sys.exit(0)

    with open(html_path, encoding='utf-8') as f:
        html = f.read()

    soup = BeautifulSoup(html, 'html.parser')
    replacements, warnings = collect_replacements(soup, translations, english_check)
    apply_replacements(replacements)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(str(soup))

    blank = len(english_check) - len(translations)
    print(f'Applied {len(replacements)}/{len(english_check)} translations ({blank} rows left blank — originals kept)')
    if warnings:
        print(f'{warnings} drift warnings — those rows were skipped')
    print(f'Output written to {output_path}')


if __name__ == '__main__':
    main()
