import sys
import time
import requests
import openpyxl

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

API_URL = 'https://api.sarvam.ai/translate'
MAX_CHARS = 1000   # mayura:v1 limit


def translate(text, api_key):
    """Translate a single string from English to Telugu via Sarvam AI."""
    payload = {
        'input': text,
        'source_language_code': 'en-IN',
        'target_language_code': 'te-IN',
        'model': 'mayura:v1',
        'mode': 'formal',
    }
    headers = {
        'api-subscription-key': api_key,
        'Content-Type': 'application/json',
    }
    resp = requests.post(API_URL, json=payload, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json()['translated_text']


def translate_long(text, api_key):
    """Split text that exceeds MAX_CHARS at sentence boundaries and rejoin."""
    if len(text) <= MAX_CHARS:
        return translate(text, api_key)

    # Split on '. ' or '\n' to keep semantic chunks
    sentences = []
    for part in text.replace('\n', '. ').split('. '):
        part = part.strip()
        if part:
            sentences.append(part)

    chunks, current = [], ''
    for s in sentences:
        if len(current) + len(s) + 2 <= MAX_CHARS:
            current = (current + '. ' + s).strip('. ')
        else:
            if current:
                chunks.append(current)
            current = s
    if current:
        chunks.append(current)

    parts = []
    for chunk in chunks:
        parts.append(translate(chunk, api_key))
        time.sleep(0.3)
    return '. '.join(parts)


def main():
    if len(sys.argv) != 3:
        print('Usage: python auto_translate.py translations.xlsx <api_key>')
        sys.exit(1)

    xlsx_path = sys.argv[1]
    api_key = sys.argv[2]

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    # Count rows that need translation
    rows_to_do = [
        r for r in ws.iter_rows(min_row=3)
        if r[4].value and str(r[4].value).strip()   # col E has text
        and not (r[5].value and str(r[5].value).strip())  # col F is blank
    ]
    total = len(rows_to_do)
    print(f'Translating {total} rows...\n')

    done, failed = 0, 0
    for row_cells in rows_to_do:
        idx = row_cells[0].value
        text = str(row_cells[4].value).strip()   # col E — text to translate
        row_num = row_cells[0].row

        try:
            telugu = translate_long(text, api_key)
            row_cells[5].value = telugu              # col F — Telugu translation
            done += 1
            print(f'  [{done}/{total}] Row {idx + 1}: {text[:45]!r}')
            print(f'            → {telugu[:55]}')
            # Save every 10 rows so progress isn't lost on interruption
            if done % 10 == 0:
                wb.save(xlsx_path)
            time.sleep(0.2)   # gentle rate limiting

        except requests.HTTPError as e:
            print(f'  [!] Row {idx + 1} failed ({e.response.status_code}): {text[:40]!r}')
            failed += 1
            if e.response.status_code == 429:
                print('  Rate limited — waiting 5s...')
                time.sleep(5)
        except Exception as e:
            print(f'  [!] Row {idx + 1} error: {e}')
            failed += 1

    wb.save(xlsx_path)
    print(f'\nDone. {done} translated, {failed} failed. Saved to {xlsx_path}')


if __name__ == '__main__':
    main()
