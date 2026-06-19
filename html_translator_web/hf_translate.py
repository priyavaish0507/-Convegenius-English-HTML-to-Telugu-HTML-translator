"""
Translation via Google Translate (deep_translator library).
Free, no API key required, reliable.
"""
import time
from deep_translator import GoogleTranslator

BATCH_SIZE = 20   # deep_translator handles batches efficiently


def translate_batch(texts: list[str], progress_cb=None) -> list[str]:
    """
    Translate all texts English → Telugu.
    Calls progress_cb(done, total) after each batch.
    """
    results = []
    total = len(texts)
    translator = GoogleTranslator(source='en', target='te')

    for i in range(0, total, BATCH_SIZE):
        batch = texts[i:i + BATCH_SIZE]
        try:
            translated = translator.translate_batch(batch)
            # translate_batch returns None for empty strings — fall back to original
            results.extend(t if t else batch[j] for j, t in enumerate(translated))
        except Exception:
            # On error, keep the original English for this batch
            results.extend(batch)

        if progress_cb:
            progress_cb(min(i + BATCH_SIZE, total), total)
        time.sleep(0.1)

    return results


def translate_text(text: str) -> str:
    return GoogleTranslator(source='en', target='te').translate(text) or text
