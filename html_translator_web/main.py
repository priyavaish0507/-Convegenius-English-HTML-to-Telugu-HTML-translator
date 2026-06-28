import os
import json
import uuid
import base64
import asyncio
from pathlib import Path
from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from translator import (extract_to_excel, load_translations_from_bytes,
                        apply_translations, walk_dom, strip_emoji,
                        extract_narr_records, apply_narr_translations,
                        update_voice_language, inject_audio_clips,
                        fix_play_welcome)
from hf_translate import translate_batch
from bs4 import BeautifulSoup

app = FastAPI(title='HTML Translator')

app.add_middleware(
    CORSMiddleware,
    allow_origins=['*'],
    allow_methods=['*'],
    allow_headers=['*'],
)

# In-memory store: job_id → (filename, html_bytes)
# Render free tier is single-worker so this is safe
_results: dict[str, tuple[str, bytes]] = {}

_TTS_VOICE_ID = 'EMxdghWQV7gqV33j4J3F'
_TTS_MODEL_ID = 'eleven_v3'
_TTS_FORMAT   = 'mp3_44100_128'

# Common Unicode punctuation → ASCII equivalents, then drop anything still outside latin-1.
# Needed because HTTP Content-Disposition filenames must be latin-1 encodable.
_UNICODE_TO_ASCII = str.maketrans({
    '–': '-',   # en-dash  –
    '—': '-',   # em-dash  —
    '‘': "'",   # left single quote  '
    '’': "'",   # right single quote '
    '“': '"',   # left double quote  "
    '”': '"',   # right double quote "
    '…': '...',  # ellipsis …
})

def _safe_filename(name: str) -> str:
    return name.translate(_UNICODE_TO_ASCII).encode('latin-1', errors='ignore').decode('latin-1')


@app.post('/extract')
async def extract(file: UploadFile = File(...)):
    """Upload HTML → download Excel template."""
    if not file.filename.endswith('.html'):
        raise HTTPException(400, 'Please upload an .html file')
    html = (await file.read()).decode('utf-8', errors='replace')
    xlsx_bytes = extract_to_excel(html)
    stem = _safe_filename(Path(file.filename).stem)
    return Response(
        content=xlsx_bytes,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{stem}_translations.xlsx"'}
    )


@app.post('/auto-translate')
async def auto_translate(file: UploadFile = File(...)):
    """Stream SSE progress; on completion store result and return download URL."""
    if not file.filename.endswith('.html'):
        raise HTTPException(400, 'Please upload an .html file')
    html = (await file.read()).decode('utf-8', errors='replace')
    stem = _safe_filename(Path(file.filename).stem)
    out_filename = stem + '_Telugu.html'

    soup = BeautifulSoup(html, 'html.parser')
    ost_records = list(walk_dom(soup))
    vo_records = extract_narr_records(html)   # [(idx, text), ...]

    if not ost_records and not vo_records:
        raise HTTPException(400, 'No translatable text found in this HTML file')

    ost_texts = [strip_emoji(text) for _, _, _, text in ost_records]
    vo_texts  = [text for _, text in vo_records]
    all_texts = ost_texts + vo_texts
    total     = len(all_texts)
    ost_count = len(ost_texts)

    async def event_stream():
        loop = asyncio.get_event_loop()
        queue = asyncio.Queue()

        def progress_cb(done, tot):
            pct = int(done / tot * 100)
            loop.call_soon_threadsafe(queue.put_nowait, ('progress', done, tot, pct))

        def run_translation():
            try:
                results = translate_batch(all_texts, progress_cb=progress_cb)
                loop.call_soon_threadsafe(queue.put_nowait, ('done', results))
            except Exception as e:
                loop.call_soon_threadsafe(queue.put_nowait, ('error', str(e)))

        loop.run_in_executor(None, run_translation)
        yield f"data: {json.dumps({'type': 'start', 'total': total})}\n\n"

        while True:
            msg = await queue.get()
            if msg[0] == 'progress':
                _, done, tot, pct = msg
                yield f"data: {json.dumps({'type': 'progress', 'done': done, 'total': tot, 'pct': pct})}\n\n"

            elif msg[0] == 'done':
                all_translated = msg[1]
                ost_translated = all_translated[:ost_count]
                vo_translated  = all_translated[ost_count:]

                ost_trans = {idx: t   for (idx,_,_,_), t   in zip(ost_records, ost_translated)}
                ost_check = {idx: orig for idx,_,_,orig    in ost_records}
                vo_trans  = {idx: t   for (idx,_), t       in zip(vo_records, vo_translated)}
                vo_check  = {idx: orig for idx, orig        in vo_records}

                new_html, _, _ = apply_translations(html, ost_trans, ost_check)
                if vo_trans:
                    new_html, _, _ = apply_narr_translations(new_html, vo_trans, vo_check)
                new_html = fix_play_welcome(new_html)
                new_html = update_voice_language(new_html)

                job_id = str(uuid.uuid4())
                _results[job_id] = (out_filename, new_html.encode('utf-8'))
                yield f"data: {json.dumps({'type': 'complete', 'url': f'/download/{job_id}', 'filename': out_filename})}\n\n"
                break

            elif msg[0] == 'error':
                yield f"data: {json.dumps({'type': 'error', 'message': msg[1]})}\n\n"
                break

    return StreamingResponse(event_stream(), media_type='text/event-stream',
                             headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})


@app.get('/download/{job_id}')
async def download(job_id: str):
    """Return the stored translated HTML file for download."""
    entry = _results.pop(job_id, None)   # pop — one-time download, cleans up memory
    if not entry:
        raise HTTPException(404, 'File not found or already downloaded')
    filename, content = entry
    return Response(
        content=content,
        media_type='text/html; charset=utf-8',
        headers={'Content-Disposition': f'attachment; filename="{filename}"'}
    )


@app.post('/apply')
async def apply(
    html_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    tts_api_key: str = Form(''),
):
    """Upload HTML + filled Excel (+ optional ElevenLabs API key) → stream progress → download Telugu HTML."""
    if not html_file.filename.endswith('.html'):
        raise HTTPException(400, 'First file must be .html')
    if not excel_file.filename.endswith('.xlsx'):
        raise HTTPException(400, 'Second file must be .xlsx')

    html         = (await html_file.read()).decode('utf-8', errors='replace')
    xlsx_bytes   = await excel_file.read()
    stem         = Path(html_file.filename).stem
    out_filename = stem + '_Telugu.html'
    tts_api_key  = tts_api_key.strip()

    try:
        ost_trans, ost_check, vo_trans, vo_check = load_translations_from_bytes(xlsx_bytes)
    except Exception as e:
        raise HTTPException(400, f'Could not read Excel file: {e}')

    if not ost_trans and not vo_trans:
        raise HTTPException(400, 'No translations found — fill column F (OST) or column D (VO) first')

    async def event_stream():
        try:
            tts_total = len(vo_trans) if (tts_api_key and vo_trans) else 0
            yield f"data: {json.dumps({'type': 'start', 'tts_total': tts_total})}\n\n"

            # Apply OST + VO text replacements (fast)
            new_html, _, _ = apply_translations(html, ost_trans, ost_check)
            if vo_trans:
                new_html, _, _ = apply_narr_translations(new_html, vo_trans, vo_check)
            new_html = fix_play_welcome(new_html)
            new_html = update_voice_language(new_html)

            # TTS audio generation — sequential, one clip at a time, progress streamed
            audio_clips = {}
            if tts_api_key and vo_trans:
                from elevenlabs.client import ElevenLabs as ELClient
                el_client = ELClient(api_key=tts_api_key)
                vo_items  = list(vo_trans.items())
                total_tts = len(vo_items)
                loop      = asyncio.get_event_loop()

                for i, (idx, telugu_text) in enumerate(vo_items):
                    pct = 10 + int(i / total_tts * 80)
                    yield f"data: {json.dumps({'type': 'tts_progress', 'done': i, 'total': total_tts, 'pct': pct})}\n\n"

                    try:
                        audio_iter = await loop.run_in_executor(
                            None,
                            lambda t=telugu_text: el_client.text_to_speech.convert(
                                text=t,
                                voice_id=_TTS_VOICE_ID,
                                model_id=_TTS_MODEL_ID,
                                output_format=_TTS_FORMAT,
                            )
                        )
                        mp3_bytes = b''.join(
                            chunk for chunk in audio_iter if isinstance(chunk, bytes)
                        )
                        if mp3_bytes:
                            b64 = base64.b64encode(mp3_bytes).decode('ascii')
                            audio_clips[telugu_text] = f'data:audio/mp3;base64,{b64}'
                    except Exception:
                        pass  # non-fatal: this clip falls back to browser TTS

            if audio_clips:
                new_html = inject_audio_clips(new_html, audio_clips)

            job_id = str(uuid.uuid4())
            _results[job_id] = (out_filename, new_html.encode('utf-8'))
            yield f"data: {json.dumps({'type': 'complete', 'url': f'/download/{job_id}', 'filename': out_filename})}\n\n"

        except Exception as e:
            yield f"data: {json.dumps({'type': 'error', 'message': str(e)})}\n\n"

    return StreamingResponse(event_stream(), media_type='text/event-stream',
                             headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})


static_dir = Path(__file__).parent / 'static'
app.mount('/', StaticFiles(directory=str(static_dir), html=True), name='static')
