import os
import json
import uuid
import asyncio
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from translator import (extract_to_excel, load_translations_from_bytes,
                        apply_translations, walk_dom, strip_emoji,
                        extract_narr_records, apply_narr_translations,
                        update_voice_language)
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


@app.post('/extract')
async def extract(file: UploadFile = File(...)):
    """Upload HTML → download Excel template."""
    if not file.filename.endswith('.html'):
        raise HTTPException(400, 'Please upload an .html file')
    html = (await file.read()).decode('utf-8', errors='replace')
    xlsx_bytes = extract_to_excel(html)
    stem = Path(file.filename).stem
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
    stem = Path(file.filename).stem
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
    excel_file: UploadFile = File(...)
):
    """Upload original HTML + filled Excel → download Telugu HTML."""
    if not html_file.filename.endswith('.html'):
        raise HTTPException(400, 'First file must be .html')
    if not excel_file.filename.endswith('.xlsx'):
        raise HTTPException(400, 'Second file must be .xlsx')

    html = (await html_file.read()).decode('utf-8', errors='replace')
    xlsx_bytes = await excel_file.read()

    try:
        ost_trans, ost_check, vo_trans, vo_check = load_translations_from_bytes(xlsx_bytes)
    except Exception as e:
        raise HTTPException(400, f'Could not read Excel file: {e}')

    if not ost_trans and not vo_trans:
        raise HTTPException(400, 'No translations found in the Excel file — fill column F (OST) or column D (VO) first')

    new_html, _, _ = apply_translations(html, ost_trans, ost_check)
    if vo_trans:
        new_html, _, _ = apply_narr_translations(new_html, vo_trans, vo_check)
    new_html = update_voice_language(new_html)

    stem = Path(html_file.filename).stem
    return Response(
        content=new_html.encode('utf-8'),
        media_type='text/html; charset=utf-8',
        headers={'Content-Disposition': f'attachment; filename="{stem}_Telugu.html"'}
    )


static_dir = Path(__file__).parent / 'static'
app.mount('/', StaticFiles(directory=str(static_dir), html=True), name='static')
