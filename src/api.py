"""
FastAPI application for uploading and parsing Excel files.

Run with:
    uvicorn xlsx_parser.api:app --reload --port 8080

Or use the xlsx-parser-api script (runs on port 8080 by default):
    xlsx-parser-api

Then open http://localhost:8080 in your browser to upload a file.
"""

from __future__ import annotations

import json
import time
import traceback
from pathlib import Path

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse

from pipeline import parse_workbook
from verification import StageVerifier

app = FastAPI(
    title="ks-xlsx-parser API",
    description="Upload an Excel file and get structured JSON output",
    version="0.1.0",
)


# ---------------------------------------------------------------------------
# HTML UI
# ---------------------------------------------------------------------------

UPLOAD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ks-xlsx-parser</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
         background: #0f172a; color: #e2e8f0; min-height: 100vh; padding: 2rem; }
  .container { max-width: 900px; margin: 0 auto; }
  h1 { font-size: 1.8rem; margin-bottom: 0.5rem; color: #f1f5f9; }
  .subtitle { color: #94a3b8; margin-bottom: 2rem; font-size: 0.95rem; }
  .upload-area { border: 2px dashed #334155; border-radius: 12px; padding: 3rem;
                 text-align: center; cursor: pointer; transition: all 0.2s;
                 background: #1e293b; margin-bottom: 1.5rem; }
  .upload-area:hover, .upload-area.dragover { border-color: #3b82f6; background: #1e3a5f; }
  .upload-area input { display: none; }
  .upload-area p { color: #94a3b8; margin-top: 0.5rem; }
  .upload-icon { font-size: 2.5rem; margin-bottom: 0.5rem; }
  .tabs { display: flex; gap: 0; margin-bottom: 0; }
  .tab { padding: 0.6rem 1.2rem; background: #1e293b; border: 1px solid #334155;
         border-bottom: none; cursor: pointer; color: #94a3b8; font-size: 0.85rem;
         border-radius: 8px 8px 0 0; }
  .tab.active { background: #0f172a; color: #f1f5f9; border-color: #475569; }
  .result-box { background: #0f172a; border: 1px solid #475569; border-radius: 0 8px 8px 8px;
                padding: 1rem; min-height: 200px; max-height: 70vh; overflow: auto;
                display: none; }
  .result-box.active { display: block; }
  pre { white-space: pre-wrap; word-break: break-word; font-size: 0.82rem;
        line-height: 1.5; color: #cbd5e1; }
  .stats { display: flex; gap: 1rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
  .stat { background: #1e293b; border: 1px solid #334155; border-radius: 8px;
          padding: 0.8rem 1.2rem; min-width: 120px; }
  .stat-label { font-size: 0.75rem; color: #64748b; text-transform: uppercase; }
  .stat-value { font-size: 1.4rem; font-weight: 600; color: #f1f5f9; }
  .loading { display: none; text-align: center; padding: 2rem; color: #94a3b8; }
  .spinner { display: inline-block; width: 24px; height: 24px; border: 3px solid #334155;
             border-top-color: #3b82f6; border-radius: 50%;
             animation: spin 0.8s linear infinite; margin-right: 0.5rem; vertical-align: middle; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .error { color: #f87171; background: #1e293b; padding: 1rem; border-radius: 8px;
           border: 1px solid #7f1d1d; margin-top: 1rem; display: none; }
  .filename { color: #3b82f6; font-weight: 500; }
</style>
</head>
<body>
<div class="container">
  <h1>ks-xlsx-parser</h1>
  <p class="subtitle">Upload an Excel file to parse it into structured JSON</p>

  <div class="upload-area" id="dropzone">
    <div class="upload-icon">&#128196;</div>
    <strong>Drop .xlsx file here or click to browse</strong>
    <p>Supports .xlsx files up to 50MB</p>
    <input type="file" id="fileInput" accept=".xlsx,.xlsm,.xls">
  </div>

  <div class="loading" id="loading">
    <span class="spinner"></span> Parsing <span class="filename" id="loadingName"></span>...
  </div>

  <div class="error" id="error"></div>

  <div id="results" style="display:none;">
    <div class="stats" id="stats"></div>
    <div class="tabs">
      <div class="tab active" data-tab="json">JSON</div>
      <div class="tab" data-tab="chunks">Chunks</div>
      <div class="tab" data-tab="verification">Verification</div>
    </div>
    <div class="result-box active" id="tab-json"><pre id="jsonOutput"></pre></div>
    <div class="result-box" id="tab-chunks"><pre id="chunksOutput"></pre></div>
    <div class="result-box" id="tab-verification"><pre id="verificationOutput"></pre></div>
  </div>
</div>

<script>
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const loading = document.getElementById('loading');
const errorDiv = document.getElementById('error');
const results = document.getElementById('results');

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', e => {
  e.preventDefault(); dropzone.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => { if (fileInput.files.length) handleFile(fileInput.files[0]); });

document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.result-box').forEach(b => b.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById('tab-' + tab.dataset.tab).classList.add('active');
  });
});

async function handleFile(file) {
  errorDiv.style.display = 'none';
  results.style.display = 'none';
  loading.style.display = 'block';
  document.getElementById('loadingName').textContent = file.name;

  const form = new FormData();
  form.append('file', file);

  try {
    const resp = await fetch('/parse', { method: 'POST', body: form });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.detail || 'Parse failed');
    loading.style.display = 'none';
    showResults(data, file.name);
  } catch (err) {
    loading.style.display = 'none';
    errorDiv.textContent = err.message;
    errorDiv.style.display = 'block';
  }
}

function showResults(data, filename) {
  results.style.display = 'block';
  const wb = data.parse_result.workbook;
  document.getElementById('stats').innerHTML = [
    stat('File', filename),
    stat('Sheets', wb.total_sheets),
    stat('Cells', wb.total_cells.toLocaleString()),
    stat('Formulas', wb.total_formulas),
    stat('Chunks', data.parse_result.total_chunks),
    stat('Tokens', data.parse_result.total_tokens.toLocaleString()),
    stat('Time', wb.parse_duration_ms?.toFixed(0) + 'ms'),
  ].join('');

  document.getElementById('jsonOutput').textContent = JSON.stringify(data.parse_result, null, 2);
  document.getElementById('chunksOutput').textContent = data.parse_result.chunks
    .map((c, i) => `--- Chunk ${i} [${c.block_type}] ${c.source_uri} ---\\n${c.render_text}`)
    .join('\\n\\n');
  document.getElementById('verificationOutput').textContent = data.verification_markdown;
}

function stat(label, value) {
  return `<div class="stat"><div class="stat-label">${label}</div><div class="stat-value">${value}</div></div>`;
}
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------


@app.get("/", response_class=HTMLResponse)
async def index():
    """Serve the upload UI."""
    return UPLOAD_HTML


@app.post("/parse")
async def parse_excel(file: UploadFile = File(...)):
    """
    Parse an uploaded Excel file and return JSON results.

    Returns the full parse result plus a verification report.
    """
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        return JSONResponse(
            status_code=400,
            content={"detail": "Only .xlsx and .xlsm files are supported"},
        )

    content = await file.read()
    if len(content) > 50 * 1024 * 1024:
        return JSONResponse(
            status_code=400,
            content={"detail": "File too large (max 50MB)"},
        )

    try:
        # Parse the workbook
        result = parse_workbook(content=content, filename=file.filename)
        parse_json = result.to_json()

        # Run stage verification
        verifier = StageVerifier(content=content, filename=file.filename)
        report = verifier.verify()

        return {
            "parse_result": parse_json,
            "verification_markdown": report.to_markdown(),
            "verification": report.to_json(),
        }
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"detail": f"Parse error: {str(e)}"},
        )


def main() -> None:
    """Run the API server on port 8080. Use: xlsx-parser-api"""
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080, reload=True)
