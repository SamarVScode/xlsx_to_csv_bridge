import os
import re
import io
import csv
import sys
import time
import uuid
import hashlib
import shutil
import logging
import threading
import requests
import xlsx2csv
import traceback
from pathlib import Path
from typing import Optional
from fastapi import FastAPI, HTTPException, Header, Query, UploadFile, File, Body
from fastapi.responses import StreamingResponse, FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# --- Logging Setup (flushes immediately for Render) ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-5s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    stream=sys.stdout,
)
log = logging.getLogger("bridge")
log.setLevel(logging.INFO)

app = FastAPI()

# --- CORS Configuration ---
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://script.google.com"],  # Only Google Apps Script needs CORS
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["X-API-KEY", "Content-Type"],
)

log.info("=" * 40)
log.info("XLSX-to-CSV Bridge Starting...")
log.info("=" * 40)

@app.get("/")
async def root():
    return {"status": "ready", "message": "XLSX-to-CSV Bridge is running. Visit /test for the UI."}

@app.get("/health")
async def health():
    """Health check endpoint for Render and uptime monitors."""
    _evict_old_jobs()
    return {"status": "ok", "cache_dir": str(CACHE_DIR), "active_jobs": len(active_jobs)}

def _evict_old_jobs(max_age_seconds: int = 7200):
    """Evict jobs older than max_age_seconds from active_jobs to prevent memory leak."""
    now = time.time()
    to_delete = []
    for jid, job in list(active_jobs.items()):
        created = job.get("created_at", now)
        if now - created > max_age_seconds:
            to_delete.append(jid)
    for jid in to_delete:
        del active_jobs[jid]
    if to_delete:
        log.info(f"[CLEANUP] Evicted {len(to_delete)} old job(s) from active_jobs")

@app.get("/test", response_class=HTMLResponse)
async def test_page():
    return """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLSX-to-CSV Bridge Test Bench</title>
    <style>
        :root {
            --primary: #2563eb;
            --bg: #f8fafc;
            --card: #ffffff;
            --text: #1e293b;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: var(--bg);
            color: var(--text);
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            margin: 0;
        }
        .container {
            background: var(--card);
            padding: 2rem;
            border-radius: 1rem;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
            width: 100%;
            max-width: 500px;
        }
        h1 {
            font-size: 1.5rem;
            margin-bottom: 1.5rem;
            text-align: center;
            color: var(--primary);
        }
        .field {
            margin-bottom: 1rem;
        }
        label {
            display: block;
            margin-bottom: 0.5rem;
            font-weight: 600;
        }
        input[type="text"], input[type="file"] {
            width: 100%;
            padding: 0.75rem;
            border: 1px solid #cbd5e1;
            border-radius: 0.5rem;
            box-sizing: border-box;
        }
        button {
            width: 100%;
            padding: 0.75rem;
            background-color: var(--primary);
            color: white;
            border: none;
            border-radius: 0.5rem;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: background 0.2s;
        }
        button:hover {
            background-color: #1d4ed8;
        }
        button:disabled {
            background-color: #94a3b8;
            cursor: not-allowed;
        }
        #status {
            margin-top: 1.5rem;
            padding: 1rem;
            border-radius: 0.5rem;
            display: none;
            white-space: pre-wrap;
            word-break: break-all;
        }
        .progress-group {
            margin-top: 1.5rem;
            display: none;
        }
        .progress-label {
            font-size: 0.875rem;
            margin-bottom: 0.5rem;
            display: flex;
            justify-content: space-between;
        }
        .progress-bar-container {
            width: 100%;
            height: 10px;
            background: #e2e8f0;
            border-radius: 5px;
            overflow: hidden;
            margin-bottom: 1rem;
        }
        .progress-bar-fill {
            height: 100%;
            background: var(--primary);
            width: 0%;
            transition: width 0.3s;
        }
        .progress-bar-fill.indeterminate {
            width: 100%;
            background: linear-gradient(90deg, #2563eb 25%, #60a5fa 50%, #2563eb 75%);
            background-size: 200% 100%;
            animation: move-bg 1.5s infinite linear;
        }
        @keyframes move-bg {
            0% { background-position: 200% 0; }
            100% { background-position: -200% 0; }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>XLSX-to-CSV Bridge Tester</h1>
        
        <div class="field">
            <label for="bridgeUrl">Bridge Base URL</label>
            <input type="text" id="bridgeUrl" readonly value="">
        </div>

        <div class="field">
            <label for="apiKey">X-API-KEY</label>
            <input type="text" id="apiKey" placeholder="Enter your X-API-KEY here">
        </div>

        <div class="field">
            <label for="fileInput">Select XLSX File (up to 100MB+)</label>
            <input type="file" id="fileInput" accept=".xlsx">
        </div>

        <div class="field">
            <label for="sheetName">Sheet Name (Optional)</label>
            <input type="text" id="sheetName" placeholder="Defaults to All Sheets Merged">
        </div>

        <button id="testBtn">Upload & Convert to CSV</button>

        <div class="progress-group" id="progressGroup">
            <div class="progress-label">
                <span>Uploading...</span>
                <span id="uploadPct">0%</span>
            </div>
            <div class="progress-bar-container">
                <div class="progress-bar-fill" id="uploadBar"></div>
            </div>

            <div class="progress-label" id="convertLabel" style="display:none">
                <span>Converting & Filtering (Please Wait)...</span>
            </div>
            <div class="progress-bar-container" id="convertBarContainer" style="display:none">
                <div class="progress-bar-fill indeterminate"></div>
            </div>
        </div>

        <div id="status"></div>
    </div>

    <script>
        // Auto-detect current host
        document.getElementById('bridgeUrl').value = window.location.origin;

        const testBtn = document.getElementById('testBtn');
        const statusDiv = document.getElementById('status');
        const progressGroup = document.getElementById('progressGroup');
        const uploadBar = document.getElementById('uploadBar');
        const uploadPct = document.getElementById('uploadPct');
        const convertLabel = document.getElementById('convertLabel');
        const convertBarContainer = document.getElementById('convertBarContainer');

        testBtn.addEventListener('click', async () => {
            const bridgeUrl = document.getElementById('bridgeUrl').value.trim();
            const apiKey = document.getElementById('apiKey').value.trim();
            const fileInput = document.getElementById('fileInput');
            const sheetName = document.getElementById('sheetName').value.trim();

            if (!fileInput.files.length) {
                showStatus('Please select a file first.', 'error');
                return;
            }

            const file = fileInput.files[0];
            const formData = new FormData();
            formData.append('file', file);

            const uploadUrl = `${bridgeUrl}/test-upload${sheetName ? `?sheet_name=${encodeURIComponent(sheetName)}` : ''}`;

            statusDiv.style.display = 'none';
            progressGroup.style.display = 'block';
            convertLabel.style.display = 'none';
            convertBarContainer.style.display = 'none';
            uploadBar.style.width = '0%';
            uploadPct.textContent = '0%';
            testBtn.disabled = true;

            const xhr = new XMLHttpRequest();
            xhr.open('POST', uploadUrl, true);
            xhr.setRequestHeader('X-API-KEY', apiKey);
            xhr.responseType = 'blob';

            xhr.upload.onprogress = (e) => {
                if (e.lengthComputable) {
                    const percent = Math.round((e.loaded / e.total) * 100);
                    uploadBar.style.width = percent + '%';
                    uploadPct.textContent = percent + '%';
                    
                    if (percent === 100) {
                        setTimeout(() => {
                            convertLabel.style.display = 'flex';
                            convertBarContainer.style.display = 'block';
                        }, 200);
                    }
                }
            };

            xhr.onload = async () => {
                progressGroup.style.display = 'none';
                if (xhr.status >= 200 && xhr.status < 300) {
                    const blob = xhr.response;
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `${file.name.replace('.xlsx', '')}_filtered.csv`;
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                    showStatus('Success! Filtered CSV has been downloaded.', 'success');
                } else {
                    const reader = new FileReader();
                    reader.onload = () => showStatus(`Error ${xhr.status}: ${reader.result}`, 'error');
                    reader.readAsText(xhr.response);
                }
                testBtn.disabled = false;
            };

            xhr.onerror = () => {
                progressGroup.style.display = 'none';
                showStatus('Network Error during upload.', 'error');
                testBtn.disabled = false;
            };

            xhr.send(formData);
        });

        function showStatus(message, type) {
            statusDiv.textContent = message;
            statusDiv.className = type;
            statusDiv.style.display = 'block';
        }
    </script>
</body>
</html>
"""

# --- Configuration ---
API_KEY_NAME = "X-API-KEY"
REQUIRED_API_KEY = os.getenv("API_KEY")
CACHE_DIR = Path("/tmp/xlsx_cache")
CACHE_DIR.mkdir(exist_ok=True)
CACHE_TTL = 1800  # 30 minutes

# --- Async Job Store ---
active_jobs = {}   # { job_id: { status, cache_path, error, progress, created_at } }
job_totals  = {}   # { job_id: int } — OFD+OFP sum from E2E_Dexter MRZ row, used for early-exit on raw sheets

# --- Concurrency Control ---
# Render free tier: 512MB RAM. A single 250MB XLSX needs ~80MB RAM during conversion.
# Two simultaneous conversions = ~160MB + overhead — technically safe, but we limit
# to 1 at a time to avoid any risk of OOM kill. Queue a second job behind the first.
conversion_semaphore = threading.Semaphore(1)

class ConversionRequest(BaseModel):
    drive_url: str
    sheet_name: str | None = None

def verify_api_key(x_api_key: str = Header(..., alias="X-API-KEY")):
    if REQUIRED_API_KEY and x_api_key != REQUIRED_API_KEY:
        raise HTTPException(status_code=403, detail="Invalid API Key")
    return x_api_key

def extract_file_id(url: str) -> str:
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)
    match = re.search(r"id=([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)
    match = re.search(r"open\?id=([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)
    raise HTTPException(status_code=400, detail="Could not extract File ID from URL")

def parse_range_header(range_header: str, content_length: int):
    match = re.search(r"bytes=(\d+)-(\d*)", range_header)
    if match:
        start = int(match.group(1))
        end_str = match.group(2)
        end = int(end_str) if end_str else content_length - 1
        return start, end
    return 0, content_length - 1

# =============================================================================
# FIXED DOWNLOAD FUNCTION — handles large files & Google's virus-scan warning
# =============================================================================

def download_drive_file(file_id: str, dest_path: Path) -> None:
    """
    Downloads a Google Drive file to dest_path, correctly handling:
      1. Google's virus-scan warning page for large files (>~100MB)
         — detects the HTML warning and re-requests with the confirm token
      2. Streams directly to disk — never loads full file into memory
      3. Validates the downloaded file is a real binary (not an HTML error page)
    """
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0.0.0 Safari/537.36"
    }

    base_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    log.info(f"[DOWNLOAD] Starting download for {file_id}...")

    # ── Step 1: Initial request (no stream yet — we need to check for warning page) ──
    # Use a small initial request to check if Google redirects us to a warning page.
    # For large files, Google returns a small HTML page asking you to confirm.
    # We use stream=True but only peek at the first chunk to detect HTML.
    response = session.get(base_url, headers=headers, stream=True, allow_redirects=True, timeout=(30, 300))
    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "")
    log.info(f"[DOWNLOAD] Initial response Content-Type: {content_type}")

    # ── Step 2: Detect Google's virus-scan warning page ──
    # If Content-Type is HTML, this is the "file too large to scan" warning.
    # We need to extract the confirm token and re-request.
    if "text/html" in content_type:
        log.info(f"[DOWNLOAD] ⚠️ Got HTML response — likely Google's large-file warning. Extracting confirm token...")

        # Read the HTML (it's small — just a warning page)
        html_content = b""
        for chunk in response.iter_content(chunk_size=4096):
            html_content += chunk
            if len(html_content) > 1_000_000:  # Safety: don't read more than 1MB of HTML
                break
        html_text = html_content.decode("utf-8", errors="ignore")

        # Try multiple token extraction patterns Google has used over the years
        confirm_token = None

        # Pattern 1: confirm=t (newer Google Drive)
        m = re.search(r'confirm=([a-zA-Z0-9_\-]+)', html_text)
        if m:
            confirm_token = m.group(1)

        # Pattern 2: Cookie-based token (Google sets download_warning_* cookie)
        if not confirm_token:
            for cookie_name, cookie_val in session.cookies.items():
                if "download_warning" in cookie_name:
                    confirm_token = cookie_val
                    log.info(f"[DOWNLOAD] Got confirm token from cookie: {cookie_name}")
                    break

        # Pattern 3: UUID-style token in form action
        if not confirm_token:
            m = re.search(r'uuid=([a-zA-Z0-9_\-]+)', html_text)
            if m:
                confirm_token = m.group(1)

        if confirm_token:
            confirmed_url = f"https://drive.google.com/uc?export=download&id={file_id}&confirm={confirm_token}"
            log.info(f"[DOWNLOAD] Re-requesting with confirm token: {confirm_token[:8]}...")
            response = session.get(confirmed_url, headers=headers, stream=True, allow_redirects=True, timeout=(30, 300))
            response.raise_for_status()
            content_type = response.headers.get("Content-Type", "")
            log.info(f"[DOWNLOAD] Confirmed response Content-Type: {content_type}")
        else:
            # Try the newer Google Drive export URL format as fallback
            log.warning(f"[DOWNLOAD] No confirm token found. Trying alternate export URL...")
            alt_url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t"
            response = session.get(alt_url, headers=headers, stream=True, allow_redirects=True, timeout=(30, 300))
            response.raise_for_status()
            content_type = response.headers.get("Content-Type", "")
            log.info(f"[DOWNLOAD] Alt URL Content-Type: {content_type}")

    # ── Step 3: Stream the actual file to disk ──
    if "text/html" in content_type:
        # Still getting HTML after confirm attempt — something is wrong
        raise HTTPException(
            status_code=500,
            detail=f"Google Drive returned an HTML page instead of the file. "
                   f"The file may not be publicly shared or the sharing link may be invalid."
        )

    log.info(f"[DOWNLOAD] Streaming file to disk: {dest_path}")
    bytes_written = 0
    with open(dest_path, 'wb') as f:
        for chunk in response.iter_content(chunk_size=32768):  # 32KB chunks
            if chunk:
                f.write(chunk)
                bytes_written += len(chunk)

    size_mb = bytes_written / (1024 * 1024)
    log.info(f"[DOWNLOAD] ✅ Download complete: {size_mb:.1f} MB written to {dest_path.name}")

    # ── Step 4: Validate downloaded file is real binary (not HTML error) ──
    if bytes_written < 100:
        dest_path.unlink(missing_ok=True)
        raise HTTPException(status_code=500, detail=f"Downloaded file is too small ({bytes_written} bytes) — likely an error page.")

    # Check magic bytes: real XLSX starts with PK\x03\x04 (ZIP format)
    with open(dest_path, 'rb') as f:
        magic = f.read(4)

    if magic != b'PK\x03\x04':
        # Read first 500 bytes to help debug what we actually got
        with open(dest_path, 'rb') as f:
            preview = f.read(500).decode('utf-8', errors='ignore')
        dest_path.unlink(missing_ok=True)
        log.error(f"[DOWNLOAD] ❌ File is not a valid XLSX/ZIP. Magic bytes: {magic}. Preview: {preview[:200]}")
        raise HTTPException(
            status_code=500,
            detail=f"Downloaded file is not a valid XLSX (got {magic} instead of ZIP magic bytes). "
                   f"Google may have returned an error page. Ensure the file is publicly shared."
        )

    log.info(f"[DOWNLOAD] ✅ Magic byte check passed — valid XLSX/ZIP confirmed.")

# =============================================================================
# CONVERSION
# =============================================================================

# Sameday-specific sheets — only these two are processed when filename contains "sameday"
SAMEDAY_SHEETS = ["Agent_view", "E2E_DC"]

# D-1 sheet roles
D1_SUMMARY_SHEET  = "E2E_Dexter"   # Sheet 1 — read OFD+OFP total for MRZ row
D1_DC_SHEET       = "E2E_DC"       # Sheet 2 — filter MRZ rows normally
D1_RAW_SHEETS     = ["E2E_Raw", "North", "East", "West", "South"]   # Sheet 3+ raw data (variable)
D1_AGENT_SHEET    = "Agent_view"   # Last sheet — always process fully

def is_sameday_file(filename: str) -> bool:
    """Returns True if the filename contains 'sameday' (case-insensitive)."""
    return "sameday" in (filename or "").lower()

def is_d1_file(filename: str) -> bool:
    """Returns True if the filename contains 'd-1' or 'd1' (case-insensitive)."""
    fn = (filename or "").lower()
    return "d-1" in fn or "_d1" in fn or "-d1" in fn

def extract_ofd_ofp_total(xlsx_path: Path, target_val: str = "MRZ") -> int:
    """
    Scans E2E_Dexter sheet (Sheet 1 of D-1 file).
    Finds the MRZ row, reads OFD and OFP column values, returns their sum.
    This total is used as the early-exit threshold for raw data sheets.
    Returns 0 if not found (disables early-exit — safe fallback).
    """
    try:
        parser = xlsx2csv.Xlsx2csv(str(xlsx_path))
        sheet_info = parser.workbook.sheets

        # Find E2E_Dexter sheet index
        dexter_idx = None
        for s in sheet_info:
            if s["name"].strip().lower() == D1_SUMMARY_SHEET.lower():
                dexter_idx = s["index"]
                break

        if dexter_idx is None:
            log.warning(f"[TOTAL] ⚠️ Sheet '{D1_SUMMARY_SHEET}' not found — early-exit disabled")
            return 0

        log.info(f"[TOTAL] 🔍 Scanning '{D1_SUMMARY_SHEET}' for OFD+OFP total (MRZ row)...")

        result = {"total": 0, "found": False, "ofd_col": None, "ofp_col": None, "dc_col": None}

        class TotalExtractor:
            def __init__(self):
                self._buf  = []
                self._done = False

            def write(self, data):
                if self._done:
                    return
                if isinstance(data, bytes):
                    data = data.decode("utf-8", errors="ignore")
                self._buf.append(data)
                if "\n" not in data:
                    return
                combined = "".join(self._buf)
                self._buf = []
                for line in combined.split("\n"):
                    if not line or self._done:
                        continue
                    try:
                        row = next(csv.reader(io.StringIO(line)))
                    except Exception:
                        continue
                    if not row:
                        continue

                    # First row = headers
                    if result["ofd_col"] is None and result["dc_col"] is None:
                        col_map = {str(h).strip().upper(): i for i, h in enumerate(row)}
                        result["dc_col"]  = col_map.get("SOURCE_DC") or col_map.get("DC")
                        result["ofd_col"] = col_map.get("OFD")
                        result["ofp_col"] = col_map.get("OFP")
                        log.info(f"[TOTAL] Headers found — DC col: {result['dc_col']}, OFD col: {result['ofd_col']}, OFP col: {result['ofp_col']}")
                        continue

                    # Data rows — find MRZ row
                    dc_col = result["dc_col"]
                    if dc_col is not None and len(row) > dc_col:
                        if str(row[dc_col]).strip().upper() == target_val.upper():
                            ofd_val = 0
                            ofp_val = 0
                            try:
                                if result["ofd_col"] is not None and len(row) > result["ofd_col"]:
                                    ofd_val = int(float(str(row[result["ofd_col"]]).strip() or "0"))
                            except (ValueError, TypeError):
                                pass
                            try:
                                if result["ofp_col"] is not None and len(row) > result["ofp_col"]:
                                    ofp_val = int(float(str(row[result["ofp_col"]]).strip() or "0"))
                            except (ValueError, TypeError):
                                pass
                            result["total"] = ofd_val + ofp_val
                            result["found"] = True
                            self._done = True
                            log.info(f"[TOTAL] ✅ MRZ row found — OFD={ofd_val}, OFP={ofp_val}, Total={result['total']}")

        extractor = TotalExtractor()
        xlsx2csv.Xlsx2csv(str(xlsx_path), skip_empty_lines=True).convert(extractor, sheetid=dexter_idx)

        if not result["found"]:
            log.warning(f"[TOTAL] ⚠️ MRZ row not found in '{D1_SUMMARY_SHEET}' — early-exit disabled")
            return 0

        return result["total"]

    except Exception as e:
        log.error(f"[TOTAL] ❌ Failed to extract OFD+OFP total: {e}")
        traceback.print_exc()
        return 0

def convert_xlsx_to_csv(xlsx_path: Path, sheet_name_req: str | None, csv_path: Path, target_val: str = "MRZ", date_val: str | None = None, source_filename: str | None = None, job_id: str | None = None):
    """
    ULTRA-PERFORMANCE VERSION using xlsx2csv (SAX Parser).
    Memory usage: < 50MB even for 100MB+ files.

    Sheet selection priority:
      1. sheet_name_req explicitly passed   → only that sheet
      2. source_filename contains "sameday" → only Agent_view + E2E_DC
      3. source_filename contains "d-1"     → D1 mode with early-exit optimization
      4. Otherwise                          → ALL sheets

    D-1 Early-Exit Optimization:
      - E2E_Dexter (Sheet 1): extract OFD+OFP total for MRZ row → store in job_totals
      - Raw sheets (Sheet 3+): once match_count == total → skip remaining raw sheets
      - Agent_view: always processed fully
    """
    try:
        parser_info = xlsx2csv.Xlsx2csv(str(xlsx_path))
        sheet_info = parser_info.workbook.sheets
        sheet_names = [s["name"] for s in sheet_info]
        log.info(f"[CONVERT] Workbook opened. Sheets found: {sheet_names}")

        target_indices = []
        d1_mode        = False
        early_exit_total = 0  # 0 = disabled

        if sheet_name_req and sheet_name_req.strip():
            # Explicit single sheet requested
            target = sheet_name_req.strip().lower()
            found_idx = None
            for s in sheet_info:
                if s["name"].strip().lower() == target:
                    found_idx = s["index"]
                    break
            if found_idx is None:
                raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name_req}' not found")
            target_indices.append(found_idx)
            log.info(f"[CONVERT] Mode: SINGLE sheet '{sheet_name_req}'")

        elif is_sameday_file(source_filename):
            # Sameday file — only scan Agent_view and E2E_DC
            sameday_lower = [s.lower() for s in SAMEDAY_SHEETS]
            for s in sheet_info:
                if s["name"].strip().lower() in sameday_lower:
                    target_indices.append(s["index"])
            found_names   = [s["name"] for s in sheet_info if s["name"].strip().lower() in sameday_lower]
            skipped_names = [s["name"] for s in sheet_info if s["name"].strip().lower() not in sameday_lower]
            log.info(f"[CONVERT] Mode: SAMEDAY — processing {found_names}, skipping {skipped_names}")
            if not target_indices:
                log.warning("[CONVERT] ⚠️ No matching sameday sheets found. Falling back to ALL sheets.")
                for s in sheet_info:
                    target_indices.append(s["index"])

        elif is_d1_file(source_filename):
            # D-1 mode — smart sheet ordering with early-exit on raw sheets
            d1_mode = True
            raw_lower    = [r.lower() for r in D1_RAW_SHEETS]
            agent_lower  = D1_AGENT_SHEET.lower()
            dc_lower     = D1_DC_SHEET.lower()
            dexter_lower = D1_SUMMARY_SHEET.lower()

            # Build ordered list: E2E_Dexter → E2E_DC → raw sheets → Agent_view
            # E2E_Dexter is BOTH extracted for OFD+OFP total AND written to CSV
            dexter_indices = [s["index"] for s in sheet_info if s["name"].strip().lower() == dexter_lower]
            dc_indices     = [s["index"] for s in sheet_info if s["name"].strip().lower() == dc_lower]
            raw_indices    = [s["index"] for s in sheet_info if s["name"].strip().lower() in raw_lower]
            agent_indices  = [s["index"] for s in sheet_info if s["name"].strip().lower() == agent_lower]

            target_indices = dexter_indices + dc_indices + raw_indices + agent_indices
            log.info(f"[CONVERT] Mode: D-1 — order: {[s['name'] for s in sheet_info if s['index'] in target_indices]}")
            log.info(f"[CONVERT] D-1 — all 4 sheet types written to CSV: E2E_Dexter + E2E_DC + raw + Agent_view")

            # ── Extract OFD+OFP total from E2E_Dexter before main loop ──
            early_exit_total = extract_ofd_ofp_total(xlsx_path, target_val)
            if early_exit_total > 0:
                log.info(f"[CONVERT] ⚡ Early-exit enabled — will stop raw sheets once {early_exit_total} MRZ matches found")
                if job_id:
                    job_totals[job_id] = early_exit_total
            else:
                log.warning("[CONVERT] ⚠️ Early-exit disabled (total=0) — will scan all raw sheets fully")

        else:
            # Default — all sheets
            for s in sheet_info:
                target_indices.append(s["index"])
            log.info(f"[CONVERT] Mode: ALL sheets ({len(target_indices)} total)")

        # Use a large write buffer on the file — reduces OS-level write syscalls significantly
        with open(csv_path, "w", encoding="utf-8", newline="", buffering=256*1024) as f_out:
            writer = csv.writer(f_out)

            # Running match count across all raw sheets (for D-1 early-exit)
            raw_matches_total = 0

            for s_idx in target_indices:
                s_name = next(s['name'] for s in sheet_info if s['index'] == s_idx)

                # D-1 early-exit check: if raw matches already hit total, skip remaining raw sheets
                is_raw_sheet = d1_mode and s_name.strip().lower() in [r.lower() for r in D1_RAW_SHEETS]
                if is_raw_sheet and early_exit_total > 0 and raw_matches_total >= early_exit_total:
                    log.info(f"[FILTER] ⚡ Early-exit triggered — skipping '{s_name}' "
                             f"(already found {raw_matches_total}/{early_exit_total} matches)")
                    continue

                log.info(f"[FILTER] ▶ Processing sheet: '{s_name}' (Index: {s_idx})")
                
                class FilteredOutput:
                    def __init__(self, csv_writer, date_val, target_val, sheet_name, early_exit=0):
                        self.writer       = csv_writer
                        self.date_val     = date_val
                        self.target_val   = target_val
                        self.sheet_name   = sheet_name
                        self.first_row    = True
                        self.target_col   = None
                        self.row_count    = 0
                        self.match_count  = 0
                        # early_exit > 0 means: stop writing once match_count == early_exit
                        # This allows within-sheet early exit on the last raw sheet
                        self.early_exit   = early_exit
                        self._done        = False
                        # ── KEY FIX: use a list instead of str += str ──
                        # str concatenation creates a new object every call (O(n²)).
                        # list.append + "".join is O(n) and allocates far less.
                        self._buf_parts   = []
                        self._buf_len     = 0

                    def write(self, data):
                        if self._done:
                            return
                        if isinstance(data, bytes):
                            data = data.decode('utf-8', errors='ignore')

                        # Accumulate in list — no string copies
                        self._buf_parts.append(data)
                        self._buf_len += len(data)

                        # Only join + process when we have at least one newline
                        # (avoids join cost on every tiny write from xlsx2csv)
                        if '\n' not in data:
                            return

                        combined = ''.join(self._buf_parts)
                        self._buf_parts = []
                        self._buf_len   = 0

                        if '\n' in combined:
                            parts = combined.split('\n')
                            # Everything except the last fragment is a complete line
                            for line in parts[:-1]:
                                if line:   # skip blank lines fast (no .strip() cost)
                                    try:
                                        self.process_row_data(
                                            next(csv.reader(io.StringIO(line)))
                                        )
                                    except Exception:
                                        pass
                            # Keep the trailing incomplete fragment
                            remainder = parts[-1]
                            if remainder:
                                self._buf_parts.append(remainder)
                                self._buf_len = len(remainder)

                    def process_row_data(self, row):
                        self.row_count += 1

                        # Log every 100k rows — 10k is too noisy for 250MB files
                        if self.row_count % 100_000 == 0:
                            log.info(
                                f"[FILTER] ⏳ {s_name}: "
                                f"{self.row_count:,} rows scanned, "
                                f"{self.match_count:,} matches"
                            )

                        if not row:
                            return

                        if self.first_row:
                            self.first_row = False
                            col_map = {str(h).strip(): i for i, h in enumerate(row) if h is not None}
                            self.target_col = col_map.get("Source_DC") if col_map.get("Source_DC") is not None else col_map.get("DC")
                            header_list = ["Sheet"] + list(row)
                            if self.date_val:
                                header_list.append("Date")
                            self.writer.writerow(header_list)
                            return

                        # Fast-path filter: avoid function call overhead on mismatch
                        tc = self.target_col
                        if tc is not None and len(row) > tc:
                            val = row[tc]
                            if val is not None and str(val).strip() == self.target_val:
                                self.match_count += 1
                                row_list = [s_name] + list(row)
                                if self.date_val:
                                    row_list.append(self.date_val)
                                self.writer.writerow(row_list)
                                # Early-exit: stop processing this sheet once quota hit
                                if self.early_exit > 0 and self.match_count >= self.early_exit:
                                    log.info(f"[FILTER] ⚡ Within-sheet early-exit — "
                                             f"hit {self.match_count}/{self.early_exit} matches in '{s_name}'. "
                                             f"Stopping sheet scan.")
                                    self._done = True

                    def finalize(self):
                        # Flush any remaining buffer content
                        if self._buf_parts:
                            remainder = ''.join(self._buf_parts)
                            if remainder.strip():
                                try:
                                    self.process_row_data(
                                        next(csv.reader(io.StringIO(remainder)))
                                    )
                                except Exception:
                                    pass
                            self._buf_parts = []
                            self._buf_len   = 0

                f_output = FilteredOutput(writer, date_val, target_val, s_name,
                                          early_exit=(early_exit_total - raw_matches_total) if is_raw_sheet else 0)
                try:
                    xlsx2csv.Xlsx2csv(str(xlsx_path), skip_empty_lines=True).convert(f_output, sheetid=s_idx)
                    f_output.finalize()
                    log.info(
                        f"[FILTER] ✅ Done '{s_name}': "
                        f"{f_output.row_count:,} rows scanned, "
                        f"{f_output.match_count:,} matches kept"
                    )
                    # Accumulate raw sheet matches for cross-sheet early-exit
                    if is_raw_sheet:
                        raw_matches_total += f_output.match_count
                        if early_exit_total > 0:
                            log.info(f"[FILTER] 📊 Raw progress: {raw_matches_total}/{early_exit_total} matches across raw sheets")
                except Exception as e:
                    log.error(f"[FILTER] ❌ Error on sheet index {s_idx}: {str(e)}")
                    traceback.print_exc()
                    raise e

        log.info("[CONVERT] ✅ Filtering & Conversion Complete.")
    except Exception as e:
        log.error(f"[CONVERT] ❌ CRITICAL ERROR: {str(e)}")
        traceback.print_exc()
        if isinstance(e, HTTPException): raise e
        raise HTTPException(status_code=500, detail=f"Xlsx2csv Error: {str(e)}")

# =============================================================================
# CORE DOWNLOAD + CONVERT (used by all endpoints)
# =============================================================================

def perform_conversion_from_url_with_filter(file_id: str, sheet_name: str | None, cache_path: Path, target_val: str, date_val: str | None, source_filename: str | None = None, job_id: str | None = None):
    """Downloads XLSX via the fixed download_drive_file(), then converts to CSV."""
    tmp_xlsx = CACHE_DIR / f"{file_id}.xlsx"

    try:
        # Use cached XLSX if fresh
        if tmp_xlsx.exists() and (time.time() - tmp_xlsx.stat().st_mtime < CACHE_TTL):
            log.info(f"[CACHE] Using cached source XLSX for {file_id}")
        else:
            # Delete stale cache if present
            if tmp_xlsx.exists():
                tmp_xlsx.unlink()
            # Download with full large-file + confirm-token handling
            download_drive_file(file_id, tmp_xlsx)

        convert_xlsx_to_csv(tmp_xlsx, sheet_name, cache_path, target_val, date_val, source_filename=source_filename, job_id=job_id)

    except Exception as e:
        if tmp_xlsx.exists():
            tmp_xlsx.unlink(missing_ok=True)
        if isinstance(e, HTTPException):
            raise e
        raise HTTPException(status_code=500, detail=str(e))

# Keep old function name for /convert endpoint compatibility
def perform_conversion_from_url(file_id: str, sheet_name: str | None, cache_path: Path):
    perform_conversion_from_url_with_filter(file_id, sheet_name, cache_path, "MRZ", None)

@app.post("/test-upload")
async def test_upload(
    file: UploadFile = File(...),
    sheet_name: Optional[str] = Query(None),
    date_str: Optional[str] = Query(None),
    target_value: str = Query("MRZ"),
    x_api_key: Optional[str] = Header(None, alias="X-API-KEY")
):
    if REQUIRED_API_KEY and x_api_key != REQUIRED_API_KEY:
        raise HTTPException(status_code=403, detail="Unauthorized")

    file_id = f"test_{int(time.time())}"
    tmp_xlsx = CACHE_DIR / f"{file_id}.xlsx"
    cache_path = CACHE_DIR / f"{file_id}.csv"
    start_time = time.time()

    sameday_mode = is_sameday_file(file.filename)
    log.info(f"[UPLOAD] 📥 Received file: '{file.filename}' (Sheet={sheet_name or ('SAMEDAY:Agent_view+E2E_DC' if sameday_mode else 'ALL')}, Target={target_value})")

    try:
        with open(tmp_xlsx, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        file_size_mb = tmp_xlsx.stat().st_size / (1024 * 1024)
        log.info(f"[UPLOAD] 💾 Saved to disk: {file_size_mb:.1f} MB")
        
        convert_xlsx_to_csv(tmp_xlsx, sheet_name, cache_path, target_val=target_value, date_val=date_str, source_filename=file.filename)
        
        if tmp_xlsx.exists(): tmp_xlsx.unlink()
        
        elapsed = time.time() - start_time
        csv_size_kb = cache_path.stat().st_size / 1024
        log.info(f"[UPLOAD] ✅ Complete in {elapsed:.1f}s → CSV output: {csv_size_kb:.1f} KB")
        
        return FileResponse(path=cache_path, media_type="text/csv", filename="converted.csv")
    except Exception as e:
        if tmp_xlsx.exists(): tmp_xlsx.unlink()
        raise HTTPException(status_code=500, detail=str(e))

async def handle_conversion_request(
    drive_url: Optional[str],
    sheet_name: Optional[str],
    api_key_provided: Optional[str],
    range_header: Optional[str],
    date_str: Optional[str] = None,
    target_value: str = "MRZ"
):
    if REQUIRED_API_KEY and api_key_provided != REQUIRED_API_KEY:
        log.warning(f"[AUTH] ⛔ Failed: key does not match")
        raise HTTPException(status_code=403, detail="Unauthorized")

    if not drive_url:
        raise HTTPException(status_code=400, detail="drive_url is required")

    file_id = extract_file_id(drive_url)
    # Try to extract filename from URL for sameday detection
    fname_match = re.search(r"([^/=&]+\.xlsx)", drive_url, re.IGNORECASE)
    source_filename = fname_match.group(1) if fname_match else None
    sameday_mode = is_sameday_file(source_filename or "")
    clean_sheet_log = sheet_name if sheet_name else ("SAMEDAY:Agent_view+E2E_DC" if sameday_mode else "ALL_SHEETS")
    log.info(f"[REQUEST] 📨 File={file_id}, Sheet={clean_sheet_log}, Target={target_value}")

    cache_key_raw = f"{file_id}_{sheet_name}_{target_value}_{date_str}"
    cache_key = hashlib.md5(cache_key_raw.encode()).hexdigest()
    cache_path = CACHE_DIR / f"{cache_key}.csv"

    if not cache_path.exists() or (time.time() - cache_path.stat().st_mtime > CACHE_TTL):
        log.info(f"[REQUEST] 🔄 Cache miss — starting filter & conversion for {file_id}")
        perform_conversion_from_url_with_filter(file_id, sheet_name, cache_path, target_value, date_str, source_filename=source_filename)

    total_size = cache_path.stat().st_size
    log.info(f"[REQUEST] ✅ Response ready: {total_size:,} bytes")

    if range_header:
        start, end = parse_range_header(range_header, total_size)
        if end >= total_size: end = total_size - 1
        if start > end: raise HTTPException(status_code=416, detail="Range Not Satisfiable")
        chunk_size = end - start + 1
        def iterfile():
            with open(cache_path, "rb") as f:
                f.seek(start)
                yield f.read(chunk_size)
        return StreamingResponse(
            iterfile(),
            status_code=206,
            media_type="text/csv",
            headers={
                "Content-Range": f"bytes {start}-{end}/{total_size}",
                "Content-Length": str(chunk_size),
                "Accept-Ranges": "bytes",
                "Content-Disposition": f"attachment; filename={file_id}.csv"
            }
        )
    else:
        return FileResponse(
            path=cache_path,
            media_type="text/csv",
            filename=f"{file_id}.csv",
            headers={"Accept-Ranges": "bytes"}
        )

@app.get("/convert")
async def convert_get(
    drive_url: Optional[str] = Query(None),
    sheet_name: Optional[str] = Query(None),
    api_key: Optional[str] = Query(None),
    date_str: Optional[str] = Query(None),
    target_value: str = Query("MRZ"),
    x_api_key: Optional[str] = Header(None, alias="X-API-KEY"),
    range_header: Optional[str] = Header(None, alias="Range")
):
    return await handle_conversion_request(drive_url, sheet_name, x_api_key or api_key, range_header, date_str, target_value)

@app.post("/convert")
async def convert_post(
    request_data: Optional[ConversionRequest] = Body(None),
    drive_url: Optional[str] = Query(None),
    sheet_name: Optional[str] = Query(None),
    api_key: Optional[str] = Query(None),
    date_str: Optional[str] = Query(None),
    target_value: str = Query("MRZ"),
    x_api_key: Optional[str] = Header(None, alias="X-API-KEY"),
    range_header: Optional[str] = Header(None, alias="Range")
):
    url = drive_url or (request_data.drive_url if request_data else None)
    sn = sheet_name or (request_data.sheet_name if request_data else None)
    return await handle_conversion_request(url, sn, x_api_key or api_key, range_header, date_str, target_value)

# =============================================================================
# ASYNC JOB SYSTEM
# =============================================================================

@app.get("/convert-async")
async def convert_async(
    drive_url: Optional[str] = Query(None),
    sheet_name: Optional[str] = Query(None),
    date_str: Optional[str] = Query(None),
    target_value: str = Query("MRZ"),
    source_filename: Optional[str] = Query(None),  # filename passed by Apps Script for sameday detection
    x_api_key: Optional[str] = Header(None, alias="X-API-KEY"),
    api_key: Optional[str] = Query(None),
):
    key = x_api_key or api_key
    if REQUIRED_API_KEY and key != REQUIRED_API_KEY:
        raise HTTPException(status_code=403, detail="Unauthorized")
    if not drive_url:
        raise HTTPException(status_code=400, detail="drive_url is required")

    file_id = extract_file_id(drive_url)
    job_id = str(uuid.uuid4())[:8]

    # Log whether sameday mode will be used
    if source_filename:
        log.info(f"[ASYNC] Job {job_id}: source_filename='{source_filename}' → sameday={is_sameday_file(source_filename)}")
    else:
        log.info(f"[ASYNC] Job {job_id}: no source_filename provided → ALL sheets mode")

    cache_key_raw = f"{file_id}_{sheet_name}_{target_value}_{date_str}_{source_filename}"
    cache_key = hashlib.md5(cache_key_raw.encode()).hexdigest()
    cache_path = CACHE_DIR / f"{cache_key}.csv"

    if cache_path.exists() and (time.time() - cache_path.stat().st_mtime < CACHE_TTL):
        log.info(f"[ASYNC] Job {job_id}: cache hit, returning done immediately")
        active_jobs[job_id] = {"status": "done", "cache_path": str(cache_path), "error": None, "progress": "Cached"}
        return {"job_id": job_id, "status": "done"}

    active_jobs[job_id] = {"status": "processing", "cache_path": str(cache_path), "error": None, "progress": "Starting...", "created_at": time.time()}
    log.info(f"[ASYNC] Job {job_id}: starting background conversion for {file_id}")

    thread = threading.Thread(
        target=background_conversion,
        args=(job_id, file_id, sheet_name, cache_path, target_value, date_str, source_filename),
        daemon=True
    )
    thread.start()

    return {"job_id": job_id, "status": "processing"}

def background_conversion(job_id, file_id, sheet_name, cache_path, target_val, date_val, source_filename=None):
    acquired = conversion_semaphore.acquire(timeout=1800)  # Wait up to 30 min for slot
    if not acquired:
        active_jobs[job_id]["status"] = "error"
        active_jobs[job_id]["error"] = "Server busy: another conversion is running. Please retry."
        active_jobs[job_id]["progress"] = "Failed: server busy"
        log.error(f"[ASYNC] Job {job_id}: ❌ timed out waiting for conversion slot")
        return
    try:
        active_jobs[job_id]["progress"] = "Downloading XLSX..."
        perform_conversion_from_url_with_filter(file_id, sheet_name, cache_path, target_val, date_val, source_filename=source_filename, job_id=job_id)
        active_jobs[job_id]["status"] = "done"
        active_jobs[job_id]["progress"] = "Complete"
        log.info(f"[ASYNC] Job {job_id}: ✅ completed successfully")
    except Exception as e:
        active_jobs[job_id]["status"] = "error"
        active_jobs[job_id]["error"] = str(e)
        active_jobs[job_id]["progress"] = f"Failed: {str(e)}"
        log.error(f"[ASYNC] Job {job_id}: ❌ failed — {str(e)}")
        traceback.print_exc()
    finally:
        conversion_semaphore.release()

@app.get("/job/{job_id}")
async def job_status(job_id: str):
    job = active_jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found (server may have restarted)")
    return {"job_id": job_id, "status": job["status"], "progress": job["progress"], "error": job["error"]}

@app.get("/job/{job_id}/result")
async def job_result(job_id: str):
    job = active_jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    if job["status"] != "done":
        raise HTTPException(status_code=400, detail=f"Job not ready. Status: {job['status']}")
    cache_path = Path(job["cache_path"])
    if not cache_path.exists():
        raise HTTPException(status_code=410, detail="Result file expired")
    return FileResponse(cache_path, media_type="text/csv", filename=f"{job_id}.csv", headers={"Accept-Ranges": "bytes"})

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
