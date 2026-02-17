import os
import re
import io
import csv
import time
import hashlib
import requests
import openpyxl
from pathlib import Path
from fastapi import FastAPI, HTTPException, Header, Query
from fastapi.responses import StreamingResponse, FileResponse
from pydantic import BaseModel

app = FastAPI()

@app.get("/")
async def root():
    return {"status": "ready", "message": "XLSX-to-CSV Bridge is running"}

# --- Configuration ---
API_KEY_NAME = "X-API-KEY"
REQUIRED_API_KEY = os.getenv("API_KEY")
CACHE_DIR = Path("/tmp/xlsx_cache")
CACHE_DIR.mkdir(exist_ok=True)

class ConversionRequest(BaseModel):
    drive_url: str
    sheet_name: str | None = None

def verify_api_key(x_api_key: str = Header(..., alias="X-API-KEY")):
    """Validates the X-API-KEY header."""
    if REQUIRED_API_KEY and x_api_key != REQUIRED_API_KEY:
        raise HTTPException(status_code=403, detail="Invalid API Key")
    return x_api_key

def extract_file_id(url: str) -> str:
    """Extracts the Google Drive file ID from various URL formats."""
    # Standard /d/ID format
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)
    # uc?id=ID format
    match = re.search(r"id=([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)
    # open?id=ID format
    match = re.search(r"open\?id=([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)
    
    raise HTTPException(status_code=400, detail="Could not extract File ID from URL")

def parse_range_header(range_header: str, content_length: int):
    """Parses the Range header and returns start, end."""
    match = re.search(r"bytes=(\d+)-(\d*)", range_header)
    if match:
        start = int(match.group(1))
        end_str = match.group(2)
        end = int(end_str) if end_str else content_length - 1
        return start, end
    return 0, content_length - 1

def perform_conversion(file_id: str, sheet_name: str | None, cache_path: Path):
    """Downloads XLSX and converts to CSV on disk."""
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"

    try:
        response = session.get(download_url, headers=headers, stream=True)
        if response.status_code == 200 and "confirm=" in response.text:
            confirm_match = re.search(r'confirm=([a-zA-Z0-9_-]+)', response.text)
            if confirm_match:
                confirm_token = confirm_match.group(1)
                download_url = f"https://drive.google.com/uc?export=download&id={file_id}&confirm={confirm_token}"
                response = session.get(download_url, headers=headers, stream=True)
        response.raise_for_status()
        
        # Load XLSX in-memory (still needed for openpyxl)
        xlsx_data = io.BytesIO(response.content)
        wb = openpyxl.load_workbook(xlsx_data, read_only=True, data_only=True)
        
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found")
            ws = wb[sheet_name]
        else:
            ws = wb.active

        # Write to local cache file
        with open(cache_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            for row in ws.iter_rows(values_only=True):
                writer.writerow(row)
        
        wb.close()
        xlsx_data.close()
    except Exception as e:
        if isinstance(e, HTTPException): raise e
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

@app.get("/convert")
@app.post("/convert")
async def convert_endpoint(
    request_data: ConversionRequest | None = None,
    drive_url: str | None = Query(None),
    sheet_name: str | None = Query(None),
    api_key: str | None = Query(None),
    x_api_key: str | None = Header(None, alias="X-API-KEY")
):
    # 1. Security Check
    provided_key = x_api_key or api_key
    if REQUIRED_API_KEY and provided_key != REQUIRED_API_KEY:
        raise HTTPException(status_code=403, detail="Unauthorized")

    # 2. Extract params from either Body or Query
    final_url = (request_data.drive_url if request_data else None) or drive_url
    final_sheet = (request_data.sheet_name if request_data else None) or sheet_name

    if not final_url:
        raise HTTPException(status_code=400, detail="drive_url is required")

    file_id = extract_file_id(final_url)
    cache_key = hashlib.md5(f"{file_id}_{final_sheet}".encode()).hexdigest()
    cache_path = CACHE_DIR / f"{cache_key}.csv"

    # 3. Check Cache (valid for 10 minutes to support chunking)
    if not cache_path.exists() or (time.time() - cache_path.stat().st_mtime > 600):
        perform_conversion(file_id, final_sheet, cache_path)

    total_size = cache_path.stat().st_size

    # 4. Handle Range Request or Full Response
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

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
