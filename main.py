import os
import re
import io
import csv
import requests
import openpyxl
from fastapi import FastAPI, HTTPException, Header, Body
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

app = FastAPI()

@app.get("/")
async def root():
    return {"status": "ready", "message": "XLSX-to-CSV Bridge is running"}

# --- Configuration ---
API_KEY_NAME = "X-API-KEY"
# In production, this should be set in environment variables.
REQUIRED_API_KEY = os.getenv("API_KEY")

class ConversionRequest(BaseModel):
    drive_url: str
    sheet_name: str | None = None

def verify_api_key(x_api_key: str = Header(...)):
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
    # Pattern: bytes=0-1023
    match = re.search(r"bytes=(\d+)-(\d*)", range_header)
    if match:
        start = int(match.group(1))
        end_str = match.group(2)
        end = int(end_str) if end_str else content_length - 1
        return start, end
    return 0, content_length - 1

@app.post("/convert")
async def convert_endpoint(
    request: ConversionRequest, 
    x_api_key: str = Header(..., alias="X-API-KEY"),
    range_header: str | None = Header(None, alias="Range")
):
    """
    Endpoint to convert a Google Drive XLSX file to CSV.
    Supports Range requests for chunked downloads.
    """
    # 1. Security Check
    if REQUIRED_API_KEY and x_api_key != REQUIRED_API_KEY:
        raise HTTPException(status_code=403, detail="Unauthorized")

    file_id = extract_file_id(request.drive_url)
    
    # 2. Download from Google Drive with large file handling
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"

    try:
        # First attempt
        response = session.get(download_url, headers=headers, stream=True)
        
        # If file is large, Google returns a confirmation page
        if response.status_code == 200 and "confirm=" in response.text:
            # Extract confirmation token
            confirm_match = re.search(r'confirm=([a-zA-Z0-9_-]+)', response.text)
            if confirm_match:
                confirm_token = confirm_match.group(1)
                download_url = f"https://drive.google.com/uc?export=download&id={file_id}&confirm={confirm_token}"
                response = session.get(download_url, headers=headers, stream=True)

        response.raise_for_status()
        content = response.content
    except requests.exceptions.RequestException as e:
        status = response.status_code if 'response' in locals() else 400
        raise HTTPException(status_code=status, detail=f"Failed to download file. Error: {str(e)}")

    # 3. Stream Processing (Memory Efficient)
    try:
        # Save bytes to a temporary seekable stream
        xlsx_file = io.BytesIO(content)
        
        # Open in read_only mode to stream rows
        wb = openpyxl.load_workbook(xlsx_file, read_only=True, data_only=True)
        
        # Select sheet
        if request.sheet_name:
            if request.sheet_name not in wb.sheetnames:
                raise HTTPException(status_code=400, detail=f"Sheet '{request.sheet_name}' not found")
            ws = wb[request.sheet_name]
        else:
            ws = wb.active

        # Convert to CSV using a buffer
        output_buffer = io.StringIO()
        csv_writer = csv.writer(output_buffer)

        # Iterate through rows and write to buffer
        for row in ws.iter_rows(values_only=True):
            csv_writer.writerow(row)
        
        csv_bytes = output_buffer.getvalue().encode('utf-8')
        total_size = len(csv_bytes)
        
        # Cleanup
        wb.close()
        xlsx_file.close()

        # 4. Handle Range Request or Full Response
        if range_header:
            start, end = parse_range_header(range_header, total_size)
            
            # Clamp end
            if end >= total_size:
                end = total_size - 1
            
            # Handle invalid range
            if start > end:
                raise HTTPException(status_code=416, detail="Range Not Satisfiable")

            chunk_size = end - start + 1
            chunk_data = csv_bytes[start : end + 1]
            
            return StreamingResponse(
                io.BytesIO(chunk_data),
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
            return StreamingResponse(
                io.BytesIO(csv_bytes),
                status_code=200,
                media_type="text/csv",
                headers={
                    "Content-Length": str(total_size),
                    "Accept-Ranges": "bytes",
                    "Content-Disposition": f"attachment; filename={file_id}.csv"
                }
            )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
