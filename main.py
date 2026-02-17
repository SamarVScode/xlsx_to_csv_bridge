import os
import re
import io
import requests
import pandas as pd
from fastapi import FastAPI, HTTPException, Header, Body
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

app = FastAPI()

# --- Configuration ---
API_KEY_NAME = "X-API-KEY"
# In production, this should be set in environment variables.
# For local listing/testing, ensure you have this set or provide a default (CAUTION with defaults in prod).
REQUIRED_API_KEY = os.getenv("MY_SECRET_API_KEY")

class ConversionRequest(BaseModel):
    drive_url: str
    sheet_name: str | None = None

def verify_api_key(x_api_key: str = Header(...)):
    """Validates the X-API-KEY header."""
    if REQUIRED_API_KEY and x_api_key != REQUIRED_API_KEY:
        raise HTTPException(status_code=403, detail="Invalid API Key")
    return x_api_key

def extract_file_id(url: str) -> str:
    """Extracts the Google Drive file ID from a URL."""
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
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
    
    # 2. Download from Google Drive
    download_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"

    try:
        response = requests.get(download_url)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        raise HTTPException(status_code=400, detail=f"Failed to download file. Ensure it is accessible. Error: {str(e)}")

    # 3. In-Memory Processing
    try:
        file_content = io.BytesIO(response.content)
        
        # Load specific sheet if requested, otherwise first sheet
        sheet_to_load = request.sheet_name if request.sheet_name else 0

        try:
            df = pd.read_excel(file_content, engine='openpyxl', sheet_name=sheet_to_load)
        except ValueError as e:
            # Handle "Worksheet not found" errors
            raise HTTPException(status_code=400, detail=f"Sheet processing error: {str(e)}")
        
        # Convert to CSV
        output_buffer = io.StringIO()
        df.to_csv(output_buffer, index=False)
        output_buffer.seek(0)
        
        csv_bytes = output_buffer.getvalue().encode('utf-8')
        total_size = len(csv_bytes)

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
