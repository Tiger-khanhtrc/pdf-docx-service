#!/usr/bin/env python3
"""
FastAPI DOCX Generation Service - Lightweight version for Render
"""

import os
import io
import json
from datetime import datetime
from typing import Optional
from zipfile import ZipFile
from fastapi import FastAPI, HTTPException
from fastapi.responses import Response, JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
import uvicorn

# ============================================================================
# CONFIGURATION
# ============================================================================
PORT = int(os.environ.get("PORT", 8001))
HOST = os.environ.get("HOST", "0.0.0.0")
ENVIRONMENT = os.environ.get("ENVIRONMENT", "production")

app = FastAPI(title="DOCX Generator", version="1.0.0")

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ============================================================================
# DATA MODELS
# ============================================================================
class DocxRequest(BaseModel):
    title: str = "PPAP REPORT"
    customer: str = ""
    content: str = ""
    filename: str = "report.docx"

# ============================================================================
# DOCX GENERATION (Simple version)
# ============================================================================
def create_simple_docx_bytes(title: str, content: str) -> bytes:
    """
    Tạo DOCX đơn giản từ template XML
    Đây là fallback khi python-docx không hoạt động
    """
    # Template DOCX cơ bản (minimal working DOCX)
    docx_structure = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>""",
        
        "word/document.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>{title}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>{content}</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>""",
        
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    }
    
    # Tạo ZIP file (DOCX là ZIP file)
    buffer = io.BytesIO()
    with ZipFile(buffer, 'w') as zip_file:
        for filename, content in docx_structure.items():
            zip_file.writestr(filename, content)
    
    buffer.seek(0)
    return buffer.getvalue()

# ============================================================================
# API ENDPOINTS
# ============================================================================
@app.get("/")
async def root():
    return {"status": "ok", "service": "DOCX Generator"}

@app.get("/health")
async def health():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

@app.post("/generate-docx")
async def generate_docx(req: DocxRequest):
    try:
        # Tạo nội dung
        full_content = f"Customer: {req.customer}\n\n{req.content}"
        
        # Tạo DOCX bytes
        docx_bytes = create_simple_docx_bytes(req.title, full_content)
        
        # Trả về file
        return Response(
            content=docx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{req.filename}"'}
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error: {str(e)}")

# ============================================================================
# MAIN
# ============================================================================
if __name__ == "__main__":
    print(f"Starting server on {HOST}:{PORT}")
    uvicorn.run(app, host=HOST, port=PORT)