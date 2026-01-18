#!/usr/bin/env python3
"""
FastAPI DOCX Generation Service - Fixed for Pydantic v2
"""

import os
import io
from datetime import datetime
from typing import Optional
from zipfile import ZipFile

from fastapi import FastAPI, HTTPException
from fastapi.responses import Response, JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, ConfigDict
import uvicorn

# ============================================================================
# CONFIGURATION
# ============================================================================
PORT = int(os.environ.get("PORT", 8001))
HOST = os.environ.get("HOST", "0.0.0.0")

app = FastAPI(
    title="DOCX Generator API",
    description="Generate DOCX files from JSON data",
    version="2.0.0"
)

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ============================================================================
# DATA MODELS (Pydantic v2 compatible)
# ============================================================================
class DocxRequest(BaseModel):
    title: str = "DOCUMENT"
    content: str = ""
    filename: str = "document.docx"
    
    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "title": "My Report",
                "content": "This is the content of the report.",
                "filename": "report.docx"
            }
        }
    )

# ============================================================================
# DOCX GENERATION
# ============================================================================
def create_docx_content(title: str, content: str) -> bytes:
    """Tạo DOCX đơn giản bằng XML"""
    
    xml_content = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:sz w:val="32"/>
                </w:rPr>
                <w:t>{title}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>{content}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>'''
    
    buffer = io.BytesIO()
    with ZipFile(buffer, 'w') as zip_file:
        # [Content_Types].xml
        zip_file.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')
        
        # word/document.xml
        zip_file.writestr('word/document.xml', xml_content)
        
        # _rels/.rels
        zip_file.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')
        
        # word/_rels/document.xml.rels
        zip_file.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')
    
    buffer.seek(0)
    return buffer.getvalue()

# ============================================================================
# API ENDPOINTS
# ============================================================================
@app.get("/")
async def root():
    return HTMLResponse("""
    <html>
        <head><title>DOCX Generator</title></head>
        <body>
            <h1>DOCX Generator API</h1>
            <p>POST /generate-docx to generate DOCX files</p>
            <p><a href="/docs">API Docs</a> | <a href="/health">Health</a></p>
        </body>
    </html>
    """)

@app.get("/health")
async def health():
    return {
        "status": "healthy",
        "service": "DOCX Generator",
        "timestamp": datetime.now().isoformat(),
        "port": PORT
    }

@app.post("/generate-docx")
async def generate_docx(request: DocxRequest):
    """Generate DOCX file"""
    try:
        docx_bytes = create_docx_content(request.title, request.content)
        
        return Response(
            content=docx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{request.filename}"',
                "Access-Control-Expose-Headers": "Content-Disposition"
            }
        )
    
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Failed to generate DOCX: {str(e)}"
        )

# ============================================================================
# STARTUP
# ============================================================================
def check_port(port: int):
    """Kiểm tra port có khả dụng không"""
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(('0.0.0.0', port))
            return True
        except socket.error:
            return False

if __name__ == "__main__":
    # Kiểm tra port
    if not check_port(PORT):
        print(f"⚠️  Port {PORT} is in use! Trying port 8002...")
        PORT = 8002
    
    print(f"🚀 Starting DOCX Generator on {HOST}:{PORT}")
    print(f"📚 API Docs: http://{HOST}:{PORT}/docs")
    
    uvicorn.run(
        app,
        host=HOST,
        port=PORT,
        log_level="info"
    )