#!/usr/bin/env python3
"""
FastAPI DOCX Generator - Compatible with Render Free Tier
"""

import os
import io
from datetime import datetime
from zipfile import ZipFile

from fastapi import FastAPI, HTTPException
from fastapi.responses import Response, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel  # Pydantic v1
import uvicorn

# ============================================================================
# CONFIGURATION
# ============================================================================
PORT = int(os.environ.get("PORT", 8000))

app = FastAPI(
    title="DOCX Generator API",
    description="Generate DOCX files from JSON data",
    version="1.0.0"
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
# DATA MODELS (Pydantic v1)
# ============================================================================
class DocxRequest(BaseModel):
    title: str = "DOCUMENT"
    content: str = ""
    filename: str = "document.docx"
    
    class Config:
        schema_extra = {
            "example": {
                "title": "My Report",
                "content": "Report content here",
                "filename": "report.docx"
            }
        }

# ============================================================================
# DOCX GENERATION (Simple)
# ============================================================================
def create_docx_content(title: str, content: str) -> bytes:
    """Tạo DOCX đơn giản"""
    
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
    </w:body>
</w:document>'''
    
    buffer = io.BytesIO()
    with ZipFile(buffer, 'w') as zipf:
        zipf.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')
        
        zipf.writestr('word/document.xml', xml_content)
        
        zipf.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
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
            <p><a href="/docs">API Docs</a></p>
        </body>
    </html>
    """)

@app.get("/health")
async def health():
    return {
        "status": "healthy",
        "service": "DOCX Generator",
        "timestamp": datetime.now().isoformat()
    }

@app.post("/generate-docx")
async def generate_docx(request: DocxRequest):
    try:
        docx_bytes = create_docx_content(request.title, request.content)
        
        return Response(
            content=docx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{request.filename}"'
            }
        )
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ============================================================================
# MAIN
# ============================================================================
if __name__ == "__main__":
    print(f"Starting DOCX Generator on port {PORT}")
    uvicorn.run(app, host="0.0.0.0", port=PORT)