from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn
import os
import json
from docx import Document
from datetime import datetime
from typing import Any

PORT = int(os.environ.get("PORT", 8000))
app = FastAPI(title="IDMEA PPAP Stable Generator")

app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

class DocxRequest(BaseModel):
    title: str = "PPAP REPORT"
    customer: str = "CUSTOMER"
    html: Any = None
    filename: str = "ppap.docx"

@app.get("/")
@app.head("/")
def root():
    return {"status": "online", "service": "IDMEA Pro"}

@app.get("/health")
def health():
    return {"status": "healthy"}

@app.post("/generate-docx")
async def generate_docx(request: DocxRequest):
    try:
        # 1. Kiểm tra và giải mã dữ liệu an toàn
        data_raw = request.html
        if isinstance(data_raw, str):
            try:
                ppap_data = json.loads(data_raw)
            except:
                return JSONResponse(status_code=400, content={"error": "Dữ liệu 'html' không phải JSON chuẩn"})
        else:
            ppap_data = data_raw if data_raw else {}

        # 2. Tạo file Word
        doc = Document()
        doc.add_heading(f"{request.title} - {request.customer}", 0)
        
        meta = ppap_data.get("Meta", {})
        doc.add_heading("I. THÔNG TIN CHUNG", level=1)
        doc.add_paragraph(f"Khách hàng: {meta.get('Customer_name', 'N/A')}")
        doc.add_paragraph(f"Linh kiện: {meta.get('Part_name', 'N/A')}")
        doc.add_paragraph(f"Ngày: {datetime.now().strftime('%d/%m/%Y')}")

        # 3. Vẽ bảng PFMEA an toàn
        pfmea_list = ppap_data.get("PFMEA", [])
        if isinstance(pfmea_list, list) and len(pfmea_list) > 0:
            doc.add_heading("II. PHÂN TÍCH PFMEA", level=1)
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(["Công đoạn", "Lỗi", "RPN", "Hành động"]):
                hdr_cells[i].text = h
            for item in pfmea_list:
                row = table.add_row().cells
                row[0].text = str(item.get("Process_step", ""))
                row[1].text = str(item.get("Failuere_mode", ""))
                row[2].text = str(item.get("rpn", ""))
                row[3].text = str(item.get("recommended_actions", ""))

        file_path = request.filename if request.filename.endswith(".docx") else "report.docx"
        doc.save(file_path)
        
        return FileResponse(path=file_path, filename=file_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    except Exception as e:
        # Trả về lỗi chi tiết để anh xem trong Dify Tracing
        return JSONResponse(status_code=500, content={"error_detail": str(e)})

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
