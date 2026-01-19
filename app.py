from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn
import os
import json
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
from typing import Any # Cần thiết để nhận diện dữ liệu từ Dify

# Lấy cổng từ môi trường Render Pro
PORT = int(os.environ.get("PORT", 8000))

app = FastAPI(title="IDMEA PPAP Professional Generator")

# Cấu hình CORS để Dify truy cập an toàn
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class DocxRequest(BaseModel):
    title: str = "PPAP REPORT"
    customer: str = "CUSTOMER"
    html: Any # Nhận trực tiếp structured_output từ Dify
    filename: str = "ppap.docx"

def add_table_data(doc, title, headers, data_list, keys):
    """Hàm tự động vẽ bảng biểu chuẩn IATF 16949"""
    doc.add_heading(title, level=1)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        run = hdr_cells[i].paragraphs[0].runs[0]
        run.bold = True
    for item in data_list:
        row_cells = table.add_row().cells
        for i, key in enumerate(keys):
            row_cells[i].text = str(item.get(key, ""))

# CỔNG KIỂM TRA - Giúp hết lỗi 404/405 trên Render
@app.get("/")
@app.head("/")
def root():
    return {"status": "online", "message": "IDMEA PPAP Generator Pro is ready"}

@app.get("/health")
def health():
    return {"status": "healthy"}

@app.post("/generate-docx")
async def generate_docx(request: DocxRequest):
    try:
        doc = Document()
        # Tiêu đề báo cáo
        header = doc.add_heading(f"{request.title} - {request.customer}", 0)
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Xử lý dữ liệu JSON (Tự động nhận diện chuỗi hoặc đối tượng)
        ppap_data = request.html if not isinstance(request.html, str) else json.loads(request.html)
        meta = ppap_data.get("Meta", {})

        # I. Thông tin chung
        doc.add_heading("I. THÔNG TIN CHUNG", level=1)
        p = doc.add_paragraph()
        p.add_run(f"Khách hàng: {meta.get('Customer_name', '')}\n").bold = True
        p.add_run(f"Tên linh kiện: {meta.get('Part_name', '')}\n")
        p.add_run(f"Mã linh kiện: {meta.get('Part_number', '')}\n")
        p.add_run(f"Ngày lập: {datetime.now().strftime('%d/%m/%Y')}")

        # II. Bảng PFMEA
        if "PFMEA" in ppap_data:
            headers = ["Công đoạn", "Lỗi tiềm ẩn", "Nguyên nhân", "RPN", "Hành động"]
            keys = ["Process_step", "Failuere_mode", "Cause", "rpn", "recommended_actions"]
            add_table_data(doc, "II. PHÂN TÍCH PFMEA", headers, ppap_data["PFMEA"], keys)

        # III. Bảng Control Plan
        if "Control_plan" in ppap_data:
            headers = ["Công đoạn", "Đặc tính SP", "Thông số KT", "Phương pháp đo", "Tần suất"]
            keys = ["Process_step", "product_characteristic", "spec", "measurement_method", "sample_size_freq"]
            add_table_data(doc, "III. KẾ HOẠCH KIỂM SOÁT", headers, ppap_data["Control_plan"], keys)

        file_path = request.filename
        doc.save(file_path)
        
        return FileResponse(path=file_path, filename=file_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi: {str(e)}")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
