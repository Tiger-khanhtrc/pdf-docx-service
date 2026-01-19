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
from typing import Any # Cần thiết để nhận diện mọi kiểu dữ liệu

PORT = int(os.environ.get("PORT", 8000))
app = FastAPI(title="IDMEA PPAP Professional Generator")

app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

class DocxRequest(BaseModel):
    title: str = "PPAP REPORT"
    customer: str = "TTI"
    html: Any # Chuyển từ str sang Any để nhận trực tiếp structured_output
    filename: str = "ppap.docx"

def add_table_data(doc, title, headers, data_list, keys):
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

@app.get("/")
def root():
    return {"status": "online", "service": "IDMEA Pro"}

@app.get("/health")
def health():
    return {"status": "healthy"}

@app.post("/generate-docx")
async def generate_docx(request: DocxRequest):
    try:
        doc = Document()
        header = doc.add_heading(f"{request.title} - {request.customer}", 0)
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # KIỂM TRA KIỂU DỮ LIỆU ĐẦU VÀO
        if isinstance(request.html, str):
            ppap_data = json.loads(request.html)
        else:
            ppap_data = request.html # Nếu là object thì dùng luôn, không cần loads()

        meta = ppap_data.get("Meta", {})

        doc.add_heading("I. THÔNG TIN CHUNG", level=1)
        p = doc.add_paragraph()
        p.add_run("Khách hàng: ").bold = True
        p.add_run(meta.get("Customer_name", ""))
        p.add_run("\nTên linh kiện: ").bold = True
        p.add_run(meta.get("Part_name", ""))
        p.add_run("\nMã linh kiện: ").bold = True
        p.add_run(meta.get("Part_number", ""))
        p.add_run("\nNgày lập báo cáo: ").bold = True
        p.add_run(datetime.now().strftime('%d/%m/%Y'))

        if "PFMEA" in ppap_data:
            headers = ["Công đoạn", "Lỗi tiềm ẩn", "Nguyên nhân", "S", "O", "D", "RPN", "Hành động"]
            keys = ["Process_step", "Failuere_mode", "Cause", "severity", "occurrence", "detection", "rpn", "recommended_actions"]
            add_table_data(doc, "II. PHÂN TÍCH PFMEA", headers, ppap_data["PFMEA"], keys)

        if "Control_plan" in ppap_data:
            headers = ["Công đoạn", "Đặc tính SP", "Thông số KT", "Phương pháp đo", "Tần suất", "Xử lý"]
            keys = ["Process_step", "product_characteristic", "spec", "measurement_method", "sample_size_freq", "reaction_plan"]
            add_table_data(doc, "III. KẾ HOẠCH KIỂM SOÁT", headers, ppap_data["Control_plan"], keys)

        file_path = request.filename
        doc.save(file_path)
        
        return FileResponse(path=file_path, filename=file_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi: {str(e)}")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
