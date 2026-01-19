from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn
import os
import json
from docx import Document
from datetime import datetime
from typing import Any # Cần thiết để nhận diện mọi kiểu dữ liệu từ Dify

PORT = int(os.environ.get("PORT", 8000))
app = FastAPI(title="IDMEA PPAP Final Fix")

app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# 1. Cấu hình nhận diện dữ liệu thông minh
class DocxRequest(BaseModel):
    title: str = "PPAP REPORT"
    customer: str = "CUSTOMER"
    html: Any # Sửa từ str sang Any để nhận trực tiếp Object từ Dify (Sửa lỗi 422)
    filename: str = "ppap.docx"

@app.get("/")
@app.head("/")
def root():
    return {"status": "online", "message": "IDMEA Pro is Ready"}

@app.post("/generate-docx")
async def generate_docx(request: DocxRequest):
    try:
        # 2. Tự động xử lý dữ liệu (Dù Dify gửi Object hay String)
        if isinstance(request.html, dict):
            ppap_data = request.html
        elif isinstance(request.html, str):
            # Tự động dọn dẹp Markdown tags (Sửa lỗi 400)
            clean_str = request.html.strip()
            if clean_str.startswith("```json"):
                clean_str = clean_str.replace("```json", "", 1)
            if clean_str.endswith("```"):
                clean_str = clean_str.rsplit("```", 1)[0]
            ppap_data = json.loads(clean_str.strip())
        else:
            ppap_data = {}

        # 3. Tạo file Word chuyên nghiệp
        doc = Document()
        doc.add_heading(f"{request.title} - {request.customer}", 0)
        
        meta = ppap_data.get("Meta", {})
        doc.add_heading("I. THÔNG TIN CHUNG", level=1)
        doc.add_paragraph(f"Linh kiện: {meta.get('Part_name', 'N/A')}")
        doc.add_paragraph(f"Mã sản phẩm: {meta.get('Part_number', 'N/A')}")
        doc.add_paragraph(f"Ngày lập: {datetime.now().strftime('%d/%m/%Y')}")

        # Vẽ bảng PFMEA (Nếu có dữ liệu)
        pfmea = ppap_data.get("PFMEA", [])
        if pfmea:
            doc.add_heading("II. PHÂN TÍCH PFMEA", level=1)
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            for i, h in enumerate(["Công đoạn", "Lỗi tiềm ẩn", "RPN", "Hành động"]):
                table.rows[0].cells[i].text = h
            for item in pfmea:
                row = table.add_row().cells
                row[0].text = str(item.get("Process_step", ""))
                row[1].text = str(item.get("Failuere_mode", ""))
                row[2].text = str(item.get("rpn", ""))
                row[3].text = str(item.get("recommended_actions", ""))

        file_path = request.filename
        doc.save(file_path)
        
        return FileResponse(path=file_path, filename=file_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
