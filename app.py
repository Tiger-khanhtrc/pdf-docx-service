from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn
import os
import json
from docx import Document
from datetime import datetime

PORT = int(os.environ.get("PORT", 8000))

app = FastAPI(title="IDMEA PPAP Word Generator", version="1.2")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Cấu trúc này khớp 100% với dữ liệu TTI anh vừa gửi trong log
class DocxRequest(BaseModel):
    title: str = "PPAP REPORT"
    customer: str = "TTI"
    html: str = "" # Đây là nơi chứa chuỗi JSON khổng lồ

@app.post("/generate-docx")
async def generate_docx(request: DocxRequest):
    try:
        doc = Document()
        doc.add_heading(f"{request.title} - {request.customer}", 0)
        
        # Giải mã chuỗi JSON trong trường 'html'
        try:
            ppap_data = json.loads(request.html)
            meta = ppap_data.get("Meta", {})
            
            # Ghi thông tin chung
            doc.add_heading("1. Thông tin chung", level=1)
            doc.add_paragraph(f"Sản phẩm: {meta.get('Part_name')}")
            doc.add_paragraph(f"Mã sản phẩm: {meta.get('Part_number')}")
            doc.add_paragraph(f"Ngày lập: {datetime.now().strftime('%d/%m/%Y')}")

            # Ghi dữ liệu PFMEA (Nếu có)
            if "PFMEA" in ppap_data:
                doc.add_heading("2. Phân tích PFMEA", level=1)
                for item in ppap_data["PFMEA"]:
                    p = doc.add_paragraph(style='List Bullet')
                    p.add_run(f"{item['Process_step']}: ").bold = True
                    p.add_run(f"Lỗi: {item['Failuere_mode']} - RPN: {item['rpn']}")

        except Exception as e:
            # Nếu chuỗi html không phải JSON, ghi thô vào file
            doc.add_paragraph("Nội dung chi tiết:")
            doc.add_paragraph(request.html)

        file_path = "Bao_cao_PPAP_IDMEA.docx"
        doc.save(file_path)
        
        if os.path.exists(file_path):
            return FileResponse(
                path=file_path,
                filename=file_path,
                media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        raise HTTPException(status_code=500, detail="Không thể tạo file")
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
