from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn
import os
from docx import Document # Thư viện để tạo file Word
from datetime import datetime

# Lấy cổng từ môi trường Render
PORT = int(os.environ.get("PORT", 8000))

app = FastAPI(title="IDMEA DOCX Generator", version="1.1")

# Cấu hình CORS để Dify có thể truy cập
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Cấu trúc dữ liệu nhận từ Dify
class DocxRequest(BaseModel):
    title: str = "BÁO CÁO SẢN XUẤT IDMEA"
    content: str = ""
    filename: str = "Bao_cao_PPAP.docx"

@app.get("/")
def root():
    return {"status": "IDMEA API is running", "time": datetime.now()}

@app.get("/health")
def health():
    return {"status": "healthy"}

@app.post("/generate-docx")
async def generate_docx(request: DocxRequest):
    try:
        # 1. Khởi tạo một tài liệu Word mới
        doc = Document()
        
        # 2. Thêm tiêu đề báo cáo
        doc.add_heading(request.title, 0)
        
        # 3. Thêm nội dung (Đây là nơi chứa phân tích AI từ Dify)
        doc.add_paragraph(f"Ngày lập: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        doc.add_paragraph("-" * 20)
        doc.add_paragraph(request.content)
        
        # 4. Lưu file tạm thời trên server Render
        file_path = request.filename
        doc.save(file_path)
        
        # 5. Kiểm tra file tồn tại và trả về định dạng File (không phải JSON)
        if os.path.exists(file_path):
            return FileResponse(
                path=file_path,
                filename=request.filename,
                # Định dạng chuẩn cho file Word (.docx)
                media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            raise HTTPException(status_code=500, detail="Lỗi tạo file trên server")
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi hệ thống: {str(e)}")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
