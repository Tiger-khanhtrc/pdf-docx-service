from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse
import json
import os
from docx import Document
from datetime import datetime
from typing import Any

app = FastAPI()

@app.post("/generate-docx")
async def generate_docx(request: Any): # Dùng Any để nhận mọi loại Request từ Dify
    try:
        # 1. Nhận dữ liệu thô
        body = await request.json()
        raw_html = body.get("html", "")
        customer = body.get("customer", "Canon")
        
        # 2. BỘ LỌC TỰ CHỮA LÀNH: Loại bỏ thẻ ```json và các ký tự lạ
        if isinstance(raw_html, str):
            clean_html = raw_html.strip()
            if clean_html.startswith("```json"):
                clean_html = clean_html.replace("```json", "", 1)
            if clean_html.endswith("```"):
                clean_html = clean_html.rsplit("```", 1)[0]
            
            try:
                ppap_data = json.loads(clean_html.strip())
            except Exception as json_err:
                # Nếu vẫn lỗi, trả về nội dung nhận được để anh kiểm tra trong TRACING
                return JSONResponse(status_code=400, content={"error": "JSON Error", "received": clean_html[:200]})
        else:
            ppap_data = raw_html

        # 3. Tạo file Word (Giữ nguyên logic cũ)
        doc = Document()
        doc.add_heading(f"PPAP REPORT - {customer}", 0)
        # ... (các phần vẽ bảng phía dưới giữ nguyên) ...

        file_path = "Canon_Report.docx"
        doc.save(file_path)
        return FileResponse(path=file_path, filename=file_path)

    except Exception as e:
        return JSONResponse(status_code=500, content={"detail": str(e)})

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
