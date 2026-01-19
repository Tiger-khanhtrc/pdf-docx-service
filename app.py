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
app = FastAPI(title="IDMEA PPAP Generator Pro")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class DocxRequest(BaseModel):
    title: str = "PPAP REPORT"
    customer: str = "TTI"
    html: str = ""

# --- PHẦN QUAN TRỌNG ĐỂ HẾT LỖI 404 ---
@app.get("/")
def root():
    return {"status": "online", "service": "IDMEA Pro"}

@app.get("/health")
def health():
    return {"status": "healthy"}
# ---------------------------------------

@app.post("/generate-docx")
async def generate_docx(request: DocxRequest):
    try:
        doc = Document()
        doc.add_heading(f"{request.title} - {request.customer}", 0)
        
        # Xử lý dữ liệu JSON khổng lồ từ Dify
        ppap_data = json.loads(request.html)
        doc.add_paragraph(f"Báo cáo tạo lúc: {datetime.now().strftime('%H:%M:%S')}")
        
        # (Anh có thể thêm logic vẽ bảng ở đây như bản trước em gửi)
        doc.add_paragraph(str(ppap_data)[:5000]) # Ghi tạm dữ liệu

        file_path = "Bao_cao_PPAP_IDMEA.docx"
        doc.save(file_path)
        
        return FileResponse(path=file_path, filename=file_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
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

PORT = int(os.environ.get("PORT", 8000))
app = FastAPI(title="IDMEA PPAP Professional Generator")

app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

class DocxRequest(BaseModel):
    title: str = "PPAP REPORT"
    customer: str = "TTI"
    html: str = ""

def add_table_data(doc, title, headers, data_list, keys):
    """Hàm hỗ trợ tạo bảng nhanh chóng"""
    doc.add_heading(title, level=1)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    # Thiết lập tiêu đề bảng (Header)
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        run = hdr_cells[i].paragraphs[0].runs[0]
        run.bold = True

    # Điền dữ liệu vào bảng
    for item in data_list:
        row_cells = table.add_row().cells
        for i, key in enumerate(keys):
            row_cells[i].text = str(item.get(key, ""))

@app.get("/")
def root():
    return {"message": "IDMEA PPAP Generator API is running"}

@app.get("/health")
def health():
    return {"status": "healthy"}
@app.post("/generate-docx")
async def generate_docx(request: DocxRequest):
    try:
        doc = Document()
        
        # 1. Tiêu đề chính
        header = doc.add_heading(f"{request.title} - {request.customer}", 0)
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Giải mã dữ liệu JSON từ trường 'html'
        ppap_data = json.loads(request.html)
        meta = ppap_data.get("Meta", {})

        # 2. Thông tin chung (Meta Section)
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

        # 3. Bảng PFMEA
        if "PFMEA" in ppap_data:
            headers = ["Công đoạn", "Lỗi tiềm ẩn", "Nguyên nhân", "S", "O", "D", "RPN", "Hành động khắc phục"]
            keys = ["Process_step", "Failuere_mode", "Cause", "severity", "occurrence", "detection", "rpn", "recommended_actions"]
            add_table_data(doc, "II. PHÂN TÍCH PFMEA", headers, ppap_data["PFMEA"], keys)

        # 4. Bảng Control Plan
        if "Control_plan" in ppap_data:
            headers = ["Công đoạn", "Đặc tính sản phẩm", "Thông số KT", "Phương pháp đo", "Tần suất", "Biện pháp xử lý"]
            keys = ["Process_step", "product_characteristic", "spec", "measurement_method", "sample_size_freq", "reaction_plan"]
            add_table_data(doc, "III. KẾ HOẠCH KIỂM SOÁT (CONTROL PLAN)", headers, ppap_data["Control_plan"], keys)

        # 5. Risk Summary (Tóm tắt rủi ro)
        if "Final_output" in ppap_data and "risk_summary" in ppap_data["Final_output"]:
            doc.add_heading("IV. TÓM TẮT RỦI RO & PHÒNG NGỪA", level=1)
            for risk in ppap_data["Final_output"]["risk_summary"]:
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(f"{risk['risk']}: ").bold = True
                p.add_run(risk['mitigation'])

        file_path = "Bao_cao_PPAP_Professional.docx"
        doc.save(file_path)
        
        return FileResponse(path=file_path, filename=file_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xử lý dữ liệu: {str(e)}")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
