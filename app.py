from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
from fastapi.responses import StreamingResponse

app = FastAPI()

class ReportRequest(BaseModel):
    title: str
    customer: str
    html: dict  # Dữ liệu JSON tổng hợp từ Dify
    filename: str = "PPAP_Report_Full.docx"

# --- HÀM THÔNG MINH: Bắt dính dữ liệu dù AI trả về key nào ---
def get_smart_value(item, keys_list, default=""):
    for key in keys_list:
        if key in item and item[key]:
            return str(item[key])
    return default

def set_cell_background(cell, color_hex):
    """Tô màu nền cho ô bảng"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:val'), 'clear')
    shading_elm.set(qn('w:color'), 'auto')
    shading_elm.set(qn('w:fill'), color_hex)
    cell._tc.get_or_add_tcPr().append(shading_elm)

@app.post("/generate-docx")
async def generate_docx(request: ReportRequest):
    doc = Document()
    
    # --- 1. HEADER & LOGO ---
    header = doc.sections[0].header
    htable = header.add_table(rows=1, cols=2, width=Inches(6))
    htable.autofit = False
    htable.columns[0].width = Inches(4)
    htable.columns[1].width = Inches(2)
    
    cell_left = htable.cell(0, 0)
    run = cell_left.paragraphs[0].add_run("IDMEA TECHNOLOGY")
    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0, 51, 102) 
    
    cell_right = htable.cell(0, 1)
    p = cell_right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_r = p.add_run("PPAP REPORT")
    run_r.bold = True
    run_r.font.color.rgb = RGBColor(128, 128, 128)

    # --- 2. TIÊU ĐỀ CHÍNH ---
    title = doc.add_heading(f"{request.title} - {request.customer}", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- 3. THÔNG TIN SẢN PHẨM (META) ---
    doc.add_heading('I. THÔNG TIN SẢN PHẨM', level=1)
    meta = request.html.get('Meta', {})
    
    table_meta = doc.add_table(rows=3, cols=2)
    table_meta.style = 'Table Grid'
    
    part_name = get_smart_value(meta, ['Part_name', 'part_name', 'Ten_linh_kien', 'Part Name'], "N/A")
    part_number = get_smart_value(meta, ['Part_number', 'part_number', 'Ma_hang', 'Part Number'], "Reviewing...")
    rev = get_smart_value(meta, ['Revise', 'drawing_rev', 'Rev', 'Revision'], "01")
    
    rows = table_meta.rows
    rows[0].cells[0].text = "Tên linh kiện:"
    rows[0].cells[1].text = part_name
    rows[1].cells[0].text = "Mã số (P/N):"
    rows[1].cells[1].text = part_number
    rows[2].cells[0].text = "Phiên bản (Rev):"
    rows[2].cells[1].text = rev
    doc.add_paragraph("\n")

    # --- 4. PFMEA ---
    doc.add_heading('II. PHÂN TÍCH RỦI RO (PFMEA)', level=1)
    pfmea_data = request.html.get('PFMEA', [])
    
    if pfmea_data:
        headers = ["Công đoạn", "Lỗi tiềm ẩn", "Nguyên nhân", "S", "O", "D", "RPN", "Hành động"]
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        
        hdr_cells = table.rows[0].cells
        for i, text in enumerate(headers):
            hdr_cells[i].text = text
            set_cell_background(hdr_cells[i], "D9EAD3")
            hdr_cells[i].paragraphs[0].runs[0].bold = True

        for item in pfmea_data:
            row_cells = table.add_row().cells
            row_cells[0].text = get_smart_value(item, ['Process_step', 'process_step', 'Op'])
            row_cells[1].text = get_smart_value(item, ['Failuere_mode', 'failure_mode', 'Failure_Mode']) 
            row_cells[2].text = get_smart_value(item, ['Cause', 'cause', 'Root_Cause'])
            
            s = get_smart_value(item, ['severity', 'S'], "0")
            o = get_smart_value(item, ['occurrence', 'O'], "0")
            d = get_smart_value(item, ['detection', 'D'], "0")
            
            row_cells[3].text = s
            row_cells[4].text = o
            row_cells[5].text = d
            
            rpn = get_smart_value(item, ['rpn', 'RPN'])
            if not rpn or rpn == "0":
                try:
                    rpn = str(int(s) * int(o) * int(d))
                except: rpn = "0"
            
            row_cells[6].text = rpn
            if rpn.isdigit() and int(rpn) > 100:
                set_cell_background(row_cells[6], "FFCCCC")

            row_cells[7].text = get_smart_value(item, ['recommended_actions', 'recommended_action', 'Action'])
    doc.add_paragraph("\n")

    # --- 5. CONTROL PLAN ---
    doc.add_heading('III. KẾ HOẠCH KIỂM SOÁT (CONTROL PLAN)', level=1)
    cp_data = request.html.get('Control_plan', [])
    
    if cp_data:
        # DÒNG NÀY LÀ DÒNG BỊ LỖI LÚC NÃY - ĐÃ SỬA LẠI
        headers_cp = ["Đặc tính", "Thông số (Spec)", "Phương pháp đo", "Tần suất", "Phản ứng"]
        table_cp = doc.add_table(rows=1, cols=len(headers_cp))
        table_cp.style = '
