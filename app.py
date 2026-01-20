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
    html: dict  # Dữ liệu JSON từ Dify
    filename: str = "PPAP_Report.docx"

# Hàm thông minh: Tự động tìm dữ liệu dù key viết hoa hay thường
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
    
    # 1. HEADER CHUYÊN NGHIỆP
    header = doc.sections[0].header
    htable = header.add_table(rows=1, cols=2, width=Inches(6))
    htable.autofit = False
    htable.columns[0].width = Inches(4)
    htable.columns[1].width = Inches(2)
    
    # Logo text bên trái
    cell_left = htable.cell(0, 0)
    run = cell_left.paragraphs[0].add_run("IDMEA TECHNOLOGY")
    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0, 51, 102) # Màu xanh đậm IDMEA
    
    # Tên báo cáo bên phải
    cell_right = htable.cell(0, 1)
    p = cell_right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_r = p.add_run("PPAP REPORT")
    run_r.bold = True
    run_r.font.color.rgb = RGBColor(128, 128, 128)

    # 2. TIÊU ĐỀ CHÍNH
    title = doc.add_heading(f"{request.title} - {request.customer}", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 3. THÔNG TIN CHUNG (META DATA)
    doc.add_heading('I. THÔNG TIN SẢN PHẨM', level=1)
    meta = request.html.get('Meta', {})
    
    table_meta = doc.add_table(rows=3, cols=2)
    table_meta.style = 'Table Grid'
    
    # Dùng hàm thông minh để bắt mọi trường hợp key
    part_name = get_smart_value(meta, ['Part_name', 'part_name', 'Ten_linh_kien'], "N/A")
    part_number = get_smart_value(meta, ['Part_number', 'part_number', 'Ma_hang'], "Reviewing...")
    rev = get_smart_value(meta, ['Revise', 'drawing_rev', 'Rev'], "01")
    
    rows = table_meta.rows
    rows[0].cells[0].text = "Tên linh kiện:"
    rows[0].cells[1].text = part_name
    rows[1].cells[0].text = "Mã số (P/N):"
    rows[1].cells[1].text = part_number
    rows[2].cells[0].text = "Phiên bản (Rev):"
    rows[2].cells[1].text = rev

    doc.add_paragraph("\n")

    # 4. BẢNG PFMEA (Full Cột S-O-D)
    doc.add_heading('II. PHÂN TÍCH RỦI RO (PFMEA)', level=1)
    pfmea_data = request.html.get('PFMEA', [])
    
    if pfmea_data:
        # Tạo bảng 8 cột chuẩn IATF
        headers = ["Công đoạn", "Lỗi tiềm ẩn", "Nguyên nhân", "S", "O", "D", "RPN", "Hành động"]
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        
        # Tô màu header bảng
        hdr_cells = table.rows[0].cells
        for i, text in enumerate(headers):
            hdr_cells[i].text = text
            set_cell_background(hdr_cells[i], "D9EAD3") # Màu xanh nhạt
            hdr_cells[i].paragraphs[0].runs[0].bold = True

        for item in pfmea_data:
            row_cells = table.add_row().cells
            
            # Smart Key Search cho từng cột
            row_cells[0].text = get_smart_value(item, ['Process_step', 'process_step', 'process_step_id'])
            row_cells[1].text = get_smart_value(item, ['Failuere_mode', 'failure_mode', 'Failure_Mode']) # Sửa lỗi trống cột
            row_cells[2].text = get_smart_value(item, ['Cause', 'cause'])
            
            # Các chỉ số S, O, D
            s = get_smart_value(item, ['severity', 'Severity', 'S'], "0")
            o = get_smart_value(item, ['occurrence', 'Occurrence', 'O'], "0")
            d = get_smart_value(item, ['detection', 'Detection', 'D'], "0")
            
            row_cells[3].text = s
            row_cells[4].text = o
            row_cells[5].text = d
            
            # RPN và Tô màu cảnh báo
            rpn = get_smart_value(item, ['rpn', 'RPN'])
            if not rpn or rpn == "0":
                try:
                    rpn = str(int(s) * int(o) * int(d))
                except:
                    rpn = "0"
            
            row_cells[6].text = rpn
            if int(rpn) > 100:
                set_cell_background(row_cells[6], "FFCCCC") # Tô đỏ nếu RPN cao

            row_cells[7].text = get_smart_value(item, ['recommended_actions', 'recommended_action', 'Action'])

    doc.add_paragraph("\n")

    # 5. BẢNG CONTROL PLAN (Mới thêm)
    # ... (Phần trên giữ nguyên)

    # 5. BẢNG CONTROL PLAN (Cập nhật mới nhất)
    doc.add_heading('III. KẾ HOẠCH KIỂM SOÁT (CONTROL PLAN)', level=1)
    cp_data = request.html.get('Control_plan', [])
    
    if cp_data:
        headers_cp = ["Đặc tính", "Thông số (Spec)", "Phương pháp đo", "Tần suất", "Phản ứng"]
        table_cp = doc.add_table(rows=1, cols=len(headers_cp))
        table_cp.style = 'Table Grid'
        
        hdr_cp = table_cp.rows[0].cells
        for i, text in enumerate(headers_cp):
            hdr_cp[i].text = text
            set_cell_background(hdr_cp[i], "CFE2F3") 
            hdr_cp[i].paragraphs[0].runs[0].bold = True
            
        for item in cp_data:
            row = table_cp.add_row().cells
            
            # CẬP NHẬT: Thêm nhiều từ khóa dự phòng để bắt dữ liệu
            # 1. Đặc tính
            row[0].text = get_smart_value(item, [
                'product_characteristic', 'Product_Characteristic', 'characteristic', 'Characteristic', 
                'feature', 'Feature', 'description'
            ])
            
            # 2. Thông số (Spec)
            row[1].text = get_smart_value(item, [
                'spec', 'Spec', 'specification', 'Specification', 
                'tolerance', 'Tolerance', 'standard'
            ])
            
            # 3. Phương pháp đo
            row[2].text = get_smart_value(item, [
                'measurement_method', 'Measurement_Method', 'method', 'Method', 
                'evaluation_measurement_technique', 'technique'
            ])
            
            # 4. Tần suất
            row[3].text = get_smart_value(item, [
                'sample_size_freq', 'Sample_Size_Freq', 'frequency', 'Frequency', 
                'sample_size', 'Sample_Size'
            ])
            
            # 5. Phản ứng
            row[4].text = get_smart_value(item, [
                'reaction_plan', 'Reaction_Plan', 'reaction', 'Reaction', 
                'action', 'Action'
            ])

    # ... (Phần save file giữ nguyên)

    # Stream file về
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    
    return StreamingResponse(
        file_stream, 
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={request.filename}"}
    )
