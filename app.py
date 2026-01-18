from fastapi import FastAPI
from fastapi.responses import Response
from pydantic import BaseModel
from docx import Document
import io

app = FastAPI()

class DocxRequest(BaseModel):
    title: str = "PPAP REPORT"
    customer: str = ""
    html: str | None = None  # nếu bạn muốn gửi HTML/markdown cũng được
    filename: str = "ppap.docx"

@app.post("/generate-docx")
def generate_docx(req: DocxRequest):
    doc = Document()
    doc.add_heading(req.title, level=1)
    doc.add_paragraph(f"Customer: {req.customer}")

    # demo: nhét raw html text cho dễ test (sau mình sẽ hướng dẫn parse html -> docx đẹp)
    if req.html:
        doc.add_paragraph("-----")
        doc.add_paragraph(req.html)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    headers = {
        "Content-Disposition": f'attachment; filename="{req.filename}"'
    }
    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers
    )
