from fastapi import FastAPI, HTTPException
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn
import os
import io
from zipfile import ZipFile
from datetime import datetime

PORT = int(os.environ.get("PORT", 8000))

app = FastAPI(title="DOCX Generator", version="1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class DocxRequest(BaseModel):
    title: str = "DOCUMENT"
    content: str = ""
    filename: str = "document.docx"

@app.get("/")
def root():
    return {"status": "ok"}

@app.get("/health")
def health():
    return {"status": "healthy"}

@app.post("/generate-docx")
def generate_docx(request: DocxRequest):
    return {"message": "API is working", "filename": request.filename}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
