@echo off
cd /d "C:\Users\khanh\OneDrive\Documents\3. PROJECT\AI SOFTWARE\pdf_service"

echo Creating correct app.py...
(
echo from fastapi import FastAPI, HTTPException
echo from fastapi.responses import Response
echo from fastapi.middleware.cors import CORSMiddleware
echo from pydantic import BaseModel
echo import uvicorn
echo import os
echo import io
echo from zipfile import ZipFile
echo from datetime import datetime
echo.
echo PORT = int(os.environ.get^("PORT", 8000^)^)
echo.
echo app = FastAPI^(title="DOCX Generator", version="1.0"^)
echo.
echo app.add_middleware^(
echo     CORSMiddleware,
echo     allow_origins=["*"],
echo     allow_credentials=True,
echo     allow_methods=["*"],
echo     allow_headers=["*"],
echo ^)
echo.
echo class DocxRequest^(BaseModel^):
echo     title: str = "DOCUMENT"
echo     content: str = ""
echo     filename: str = "document.docx"
echo.
echo @app.get^("/"^)
echo def root^(^):
echo     return {"status": "ok"}
echo.
echo @app.get^("/health"^)
echo def health^(^):
echo     return {"status": "healthy"}
echo.
echo @app.post^("/generate-docx"^)
echo def generate_docx^(request: DocxRequest^):
echo     return {"message": "API is working", "filename": request.filename}
echo.
echo if __name__ == "__main__":
echo     uvicorn.run^(app, host="0.0.0.0", port=PORT^)
) > app.py

echo Pushing to GitHub...
git add app.py
git commit -m "FIX: Add FastAPI app variable"
git push

echo.
echo ✅ Done! Now on Render:
echo 1. Clear Build Cache
echo 2. Manual Deploy
pause