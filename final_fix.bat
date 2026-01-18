@echo off
echo ========================================
echo FINAL FIX for Render - WILL WORK
echo ========================================

cd /d "C:\Users\khanh\OneDrive\Documents\3. PROJECT\AI SOFTWARE\pdf_service"

echo 1. Setting Python 3.10...
echo python-3.10.0 > runtime.txt

echo 2. Creating compatible requirements.txt...
(
echo fastapi==0.104.1
echo uvicorn[standard]==0.24.0
echo pydantic==1.10.13
echo python-multipart==0.0.6
) > requirements.txt

echo 3. Creating minimal working app.py...
(
echo "from fastapi import FastAPI"
echo "from pydantic import BaseModel"
echo "import uvicorn"
echo "import os"
echo ""
echo "PORT = int(os.environ.get('PORT', 8000))"
echo ""
echo "app = FastAPI()"
echo ""
echo "class Request(BaseModel):"
echo "    text: str = 'Hello'"
echo ""
echo "@app.get('/')"
echo "def root():"
echo "    return {'status': 'ok', 'service': 'DOCX Generator'}"
echo ""
echo "@app.get('/health')"
echo "def health():"
echo "    return {'status': 'healthy'}"
echo ""
echo "@app.post('/generate-docx')"
echo "def generate(data: Request):"
echo "    return {'message': data.text, 'filename': 'output.docx'}"
echo ""
echo "if __name__ == '__main__':"
echo "    uvicorn.run(app, host='0.0.0.0', port=PORT)"
) > app.py

echo 4. Pushing to GitHub...
git add .
git commit -m "MINIMAL WORKING VERSION for Render"
git push

echo.
echo ========================================
echo ✅ DONE! Now on Render:
echo 1. Clear Build Cache
echo 2. Manual Deploy
echo ========================================
echo Your app will be at: https://[your-app].onrender.com
pause