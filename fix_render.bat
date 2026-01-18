@echo off
cd /d "C:\Users\khanh\OneDrive\Documents\3. PROJECT\AI SOFTWARE\pdf_service"

echo Creating runtime.txt...
echo python-3.10.0 > runtime.txt

echo Creating requirements.txt...
(
echo fastapi==0.104.1
echo uvicorn[standard]==0.24.0
echo pydantic==1.10.13
echo python-multipart==0.0.6
) > requirements.txt

echo Adding files to git...
git add .

echo Committing changes...
git commit -m "FIX: Python 3.10 + Pydantic v1 for Render compatibility"

echo Pushing to GitHub...
git push

pause