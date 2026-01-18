@echo off
echo ========================================
echo Fixing Git Push Error
echo ========================================

cd /d "C:\Users\khanh\OneDrive\Documents\3. PROJECT\AI SOFTWARE\pdf_service"

echo.
echo 1. Pulling changes from GitHub...
git pull origin main --allow-unrelated-histories --no-edit

if %errorlevel% neq 0 (
    echo.
    echo Conflict detected! Opening merge tool...
    echo Please resolve conflicts manually.
    pause
    exit /b
)

echo.
echo 2. Adding any new files...
git add .

echo.
echo 3. Committing changes...
git commit -m "Merge from GitHub" --no-edit

echo.
echo 4. Pushing to GitHub...
git push -u origin main

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo SUCCESS: Code pushed to GitHub!
    echo ========================================
    echo.
    echo Repository: https://github.com/Tiger-khanhtrc/pdf-docx-service
) else (
    echo.
    echo ========================================
    echo ERROR: Push failed
    echo ========================================
)

pause