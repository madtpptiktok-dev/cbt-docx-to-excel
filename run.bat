@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Aplikasi belum di-install.
    echo Jalankan install.bat dulu.
    pause
    exit /b
)

call .venv\Scripts\activate
python cbt_docx_to_excel_tool.py

pause
