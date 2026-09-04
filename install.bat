@echo off
echo ==============================
echo CBT DOCX TO EXCEL - INSTALLER
echo ==============================

python -m venv .venv

call .venv\Scripts\activate

python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo Install selesai.
echo Jalankan aplikasi dengan:
echo run.bat

pause
