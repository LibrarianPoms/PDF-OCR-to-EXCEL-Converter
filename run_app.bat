@echo off
echo Starting PDF OCR to Excel Converter...
echo.
echo Installing required packages...
pip install -r requirements.txt
echo.
echo You may also need to install poppler-utils for PDF conversion.
echo Download from: https://github.com/oschwartz10612/poppler-windows/releases/
echo.
echo Starting Flask application on port 5001...
python app.py
pause
