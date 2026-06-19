@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv\Scripts\streamlit.exe" (
    echo Dang tao moi truong ao va cai dat thu vien...
    python -m venv .venv
    call ".venv\Scripts\activate.bat"
    pip install -r requirements.txt -q
) else (
    call ".venv\Scripts\activate.bat"
)

echo.
echo Dang mo ung dung So do thuy chuan...
echo Trinh duyet se mo tai: http://localhost:8501
echo Nhan Ctrl+C de dung chuong trinh.
echo.

streamlit run app.py

pause
