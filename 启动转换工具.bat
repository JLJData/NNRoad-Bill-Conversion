@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo Starting bill convert tool...

where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "BillConvert" pythonw "%~dp0bill_convert_ui.py"
) else (
    start "BillConvert" python "%~dp0bill_convert_ui.py"
)

echo Window started. Check taskbar for BillConvert.
ping -n 3 127.0.0.1 >nul
