@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Starting bill convert service on http://127.0.0.1:8765 ...
python -m uvicorn convert_api:app --host 0.0.0.0 --port 8765
pause
