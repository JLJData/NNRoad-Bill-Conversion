@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM 默认仅本机监听；勿用 0.0.0.0 除非已设 CONVERT_API_KEY 且有防火墙
if not defined CONVERT_HOST set CONVERT_HOST=127.0.0.1
if not defined CONVERT_PORT set CONVERT_PORT=8765

REM 与 Office hrone.bill-convert.api-key 明文一致（开发可用下方默认；生产请改环境变量）
if not defined CONVERT_API_KEY set CONVERT_API_KEY=nd6AReTl-bdt09-f8d_15Bdx65wEps0KuxGJzQ2VppA

echo Starting bill convert service on http://%CONVERT_HOST%:%CONVERT_PORT% ...
echo Auth: CONVERT_API_KEY is set (X-Api-Key required except /health)
python -m uvicorn convert_api:app --host %CONVERT_HOST% --port %CONVERT_PORT%
pause
