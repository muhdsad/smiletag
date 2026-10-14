@echo off
REM Retail Price Tag Generator Launcher for Windows
cd /d "%~dp0"

echo =======================================================
echo  Smile PriceTag Master - Starting...
echo =======================================================

where python >nul 2>nul
if %errorlevel% equ 0 (
    start http://localhost:8080
    python server.py
) else (
    start index.html
)
pause
