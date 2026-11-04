@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ELKA v8.2 - installing USER workbook (viewer)...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup\Install_Viewer.ps1"
echo.
pause
