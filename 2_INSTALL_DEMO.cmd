@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ELKA v8.2 - installing DEMO (operator + viewer)...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup\Install_Operator.ps1" -Mode 2
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup\Install_Viewer.ps1" -Mode 2
echo.
pause
