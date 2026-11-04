@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ELKA v8.2 - updating macros in existing operator workbook (data is kept)...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup\Install_Operator.ps1" -Mode 3
echo.
pause
