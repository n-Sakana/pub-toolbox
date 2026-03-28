@echo off
chcp 65001 >nul
setlocal
if "%~1"=="" (
    echo Drop an Excel file to remove VBA project password.
    echo Supported: .xls / .xlsm / .xlam
    pause
    exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\RemovePassword.ps1" "%~1"
pause
