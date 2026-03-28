@echo off
chcp 65001 >nul
setlocal
if "%~1"=="" (
    echo Drop an Excel file to remove VBA project password.
    echo Supported: .xls / .xlsm / .xlam
    pause
    exit /b 1
)
for %%F in (%*) do (
    echo.
    echo ========================================
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Unlock.ps1" "%%~F"
)
pause
