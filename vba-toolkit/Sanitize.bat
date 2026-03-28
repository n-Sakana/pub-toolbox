@echo off
chcp 65001 >nul
setlocal
if "%~1"=="" (
    echo Drop an Excel file to comment out Win32 API declarations.
    echo Supported: .xls / .xlsm / .xlam
    pause
    exit /b 1
)
for %%F in (%*) do (
    echo.
    echo ========================================
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Sanitize.ps1" "%%~F"
)
pause
