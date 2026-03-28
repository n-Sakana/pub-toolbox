@echo off
chcp 65001 >nul
setlocal
if "%~1"=="" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Sanitize.ps1"
    pause
    exit /b
)
for %%F in (%*) do (
    echo.
    echo ========================================
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Sanitize.ps1" "%%~F"
)
pause
