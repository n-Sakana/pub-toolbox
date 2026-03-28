@echo off
chcp 65001 >nul
setlocal
if "%~2"=="" (
    echo Usage: Diff.bat file1.xlsm file2.xlsm
    echo Compares VBA code between two Excel files without opening them.
    pause
    exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Diff.ps1" "%~1" "%~2"
pause
