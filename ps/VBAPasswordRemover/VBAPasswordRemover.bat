@echo off
setlocal
if "%~1"=="" (
    echo ファイルをこの BAT にドラッグ＆ドロップしてください。
    echo 対応形式: .xls / .xlsm / .xlam
    pause
    exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0VBAPasswordRemover.ps1" "%~1"
pause
