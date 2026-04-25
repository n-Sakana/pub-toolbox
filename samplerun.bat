@echo off

echo === Build + Run vba-toolbox (sample) ===
powershell -NoProfile -ExecutionPolicy Bypass -Command "if (Get-Process EXCEL -ErrorAction SilentlyContinue) { Write-Host 'Close all Excel windows before running samplerun.bat.' -ForegroundColor Red; exit 1 }"
if errorlevel 1 (
    pause
    exit /b 1
)

powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Build-Addin.ps1"
if errorlevel 1 (
    echo Build add-in failed.
    pause
    exit /b 1
)

powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Build-Sample.ps1"
if errorlevel 1 (
    echo Build sample failed.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$deadline=(Get-Date).AddSeconds(15); while(Get-Process EXCEL -ErrorAction SilentlyContinue){ if((Get-Date) -ge $deadline){ Write-Host 'Excel did not shut down after build.' -ForegroundColor Red; exit 1 }; Start-Sleep -Milliseconds 500 }"
if errorlevel 1 (
    pause
    exit /b 1
)

echo.
echo Opening sample workbook and installing add-in...
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Run-Sample.ps1"
if errorlevel 1 (
    echo Run sample failed.
    pause
    exit /b 1
)

echo.
echo Ready.
