$ErrorActionPreference = 'Stop'

$projectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$xlamPath = Join-Path $projectDir 'dist\vba-toolbox.xlam'
$samplePath = Join-Path $projectDir 'sample\vba-toolbox-sample.xlsx'

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    $deadline = (Get-Date).AddSeconds(15)
    while (Get-Process EXCEL -ErrorAction SilentlyContinue) {
        if ((Get-Date) -ge $deadline) {
            Write-Host 'ERROR: Close all Excel windows before running the sample launcher.' -ForegroundColor Red
            exit 1
        }
        Start-Sleep -Milliseconds 500
    }
}

if (-not (Test-Path -LiteralPath $xlamPath)) {
    Write-Host "ERROR: $xlamPath not found. Run Build-Addin.ps1 first." -ForegroundColor Red
    exit 1
}
if (-not (Test-Path -LiteralPath $samplePath)) {
    Write-Host "ERROR: $samplePath not found. Run Build-Sample.ps1 first." -ForegroundColor Red
    exit 1
}

Write-Host 'Starting Excel...' -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$sampleWb = $null
$addin = $null

try {
    Write-Host "  Opening sample: $samplePath"
    $sampleWb = $excel.Workbooks.Open($samplePath)

    Write-Host "  Reloading add-in: $xlamPath"
    foreach ($candidate in @($excel.AddIns)) {
        if ($candidate.FullName -eq $xlamPath) {
            $addin = $candidate
            break
        }
    }
    if ($null -eq $addin) {
        $addin = $excel.AddIns.Add($xlamPath, $false)
    } elseif ($addin.Installed) {
        $addin.Installed = $false
    }
    $addin.Installed = $true

    [void]$sampleWb.Activate()
    Write-Host ''
    Write-Host 'Ready. Use the vba-tools ribbon tab.' -ForegroundColor Green
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
} finally {
    if ($null -ne $sampleWb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sampleWb) }
    if ($null -ne $addin) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($addin) }
    if ($null -ne $excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
