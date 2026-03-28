param(
    [Parameter(Mandatory)][string]$FilePath
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $FilePath)) {
    Write-Host "Error: file not found: $FilePath" -ForegroundColor Red
    exit 1
}

$FilePath = (Resolve-Path $FilePath).Path
$ext = [IO.Path]::GetExtension($FilePath).ToLower()
if ($ext -notin '.xls', '.xlsm', '.xlam') {
    Write-Host "Error: unsupported format: $ext" -ForegroundColor Red
    Write-Host "Supported: .xls / .xlsm / .xlam"
    exit 1
}

# Backup
$bakPath = "$FilePath.bak"
Copy-Item $FilePath $bakPath -Force
Write-Host "Backup created: $bakPath"
Write-Host "Processing: $FilePath"

function Find-DPB([byte[]]$data) {
    $pattern = [System.Text.Encoding]::ASCII.GetBytes('DPB=')
    for ($i = 0; $i -le $data.Length - $pattern.Length; $i++) {
        $match = $true
        for ($j = 0; $j -lt $pattern.Length; $j++) {
            if ($data[$i + $j] -ne $pattern[$j]) { $match = $false; break }
        }
        if ($match) { return $i }
    }
    return -1
}

function Patch-XlsFile([string]$path) {
    $data = [IO.File]::ReadAllBytes($path)
    $pos = Find-DPB $data
    if ($pos -eq -1) { return $false }
    # DPB= -> DPx=
    $data[$pos + 2] = 0x78
    [IO.File]::WriteAllBytes($path, $data)
    return $true
}

if ($ext -eq '.xls') {
    # OLE2: patch directly
    $result = Patch-XlsFile $FilePath
} else {
    # OOXML: use Excel to convert to .xls, patch, convert back
    $tempXls = Join-Path ([IO.Path]::GetTempPath()) "VBAPwdRemover_$(Get-Date -Format yyyyMMddHHmmss).xls"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    try {
        # Open and save as .xls (OLE2) - xlExcel8 = 56
        $wb = $excel.Workbooks.Open($FilePath, 0, $false)
        $wb.SaveAs($tempXls, 56)
        $wb.Close($false)

        # Patch the .xls binary
        $result = Patch-XlsFile $tempXls

        if ($result) {
            # Open patched .xls and save back as original format
            $wb = $excel.Workbooks.Open($tempXls, 0, $false)
            if ($ext -eq '.xlam') {
                $wb.SaveAs($FilePath, 55)  # xlOpenXMLAddIn
            } else {
                $wb.SaveAs($FilePath, 52)  # xlOpenXMLWorkbookMacroEnabled
            }
            $wb.Close($false)
        }
    }
    finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Remove-Item $tempXls -Force -ErrorAction SilentlyContinue
    }
}

if ($result) {
    Write-Host ""
    Write-Host "VBA password protection disabled." -ForegroundColor Green
    Write-Host ""
    Write-Host "To fully remove, open the file and:"
    Write-Host "  1. Open VBE (Alt+F11)"
    Write-Host "  2. Tools > VBAProject Properties > Protection tab"
    Write-Host "  3. Clear the password fields and click OK"
    Write-Host "  4. Save the file"
} else {
    Write-Host ""
    Write-Host "No VBA password hash (DPB=) found in this file." -ForegroundColor Yellow
    Write-Host "The file may not contain a VBA project."
}
