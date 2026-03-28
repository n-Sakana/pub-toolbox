param(
    [Parameter(Mandatory)][string]$FilePath
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $FilePath)) {
    Write-Host "Error: file not found: $FilePath" -ForegroundColor Red
    exit 1
}

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

# Read entire file as bytes
$data = [IO.File]::ReadAllBytes($FilePath)

# Search for DPB= (ASCII bytes: 0x44 0x50 0x42 0x3D)
$pattern = [System.Text.Encoding]::ASCII.GetBytes('DPB=')
$pos = -1
for ($i = 0; $i -le $data.Length - $pattern.Length; $i++) {
    $match = $true
    for ($j = 0; $j -lt $pattern.Length; $j++) {
        if ($data[$i + $j] -ne $pattern[$j]) { $match = $false; break }
    }
    if ($match) { $pos = $i; break }
}

if ($pos -eq -1) {
    Write-Host ""
    Write-Host "No VBA password hash (DPB=) found in this file." -ForegroundColor Yellow
    Write-Host "The file may not contain a VBA project."
    exit 0
}

# Patch: DPB= -> DPx= (change byte at pos+2 from 0x42 to 0x78)
$data[$pos + 2] = 0x78

[IO.File]::WriteAllBytes($FilePath, $data)

Write-Host ""
Write-Host "VBA password protection disabled." -ForegroundColor Green
Write-Host ""
Write-Host "To fully remove, open the file and:"
Write-Host "  1. Open VBE (Alt+F11)"
Write-Host "  2. Tools > VBAProject Properties > Protection tab"
Write-Host "  3. Clear the password fields and click OK"
Write-Host "  4. Save the file"
