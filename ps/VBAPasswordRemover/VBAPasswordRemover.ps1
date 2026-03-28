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

function Remove-PasswordXls([string]$path) {
    $data = [IO.File]::ReadAllBytes($path)
    $pos = Find-DPB $data
    if ($pos -eq -1) { return $false }
    $data[$pos + 2] = 0x78
    [IO.File]::WriteAllBytes($path, $data)
    return $true
}

Write-Host "Processing: $FilePath"

if ($ext -eq '.xls') {
    $result = Remove-PasswordXls $FilePath
} else {
    # Load assemblies, then use types in a ScriptBlock created AFTER loading
    # so PowerShell can resolve them at parse time of the block
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $result = & {
        param($path)
        $zip = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $entry = $zip.Entries | Where-Object { $_.Name -eq 'vbaProject.bin' } | Select-Object -First 1
            if (-not $entry) { return $false }

            $stream = $entry.Open()
            $ms = New-Object IO.MemoryStream
            $stream.CopyTo($ms)
            $stream.Close()
            $data = $ms.ToArray()
            $ms.Close()

            $pos = Find-DPB $data
            if ($pos -eq -1) { return $false }

            $data[$pos + 2] = 0x78

            $stream = $entry.Open()
            $stream.SetLength(0)
            $stream.Write($data, 0, $data.Length)
            $stream.Close()

            return $true
        }
        finally {
            $zip.Dispose()
        }
    } $FilePath
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
