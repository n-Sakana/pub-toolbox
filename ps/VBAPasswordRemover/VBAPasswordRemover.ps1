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

function Patch-DPB([byte[]]$data, [long]$pos) {
    $data[$pos + 2] = 0x78
}

if ($ext -eq '.xls') {
    # OLE2: DPB= is in raw bytes
    $data = [IO.File]::ReadAllBytes($FilePath)
    $pos = Find-DPB $data
    if ($pos -eq -1) {
        Write-Host "`nNo VBA password hash (DPB=) found." -ForegroundColor Yellow
        exit 0
    }
    Patch-DPB $data $pos
    [IO.File]::WriteAllBytes($FilePath, $data)
} else {
    # OOXML: vbaProject.bin is compressed inside ZIP
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    # Use ScriptBlock::Create to defer type resolution past Add-Type
    $block = [ScriptBlock]::Create('
        param($path, $findDPB, $patchDPB)
        $zip = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $entry = $zip.Entries | Where-Object { $_.Name -eq "vbaProject.bin" } | Select-Object -First 1
            if (-not $entry) { return $false }

            $stream = $entry.Open()
            $ms = New-Object IO.MemoryStream
            $stream.CopyTo($ms)
            $stream.Close()
            $data = $ms.ToArray()
            $ms.Close()

            $pos = & $findDPB $data
            if ($pos -eq -1) { return $false }

            & $patchDPB $data $pos

            $stream = $entry.Open()
            $stream.SetLength(0)
            $stream.Write($data, 0, $data.Length)
            $stream.Close()

            return $true
        }
        finally {
            $zip.Dispose()
        }
    ')

    $result = & $block $FilePath ${function:Find-DPB} ${function:Patch-DPB}

    if (-not $result) {
        Write-Host "`nNo VBA password hash (DPB=) found." -ForegroundColor Yellow
        exit 0
    }
}

Write-Host ""
Write-Host "VBA password protection disabled." -ForegroundColor Green
Write-Host ""
Write-Host "To fully remove, open the file and:"
Write-Host "  1. Open VBE (Alt+F11)"
Write-Host "  2. Tools > VBAProject Properties > Protection tab"
Write-Host "  3. Clear the password fields and click OK"
Write-Host "  4. Save the file"
