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
    exit 1
}

# ============================================================================
# OLE2 Compound Document Parser
# ============================================================================

function Read-SectorChain([byte[]]$bytes, [int]$startSector, [int]$sectorSize, [int[]]$fat) {
    $ms = New-Object System.IO.MemoryStream
    $s = $startSector
    $visited = @{}
    while ($s -ge 0 -and $s -ne -2 -and $s -ne -1 -and -not $visited.ContainsKey($s)) {
        $visited[$s] = $true
        $off = ($s + 1) * $sectorSize
        if ($off + $sectorSize -gt $bytes.Length) { break }
        $ms.Write($bytes, $off, $sectorSize)
        if ($s -lt $fat.Length) { $s = $fat[$s] } else { break }
    }
    return ,$ms.ToArray()
}

function Read-MiniStream([byte[]]$miniStreamData, [int]$startSector, [int]$size, [int]$miniSectorSize, [int[]]$miniFat) {
    $data = New-Object byte[] $size
    $s = $startSector
    $written = 0
    while ($s -ge 0 -and $s -ne -2 -and $written -lt $size) {
        $off = $s * $miniSectorSize
        $toRead = [Math]::Min($miniSectorSize, $size - $written)
        if ($off + $toRead -le $miniStreamData.Length) {
            [Array]::Copy($miniStreamData, $off, $data, $written, $toRead)
        }
        $written += $toRead
        if ($s -lt $miniFat.Length) { $s = $miniFat[$s] } else { break }
    }
    return ,$data
}

function Read-Ole2([byte[]]$bytes) {
    $sectorPow = [BitConverter]::ToUInt16($bytes, 30)
    $sectorSize = [int][Math]::Pow(2, $sectorPow)
    $miniSectorPow = [BitConverter]::ToUInt16($bytes, 32)
    $miniSectorSize = [int][Math]::Pow(2, $miniSectorPow)
    $firstDirSector = [BitConverter]::ToInt32($bytes, 48)
    $miniStreamCutoff = [BitConverter]::ToUInt32($bytes, 56)
    $firstMiniFatSector = [BitConverter]::ToInt32($bytes, 60)
    $firstDifatSector = [BitConverter]::ToInt32($bytes, 68)

    $difat = [System.Collections.ArrayList]::new()
    for ($i = 0; $i -lt 109; $i++) {
        $v = [BitConverter]::ToInt32($bytes, 76 + $i * 4)
        if ($v -ge 0) { [void]$difat.Add($v) }
    }
    $nextDifat = $firstDifatSector
    while ($nextDifat -ge 0 -and $nextDifat -ne -2) {
        $off = ($nextDifat + 1) * $sectorSize
        $ePS = $sectorSize / 4 - 1
        for ($i = 0; $i -lt $ePS; $i++) {
            $v = [BitConverter]::ToInt32($bytes, $off + $i * 4)
            if ($v -ge 0) { [void]$difat.Add($v) }
        }
        $nextDifat = [BitConverter]::ToInt32($bytes, $off + $ePS * 4)
    }

    $fatEntries = [int]($bytes.Length / $sectorSize)
    [int[]]$fat = New-Object int[] $fatEntries
    for ($i = 0; $i -lt $fatEntries; $i++) { $fat[$i] = -1 }
    $idx = 0
    foreach ($ds in $difat) {
        $off = ($ds + 1) * $sectorSize
        for ($i = 0; $i -lt ($sectorSize / 4) -and $idx -lt $fatEntries; $i++) {
            $fat[$idx++] = [BitConverter]::ToInt32($bytes, $off + $i * 4)
        }
    }

    $dirData = Read-SectorChain $bytes $firstDirSector $sectorSize $fat
    $entries = [System.Collections.ArrayList]::new()
    for ($i = 0; $i -lt [int]($dirData.Length / 128); $i++) {
        $eOff = $i * 128
        $nameLen = [BitConverter]::ToUInt16($dirData, $eOff + 64)
        $name = ''
        if ($nameLen -gt 2) {
            $name = [System.Text.Encoding]::Unicode.GetString($dirData, $eOff, $nameLen - 2)
        }
        [void]$entries.Add([PSCustomObject]@{
            Name = $name; ObjType = $dirData[$eOff + 66]
            Start = [BitConverter]::ToInt32($dirData, $eOff + 116)
            Size = [BitConverter]::ToUInt32($dirData, $eOff + 120)
        })
    }

    [int[]]$miniFat = @()
    if ($firstMiniFatSector -ge 0 -and $firstMiniFatSector -ne -2) {
        $mfData = Read-SectorChain $bytes $firstMiniFatSector $sectorSize $fat
        $miniFat = New-Object int[] ([int]($mfData.Length / 4))
        for ($i = 0; $i -lt $miniFat.Length; $i++) {
            $miniFat[$i] = [BitConverter]::ToInt32($mfData, $i * 4)
        }
    }

    $rootEntry = $entries | Where-Object { $_.ObjType -eq 5 } | Select-Object -First 1
    [byte[]]$miniStreamData = @()
    if ($rootEntry -and $rootEntry.Start -ge 0) {
        $miniStreamData = Read-SectorChain $bytes $rootEntry.Start $sectorSize $fat
    }

    return @{
        Entries = $entries; Bytes = $bytes; SectorSize = $sectorSize
        MiniSectorSize = $miniSectorSize; MiniStreamCutoff = $miniStreamCutoff
        Fat = $fat; MiniFat = $miniFat; MiniStreamData = $miniStreamData
    }
}

function Read-Ole2Stream($ole2, $entry) {
    if ($entry.Size -lt $ole2.MiniStreamCutoff -and $ole2.MiniStreamData.Length -gt 0) {
        return Read-MiniStream $ole2.MiniStreamData $entry.Start $entry.Size $ole2.MiniSectorSize $ole2.MiniFat
    } else {
        $raw = Read-SectorChain $ole2.Bytes $entry.Start $ole2.SectorSize $ole2.Fat
        if ($raw.Length -gt $entry.Size) {
            $trimmed = New-Object byte[] $entry.Size
            [Array]::Copy($raw, $trimmed, $entry.Size)
            return ,$trimmed
        }
        return ,$raw
    }
}

# ============================================================================
# VBA Decompression (MS-OVBA 2.4.1)
# ============================================================================

function Decompress-VBA([byte[]]$data, [int]$offset) {
    if ($offset -ge $data.Length) { return ,[byte[]]@() }
    if ($data[$offset] -ne 1) { return ,[byte[]]@() }

    $pos = $offset + 1
    $result = New-Object System.IO.MemoryStream

    while ($pos -lt $data.Length - 1) {
        if ($pos + 1 -ge $data.Length) { break }
        $header = [BitConverter]::ToUInt16($data, $pos); $pos += 2
        $chunkSize = ($header -band 0x0FFF) + 3
        $isCompressed = ($header -band 0x8000) -ne 0

        if (-not $isCompressed) {
            $toCopy = [Math]::Min(4096, $data.Length - $pos)
            $result.Write($data, $pos, $toCopy)
            $pos += $toCopy
            continue
        }

        $chunkEnd = $pos + $chunkSize - 2
        if ($chunkEnd -gt $data.Length) { $chunkEnd = $data.Length }
        $decompStart = $result.Length

        while ($pos -lt $chunkEnd) {
            if ($pos -ge $data.Length) { break }
            $flagByte = $data[$pos]; $pos++

            for ($bit = 0; $bit -lt 8 -and $pos -lt $chunkEnd; $bit++) {
                if (($flagByte -band (1 -shl $bit)) -eq 0) {
                    $result.WriteByte($data[$pos]); $pos++
                } else {
                    if ($pos + 1 -ge $data.Length) { $pos = $chunkEnd; break }
                    $token = [BitConverter]::ToUInt16($data, $pos); $pos += 2

                    $dPos = [int]($result.Length - $decompStart)
                    if ($dPos -lt 1) { $dPos = 1 }
                    $bitCount = 4
                    while ((1 -shl $bitCount) -lt $dPos) { $bitCount++ }
                    if ($bitCount -gt 12) { $bitCount = 12 }

                    $lengthMask = 0xFFFF -shr $bitCount
                    $copyLen = ($token -band $lengthMask) + 3
                    $copyOff = ($token -shr (16 - $bitCount)) + 1

                    $buf = $result.ToArray()
                    for ($c = 0; $c -lt $copyLen; $c++) {
                        $srcIdx = $buf.Length - $copyOff
                        if ($srcIdx -ge 0 -and $srcIdx -lt $buf.Length) {
                            $result.WriteByte($buf[$srcIdx])
                            $buf = $result.ToArray()
                        }
                    }
                }
            }
        }
        $pos = $chunkEnd
    }
    return ,$result.ToArray()
}

# ============================================================================
# Main
# ============================================================================

Write-Host "Extracting VBA code from: $FilePath"

# Get vbaProject.bin bytes
if ($ext -eq '.xls') {
    $ole2Bytes = [IO.File]::ReadAllBytes($FilePath)
} else {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $block = [ScriptBlock]::Create('
        param($path)
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            $entry = $zip.Entries | Where-Object { $_.Name -eq "vbaProject.bin" } | Select-Object -First 1
            if (-not $entry) { return $null }
            $s = $entry.Open(); $ms = New-Object IO.MemoryStream; $s.CopyTo($ms); $s.Close()
            return ,$ms.ToArray()
        } finally { $zip.Dispose() }
    ')
    $ole2Bytes = & $block $FilePath
    if (-not $ole2Bytes) {
        Write-Host "No vbaProject.bin found." -ForegroundColor Yellow
        exit 0
    }
}

$ole2 = Read-Ole2 $ole2Bytes

# Read PROJECT stream (plain text with module list)
$projEntry = $ole2.Entries | Where-Object { $_.Name -eq 'PROJECT' -and $_.ObjType -eq 2 } | Select-Object -First 1
if (-not $projEntry) {
    Write-Host "No PROJECT stream found." -ForegroundColor Red
    exit 1
}
$projData = Read-Ole2Stream $ole2 $projEntry
$projText = [System.Text.Encoding]::GetEncoding(932).GetString($projData)

# Parse PROJECT text for module definitions
$modules = [System.Collections.ArrayList]::new()
foreach ($line in $projText -split "`r`n|`n") {
    $line = $line.Trim()
    if ($line -match '^Module=(.+)$') {
        [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'bas' })
    } elseif ($line -match '^Class=(.+)$') {
        [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'cls' })
    } elseif ($line -match '^BaseClass=(.+)$') {
        [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'frm' })
    } elseif ($line -match '^Document=(.+?)/') {
        [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'cls' })
    }
}

if ($modules.Count -eq 0) {
    Write-Host "No VBA modules found in PROJECT stream." -ForegroundColor Yellow
    exit 0
}

# Output directory
$baseName = [IO.Path]::GetFileNameWithoutExtension($FilePath)
$outDir = Join-Path ([IO.Path]::GetDirectoryName($FilePath)) "${baseName}_vba"
if (-not (Test-Path $outDir)) { New-Item $outDir -ItemType Directory | Out-Null }

$extracted = 0
foreach ($mod in $modules) {
    $streamEntry = $ole2.Entries | Where-Object { $_.Name -eq $mod.Name -and $_.ObjType -eq 2 } | Select-Object -First 1
    if (-not $streamEntry) {
        $streamEntry = $ole2.Entries | Where-Object { $_.Name -ieq $mod.Name -and $_.ObjType -eq 2 } | Select-Object -First 1
    }
    if (-not $streamEntry -or $streamEntry.Size -eq 0) {
        Write-Host "  SKIP: $($mod.Name) (stream not found)" -ForegroundColor Yellow
        continue
    }

    $streamData = Read-Ole2Stream $ole2 $streamEntry

    # Find compression signature (0x01) - scan backwards from end
    # The source code starts after the p-code performance cache
    $codeText = $null
    for ($tryOff = $streamData.Length - 2; $tryOff -ge 0; $tryOff--) {
        if ($streamData[$tryOff] -eq 0x01) {
            # Verify: next 2 bytes should be a valid chunk header with signature bits = 011
            if ($tryOff + 2 -lt $streamData.Length) {
                $hdr = [BitConverter]::ToUInt16($streamData, $tryOff + 1)
                $sig = ($hdr -shr 12) -band 0x07
                if ($sig -eq 3) {
                    $code = Decompress-VBA $streamData $tryOff
                    if ($code.Length -gt 0) {
                        $text = [System.Text.Encoding]::GetEncoding(932).GetString($code)
                        if ($text -match 'Attribute\s+VB_Name') {
                            $codeText = $text
                            break
                        }
                    }
                }
            }
        }
    }

    if (-not $codeText) {
        Write-Host "  SKIP: $($mod.Name) (could not decompress)" -ForegroundColor Yellow
        continue
    }

    $outPath = Join-Path $outDir "$($mod.Name).$($mod.Ext)"
    [IO.File]::WriteAllText($outPath, $codeText, [System.Text.Encoding]::UTF8)
    Write-Host "  $($mod.Name).$($mod.Ext)" -ForegroundColor Green
    $extracted++
}

Write-Host ""
if ($extracted -eq 0) {
    Write-Host "No VBA modules extracted." -ForegroundColor Yellow
    exit 0
}

Write-Host "$extracted module(s) extracted to: $outDir" -ForegroundColor Green

# ============================================================================
# Code Analysis Report
# ============================================================================

Write-Host ""
Write-Host "=== Code Analysis ===" -ForegroundColor Cyan

$report = [System.Text.StringBuilder]::new()
[void]$report.AppendLine("# VBA Code Analysis Report")
[void]$report.AppendLine("# Source: $([IO.Path]::GetFileName($FilePath))")
[void]$report.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$report.AppendLine("")

$allFiles = Get-ChildItem $outDir -File
$totalLines = 0

# --- Module summary ---
[void]$report.AppendLine("## Modules ($($allFiles.Count))")
[void]$report.AppendLine("")
foreach ($f in $allFiles) {
    $lines = (Get-Content $f.FullName -Encoding UTF8).Count
    $totalLines += $lines
    [void]$report.AppendLine("  $($f.Name) ($lines lines)")
}
[void]$report.AppendLine("")
[void]$report.AppendLine("  Total: $totalLines lines")
[void]$report.AppendLine("")

# --- Analysis patterns ---
$patterns = [ordered]@{
    'Win32 API (Declare)' = @{
        Pattern = '(?m)^[^'']*\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)'
        Extract = { param($m) "$($m.Groups[3].Value) ($(if($m.Groups[1].Value){'PtrSafe'} else {'Legacy'}))" }
    }
    'COM / CreateObject' = @{
        Pattern = '(?m)^[^'']*\bCreateObject\s*\(\s*"([^"]+)"'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'COM / GetObject' = @{
        Pattern = '(?m)^[^'']*\bGetObject\s*\(\s*"?([^")\s]+)"?'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Shell / process execution' = @{
        Pattern = '(?m)^[^'']*\b(Shell\s*[\("]|WScript\.Shell|cmd\s*/[ck])'
        Extract = { param($m) $m.Groups[1].Value.Trim() }
    }
    'File I/O (Open/Kill/FileCopy)' = @{
        Pattern = '(?m)^[^'']*\b(Open\s+\S+\s+For\s+(Input|Output|Append|Binary|Random)|Kill\s|FileCopy\s|MkDir\s|RmDir\s)'
        Extract = { param($m) if ($m.Groups[2].Value) { "Open For $($m.Groups[2].Value)" } else { $m.Groups[1].Value.Trim() } }
    }
    'FileSystemObject' = @{
        Pattern = '(?m)^[^'']*\b(Scripting\.FileSystemObject)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Registry access' = @{
        Pattern = '(?m)^[^'']*\b(GetSetting|SaveSetting|DeleteSetting|RegRead|RegWrite|RegDelete)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'SendKeys' = @{
        Pattern = '(?m)^[^'']*\b(SendKeys)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Network / HTTP' = @{
        Pattern = '(?m)^[^'']*\b(MSXML2\.XMLHTTP|WinHttp\.WinHttpRequest|Inet|URLDownloadToFile|InternetOpen|HttpSendRequest|MSXML2\.ServerXMLHTTP)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'PowerShell / WScript' = @{
        Pattern = '(?m)^[^'']*\b(powershell|wscript|cscript|mshta)\b'
        Extract = { param($m) $m.Groups[1].Value }
        Flags = 'IgnoreCase'
    }
    'Process / WMI' = @{
        Pattern = '(?m)^[^'']*\b(winmgmts|Win32_Process|WbemScripting|ExecQuery)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Clipboard' = @{
        Pattern = '(?m)^[^'']*\b(MSForms\.DataObject|GetClipboardData|SetClipboardData|OpenClipboard)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Environment variables' = @{
        Pattern = '(?m)^[^'']*\b(Environ\s*\$?\s*\()'
        Extract = { param($m) "Environ" }
    }
    'Macro auto-execution' = @{
        Pattern = '(?m)^\s*(Sub\s+(Auto_Open|Auto_Close|Workbook_Open|Workbook_BeforeClose|Document_Open|Document_Close)\b)'
        Extract = { param($m) $m.Groups[2].Value }
    }
    'DLL loading' = @{
        Pattern = '(?m)^[^'']*\b(LoadLibrary|GetProcAddress|FreeLibrary|CallByName)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Encoding / obfuscation' = @{
        Pattern = '(?m)^[^'']*\b(Chr\s*\$?\s*\(\s*\d+\s*\)|ChrW?\s*\$?\s*\(\s*\d+\s*\))'
        Extract = { param($m) $m.Groups[1].Value }
        Aggregate = $true
    }
    'Late-bound object calls' = @{
        Pattern = '(?m)^[^'']*\bSet\s+\w+\s*=\s*CreateObject\s*\(\s*"([^"]+)"'
        Extract = { param($m) $m.Groups[1].Value }
        Skip = $true
    }
}

$issueCount = 0
foreach ($category in $patterns.Keys) {
    if ($patterns[$category].Skip) { continue }
    $p = $patterns[$category].Pattern
    $extractFn = $patterns[$category].Extract
    $findings = [System.Collections.ArrayList]::new()

    foreach ($f in $allFiles) {
        $content = [IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
        $matches2 = [regex]::Matches($content, $p)
        foreach ($m in $matches2) {
            $detail = & $extractFn $m
            [void]$findings.Add("$($f.Name): $detail")
        }
    }

    if ($findings.Count -gt 0) {
        $issueCount += $findings.Count
        [void]$report.AppendLine("## $category ($($findings.Count))")
        [void]$report.AppendLine("")
        if ($patterns[$category].Aggregate) {
            # Group by file, show count only
            $grouped = $findings | Group-Object { $_ -replace ':.*', '' }
            foreach ($g in $grouped) {
                [void]$report.AppendLine("  $($g.Name): $($g.Count) occurrence(s)")
            }
        } else {
            $unique = $findings | Sort-Object -Unique
            foreach ($item in $unique) {
                [void]$report.AppendLine("  $item")
            }
        }
        [void]$report.AppendLine("")
    }
}

# --- External references from PROJECT stream ---
$projEntry2 = $ole2.Entries | Where-Object { $_.Name -eq 'PROJECT' -and $_.ObjType -eq 2 } | Select-Object -First 1
if ($projEntry2) {
    $projData2 = Read-Ole2Stream $ole2 $projEntry2
    $projText2 = [System.Text.Encoding]::GetEncoding(932).GetString($projData2)
    $refs = [System.Collections.ArrayList]::new()
    foreach ($line in $projText2 -split "`r`n|`n") {
        if ($line -match '^Reference=') {
            # Format: Reference=*\G{GUID}#ver#0#path#description
            if ($line -match '#([^#]+)$') {
                [void]$refs.Add($Matches[1])
            } else {
                [void]$refs.Add($line.Substring(10))
            }
        }
    }
    if ($refs.Count -gt 0) {
        [void]$report.AppendLine("## External References ($($refs.Count))")
        [void]$report.AppendLine("")
        foreach ($r in $refs) {
            [void]$report.AppendLine("  $r")
        }
        [void]$report.AppendLine("")
    }
}

# --- Summary ---
if ($issueCount -eq 0) {
    [void]$report.AppendLine("## Result")
    [void]$report.AppendLine("")
    [void]$report.AppendLine("  No Win32 API, COM, Shell, or other external dependencies detected.")
    [void]$report.AppendLine("  Migration risk: LOW")
} else {
    [void]$report.AppendLine("## Summary")
    [void]$report.AppendLine("")
    [void]$report.AppendLine("  $issueCount potential migration issue(s) detected.")
    [void]$report.AppendLine("  Review items above before deploying to restricted environments.")
}

$reportText = $report.ToString()
$reportPath = Join-Path $outDir "_analysis.txt"
[IO.File]::WriteAllText($reportPath, $reportText, [System.Text.Encoding]::UTF8)

# Print to console
Write-Host ""
Write-Host $reportText
Write-Host "Report saved to: $reportPath" -ForegroundColor Green
