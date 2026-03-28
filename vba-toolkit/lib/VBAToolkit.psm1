# ============================================================================
# VBAToolkit - Common module for binary-level VBA file manipulation
# OLE2 parser, VBA compression/decompression (MS-OVBA 2.4.1), ZIP helpers
# ============================================================================

$ErrorActionPreference = 'Stop'

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
            DirOffset = $eOff
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
        FirstDirSector = $firstDirSector
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

function Write-Ole2Stream([byte[]]$ole2Bytes, $ole2, $entry, [byte[]]$newData) {
    $sectorSize = $ole2.SectorSize
    $fat = $ole2.Fat

    if ($entry.Size -lt $ole2.MiniStreamCutoff) {
        $miniSectorSize = $ole2.MiniSectorSize
        $miniFat = $ole2.MiniFat
        $s = $entry.Start; $written = 0
        while ($s -ge 0 -and $s -ne -2 -and $written -lt $newData.Length) {
            $off = $s * $miniSectorSize
            $toWrite = [Math]::Min($miniSectorSize, $newData.Length - $written)
            [Array]::Copy($newData, $written, $ole2.MiniStreamData, $off, $toWrite)
            if ($toWrite -lt $miniSectorSize) {
                for ($p = $toWrite; $p -lt $miniSectorSize; $p++) { $ole2.MiniStreamData[$off + $p] = 0 }
            }
            $written += $miniSectorSize
            if ($s -lt $miniFat.Length) { $s = $miniFat[$s] } else { break }
        }
        $rootEntry = $ole2.Entries | Where-Object { $_.ObjType -eq 5 } | Select-Object -First 1
        $s2 = $rootEntry.Start; $written2 = 0; $visited = @{}
        while ($s2 -ge 0 -and $s2 -ne -2 -and -not $visited.ContainsKey($s2) -and $written2 -lt $ole2.MiniStreamData.Length) {
            $visited[$s2] = $true
            $off2 = ($s2 + 1) * $sectorSize
            [Array]::Copy($ole2.MiniStreamData, $written2, $ole2Bytes, $off2, [Math]::Min($sectorSize, $ole2.MiniStreamData.Length - $written2))
            $written2 += $sectorSize; $s2 = $fat[$s2]
        }
    } else {
        $s = $entry.Start; $written = 0; $visited = @{}
        while ($s -ge 0 -and $s -ne -2 -and -not $visited.ContainsKey($s) -and $written -lt $newData.Length) {
            $visited[$s] = $true
            $off = ($s + 1) * $sectorSize
            $toWrite = [Math]::Min($sectorSize, $newData.Length - $written)
            [Array]::Copy($newData, $written, $ole2Bytes, $off, $toWrite)
            if ($toWrite -lt $sectorSize) {
                for ($p = $toWrite; $p -lt $sectorSize; $p++) { $ole2Bytes[$off + $p] = 0 }
            }
            $written += $sectorSize; $s = $fat[$s]
        }
    }

    # Update size in directory
    $dirSectorData = Read-SectorChain $ole2Bytes $ole2.FirstDirSector $sectorSize $fat
    [Array]::Copy([BitConverter]::GetBytes([uint32]$newData.Length), 0, $dirSectorData, $entry.DirOffset + 120, 4)
    $s3 = $ole2.FirstDirSector; $written3 = 0; $visited3 = @{}
    while ($s3 -ge 0 -and $s3 -ne -2 -and -not $visited3.ContainsKey($s3)) {
        $visited3[$s3] = $true
        [Array]::Copy($dirSectorData, $written3, $ole2Bytes, ($s3 + 1) * $sectorSize, [Math]::Min($sectorSize, $dirSectorData.Length - $written3))
        $written3 += $sectorSize; $s3 = $fat[$s3]
    }
}

# ============================================================================
# VBA Decompression (MS-OVBA 2.4.1)
# ============================================================================

function Decompress-VBA([byte[]]$data, [int]$offset) {
    if ($offset -ge $data.Length -or $data[$offset] -ne 1) { return ,[byte[]]@() }
    $pos = $offset + 1
    $result = New-Object System.IO.MemoryStream
    while ($pos -lt $data.Length - 1) {
        if ($pos + 1 -ge $data.Length) { break }
        $header = [BitConverter]::ToUInt16($data, $pos); $pos += 2
        $chunkSize = ($header -band 0x0FFF) + 3
        $isCompressed = ($header -band 0x8000) -ne 0
        if (-not $isCompressed) {
            $toCopy = [Math]::Min(4096, $data.Length - $pos)
            $result.Write($data, $pos, $toCopy); $pos += $toCopy; continue
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
                            $result.WriteByte($buf[$srcIdx]); $buf = $result.ToArray()
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
# VBA Compression (MS-OVBA 2.4.1)
# ============================================================================

function Compress-VBA([byte[]]$data) {
    $result = New-Object System.IO.MemoryStream
    $result.WriteByte(0x01)
    $srcPos = 0
    while ($srcPos -lt $data.Length) {
        $chunkStart = $srcPos
        $chunkEnd = [Math]::Min($srcPos + 4096, $data.Length)
        $chunkBuf = New-Object System.IO.MemoryStream
        $dPos = $srcPos
        while ($dPos -lt $chunkEnd) {
            $flagPos = $chunkBuf.Position
            $chunkBuf.WriteByte(0)
            $flagByte = 0
            for ($bit = 0; $bit -lt 8 -and $dPos -lt $chunkEnd; $bit++) {
                $bestLen = 0; $bestOff = 0
                $decompPos = $dPos - $chunkStart
                if ($decompPos -lt 1) { $decompPos = 1 }
                $bitCount = 4
                while ((1 -shl $bitCount) -lt $decompPos) { $bitCount++ }
                if ($bitCount -gt 12) { $bitCount = 12 }
                $maxOff = (1 -shl $bitCount)
                $maxLen = (0xFFFF -shr $bitCount) + 3
                for ($off = 1; $off -le [Math]::Min($maxOff, $dPos - $chunkStart); $off++) {
                    $matchLen = 0
                    while ($matchLen -lt $maxLen -and ($dPos + $matchLen) -lt $chunkEnd) {
                        if ($data[$dPos - $off + ($matchLen % $off)] -ne $data[$dPos + $matchLen]) { break }
                        $matchLen++
                    }
                    if ($matchLen -ge 3 -and $matchLen -gt $bestLen) { $bestLen = $matchLen; $bestOff = $off }
                }
                if ($bestLen -ge 3) {
                    $flagByte = $flagByte -bor (1 -shl $bit)
                    $token = (($bestOff - 1) -shl (16 - $bitCount)) -bor ($bestLen - 3)
                    $chunkBuf.WriteByte([byte]($token -band 0xFF))
                    $chunkBuf.WriteByte([byte](($token -shr 8) -band 0xFF))
                    $dPos += $bestLen
                } else {
                    $chunkBuf.WriteByte($data[$dPos]); $dPos++
                }
            }
            $savedPos = $chunkBuf.Position
            $chunkBuf.Position = $flagPos
            $chunkBuf.WriteByte($flagByte)
            $chunkBuf.Position = $savedPos
        }
        $compressed = $chunkBuf.ToArray()
        $srcPos = $dPos
        if ($compressed.Length -lt 4096) {
            $hdr = [uint16](0x8000 -bor 0x3000 -bor ($compressed.Length + 2 - 3))
            $result.WriteByte([byte]($hdr -band 0xFF))
            $result.WriteByte([byte](($hdr -shr 8) -band 0xFF))
            $result.Write($compressed, 0, $compressed.Length)
        } else {
            $hdr = [uint16](0x3000 -bor (4096 + 2 - 3))
            $result.WriteByte([byte]($hdr -band 0xFF))
            $result.WriteByte([byte](($hdr -shr 8) -band 0xFF))
            $result.Write($data, $chunkStart, $chunkEnd - $chunkStart)
            if ($chunkEnd - $chunkStart -lt 4096) {
                $pad = New-Object byte[] (4096 - ($chunkEnd - $chunkStart))
                $result.Write($pad, 0, $pad.Length)
            }
        }
    }
    return ,$result.ToArray()
}

# ============================================================================
# High-level helpers
# ============================================================================

function Get-VbaProjectBytes([string]$filePath) {
    $ext = [IO.Path]::GetExtension($filePath).ToLower()
    if ($ext -eq '.xls') {
        return @{ Bytes = [IO.File]::ReadAllBytes($filePath); IsZip = $false }
    }
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
    $bytes = & $block $filePath
    return @{ Bytes = $bytes; IsZip = $true }
}

function Save-VbaProjectBytes([string]$filePath, [byte[]]$ole2Bytes, [bool]$isZip) {
    if ($isZip) {
        $block = [ScriptBlock]::Create('
            param($path, $data)
            $zip = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Update)
            try {
                $entry = $zip.Entries | Where-Object { $_.Name -eq "vbaProject.bin" } | Select-Object -First 1
                $stream = $entry.Open(); $stream.SetLength(0)
                $stream.Write($data, 0, $data.Length); $stream.Close()
            } finally { $zip.Dispose() }
        ')
        & $block $filePath $ole2Bytes
    } else {
        [IO.File]::WriteAllBytes($filePath, $ole2Bytes)
    }
}

function Get-VbaModuleList($ole2) {
    $projEntry = $ole2.Entries | Where-Object { $_.Name -eq 'PROJECT' -and $_.ObjType -eq 2 } | Select-Object -First 1
    if (-not $projEntry) { return @() }
    $projData = Read-Ole2Stream $ole2 $projEntry
    $projText = [System.Text.Encoding]::GetEncoding(932).GetString($projData)
    $modules = [System.Collections.ArrayList]::new()
    foreach ($line in $projText -split "`r`n|`n") {
        if ($line -match '^Module=(.+)$') { [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'bas' }) }
        elseif ($line -match '^Class=(.+)$') { [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'cls' }) }
        elseif ($line -match '^BaseClass=(.+)$') { [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'frm' }) }
        elseif ($line -match '^Document=(.+?)/') { [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'cls' }) }
    }
    return ,$modules
}

function Get-VbaModuleCode($ole2, [string]$moduleName) {
    $streamEntry = $ole2.Entries | Where-Object { $_.Name -eq $moduleName -and $_.ObjType -eq 2 } | Select-Object -First 1
    if (-not $streamEntry) {
        $streamEntry = $ole2.Entries | Where-Object { $_.Name -ieq $moduleName -and $_.ObjType -eq 2 } | Select-Object -First 1
    }
    if (-not $streamEntry -or $streamEntry.Size -eq 0) { return $null }
    $streamData = Read-Ole2Stream $ole2 $streamEntry
    for ($tryOff = $streamData.Length - 2; $tryOff -ge 0; $tryOff--) {
        if ($streamData[$tryOff] -eq 0x01 -and $tryOff + 2 -lt $streamData.Length) {
            $hdr = [BitConverter]::ToUInt16($streamData, $tryOff + 1)
            if ((($hdr -shr 12) -band 0x07) -eq 3) {
                $code = Decompress-VBA $streamData $tryOff
                if ($code.Length -gt 0) {
                    $text = [System.Text.Encoding]::GetEncoding(932).GetString($code)
                    if ($text -match 'Attribute\s+VB_Name') {
                        return @{ Code = $text; Offset = $tryOff; Entry = $streamEntry; StreamData = $streamData }
                    }
                }
            }
        }
    }
    return $null
}

# ============================================================================
# HTML Code Viewer (shared by Extract, Sanitize)
# ============================================================================

# $moduleData: ordered hashtable of name -> @{ Ext; Lines = string[]; Highlights = @{ lineIndex -> cssClass } }
# $highlightLabel: e.g. "EDR Detection" or "Sanitized"
function New-HtmlCodeView {
    param(
        [string]$title,
        [string]$subtitle,
        [System.Collections.Specialized.OrderedDictionary]$moduleData,
        [string]$highlightClass,   # CSS class for highlighted lines (e.g. 'hl-edr', 'hl-sanitized')
        [string]$highlightColor,   # CSS color (e.g. '#1b3a5c' for blue, '#4b3a00' for yellow)
        [string]$highlightText,    # CSS text color
        [string]$markerColor,      # minimap marker color
        [string]$outputPath
    )

    $he = { param($s) [System.Net.WebUtility]::HtmlEncode($s) }

    $html = [System.Text.StringBuilder]::new()
    [void]$html.Append(@"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>$(& $he $title)</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: Consolas, 'Courier New', monospace; font-size: 13px; background: #1e1e1e; color: #d4d4d4; }
.header { background: #252526; padding: 10px 20px; border-bottom: 1px solid #3c3c3c; }
.header h1 { font-size: 15px; font-weight: normal; color: #cccccc; }
.header .sub { margin-top: 4px; font-size: 12px; color: #888; }
.main { display: flex; height: calc(100vh - 52px); }
.sidebar { width: 200px; min-width: 200px; background: #252526; border-right: 1px solid #3c3c3c; overflow-y: auto; padding: 8px 0; }
.sidebar .item { padding: 5px 16px; cursor: pointer; color: #888; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sidebar .item:hover { color: #d4d4d4; background: #2a2d2e; }
.sidebar .item.active { color: #ffffff; background: #37373d; border-left: 2px solid #0078d4; }
.sidebar .item.has-hl { color: $markerColor; }
.sidebar .item.no-hl { color: #606060; }
.content { flex: 1; overflow: auto; position: relative; }
.module { display: none; }
.module.active { display: block; }
.code-table { width: 100%; border-collapse: collapse; }
.code-table td { padding: 0 8px; line-height: 20px; vertical-align: top; white-space: pre; overflow: hidden; text-overflow: ellipsis; }
.code-table .ln { width: 50px; min-width: 50px; text-align: right; color: #606060; padding-right: 12px; user-select: none; border-right: 1px solid #3c3c3c; }
.code-table .code { color: #d4d4d4; }
tr.$highlightClass td.code { background: $highlightColor; color: $highlightText; }
tr.$highlightClass td.ln { color: #cccccc; }
.minimap { position: fixed; right: 0; top: 52px; width: 14px; bottom: 0; background: #1e1e1e; border-left: 1px solid #3c3c3c; z-index: 20; cursor: pointer; }
.minimap .mark { position: absolute; right: 2px; width: 10px; height: 3px; border-radius: 1px; background: $markerColor; }
.minimap .viewport { position: absolute; right: 0; width: 14px; background: rgba(255,255,255,0.25); border-radius: 2px; pointer-events: none; }
</style>
</head>
<body>
<div class="header">
  <h1>$(& $he $title)</h1>
  <div class="sub">$(& $he $subtitle)</div>
</div>
<div class="main">
<div class="sidebar" id="sidebar">
"@)

    $tabIdx = 0; $firstHlIdx = -1
    foreach ($modName in $moduleData.Keys) {
        $md = $moduleData[$modName]
        $hlCount = 0
        if ($md.Highlights) { $hlCount = $md.Highlights.Count }
        $cls = if ($hlCount -gt 0) { 'has-hl' } else { 'no-hl' }
        if ($firstHlIdx -eq -1 -and $hlCount -gt 0) { $firstHlIdx = $tabIdx }
        $label = "$modName.$($md.Ext)"
        if ($hlCount -gt 0) { $label += " ($hlCount)" }
        [void]$html.Append("<div class=`"item $cls`" onclick=`"showTab($tabIdx)`" id=`"tab$tabIdx`">$(& $he $label)</div>")
        $tabIdx++
    }
    if ($firstHlIdx -eq -1) { $firstHlIdx = 0 }

    [void]$html.Append("</div><div class=`"content`">")

    $tabIdx = 0
    foreach ($modName in $moduleData.Keys) {
        $md = $moduleData[$modName]
        [void]$html.Append("<div class=`"module`" id=`"mod$tabIdx`"><table class=`"code-table`">")
        for ($i = 0; $i -lt $md.Lines.Count; $i++) {
            $trClass = ''
            if ($md.Highlights -and $md.Highlights.ContainsKey($i)) { $trClass = $highlightClass }
            $ln = $i + 1
            $code = & $he $md.Lines[$i]
            [void]$html.Append("<tr class=`"$trClass`"><td class=`"ln`">$ln</td><td class=`"code`">$code</td></tr>")
        }
        [void]$html.Append("</table></div>")
        $tabIdx++
    }

    [void]$html.Append(@"
<div class="minimap" id="minimap"><div class="viewport" id="viewport"></div></div>
</div></div>
<script>
const content = document.querySelector('.content');
const minimap = document.getElementById('minimap');
const viewport = document.getElementById('viewport');
function showTab(idx) {
  document.querySelectorAll('.module').forEach(m => m.classList.remove('active'));
  document.querySelectorAll('.item').forEach(t => t.classList.remove('active'));
  document.getElementById('mod' + idx).classList.add('active');
  document.getElementById('tab' + idx).classList.add('active');
  updateMinimap();
}
function updateMinimap() {
  minimap.querySelectorAll('.mark').forEach(m => m.remove());
  const mod = document.querySelector('.module.active');
  if (!mod) return;
  const rows = mod.querySelectorAll('tr.$highlightClass');
  const allRows = mod.querySelectorAll('tr');
  if (allRows.length === 0) return;
  const mapH = minimap.clientHeight;
  rows.forEach(r => {
    const idx = Array.from(allRows).indexOf(r);
    const mark = document.createElement('div');
    mark.className = 'mark';
    mark.style.top = (idx / allRows.length * mapH) + 'px';
    mark.addEventListener('click', () => r.scrollIntoView({block:'center'}));
    minimap.appendChild(mark);
  });
  updateViewport();
}
function updateViewport() {
  const sh = content.scrollHeight, ch = content.clientHeight, st = content.scrollTop;
  const mapH = minimap.clientHeight;
  if (sh <= ch) { viewport.style.display = 'none'; return; }
  viewport.style.display = '';
  viewport.style.top = (st / sh * mapH) + 'px';
  viewport.style.height = (ch / sh * mapH) + 'px';
}
content.addEventListener('scroll', updateViewport);
minimap.addEventListener('click', (e) => {
  if (e.target.classList.contains('mark')) return;
  content.scrollTop = e.offsetY / minimap.clientHeight * content.scrollHeight - content.clientHeight / 2;
});
showTab($firstHlIdx);
</script>
</body></html>
"@)

    [IO.File]::WriteAllText($outputPath, $html.ToString(), [System.Text.Encoding]::UTF8)
}

Export-ModuleMember -Function Read-Ole2, Read-Ole2Stream, Write-Ole2Stream,
    Decompress-VBA, Compress-VBA,
    Get-VbaProjectBytes, Save-VbaProjectBytes,
    Get-VbaModuleList, Get-VbaModuleCode,
    New-HtmlCodeView
