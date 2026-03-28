# ============================================================================
# VBAToolkit - Common module for binary-level VBA file manipulation
# OLE2 parser, VBA compression/decompression (MS-OVBA 2.4.1), ZIP helpers
# ============================================================================

$ErrorActionPreference = 'Stop'

# ============================================================================
# C# Native Implementation (high-performance byte operations)
# ============================================================================

if (-not ([System.Management.Automation.PSTypeName]'VbaToolkitNative').Type) {
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Collections.Generic;

public static class VbaToolkitNative
{
    public static byte[] ReadSectorChain(byte[] bytes, int startSector, int sectorSize, int[] fat)
    {
        var ms = new MemoryStream();
        var visited = new HashSet<int>();
        int s = startSector;
        while (s >= 0 && s != -2 && s != -1 && !visited.Contains(s))
        {
            visited.Add(s);
            int off = (s + 1) * sectorSize;
            if (off + sectorSize > bytes.Length) break;
            ms.Write(bytes, off, sectorSize);
            s = (s < fat.Length) ? fat[s] : -1;
        }
        return ms.ToArray();
    }

    public static byte[] ReadMiniStream(byte[] miniStreamData, int startSector, int size, int miniSectorSize, int[] miniFat)
    {
        var data = new byte[size];
        int s = startSector;
        int written = 0;
        while (s >= 0 && s != -2 && written < size)
        {
            int off = s * miniSectorSize;
            int toRead = Math.Min(miniSectorSize, size - written);
            if (off + toRead <= miniStreamData.Length)
                Array.Copy(miniStreamData, off, data, written, toRead);
            written += miniSectorSize;
            s = (s < miniFat.Length) ? miniFat[s] : -1;
        }
        return data;
    }

    public static byte[] DecompressVba(byte[] data, int offset)
    {
        if (offset >= data.Length || data[offset] != 1) return new byte[0];
        var result = new List<byte>(data.Length * 2);
        int pos = offset + 1;
        while (pos < data.Length - 1)
        {
            if (pos + 1 >= data.Length) break;
            ushort header = BitConverter.ToUInt16(data, pos); pos += 2;
            int chunkSize = (header & 0x0FFF) + 3;
            bool isCompressed = (header & 0x8000) != 0;
            if (!isCompressed)
            {
                int toCopy = Math.Min(4096, data.Length - pos);
                for (int c = 0; c < toCopy; c++) result.Add(data[pos + c]);
                pos += toCopy;
                continue;
            }
            int chunkEnd = pos + chunkSize - 2;
            if (chunkEnd > data.Length) chunkEnd = data.Length;
            int decompStart = result.Count;
            while (pos < chunkEnd)
            {
                if (pos >= data.Length) break;
                byte flagByte = data[pos]; pos++;
                for (int bit = 0; bit < 8 && pos < chunkEnd; bit++)
                {
                    if ((flagByte & (1 << bit)) == 0)
                    {
                        result.Add(data[pos]); pos++;
                    }
                    else
                    {
                        if (pos + 1 >= data.Length) { pos = chunkEnd; break; }
                        ushort token = BitConverter.ToUInt16(data, pos); pos += 2;
                        int dPos = result.Count - decompStart;
                        if (dPos < 1) dPos = 1;
                        int bitCount = 4;
                        while ((1 << bitCount) < dPos) bitCount++;
                        if (bitCount > 12) bitCount = 12;
                        int lengthMask = 0xFFFF >> bitCount;
                        int copyLen = (token & lengthMask) + 3;
                        int copyOff = (token >> (16 - bitCount)) + 1;
                        for (int c = 0; c < copyLen; c++)
                        {
                            int srcIdx = result.Count - copyOff;
                            if (srcIdx >= 0 && srcIdx < result.Count)
                                result.Add(result[srcIdx]);
                        }
                    }
                }
            }
            pos = chunkEnd;
        }
        return result.ToArray();
    }

    public static byte[] CompressVba(byte[] data)
    {
        var result = new MemoryStream();
        result.WriteByte(1);
        int srcPos = 0;
        while (srcPos < data.Length)
        {
            int chunkStart = srcPos;
            int chunkEnd = Math.Min(srcPos + 4096, data.Length);
            var chunkBuf = new MemoryStream();
            int dPos = srcPos;
            while (dPos < chunkEnd)
            {
                long flagPos = chunkBuf.Position;
                chunkBuf.WriteByte(0);
                byte flagByte = 0;
                for (int bit = 0; bit < 8 && dPos < chunkEnd; bit++)
                {
                    int bestLen = 0, bestOff = 0;
                    int decompPos = dPos - chunkStart;
                    if (decompPos < 1) decompPos = 1;
                    int bitCount = 4;
                    while ((1 << bitCount) < decompPos) bitCount++;
                    if (bitCount > 12) bitCount = 12;
                    int maxOff = 1 << bitCount;
                    int maxLen = (0xFFFF >> bitCount) + 3;
                    for (int off = 1; off <= Math.Min(maxOff, dPos - chunkStart); off++)
                    {
                        int matchLen = 0;
                        while (matchLen < maxLen && dPos + matchLen < chunkEnd)
                        {
                            if (data[dPos - off + (matchLen % off)] != data[dPos + matchLen]) break;
                            matchLen++;
                        }
                        if (matchLen >= 3 && matchLen > bestLen) { bestLen = matchLen; bestOff = off; }
                    }
                    if (bestLen >= 3)
                    {
                        flagByte |= (byte)(1 << bit);
                        int token = ((bestOff - 1) << (16 - bitCount)) | (bestLen - 3);
                        chunkBuf.WriteByte((byte)(token & 0xFF));
                        chunkBuf.WriteByte((byte)((token >> 8) & 0xFF));
                        dPos += bestLen;
                    }
                    else
                    {
                        chunkBuf.WriteByte(data[dPos]); dPos++;
                    }
                }
                long savedPos = chunkBuf.Position;
                chunkBuf.Position = flagPos;
                chunkBuf.WriteByte(flagByte);
                chunkBuf.Position = savedPos;
            }
            byte[] compressed = chunkBuf.ToArray();
            srcPos = dPos;
            ushort hdr = (ushort)(0x8000 | 0x3000 | (compressed.Length + 2 - 3));
            result.WriteByte((byte)(hdr & 0xFF));
            result.WriteByte((byte)((hdr >> 8) & 0xFF));
            result.Write(compressed, 0, compressed.Length);
        }
        return result.ToArray();
    }

    public static int FindPattern(byte[] data, byte[] pattern)
    {
        for (int i = 0; i <= data.Length - pattern.Length; i++)
        {
            bool match = true;
            for (int j = 0; j < pattern.Length; j++)
            {
                if (data[i + j] != pattern[j]) { match = false; break; }
            }
            if (match) return i;
        }
        return -1;
    }
}
'@
}

# ============================================================================
# OLE2 Compound Document Parser
# ============================================================================

function Read-SectorChain([byte[]]$bytes, [int]$startSector, [int]$sectorSize, [int[]]$fat) {
    return ,[VbaToolkitNative]::ReadSectorChain($bytes, $startSector, $sectorSize, $fat)
}

function Read-MiniStream([byte[]]$miniStreamData, [int]$startSector, [int]$size, [int]$miniSectorSize, [int[]]$miniFat) {
    return ,[VbaToolkitNative]::ReadMiniStream($miniStreamData, $startSector, $size, $miniSectorSize, $miniFat)
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
        # Validate mini stream write
        $actualWritten = [Math]::Min($written, $newData.Length)
        if ($actualWritten -lt $newData.Length) {
            throw "Write-Ole2Stream: data truncated (mini stream). Wrote $actualWritten of $($newData.Length) bytes. Sector chain too short."
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
        # Validate write
        $actualWritten = [Math]::Min($written, $newData.Length)
        if ($actualWritten -lt $newData.Length) {
            throw "Write-Ole2Stream: data truncated. Wrote $actualWritten of $($newData.Length) bytes. Sector chain too short."
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
    return ,[VbaToolkitNative]::DecompressVba($data, $offset)
}

function Compress-VBA([byte[]]$data) {
    return ,[VbaToolkitNative]::CompressVba($data)
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

# ============================================================================
# Input validation
# ============================================================================

function Resolve-VbaFilePath {
    param([string]$Path, [string[]]$Supported = @('.xls','.xlsm','.xlam'))
    if (-not (Test-Path $Path)) { throw "File not found: $Path" }
    $resolved = (Resolve-Path $Path).Path
    $ext = [IO.Path]::GetExtension($resolved).ToLower()
    if ($ext -notin $Supported) { throw "Unsupported format: $ext (supported: $($Supported -join ', '))" }
    return $resolved
}

# ============================================================================
# Output management
# ============================================================================

function New-VbaOutputDir {
    param([string]$InputFilePath, [string]$ToolName)
    $inputDir = [IO.Path]::GetDirectoryName($InputFilePath)
    $outputRoot = Join-Path $inputDir 'output'
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $runDir = Join-Path $outputRoot "${timestamp}_${ToolName}"
    New-Item $runDir -ItemType Directory -Force | Out-Null
    return $runDir
}

# ============================================================================
# Logging and terminal display
# ============================================================================

function Write-VbaLog {
    param([string]$ToolName, [string]$InputFile, [string]$Message, [string]$Level = 'INFO')
    $logPath = Join-Path $PSScriptRoot '..\vba-toolkit.log'
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $fileName = [IO.Path]::GetFileName($InputFile)
    $entry = "[$timestamp] [$Level] [$ToolName] $fileName | $Message"
    [IO.File]::AppendAllText($logPath, "$entry`r`n", [System.Text.Encoding]::UTF8)
}

function Write-VbaStatus {
    param([string]$ToolName, [string]$FileName, [string]$Message)
    Write-Host "  $Message" -ForegroundColor Gray
}

function Write-VbaResult {
    param([string]$ToolName, [string]$FileName, [string]$Summary, [string]$OutputDir, [double]$ElapsedSec)
    Write-Host "  $Summary" -ForegroundColor Green
    if ($OutputDir) { Write-Host "  Output: $OutputDir" -ForegroundColor Gray }
    if ($ElapsedSec -gt 0) { Write-Host "  Done ($([Math]::Round($ElapsedSec, 1))s)" -ForegroundColor Gray }
}

function Write-VbaError {
    param([string]$ToolName, [string]$FileName, [string]$Message)
    Write-Host "  ERROR: $Message" -ForegroundColor Red
    Write-VbaLog $ToolName $FileName $Message 'ERROR'
}

function Write-VbaHeader {
    param([string]$ToolName, [string]$FileName)
    Write-Host "[$ToolName] $FileName" -ForegroundColor Cyan
}

# ============================================================================
# Codepage detection
# ============================================================================

function Get-VbaCodepage($ole2) {
    $codepage = 932  # default fallback (Shift-JIS)
    try {
        $dirEntry = $ole2.Entries | Where-Object { $_.Name -eq 'dir' } | Select-Object -First 1
        if (-not $dirEntry) { return $codepage }
        $dirRaw = Read-Ole2Stream $ole2 $dirEntry
        $dirData = Decompress-VBA $dirRaw 0
        if ($dirData.Length -lt 8) { return $codepage }
        # Scan for PROJECTCODEPAGE record (ID=0x0003)
        $pos = 0
        while ($pos + 6 -le $dirData.Length) {
            $id = [BitConverter]::ToUInt16($dirData, $pos)
            $size = [BitConverter]::ToInt32($dirData, $pos + 2)
            if ($size -lt 0 -or $pos + 6 + $size -gt $dirData.Length) { break }
            if ($id -eq 0x0003 -and $size -ge 2) {
                $codepage = [BitConverter]::ToUInt16($dirData, $pos + 6)
                break
            }
            $pos += 6 + $size
            # Handle PROJECTVERSION (0x0009) non-standard format: 2 extra bytes for MinorVersion
            if ($id -eq 0x0009) { $pos += 2 }
        }
    } catch {}
    # Validate codepage
    try { [System.Text.Encoding]::GetEncoding($codepage) | Out-Null } catch { $codepage = 932 }
    return $codepage
}

# ============================================================================
# Bulk module extraction
# ============================================================================

function Get-AllModuleCode {
    param(
        [string]$FilePath,
        [switch]$StripAttributes,
        [switch]$IncludeRawData
    )
    $proj = Get-VbaProjectBytes $FilePath
    if (-not $proj.Bytes) { return $null }
    $ole2 = Read-Ole2 $proj.Bytes
    $codepage = Get-VbaCodepage $ole2
    $encoding = [System.Text.Encoding]::GetEncoding($codepage)
    $modules = Get-VbaModuleList $ole2
    $result = [ordered]@{}
    foreach ($mod in $modules) {
        $mc = Get-VbaModuleCode $ole2 $mod.Name
        if (-not $mc) { continue }
        $code = $encoding.GetString((Decompress-VBA $mc.StreamData $mc.Offset))
        # Re-decode with detected codepage if different from what Get-VbaModuleCode used
        if ($codepage -ne 932) {
            $rawBytes = Decompress-VBA $mc.StreamData $mc.Offset
            $code = $encoding.GetString($rawBytes)
        }
        $lines = $code -split "`r`n|`n"
        if ($StripAttributes) {
            $lines = @($lines | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_' })
        }
        $entry = @{ Code = ($lines -join "`n"); Ext = $mod.Ext; Lines = $lines; Name = $mod.Name }
        if ($IncludeRawData) {
            $entry.Entry = $mc.Entry
            $entry.Offset = $mc.Offset
            $entry.StreamData = $mc.StreamData
        }
        $result[$mod.Name] = $entry
    }
    return @{ Modules = $result; Ole2 = $ole2; Ole2Bytes = $proj.Bytes; IsZip = $proj.IsZip; Codepage = $codepage; FilePath = $FilePath }
}

Export-ModuleMember -Function Read-Ole2, Read-Ole2Stream, Write-Ole2Stream,
    Decompress-VBA, Compress-VBA,
    Get-VbaProjectBytes, Save-VbaProjectBytes,
    Get-VbaModuleList, Get-VbaModuleCode,
    New-HtmlCodeView,
    Resolve-VbaFilePath, New-VbaOutputDir,
    Write-VbaLog, Write-VbaStatus, Write-VbaResult, Write-VbaError, Write-VbaHeader,
    Get-VbaCodepage, Get-AllModuleCode
