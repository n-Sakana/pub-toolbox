param([Parameter(Mandatory)][string]$FilePath)
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

if (-not (Test-Path $FilePath)) { Write-Host "Error: file not found: $FilePath" -ForegroundColor Red; exit 1 }
$FilePath = (Resolve-Path $FilePath).Path
$ext = [IO.Path]::GetExtension($FilePath).ToLower()
if ($ext -notin '.xls', '.xlsm', '.xlam') { Write-Host "Error: unsupported format: $ext" -ForegroundColor Red; exit 1 }

# ============================================================================
# Sanitize rules — edit here to customize
# ============================================================================

$rules = @(
    @{ Name = 'Win32 API (Declare)';  Pattern = '^\s*(Public\s+|Private\s+)?Declare\s+(PtrSafe\s+)?(Function|Sub)\b' }
    @{ Name = 'DLL loading';          Pattern = '^\s*[^'']*\b(LoadLibrary|GetProcAddress|FreeLibrary)\b' }
)

$commentPrefix = "' [SANITIZED] "

# ============================================================================

Write-Host "Sanitizing VBA code in: $FilePath"

$bakPath = "$FilePath.bak"
Copy-Item $FilePath $bakPath -Force
Write-Host "Backup created: $bakPath"

$proj = Get-VbaProjectBytes $FilePath
if (-not $proj.Bytes) { Write-Host "No vbaProject.bin found." -ForegroundColor Yellow; exit 0 }
$ole2Bytes = $proj.Bytes
$ole2 = Read-Ole2 $ole2Bytes
$modules = Get-VbaModuleList $ole2

# Pass 1: collect declared API names
$declaredNames = [System.Collections.ArrayList]::new()
$moduleData = [ordered]@{}

foreach ($mod in $modules) {
    $result = Get-VbaModuleCode $ole2 $mod.Name
    if (-not $result) { continue }
    $moduleData[$mod.Name] = $result
    foreach ($m in [regex]::Matches($result.Code, '(?im)^\s*(?:Public\s+|Private\s+)?Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+(\w+)')) {
        $n = $m.Groups[1].Value
        if ($declaredNames -notcontains $n) { [void]$declaredNames.Add($n) }
    }
}

if ($declaredNames.Count -gt 0) {
    Write-Host "  Found API declarations: $($declaredNames -join ', ')"
}

# Pass 2: sanitize
$totalChanges = 0
$report = [System.Text.StringBuilder]::new()
[void]$report.AppendLine("# VBA Sanitizer Report")
[void]$report.AppendLine("# Source: $([IO.Path]::GetFileName($FilePath))")
[void]$report.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$report.AppendLine("")

foreach ($modName in $moduleData.Keys) {
    $md = $moduleData[$modName]
    $lines = $md.Code -split "`r`n|`n"
    $changes = 0

    for ($i = 0; $i -lt $lines.Length; $i++) {
        $line = $lines[$i]
        if ($line -match '^\s*''') { continue }

        $matched = $false; $ruleName = ''
        foreach ($rule in $rules) {
            if ($line -match $rule.Pattern) { $matched = $true; $ruleName = $rule.Name; break }
        }
        if (-not $matched -and $declaredNames.Count -gt 0) {
            foreach ($apiName in $declaredNames) {
                if ($line -match "\b$([regex]::Escape($apiName))\b") {
                    $matched = $true; $ruleName = "Call to $apiName"; break
                }
            }
        }

        if ($matched) {
            $lines[$i] = $commentPrefix + $line.TrimStart()
            $changes++
            [void]$report.AppendLine("  $modName L$($i+1) [$ruleName]")
            [void]$report.AppendLine("    $($line.Trim())")
        }
    }

    if ($changes -eq 0) { continue }
    $totalChanges += $changes
    Write-Host "  $modName : $changes line(s) commented out" -ForegroundColor Yellow

    # Recompress and write back
    $newCode = $lines -join "`r`n"
    $newCodeBytes = [System.Text.Encoding]::GetEncoding(932).GetBytes($newCode)
    $newCompressed = Compress-VBA $newCodeBytes

    $prefix = New-Object byte[] $md.Offset
    [Array]::Copy($md.StreamData, $prefix, $md.Offset)
    $newStream = New-Object byte[] ($prefix.Length + $newCompressed.Length)
    [Array]::Copy($prefix, $newStream, $prefix.Length)
    [Array]::Copy($newCompressed, 0, $newStream, $prefix.Length, $newCompressed.Length)

    Write-Ole2Stream $ole2Bytes $ole2 $md.Entry $newStream
}

if ($totalChanges -gt 0) {
    Save-VbaProjectBytes $FilePath $ole2Bytes $proj.IsZip
    Write-Host "`n$totalChanges line(s) sanitized." -ForegroundColor Green
} else {
    Write-Host "`nNo lines matched sanitize rules. File unchanged." -ForegroundColor Cyan
}

[void]$report.AppendLine("")
[void]$report.AppendLine("## Summary")
[void]$report.AppendLine("  $totalChanges line(s) commented out.")
if ($declaredNames.Count -gt 0) {
    [void]$report.AppendLine("")
    [void]$report.AppendLine("## Declared APIs")
    foreach ($n in $declaredNames) { [void]$report.AppendLine("  $n") }
}
$reportPath = [IO.Path]::ChangeExtension($FilePath, $null).TrimEnd('.') + "_sanitized.txt"
[IO.File]::WriteAllText($reportPath, $report.ToString(), [System.Text.Encoding]::UTF8)
Write-Host "Report: $reportPath"

# ============================================================================
# HTML viewer with sanitized line highlights (yellow)
# ============================================================================

# Re-read the sanitized code for HTML
$htmlModules = [ordered]@{}
$proj2 = Get-VbaProjectBytes $FilePath
if ($proj2.Bytes) {
    $ole2_2 = Read-Ole2 $proj2.Bytes
    $modules2 = Get-VbaModuleList $ole2_2
    foreach ($mod in $modules2) {
        $result = Get-VbaModuleCode $ole2_2 $mod.Name
        if (-not $result) { continue }
        $lines = ($result.Code -split "`r`n|`n") | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_' }
        $lineArr = [System.Collections.ArrayList]::new()
        $highlights = @{}
        $idx = 0
        foreach ($line in $lines) {
            [void]$lineArr.Add($line)
            if ($line -match [regex]::Escape($commentPrefix)) { $highlights[$idx] = $true }
            $idx++
        }
        $htmlModules[$mod.Name] = @{ Ext = $mod.Ext; Lines = $lineArr; Highlights = $highlights }
    }
}

if ($htmlModules.Count -gt 0) {
    $baseName = [IO.Path]::GetFileNameWithoutExtension($FilePath)
    $htmlPath = [IO.Path]::ChangeExtension($FilePath, $null).TrimEnd('.') + "_sanitized.html"
    New-HtmlCodeView `
        -title "VBA Sanitize: $([IO.Path]::GetFileName($FilePath))" `
        -subtitle "$totalChanges line(s) commented out" `
        -moduleData $htmlModules `
        -highlightClass 'hl-sanitized' `
        -highlightColor '#4b3a00' `
        -highlightText '#f0d870' `
        -markerColor '#e8ab53' `
        -outputPath $htmlPath

    Start-Process $htmlPath
    Write-Host "HTML viewer: $htmlPath" -ForegroundColor Green
}
