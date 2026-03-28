param([Parameter(Mandatory)][string]$FilePath)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

$FilePath = Resolve-VbaFilePath $FilePath
$fileName = [IO.Path]::GetFileName($FilePath)
$sw = [System.Diagnostics.Stopwatch]::StartNew()

Write-VbaHeader 'Sanitize' $fileName
Write-VbaLog 'Sanitize' $FilePath 'Started'

# ============================================================================
# Sanitize rules
# ============================================================================

$rules = @(
    @{ Name = 'Win32 API (Declare)';  Pattern = '^\s*(Public\s+|Private\s+)?Declare\s+(PtrSafe\s+)?(Function|Sub)\b' }
    @{ Name = 'DLL loading';          Pattern = '^\s*[^'']*\b(LoadLibrary|GetProcAddress|FreeLibrary)\b' }
)
$commentPrefix = "' [SANITIZED] "

# ============================================================================
# Non-destructive: copy to output, modify the copy
# ============================================================================

$outDir = New-VbaOutputDir $FilePath 'sanitize'
$copyPath = Join-Path $outDir $fileName
Copy-Item $FilePath $copyPath -Force
Write-VbaStatus 'Sanitize' $fileName "Copy created in output folder"

# Read from the copy
$project = Get-AllModuleCode $copyPath -IncludeRawData
if (-not $project) { Write-VbaError 'Sanitize' $fileName 'No vbaProject.bin found'; exit 0 }

# Pass 1: collect declared API names
$declaredNames = [System.Collections.ArrayList]::new()
foreach ($modName in $project.Modules.Keys) {
    $mod = $project.Modules[$modName]
    foreach ($m in [regex]::Matches($mod.Code, '(?im)^\s*(?:Public\s+|Private\s+)?Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+(\w+)')) {
        $n = $m.Groups[1].Value
        if ($declaredNames -notcontains $n) { [void]$declaredNames.Add($n) }
    }
}
if ($declaredNames.Count -gt 0) {
    Write-VbaStatus 'Sanitize' $fileName "Found API declarations: $($declaredNames -join ', ')"
}

# Pass 2: sanitize
$totalChanges = 0
$report = [System.Text.StringBuilder]::new()
[void]$report.AppendLine("# VBA Sanitizer Report")
[void]$report.AppendLine("# Source: $fileName")
[void]$report.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$report.AppendLine("")

$encoding = [System.Text.Encoding]::GetEncoding($project.Codepage)
$ole2Bytes = $project.Ole2Bytes

foreach ($modName in $project.Modules.Keys) {
    $md = $project.Modules[$modName]
    if (-not $md.Entry) { continue }  # no raw data
    $lines = $md.Code -split "`r`n|`n"
    $changes = 0

    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ($lines[$i] -match '^\s*''') { continue }
        $matched = $false; $ruleName = ''
        foreach ($rule in $rules) {
            if ($lines[$i] -match $rule.Pattern) { $matched = $true; $ruleName = $rule.Name; break }
        }
        if (-not $matched -and $declaredNames.Count -gt 0) {
            foreach ($apiName in $declaredNames) {
                if ($lines[$i] -match "\b$([regex]::Escape($apiName))\b") { $matched = $true; $ruleName = "Call to $apiName"; break }
            }
        }
        if ($matched) {
            $lines[$i] = $commentPrefix + $lines[$i].TrimStart()
            $changes++
            [void]$report.AppendLine("  $modName L$($i+1) [$ruleName]")
            [void]$report.AppendLine("    $($lines[$i].Trim())")
        }
    }

    if ($changes -eq 0) { continue }
    $totalChanges += $changes
    Write-VbaStatus 'Sanitize' $fileName "${modName}: $changes line(s) commented out"

    $newCode = $lines -join "`r`n"
    $newCodeBytes = $encoding.GetBytes($newCode)
    $newCompressed = Compress-VBA $newCodeBytes

    $prefix = New-Object byte[] $md.Offset
    [Array]::Copy($md.StreamData, $prefix, $md.Offset)
    $newStream = New-Object byte[] ($prefix.Length + $newCompressed.Length)
    [Array]::Copy($prefix, $newStream, $prefix.Length)
    [Array]::Copy($newCompressed, 0, $newStream, $prefix.Length, $newCompressed.Length)

    Write-Ole2Stream $ole2Bytes $project.Ole2 $md.Entry $newStream
}

if ($totalChanges -gt 0) {
    Save-VbaProjectBytes $copyPath $ole2Bytes $project.IsZip
} else {
    Remove-Item $copyPath -Force  # no changes, remove the copy
}

[void]$report.AppendLine("")
[void]$report.AppendLine("## Summary")
[void]$report.AppendLine("  $totalChanges line(s) commented out.")
if ($declaredNames.Count -gt 0) {
    [void]$report.AppendLine("")
    [void]$report.AppendLine("## Declared APIs")
    foreach ($n in $declaredNames) { [void]$report.AppendLine("  $n") }
}
[IO.File]::WriteAllText((Join-Path $outDir 'sanitize.txt'), $report.ToString(), [System.Text.Encoding]::UTF8)

# HTML viewer
if ($totalChanges -gt 0) {
    $project2 = Get-AllModuleCode $copyPath -StripAttributes
    if ($project2) {
        $htmlModules = [ordered]@{}
        foreach ($modName in $project2.Modules.Keys) {
            $mod = $project2.Modules[$modName]
            $highlights = @{}
            for ($i = 0; $i -lt $mod.Lines.Count; $i++) {
                if ($mod.Lines[$i] -match [regex]::Escape($commentPrefix)) { $highlights[$i] = $true }
            }
            $htmlModules[$modName] = @{ Ext = $mod.Ext; Lines = [System.Collections.ArrayList]::new($mod.Lines); Highlights = $highlights }
        }
        New-HtmlCodeView -title "VBA Sanitize: $fileName" -subtitle "$totalChanges line(s) commented out" `
            -moduleData $htmlModules -highlightClass 'hl-sanitized' -highlightColor '#4b3a00' -highlightText '#f0d870' -markerColor '#e8ab53' `
            -outputPath (Join-Path $outDir 'sanitize.html')
        Start-Process (Join-Path $outDir 'sanitize.html')
    }
}

$sw.Stop()
$msg = if ($totalChanges -gt 0) { "$totalChanges line(s) sanitized" } else { "No changes needed" }
Write-VbaResult 'Sanitize' $fileName $msg $outDir $sw.Elapsed.TotalSeconds
Write-VbaLog 'Sanitize' $FilePath "$totalChanges lines sanitized | -> $outDir"
