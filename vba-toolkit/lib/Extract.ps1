param([Parameter(Mandatory)][string]$FilePath)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

$FilePath = Resolve-VbaFilePath $FilePath
$fileName = [IO.Path]::GetFileName($FilePath)
$baseName = [IO.Path]::GetFileNameWithoutExtension($FilePath)
$sw = [System.Diagnostics.Stopwatch]::StartNew()

Write-VbaHeader 'Extract' $fileName
Write-VbaLog 'Extract' $FilePath 'Started'

$project = Get-AllModuleCode $FilePath -StripAttributes
if (-not $project) { Write-VbaError 'Extract' $fileName 'No vbaProject.bin found'; exit 0 }

$outDir = New-VbaOutputDir $FilePath 'extract'
$modulesDir = Join-Path $outDir 'modules'
New-Item $modulesDir -ItemType Directory -Force | Out-Null

# Write individual module files
$extracted = 0
foreach ($modName in $project.Modules.Keys) {
    $mod = $project.Modules[$modName]
    $outPath = Join-Path $modulesDir "$modName.$($mod.Ext)"
    [IO.File]::WriteAllText($outPath, ($mod.Lines -join "`r`n"), [System.Text.Encoding]::UTF8)
    Write-VbaStatus 'Extract' $fileName "  $modName.$($mod.Ext)"
    $extracted++
}
Write-VbaStatus 'Extract' $fileName "$extracted module(s) extracted"

# ============================================================================
# Code Analysis (via shared engine)
# ============================================================================

$analysis = Get-VbaAnalysis -Project $project

$allFiles = Get-ChildItem $modulesDir -File
$totalLines = 0
$report = [System.Text.StringBuilder]::new()
[void]$report.AppendLine("# VBA Code Analysis Report")
[void]$report.AppendLine("# Source: $fileName")
[void]$report.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$report.AppendLine("")
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

# Pattern findings
foreach ($cat in $analysis.Findings.Keys) {
    $f = $analysis.Findings[$cat]
    [void]$report.AppendLine("## $cat ($($f.Findings.Count))")
    [void]$report.AppendLine("")
    if ($f.Aggregate) {
        foreach ($g in ($f.Findings | Group-Object { $_ -replace ':.*', '' })) {
            [void]$report.AppendLine("  $($g.Name): $($g.Count) occurrence(s)")
        }
    } else {
        foreach ($item in ($f.Findings | Sort-Object -Unique)) { [void]$report.AppendLine("  $item") }
    }
    [void]$report.AppendLine("")
}

# COM usage details
if ($analysis.ComBindings.Count -gt 0) {
    [void]$report.AppendLine("## COM Object Usage Details")
    [void]$report.AppendLine("")
    foreach ($g in ($analysis.ComBindings | Group-Object { $_.ProgId } | Sort-Object Name)) {
        [void]$report.AppendLine("  $($g.Name)")
        foreach ($b in $g.Group) { [void]$report.AppendLine("    $($b.File) L$($b.Line): Set $($b.VarName) = ...") }
        $varNames = ($g.Group | ForEach-Object { $_.VarName }) | Sort-Object -Unique
        foreach ($fn in $analysis.AllCode.Keys) {
            $lines = $analysis.AllCode[$fn]
            for ($li = 0; $li -lt $lines.Count; $li++) {
                if ($lines[$li] -match '^\s*''') { continue }
                foreach ($vn in $varNames) {
                    if ($lines[$li] -match "\b$([regex]::Escape($vn))\.(\w+)") {
                        $trimmed = $lines[$li].Trim(); if ($trimmed.Length -gt 80) { $trimmed = $trimmed.Substring(0, 77) + '...' }
                        [void]$report.AppendLine("    $fn L$($li+1): .$($Matches[1])  -- $trimmed")
                        break
                    }
                }
            }
        }
        [void]$report.AppendLine("")
    }
}

# API usage details
if ($analysis.ApiDecls.Count -gt 0) {
    [void]$report.AppendLine("## Win32 API Usage Details")
    [void]$report.AppendLine("")
    foreach ($api in $analysis.ApiDecls) {
        [void]$report.AppendLine("  $($api.Name)")
        [void]$report.AppendLine("    $($api.File) L$($api.Line): $($api.Sig)")
        foreach ($fn in $analysis.AllCode.Keys) {
            $lines = $analysis.AllCode[$fn]
            for ($li = 0; $li -lt $lines.Count; $li++) {
                if ($lines[$li] -match '^\s*''' -or $lines[$li] -match '(?i)\bDeclare\s') { continue }
                if ($lines[$li] -match "\b$([regex]::Escape($api.Name))\b") {
                    $trimmed = $lines[$li].Trim(); if ($trimmed.Length -gt 80) { $trimmed = $trimmed.Substring(0, 77) + '...' }
                    [void]$report.AppendLine("    $fn L$($li+1): $trimmed")
                }
            }
        }
        [void]$report.AppendLine("")
    }
}

# External references
$projEntry = $project.Ole2.Entries | Where-Object { $_.Name -eq 'PROJECT' -and $_.ObjType -eq 2 } | Select-Object -First 1
if ($projEntry) {
    $projData = Read-Ole2Stream $project.Ole2 $projEntry
    $refs = [System.Collections.ArrayList]::new()
    foreach ($line in ([System.Text.Encoding]::GetEncoding($project.Codepage).GetString($projData)) -split "`r`n|`n") {
        if ($line -match '^Reference=' -and $line -match '#([^#]+)$') { [void]$refs.Add($Matches[1]) }
    }
    if ($refs.Count -gt 0) {
        [void]$report.AppendLine("## External References ($($refs.Count))")
        [void]$report.AppendLine("")
        foreach ($r in $refs) { [void]$report.AppendLine("  $r") }
        [void]$report.AppendLine("")
    }
}

$issueCount = $analysis.IssueCount
if ($issueCount -eq 0) {
    [void]$report.AppendLine("## Result"); [void]$report.AppendLine("")
    [void]$report.AppendLine("  No external dependencies detected. Migration risk: LOW")
} else {
    [void]$report.AppendLine("## Summary"); [void]$report.AppendLine("")
    [void]$report.AppendLine("  $issueCount potential migration issue(s) detected.")
}

$reportText = $report.ToString()
[IO.File]::WriteAllText((Join-Path $outDir 'analysis.txt'), $reportText, [System.Text.Encoding]::UTF8)
Write-Host ""; Write-Host $reportText

# Combined source
$combined = [System.Text.StringBuilder]::new()
[void]$combined.AppendLine("=" * 80)
[void]$combined.AppendLine(" $fileName - VBA Source Code (Combined)")
[void]$combined.AppendLine(" Extracted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$combined.AppendLine("=" * 80)
[void]$combined.AppendLine("")
$order = @('bas','cls','frm')
$sorted = $allFiles | Sort-Object { $o = [Array]::IndexOf($order, $_.Extension.TrimStart('.')); if($o -lt 0){99}else{$o} }, Name
foreach ($f in $sorted) {
    $c = [IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
    [void]$combined.AppendLine("=" * 80)
    [void]$combined.AppendLine(" $($f.Name)")
    [void]$combined.AppendLine("=" * 80)
    [void]$combined.AppendLine("")
    [void]$combined.AppendLine($c.TrimStart("`r`n"))
    [void]$combined.AppendLine("")
}
[IO.File]::WriteAllText((Join-Path $outDir 'combined.txt'), $combined.ToString(), [System.Text.Encoding]::UTF8)

# HTML viewer
$hlPatterns = [System.Collections.ArrayList]::new()
foreach ($cat in $analysis.Patterns.Keys) { if (-not $analysis.Patterns[$cat].Aggregate) { [void]$hlPatterns.Add($analysis.Patterns[$cat].Pattern) } }

$htmlModules = [ordered]@{}
foreach ($f in $sorted) {
    $modName = [IO.Path]::GetFileNameWithoutExtension($f.Name)
    $modExt = $f.Extension.TrimStart('.')
    $lines = ([IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)) -split "`r`n|`n"
    $highlights = @{}
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*''') { continue }
        $hit = $false
        foreach ($pat in $hlPatterns) { if ($lines[$i] -match $pat) { $hit = $true; break } }
        if (-not $hit) { foreach ($n in $analysis.ApiCallNames) { if ($lines[$i] -match "\b$([regex]::Escape($n))\b" -and $lines[$i] -notmatch '(?i)\bDeclare\s') { $hit = $true; break } } }
        if (-not $hit) { foreach ($vn in $analysis.ComVarNames) { if ($lines[$i] -match "\b$([regex]::Escape($vn))\.") { $hit = $true; break } } }
        if ($hit) { $highlights[$i] = $true }
    }
    $htmlModules[$modName] = @{ Ext = $modExt; Lines = [System.Collections.ArrayList]::new($lines); Highlights = $highlights }
}

New-HtmlCodeView -title "VBA Extract: $fileName" -subtitle "$($allFiles.Count) modules, $totalLines lines, $issueCount issue(s)" `
    -moduleData $htmlModules -highlightClass 'hl-edr' -highlightColor '#1b2e4a' -highlightText '#a0c4f0' -markerColor '#4a9eff' `
    -outputPath (Join-Path $outDir 'extract.html')
Start-Process (Join-Path $outDir 'extract.html')

$sw.Stop()
Write-VbaResult 'Extract' $fileName "$extracted module(s), $issueCount issue(s)" $outDir $sw.Elapsed.TotalSeconds
Write-VbaLog 'Extract' $FilePath "$extracted modules, $issueCount issues | -> $outDir"
