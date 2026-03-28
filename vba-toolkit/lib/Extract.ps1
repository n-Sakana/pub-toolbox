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
# Code Analysis
# ============================================================================

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

# Analysis patterns
$patterns = [ordered]@{
    'Win32 API (Declare)' = @{
        Pattern = '(?m)^[^''\r\n]*\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)'
        Extract = { param($m) "$($m.Groups[3].Value) ($(if($m.Groups[1].Value){'PtrSafe'} else {'Legacy'}))" }
    }
    'COM / CreateObject' = @{ Pattern = '(?m)^[^''\r\n]*\bCreateObject\s*\(\s*"([^"]+)"'; Extract = { param($m) $m.Groups[1].Value } }
    'COM / GetObject' = @{ Pattern = '(?m)^[^''\r\n]*\bGetObject\s*\(\s*"?([^")\s]+)"?'; Extract = { param($m) $m.Groups[1].Value } }
    'Shell / process' = @{ Pattern = '(?m)^[^''\r\n]*\b(Shell\s*[\("]|WScript\.Shell|cmd\s*/[ck])'; Extract = { param($m) $m.Groups[1].Value.Trim() } }
    'File I/O' = @{
        Pattern = '(?m)^[^''\r\n]*\b(Open\s+\S+\s+For\s+(Input|Output|Append|Binary|Random)|Kill\s|FileCopy\s|MkDir\s|RmDir\s)'
        Extract = { param($m) if ($m.Groups[2].Value) { "Open For $($m.Groups[2].Value)" } else { $m.Groups[1].Value.Trim() } }
    }
    'FileSystemObject' = @{ Pattern = '(?m)^[^''\r\n]*\b(Scripting\.FileSystemObject)\b'; Extract = { param($m) $m.Groups[1].Value } }
    'Registry' = @{ Pattern = '(?m)^[^''\r\n]*\b(GetSetting|SaveSetting|DeleteSetting|RegRead|RegWrite|RegDelete)\b'; Extract = { param($m) $m.Groups[1].Value } }
    'SendKeys' = @{ Pattern = '(?m)^[^''\r\n]*\b(SendKeys)\b'; Extract = { param($m) $m.Groups[1].Value } }
    'Network / HTTP' = @{ Pattern = '(?m)^[^''\r\n]*\b(MSXML2\.XMLHTTP|WinHttp\.WinHttpRequest|URLDownloadToFile|MSXML2\.ServerXMLHTTP)\b'; Extract = { param($m) $m.Groups[1].Value } }
    'PowerShell / WScript' = @{ Pattern = '(?mi)^[^''\r\n]*\b(powershell|wscript|cscript|mshta)\b'; Extract = { param($m) $m.Groups[1].Value } }
    'Process / WMI' = @{ Pattern = '(?m)^[^''\r\n]*\b(winmgmts|Win32_Process|WbemScripting|ExecQuery)\b'; Extract = { param($m) $m.Groups[1].Value } }
    'DLL loading' = @{ Pattern = '(?m)^[^''\r\n]*\b(LoadLibrary|GetProcAddress|FreeLibrary|CallByName)\b'; Extract = { param($m) $m.Groups[1].Value } }
    'Clipboard' = @{ Pattern = '(?m)^[^''\r\n]*\b(MSForms\.DataObject|GetClipboardData|SetClipboardData)\b'; Extract = { param($m) $m.Groups[1].Value } }
    'Environment' = @{ Pattern = '(?m)^[^''\r\n]*\b(Environ\s*\$?\s*\()'; Extract = { param($m) "Environ" } }
    'Auto-execution' = @{ Pattern = '(?m)^\s*(Sub\s+(Auto_Open|Auto_Close|Workbook_Open|Workbook_BeforeClose|Document_Open|Document_Close)\b)'; Extract = { param($m) $m.Groups[2].Value } }
    'Encoding / obfuscation' = @{ Pattern = '(?m)^[^''\r\n]*\b(Chr\s*\$?\s*\(\s*\d+\s*\))'; Extract = { param($m) $m.Groups[1].Value }; Aggregate = $true }
}

$issueCount = 0
foreach ($cat in $patterns.Keys) {
    $p = $patterns[$cat]
    $findings = [System.Collections.ArrayList]::new()
    foreach ($f in $allFiles) {
        $content = [IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
        foreach ($m in [regex]::Matches($content, $p.Pattern)) {
            [void]$findings.Add("$($f.Name): $(& $p.Extract $m)")
        }
    }
    if ($findings.Count -gt 0) {
        $issueCount += $findings.Count
        [void]$report.AppendLine("## $cat ($($findings.Count))")
        [void]$report.AppendLine("")
        if ($p.Aggregate) {
            foreach ($g in ($findings | Group-Object { $_ -replace ':.*', '' })) {
                [void]$report.AppendLine("  $($g.Name): $($g.Count) occurrence(s)")
            }
        } else {
            foreach ($item in ($findings | Sort-Object -Unique)) { [void]$report.AppendLine("  $item") }
        }
        [void]$report.AppendLine("")
    }
}

# COM usage trace
$allCode = @{}
foreach ($f in $allFiles) { $allCode[$f.Name] = ([IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)) -split "`r`n|`n" }

$comBindings = [System.Collections.ArrayList]::new()
foreach ($fn in $allCode.Keys) {
    $lines = $allCode[$fn]
    for ($li = 0; $li -lt $lines.Count; $li++) {
        if ($lines[$li] -match '^\s*''') { continue }
        if ($lines[$li] -match '\bSet\s+(\w+)\s*=\s*CreateObject\s*\(\s*"([^"]+)"') {
            [void]$comBindings.Add(@{ VarName = $Matches[1]; ProgId = $Matches[2]; File = $fn; Line = $li + 1 })
        } elseif ($lines[$li] -match '\bSet\s+(\w+)\s*=\s*GetObject\s*\(') {
            [void]$comBindings.Add(@{ VarName = $Matches[1]; ProgId = '(GetObject)'; File = $fn; Line = $li + 1 })
        }
    }
}

if ($comBindings.Count -gt 0) {
    [void]$report.AppendLine("## COM Object Usage Details")
    [void]$report.AppendLine("")
    foreach ($g in ($comBindings | Group-Object { $_.ProgId } | Sort-Object Name)) {
        [void]$report.AppendLine("  $($g.Name)")
        foreach ($b in $g.Group) { [void]$report.AppendLine("    $($b.File) L$($b.Line): Set $($b.VarName) = ...") }
        $varNames = ($g.Group | ForEach-Object { $_.VarName }) | Sort-Object -Unique
        foreach ($fn in $allCode.Keys) {
            $lines = $allCode[$fn]
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

# API usage trace
$apiDecls = [System.Collections.ArrayList]::new()
foreach ($fn in $allCode.Keys) {
    $lines = $allCode[$fn]
    for ($li = 0; $li -lt $lines.Count; $li++) {
        if ($lines[$li] -match '^\s*''') { continue }
        if ($lines[$li] -match '(?i)\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)\s+Lib\s+(.*)') {
            $trimmed = $lines[$li].Trim(); if ($trimmed.Length -gt 100) { $trimmed = $trimmed.Substring(0, 97) + '...' }
            [void]$apiDecls.Add(@{ Name = $Matches[3]; File = $fn; Line = $li + 1; Sig = $trimmed })
        }
    }
}
if ($apiDecls.Count -gt 0) {
    [void]$report.AppendLine("## Win32 API Usage Details")
    [void]$report.AppendLine("")
    foreach ($api in $apiDecls) {
        [void]$report.AppendLine("  $($api.Name)")
        [void]$report.AppendLine("    $($api.File) L$($api.Line): $($api.Sig)")
        foreach ($fn in $allCode.Keys) {
            $lines = $allCode[$fn]
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

if ($issueCount -eq 0) {
    [void]$report.AppendLine("## Result"); [void]$report.AppendLine("")
    [void]$report.AppendLine("  No external dependencies detected. Migration risk: LOW")
} else {
    [void]$report.AppendLine("## Summary"); [void]$report.AppendLine("")
    [void]$report.AppendLine("  $issueCount potential migration issue(s) detected.")
}

$reportText = $report.ToString()
$reportPath = Join-Path $outDir 'analysis.txt'
[IO.File]::WriteAllText($reportPath, $reportText, [System.Text.Encoding]::UTF8)
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
foreach ($cat in $patterns.Keys) { if (-not $patterns[$cat].Aggregate) { [void]$hlPatterns.Add($patterns[$cat].Pattern) } }
$apiCallNames = [System.Collections.ArrayList]::new()
foreach ($f in $allFiles) {
    $c = [IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
    foreach ($m in [regex]::Matches($c, '(?m)^[^''\r\n]*\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)')) {
        $n = $m.Groups[3].Value; if ($apiCallNames -notcontains $n) { [void]$apiCallNames.Add($n) }
    }
}
$comVarNames = [System.Collections.ArrayList]::new()
foreach ($b in $comBindings) { if ($comVarNames -notcontains $b.VarName) { [void]$comVarNames.Add($b.VarName) } }

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
        if (-not $hit) { foreach ($n in $apiCallNames) { if ($lines[$i] -match "\b$([regex]::Escape($n))\b" -and $lines[$i] -notmatch '(?i)\bDeclare\s') { $hit = $true; break } } }
        if (-not $hit) { foreach ($vn in $comVarNames) { if ($lines[$i] -match "\b$([regex]::Escape($vn))\.") { $hit = $true; break } } }
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
