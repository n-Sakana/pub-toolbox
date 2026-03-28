param([Parameter(Mandatory)][string]$FilePath)
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

if (-not (Test-Path $FilePath)) { Write-Host "Error: file not found: $FilePath" -ForegroundColor Red; exit 1 }
$FilePath = (Resolve-Path $FilePath).Path
$ext = [IO.Path]::GetExtension($FilePath).ToLower()
if ($ext -notin '.xls', '.xlsm', '.xlam') { Write-Host "Error: unsupported format: $ext" -ForegroundColor Red; exit 1 }

Write-Host "Extracting VBA code from: $FilePath"

$proj = Get-VbaProjectBytes $FilePath
if (-not $proj.Bytes) { Write-Host "No vbaProject.bin found." -ForegroundColor Yellow; exit 0 }
$ole2 = Read-Ole2 $proj.Bytes
$modules = Get-VbaModuleList $ole2
if ($modules.Count -eq 0) { Write-Host "No VBA modules found." -ForegroundColor Yellow; exit 0 }

# Output directory
$baseName = [IO.Path]::GetFileNameWithoutExtension($FilePath)
$outDir = Join-Path ([IO.Path]::GetDirectoryName($FilePath)) "${baseName}_vba"
if (-not (Test-Path $outDir)) { New-Item $outDir -ItemType Directory | Out-Null }

$extracted = 0
foreach ($mod in $modules) {
    $result = Get-VbaModuleCode $ole2 $mod.Name
    if (-not $result) { Write-Host "  SKIP: $($mod.Name)" -ForegroundColor Yellow; continue }
    $outPath = Join-Path $outDir "$($mod.Name).$($mod.Ext)"
    [IO.File]::WriteAllText($outPath, $result.Code, [System.Text.Encoding]::UTF8)
    Write-Host "  $($mod.Name).$($mod.Ext)" -ForegroundColor Green
    $extracted++
}

Write-Host ""
if ($extracted -eq 0) { Write-Host "No VBA modules extracted." -ForegroundColor Yellow; exit 0 }
Write-Host "$extracted module(s) extracted to: $outDir" -ForegroundColor Green

# ============================================================================
# Code Analysis
# ============================================================================

Write-Host ""
Write-Host "=== Code Analysis ===" -ForegroundColor Cyan

$allFiles = Get-ChildItem $outDir -File
$totalLines = 0
$report = [System.Text.StringBuilder]::new()
[void]$report.AppendLine("# VBA Code Analysis Report")
[void]$report.AppendLine("# Source: $([IO.Path]::GetFileName($FilePath))")
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
    'File I/O' = @{
        Pattern = '(?m)^[^'']*\b(Open\s+\S+\s+For\s+(Input|Output|Append|Binary|Random)|Kill\s|FileCopy\s|MkDir\s|RmDir\s)'
        Extract = { param($m) if ($m.Groups[2].Value) { "Open For $($m.Groups[2].Value)" } else { $m.Groups[1].Value.Trim() } }
    }
    'FileSystemObject' = @{
        Pattern = '(?m)^[^'']*\b(Scripting\.FileSystemObject)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Registry' = @{
        Pattern = '(?m)^[^'']*\b(GetSetting|SaveSetting|DeleteSetting|RegRead|RegWrite|RegDelete)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'SendKeys' = @{ Pattern = '(?m)^[^'']*\b(SendKeys)\b'; Extract = { param($m) $m.Groups[1].Value } }
    'Network / HTTP' = @{
        Pattern = '(?m)^[^'']*\b(MSXML2\.XMLHTTP|WinHttp\.WinHttpRequest|URLDownloadToFile|MSXML2\.ServerXMLHTTP)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'PowerShell / WScript' = @{
        Pattern = '(?mi)^[^'']*\b(powershell|wscript|cscript|mshta)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Process / WMI' = @{
        Pattern = '(?m)^[^'']*\b(winmgmts|Win32_Process|WbemScripting|ExecQuery)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'DLL loading' = @{
        Pattern = '(?m)^[^'']*\b(LoadLibrary|GetProcAddress|FreeLibrary|CallByName)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Clipboard' = @{
        Pattern = '(?m)^[^'']*\b(MSForms\.DataObject|GetClipboardData|SetClipboardData)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Environment' = @{
        Pattern = '(?m)^[^'']*\b(Environ\s*\$?\s*\()'
        Extract = { param($m) "Environ" }
    }
    'Auto-execution' = @{
        Pattern = '(?m)^\s*(Sub\s+(Auto_Open|Auto_Close|Workbook_Open|Workbook_BeforeClose|Document_Open|Document_Close)\b)'
        Extract = { param($m) $m.Groups[2].Value }
    }
    'Encoding / obfuscation' = @{
        Pattern = '(?m)^[^'']*\b(Chr\s*\$?\s*\(\s*\d+\s*\))'
        Extract = { param($m) $m.Groups[1].Value }
        Aggregate = $true
    }
}

$issueCount = 0
foreach ($cat in $patterns.Keys) {
    $p = $patterns[$cat]
    if ($p.Skip) { continue }
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

# External references
$projEntry2 = $ole2.Entries | Where-Object { $_.Name -eq 'PROJECT' -and $_.ObjType -eq 2 } | Select-Object -First 1
if ($projEntry2) {
    $projData2 = Read-Ole2Stream $ole2 $projEntry2
    $refs = [System.Collections.ArrayList]::new()
    foreach ($line in ([System.Text.Encoding]::GetEncoding(932).GetString($projData2)) -split "`r`n|`n") {
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
$reportPath = Join-Path $outDir "_analysis.txt"
[IO.File]::WriteAllText($reportPath, $reportText, [System.Text.Encoding]::UTF8)
Write-Host ""; Write-Host $reportText
Write-Host "Report saved to: $reportPath" -ForegroundColor Green

# ============================================================================
# Combined source
# ============================================================================

$combined = [System.Text.StringBuilder]::new()
[void]$combined.AppendLine("=" * 80)
[void]$combined.AppendLine(" $([IO.Path]::GetFileName($FilePath)) - VBA Source Code (Combined)")
[void]$combined.AppendLine(" Extracted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$combined.AppendLine("=" * 80)
[void]$combined.AppendLine("")
[void]$combined.AppendLine("MODULE INDEX")
[void]$combined.AppendLine("-" * 40)

$groups = @{ bas = @(); cls = @(); frm = @(); doc = @() }
foreach ($mod in $modules) {
    $f = Get-Item (Join-Path $outDir "$($mod.Name).$($mod.Ext)") -ErrorAction SilentlyContinue
    if (-not $f) { continue }
    $lc = (Get-Content $f.FullName -Encoding UTF8).Count
    $label = "$($mod.Name) ($lc lines)"
    if ($mod.Ext -eq 'bas') { $groups.bas += $label }
    elseif ($mod.Ext -eq 'frm') { $groups.frm += $label }
    elseif ($mod.Name -match '^(ThisWorkbook|Sheet\d+)$') { $groups.doc += $label }
    else { $groups.cls += $label }
}
if ($groups.bas.Count) { [void]$combined.AppendLine("`n  Standard Modules:"); $groups.bas | ForEach-Object { [void]$combined.AppendLine("    $_") } }
if ($groups.cls.Count) { [void]$combined.AppendLine("`n  Class Modules:"); $groups.cls | ForEach-Object { [void]$combined.AppendLine("    $_") } }
if ($groups.frm.Count) { [void]$combined.AppendLine("`n  UserForms:"); $groups.frm | ForEach-Object { [void]$combined.AppendLine("    $_") } }
if ($groups.doc.Count) { [void]$combined.AppendLine("`n  Document Modules:"); $groups.doc | ForEach-Object { [void]$combined.AppendLine("    $_") } }
[void]$combined.AppendLine("`n  Total: $totalLines lines across $($allFiles.Count) modules")
[void]$combined.AppendLine("")

# Dependencies
[void]$combined.AppendLine("DEPENDENCIES"); [void]$combined.AppendLine("-" * 40)
$allProgIds = [System.Collections.ArrayList]::new()
$allApis = [System.Collections.ArrayList]::new()
foreach ($f in $allFiles) {
    $c = [IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
    foreach ($m in [regex]::Matches($c, '(?m)^[^'']*\bCreateObject\s*\(\s*"([^"]+)"')) {
        $v = $m.Groups[1].Value; if ($allProgIds -notcontains $v) { [void]$allProgIds.Add($v) }
    }
    foreach ($m in [regex]::Matches($c, '(?m)^[^'']*\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)')) {
        [void]$allApis.Add($m.Groups[3].Value)
    }
}
if ($allProgIds.Count) { [void]$combined.AppendLine("`n  COM Objects:"); $allProgIds | Sort-Object | ForEach-Object { [void]$combined.AppendLine("    $_") } }
if ($allApis.Count) { [void]$combined.AppendLine("`n  Win32 API:"); $allApis | Sort-Object -Unique | ForEach-Object { [void]$combined.AppendLine("    $_") } }
if (-not $allProgIds.Count -and -not $allApis.Count) { [void]$combined.AppendLine("`n  (none)") }
[void]$combined.AppendLine("`n")

# All modules
$order = @('bas','cls','frm')
$sorted = $allFiles | Sort-Object { $o = [Array]::IndexOf($order, $_.Extension.TrimStart('.')); if($o -lt 0){99}else{$o} }, Name
foreach ($f in $sorted) {
    $c = [IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
    $clean = (($c -split "`r`n|`n") | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_' }) -join "`r`n"
    [void]$combined.AppendLine("=" * 80)
    [void]$combined.AppendLine(" $($f.Name)")
    [void]$combined.AppendLine("=" * 80)
    [void]$combined.AppendLine("")
    [void]$combined.AppendLine($clean.TrimStart("`r`n"))
    [void]$combined.AppendLine("")
}

$combinedPath = Join-Path ([IO.Path]::GetDirectoryName($FilePath)) "${baseName}_combined.txt"
[IO.File]::WriteAllText($combinedPath, $combined.ToString(), [System.Text.Encoding]::UTF8)
Write-Host "Combined source: $combinedPath" -ForegroundColor Green
