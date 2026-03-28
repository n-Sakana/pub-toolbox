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
        Pattern = '(?m)^[^''\r\n]*\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)'
        Extract = { param($m) "$($m.Groups[3].Value) ($(if($m.Groups[1].Value){'PtrSafe'} else {'Legacy'}))" }
    }
    'COM / CreateObject' = @{
        Pattern = '(?m)^[^''\r\n]*\bCreateObject\s*\(\s*"([^"]+)"'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'COM / GetObject' = @{
        Pattern = '(?m)^[^''\r\n]*\bGetObject\s*\(\s*"?([^")\s]+)"?'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Shell / process execution' = @{
        Pattern = '(?m)^[^''\r\n]*\b(Shell\s*[\("]|WScript\.Shell|cmd\s*/[ck])'
        Extract = { param($m) $m.Groups[1].Value.Trim() }
    }
    'File I/O' = @{
        Pattern = '(?m)^[^''\r\n]*\b(Open\s+\S+\s+For\s+(Input|Output|Append|Binary|Random)|Kill\s|FileCopy\s|MkDir\s|RmDir\s)'
        Extract = { param($m) if ($m.Groups[2].Value) { "Open For $($m.Groups[2].Value)" } else { $m.Groups[1].Value.Trim() } }
    }
    'FileSystemObject' = @{
        Pattern = '(?m)^[^''\r\n]*\b(Scripting\.FileSystemObject)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Registry' = @{
        Pattern = '(?m)^[^''\r\n]*\b(GetSetting|SaveSetting|DeleteSetting|RegRead|RegWrite|RegDelete)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'SendKeys' = @{ Pattern = '(?m)^[^''\r\n]*\b(SendKeys)\b'; Extract = { param($m) $m.Groups[1].Value } }
    'Network / HTTP' = @{
        Pattern = '(?m)^[^''\r\n]*\b(MSXML2\.XMLHTTP|WinHttp\.WinHttpRequest|URLDownloadToFile|MSXML2\.ServerXMLHTTP)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'PowerShell / WScript' = @{
        Pattern = '(?mi)^[^''\r\n]*\b(powershell|wscript|cscript|mshta)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Process / WMI' = @{
        Pattern = '(?m)^[^''\r\n]*\b(winmgmts|Win32_Process|WbemScripting|ExecQuery)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'DLL loading' = @{
        Pattern = '(?m)^[^''\r\n]*\b(LoadLibrary|GetProcAddress|FreeLibrary|CallByName)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Clipboard' = @{
        Pattern = '(?m)^[^''\r\n]*\b(MSForms\.DataObject|GetClipboardData|SetClipboardData)\b'
        Extract = { param($m) $m.Groups[1].Value }
    }
    'Environment' = @{
        Pattern = '(?m)^[^''\r\n]*\b(Environ\s*\$?\s*\()'
        Extract = { param($m) "Environ" }
    }
    'Auto-execution' = @{
        Pattern = '(?m)^\s*(Sub\s+(Auto_Open|Auto_Close|Workbook_Open|Workbook_BeforeClose|Document_Open|Document_Close)\b)'
        Extract = { param($m) $m.Groups[2].Value }
    }
    'Encoding / obfuscation' = @{
        Pattern = '(?m)^[^''\r\n]*\b(Chr\s*\$?\s*\(\s*\d+\s*\))'
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

# ============================================================================
# Detailed usage trace: COM objects and Win32 API call sites
# ============================================================================

# Collect all code as lines per file
$allCode = @{}
foreach ($f in $allFiles) {
    $allCode[$f.Name] = ([IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)) -split "`r`n|`n"
}

# --- COM object usage trace ---
# 1. Find Set <var> = CreateObject("<progid>") or GetObject(...)
# 2. Track <var>.method calls across all modules
$comBindings = [System.Collections.ArrayList]::new()  # @{ ProgId; VarName; File; Line }
foreach ($f in $allFiles) {
    $lines = $allCode[$f.Name]
    for ($li = 0; $li -lt $lines.Count; $li++) {
        $line = $lines[$li]
        if ($line -match '^\s*''') { continue }
        if ($line -match '\bSet\s+(\w+)\s*=\s*CreateObject\s*\(\s*"([^"]+)"') {
            [void]$comBindings.Add(@{ VarName = $Matches[1]; ProgId = $Matches[2]; File = $f.Name; Line = $li + 1 })
        }
        elseif ($line -match '\bSet\s+(\w+)\s*=\s*GetObject\s*\(') {
            [void]$comBindings.Add(@{ VarName = $Matches[1]; ProgId = '(GetObject)'; File = $f.Name; Line = $li + 1 })
        }
        # Also catch: CreateObject without Set (function return)
        elseif ($line -match '(\w+)\.(\w+).*CreateObject\s*\(\s*"([^"]+)"' -and $line -notmatch '^\s*Set\s') {
            # inline usage like: CreateObject("...").Method
        }
    }
}

if ($comBindings.Count -gt 0) {
    [void]$report.AppendLine("## COM Object Usage Details")
    [void]$report.AppendLine("")

    $grouped = $comBindings | Group-Object { $_.ProgId } | Sort-Object Name
    foreach ($g in $grouped) {
        [void]$report.AppendLine("  $($g.Name)")
        foreach ($b in $g.Group) {
            [void]$report.AppendLine("    $($b.File) L$($b.Line): Set $($b.VarName) = ...")
        }

        # Find all method/property calls on these variable names across all files
        $varNames = ($g.Group | ForEach-Object { $_.VarName }) | Sort-Object -Unique
        $methodCalls = [System.Collections.ArrayList]::new()
        foreach ($fn in $allCode.Keys) {
            $lines = $allCode[$fn]
            for ($li = 0; $li -lt $lines.Count; $li++) {
                $line = $lines[$li]
                if ($line -match '^\s*''') { continue }
                foreach ($vn in $varNames) {
                    if ($line -match "\b$([regex]::Escape($vn))\.(\w+)") {
                        $method = $Matches[1]
                        $trimmed = $line.Trim()
                        if ($trimmed.Length -gt 80) { $trimmed = $trimmed.Substring(0, 77) + '...' }
                        $entry = "$fn L$($li+1): .$method  -- $trimmed"
                        if ($methodCalls -notcontains $entry) { [void]$methodCalls.Add($entry) }
                    }
                }
            }
        }
        if ($methodCalls.Count -gt 0) {
            foreach ($mc in $methodCalls) { [void]$report.AppendLine("    $mc") }
        }
        [void]$report.AppendLine("")
    }
}

# --- Win32 API usage trace ---
$apiDecls = [System.Collections.ArrayList]::new()  # @{ Name; File; Line; Signature }
foreach ($f in $allFiles) {
    $lines = $allCode[$f.Name]
    for ($li = 0; $li -lt $lines.Count; $li++) {
        $line = $lines[$li]
        if ($line -match '^\s*''') { continue }
        if ($line -match '(?i)\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)\s+Lib\s+(.*)') {
            $trimmed = $line.Trim()
            if ($trimmed.Length -gt 100) { $trimmed = $trimmed.Substring(0, 97) + '...' }
            [void]$apiDecls.Add(@{ Name = $Matches[3]; File = $f.Name; Line = $li + 1; Sig = $trimmed })
        }
    }
}

if ($apiDecls.Count -gt 0) {
    [void]$report.AppendLine("## Win32 API Usage Details")
    [void]$report.AppendLine("")

    foreach ($api in $apiDecls) {
        [void]$report.AppendLine("  $($api.Name)")
        [void]$report.AppendLine("    $($api.File) L$($api.Line): $($api.Sig)")

        # Find call sites across all files
        foreach ($fn in $allCode.Keys) {
            $lines = $allCode[$fn]
            for ($li = 0; $li -lt $lines.Count; $li++) {
                $line = $lines[$li]
                if ($line -match '^\s*''') { continue }
                if ($line -match '(?i)\bDeclare\s') { continue }  # skip the declaration itself
                if ($line -match "\b$([regex]::Escape($api.Name))\b") {
                    $trimmed = $line.Trim()
                    if ($trimmed.Length -gt 80) { $trimmed = $trimmed.Substring(0, 77) + '...' }
                    [void]$report.AppendLine("    $fn L$($li+1): $trimmed")
                }
            }
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
    foreach ($m in [regex]::Matches($c, '(?m)^[^''\r\n]*\bCreateObject\s*\(\s*"([^"]+)"')) {
        $v = $m.Groups[1].Value; if ($allProgIds -notcontains $v) { [void]$allProgIds.Add($v) }
    }
    foreach ($m in [regex]::Matches($c, '(?m)^[^''\r\n]*\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)')) {
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
