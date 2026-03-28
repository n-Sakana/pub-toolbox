param([string]$FilePath)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

$configDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'config'
$configPath = Join-Path $configDir 'sanitize.json'

# ============================================================================
# Load analysis patterns from shared engine (for labels)
# ============================================================================

# Dummy project to get pattern names
$dummyProject = @{ Modules = [ordered]@{}; Ole2 = $null }
$dummyAnalysis = Get-VbaAnalysis -Project $dummyProject
$edrPatternNames = @($dummyAnalysis.Patterns.Keys)
$compatPatternNames = @($dummyAnalysis.CompatPatterns.Keys)

# ============================================================================
# Config load/save helpers
# ============================================================================

function Load-SanitizeConfig {
    param([string]$Path)
    if (Test-Path $Path) {
        $json = Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        $cfg = @{ edr = @{}; compat = @{} }
        if ($json.edr) {
            $json.edr.PSObject.Properties | ForEach-Object { $cfg.edr[$_.Name] = [bool]$_.Value }
        }
        if ($json.compat) {
            $json.compat.PSObject.Properties | ForEach-Object { $cfg.compat[$_.Name] = [bool]$_.Value }
        }
        return $cfg
    }
    # Default config
    return @{
        edr = @{ 'Win32 API (Declare)' = $true; 'DLL loading' = $true }
        compat = @{}
    }
}

function Save-SanitizeConfig {
    param([hashtable]$Config, [string]$Path)
    $dir = Split-Path $Path -Parent
    if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory -Force | Out-Null }
    $obj = [ordered]@{ edr = [ordered]@{}; compat = [ordered]@{} }
    foreach ($k in ($Config.edr.Keys | Sort-Object)) { $obj.edr[$k] = $Config.edr[$k] }
    foreach ($k in ($Config.compat.Keys | Sort-Object)) { $obj.compat[$k] = $Config.compat[$k] }
    $obj | ConvertTo-Json -Depth 3 | Set-Content $Path -Encoding UTF8
}

# ============================================================================
# Mode 1: No file argument → Settings GUI
# ============================================================================

if (-not $FilePath) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $cfg = Load-SanitizeConfig $configPath

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Sanitize Settings'
    $form.Size = New-Object System.Drawing.Size(520, 620)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#252526')
    $form.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    # --- EDR Risks GroupBox ---
    $grpEdr = New-Object System.Windows.Forms.GroupBox
    $grpEdr.Text = 'EDR Risks'
    $grpEdr.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
    $grpEdr.Location = New-Object System.Drawing.Point(12, 12)
    $grpEdr.Size = New-Object System.Drawing.Size(480, 240)
    $form.Controls.Add($grpEdr)

    $edrCheckboxes = @{}
    $row = 0; $col = 0
    foreach ($name in $edrPatternNames) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $name
        $cb.Size = New-Object System.Drawing.Size(225, 22)
        $x = 12 + ($col * 236)
        $y = 22 + ($row * 24)
        $cb.Location = New-Object System.Drawing.Point($x, $y)
        $cb.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
        if ($cfg.edr.ContainsKey($name) -and $cfg.edr[$name]) { $cb.Checked = $true }
        $grpEdr.Controls.Add($cb)
        $edrCheckboxes[$name] = $cb
        $col++
        if ($col -ge 2) { $col = 0; $row++ }
    }

    # --- Compat Risks GroupBox ---
    $grpCompat = New-Object System.Windows.Forms.GroupBox
    $grpCompat.Text = 'Compatibility Risks'
    $grpCompat.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
    $grpCompat.Location = New-Object System.Drawing.Point(12, 264)
    $grpCompat.Size = New-Object System.Drawing.Size(480, 220)
    $form.Controls.Add($grpCompat)

    $compatCheckboxes = @{}
    $row = 0; $col = 0
    foreach ($name in $compatPatternNames) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $name
        $cb.Size = New-Object System.Drawing.Size(225, 22)
        $x = 12 + ($col * 236)
        $y = 22 + ($row * 24)
        $cb.Location = New-Object System.Drawing.Point($x, $y)
        $cb.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
        if ($cfg.compat.ContainsKey($name) -and $cfg.compat[$name]) { $cb.Checked = $true }
        $grpCompat.Controls.Add($cb)
        $compatCheckboxes[$name] = $cb
        $col++
        if ($col -ge 2) { $col = 0; $row++ }
    }

    # --- OK / Cancel buttons ---
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Size = New-Object System.Drawing.Size(90, 30)
    $btnOk.Location = New-Object System.Drawing.Point(300, 500)
    $btnOk.FlatStyle = 'Flat'
    $btnOk.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#0e639c')
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)
    $form.AcceptButton = $btnOk

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(400, 500)
    $btnCancel.FlatStyle = 'Flat'
    $btnCancel.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#3c3c3c')
    $btnCancel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $form.CancelButton = $btnCancel

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $newCfg = @{ edr = @{}; compat = @{} }
        foreach ($name in $edrCheckboxes.Keys) {
            if ($edrCheckboxes[$name].Checked) { $newCfg.edr[$name] = $true }
        }
        foreach ($name in $compatCheckboxes.Keys) {
            if ($compatCheckboxes[$name].Checked) { $newCfg.compat[$name] = $true }
        }
        Save-SanitizeConfig $newCfg $configPath
        Write-Host "Settings saved to $configPath" -ForegroundColor Green
    } else {
        Write-Host "Cancelled." -ForegroundColor Yellow
    }
    $form.Dispose()
    exit 0
}

# ============================================================================
# Mode 2: File argument → Run sanitization
# ============================================================================

$FilePath = Resolve-VbaFilePath $FilePath
$fileName = [IO.Path]::GetFileName($FilePath)
$sw = [System.Diagnostics.Stopwatch]::StartNew()

Write-VbaHeader 'Sanitize' $fileName
Write-VbaLog 'Sanitize' $FilePath 'Started'

# Load config
$cfg = Load-SanitizeConfig $configPath

# Build active rules from shared analysis patterns
$analysis = Get-VbaAnalysis -Project @{ Modules = [ordered]@{}; Ole2 = $null }
$rules = [System.Collections.ArrayList]::new()
foreach ($name in $analysis.Patterns.Keys) {
    if ($cfg.edr.ContainsKey($name) -and $cfg.edr[$name]) {
        [void]$rules.Add(@{ Name = $name; Pattern = $analysis.Patterns[$name].Pattern; Prefix = "' [EDR] " })
    }
}
foreach ($name in $analysis.CompatPatterns.Keys) {
    if ($cfg.compat.ContainsKey($name) -and $cfg.compat[$name]) {
        [void]$rules.Add(@{ Name = $name; Pattern = $analysis.CompatPatterns[$name].Pattern; Prefix = "' [COMPAT] " })
    }
}

if ($rules.Count -eq 0) {
    Write-Host "No sanitization rules enabled. Run Sanitize.bat without arguments to configure." -ForegroundColor Yellow
    exit 0
}

Write-VbaStatus 'Sanitize' $fileName "Active rules: $($rules.Count)"

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

# Pass 1: collect declared API names (for call-site commenting)
$declaredNames = [System.Collections.ArrayList]::new()
$hasApiDeclareRule = $false
foreach ($r in $rules) { if ($r.Name -eq 'Win32 API (Declare)') { $hasApiDeclareRule = $true; break } }
if ($hasApiDeclareRule) {
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
    if (-not $md.Entry) { continue }
    $lines = $md.Code -split "`r`n|`n"
    $changes = 0

    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ($lines[$i] -match '^\s*''') { continue }
        $matched = $false; $ruleName = ''; $prefix = "' [EDR] "
        foreach ($rule in $rules) {
            if ($lines[$i] -match $rule.Pattern) { $matched = $true; $ruleName = $rule.Name; $prefix = $rule.Prefix; break }
        }
        if (-not $matched -and $declaredNames.Count -gt 0) {
            foreach ($apiName in $declaredNames) {
                if ($lines[$i] -match "\b$([regex]::Escape($apiName))\b") { $matched = $true; $ruleName = "Call to $apiName"; $prefix = "' [EDR] "; break }
            }
        }
        if ($matched) {
            $lines[$i] = $prefix + $lines[$i].TrimStart()
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

    $prefixBytes = New-Object byte[] $md.Offset
    [Array]::Copy($md.StreamData, $prefixBytes, $md.Offset)
    $newStream = New-Object byte[] ($prefixBytes.Length + $newCompressed.Length)
    [Array]::Copy($prefixBytes, $newStream, $prefixBytes.Length)
    [Array]::Copy($newCompressed, 0, $newStream, $prefixBytes.Length, $newCompressed.Length)

    Write-Ole2Stream $ole2Bytes $project.Ole2 $md.Entry $newStream
}

if ($totalChanges -gt 0) {
    Save-VbaProjectBytes $copyPath $ole2Bytes $project.IsZip
} else {
    Remove-Item $copyPath -Force
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
                if ($mod.Lines[$i] -match "^\s*'\s*\[EDR\]") { $highlights[$i] = 'hl-edr' }
                elseif ($mod.Lines[$i] -match "^\s*'\s*\[COMPAT\]") { $highlights[$i] = 'hl-compat' }
            }
            $htmlModules[$modName] = @{ Ext = $mod.Ext; Lines = [System.Collections.ArrayList]::new($mod.Lines); Highlights = $highlights }
        }
        New-HtmlCodeView -title "VBA Sanitize: $fileName" -subtitle "$totalChanges line(s) commented out" `
            -moduleData $htmlModules -highlightClass 'hl-edr' -highlightColor '#1b2e4a' -highlightText '#a0c4f0' -markerColor '#e8ab53' `
            -outputPath (Join-Path $outDir 'sanitize.html')

        # Inject compat CSS
        $htmlPath = Join-Path $outDir 'sanitize.html'
        $htmlContent = [IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
        $compatCss = "tr.hl-compat td.code { background: #3a1b4a; color: #c4a0f0; }`ntr.hl-compat td.ln { color: #cccccc; }`n.minimap .mark.m-hl-compat { background: #9a5eff; }"
        $htmlContent = $htmlContent.Replace('</style>', "$compatCss`n</style>")
        [IO.File]::WriteAllText($htmlPath, $htmlContent, [System.Text.Encoding]::UTF8)

        Start-Process $htmlPath
    }
}

$sw.Stop()
$msg = if ($totalChanges -gt 0) { "$totalChanges line(s) sanitized" } else { "No changes needed" }
Write-VbaResult 'Sanitize' $fileName $msg $outDir $sw.Elapsed.TotalSeconds
Write-VbaLog 'Sanitize' $FilePath "$totalChanges lines sanitized | -> $outDir"
