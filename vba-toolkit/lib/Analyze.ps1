param([string[]]$Paths)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

$configDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'config'
$configPath = Join-Path $configDir 'analyze.json'
$oldConfigPath = Join-Path $configDir 'sanitize.json'

# ============================================================================
# Config helpers
# ============================================================================

function Get-DefaultConfig {
    return @{
        edr = [ordered]@{
            'Win32 API (Declare)' = @{ detect = $true; sanitize = $true }
            'DLL loading' = @{ detect = $true; sanitize = $true }
            'COM / CreateObject' = @{ detect = $true; sanitize = $false }
            'Shell / process' = @{ detect = $true; sanitize = $false }
            'File I/O' = @{ detect = $true; sanitize = $false }
            'FileSystemObject' = @{ detect = $true; sanitize = $false }
            'Registry' = @{ detect = $true; sanitize = $false }
            'SendKeys' = @{ detect = $true; sanitize = $false }
            'Network / HTTP' = @{ detect = $true; sanitize = $false }
            'PowerShell / WScript' = @{ detect = $true; sanitize = $false }
            'Process / WMI' = @{ detect = $true; sanitize = $false }
            'Clipboard' = @{ detect = $true; sanitize = $false }
            'Environment' = @{ detect = $true; sanitize = $false }
            'Auto-execution' = @{ detect = $true; sanitize = $false }
            'Encoding / obfuscation' = @{ detect = $true; sanitize = $false }
        }
        compat = [ordered]@{
            '64-bit: Missing PtrSafe' = @{ detect = $true; sanitize = $false }
            '64-bit: Long for handles' = @{ detect = $true; sanitize = $false }
            '64-bit: VarPtr/ObjPtr/StrPtr' = @{ detect = $true; sanitize = $false }
            'Deprecated: DDE' = @{ detect = $true; sanitize = $false }
            'Deprecated: IE Automation' = @{ detect = $true; sanitize = $false }
            'Deprecated: Legacy Controls' = @{ detect = $true; sanitize = $false }
            'Deprecated: DAO' = @{ detect = $true; sanitize = $false }
            'Legacy: DefType' = @{ detect = $true; sanitize = $false }
            'Legacy: GoSub' = @{ detect = $true; sanitize = $false }
            'Legacy: While/Wend' = @{ detect = $true; sanitize = $false }
        }
    }
}

function Migrate-OldConfig {
    param([string]$OldPath, [string]$NewPath)
    if (-not (Test-Path $OldPath)) { return $null }
    if (Test-Path $NewPath) { return $null }
    $json = Get-Content $OldPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $def = Get-DefaultConfig
    if ($json.edr) {
        $json.edr.PSObject.Properties | ForEach-Object {
            if ($def.edr.Contains($_.Name)) {
                $def.edr[$_.Name] = @{ detect = $true; sanitize = [bool]$_.Value }
            }
        }
    }
    if ($json.compat) {
        $json.compat.PSObject.Properties | ForEach-Object {
            if ($def.compat.Contains($_.Name)) {
                $def.compat[$_.Name] = @{ detect = $true; sanitize = [bool]$_.Value }
            }
        }
    }
    return $def
}

function Load-AnalyzeConfig {
    param([string]$Path)
    # Try migration first
    $migrated = Migrate-OldConfig $oldConfigPath $Path
    if ($migrated) {
        Save-AnalyzeConfig $migrated $Path
        return $migrated
    }
    if (Test-Path $Path) {
        $json = Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        $cfg = Get-DefaultConfig
        if ($json.edr) {
            $json.edr.PSObject.Properties | ForEach-Object {
                if ($cfg.edr.Contains($_.Name)) {
                    $d = $true; $s = $false
                    if ($null -ne $_.Value.detect) { $d = [bool]$_.Value.detect }
                    if ($null -ne $_.Value.sanitize) { $s = [bool]$_.Value.sanitize }
                    $cfg.edr[$_.Name] = @{ detect = $d; sanitize = $s }
                }
            }
        }
        if ($json.compat) {
            $json.compat.PSObject.Properties | ForEach-Object {
                if ($cfg.compat.Contains($_.Name)) {
                    $d = $true; $s = $false
                    if ($null -ne $_.Value.detect) { $d = [bool]$_.Value.detect }
                    if ($null -ne $_.Value.sanitize) { $s = [bool]$_.Value.sanitize }
                    $cfg.compat[$_.Name] = @{ detect = $d; sanitize = $s }
                }
            }
        }
        return $cfg
    }
    return Get-DefaultConfig
}

function Save-AnalyzeConfig {
    param([hashtable]$Config, [string]$Path)
    $dir = Split-Path $Path -Parent
    if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory -Force | Out-Null }
    $obj = [ordered]@{ edr = [ordered]@{}; compat = [ordered]@{} }
    foreach ($k in $Config.edr.Keys) {
        $obj.edr[$k] = [ordered]@{ detect = $Config.edr[$k].detect; sanitize = $Config.edr[$k].sanitize }
    }
    foreach ($k in $Config.compat.Keys) {
        $obj.compat[$k] = [ordered]@{ detect = $Config.compat[$k].detect; sanitize = $Config.compat[$k].sanitize }
    }
    $obj | ConvertTo-Json -Depth 3 | Set-Content $Path -Encoding UTF8
}

# ============================================================================
# Mode 1: No args -> Settings GUI
# ============================================================================

if (-not $Paths -or $Paths.Count -eq 0) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $cfg = Load-AnalyzeConfig $configPath

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Analyze Settings'
    $form.Size = New-Object System.Drawing.Size(600, 700)
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
    $grpEdr.Size = New-Object System.Drawing.Size(560, 290)
    $form.Controls.Add($grpEdr)

    # Column headers
    $lblDetect1 = New-Object System.Windows.Forms.Label
    $lblDetect1.Text = 'Detect'
    $lblDetect1.Location = New-Object System.Drawing.Point(220, 18)
    $lblDetect1.Size = New-Object System.Drawing.Size(50, 16)
    $lblDetect1.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#888888')
    $lblDetect1.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $grpEdr.Controls.Add($lblDetect1)

    $lblSanitize1 = New-Object System.Windows.Forms.Label
    $lblSanitize1.Text = 'Sanitize'
    $lblSanitize1.Location = New-Object System.Drawing.Point(280, 18)
    $lblSanitize1.Size = New-Object System.Drawing.Size(55, 16)
    $lblSanitize1.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#888888')
    $lblSanitize1.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $grpEdr.Controls.Add($lblSanitize1)

    $edrControls = @{}
    $row = 0
    foreach ($name in $cfg.edr.Keys) {
        $y = 36 + ($row * 24)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $name
        $lbl.Location = New-Object System.Drawing.Point(12, ($y + 2))
        $lbl.Size = New-Object System.Drawing.Size(200, 20)
        $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
        $grpEdr.Controls.Add($lbl)

        $cbD = New-Object System.Windows.Forms.CheckBox
        $cbD.Location = New-Object System.Drawing.Point(232, $y)
        $cbD.Size = New-Object System.Drawing.Size(20, 20)
        $cbD.Checked = $cfg.edr[$name].detect
        $grpEdr.Controls.Add($cbD)

        $cbS = New-Object System.Windows.Forms.CheckBox
        $cbS.Location = New-Object System.Drawing.Point(295, $y)
        $cbS.Size = New-Object System.Drawing.Size(20, 20)
        $cbS.Checked = $cfg.edr[$name].sanitize
        $grpEdr.Controls.Add($cbS)

        $edrControls[$name] = @{ Detect = $cbD; Sanitize = $cbS }
        $row++
    }

    # --- Compat Risks GroupBox ---
    $grpCompat = New-Object System.Windows.Forms.GroupBox
    $grpCompat.Text = 'Compatibility Risks'
    $grpCompat.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
    $grpCompat.Location = New-Object System.Drawing.Point(12, 314)
    $grpCompat.Size = New-Object System.Drawing.Size(560, 260)
    $form.Controls.Add($grpCompat)

    $lblDetect2 = New-Object System.Windows.Forms.Label
    $lblDetect2.Text = 'Detect'
    $lblDetect2.Location = New-Object System.Drawing.Point(220, 18)
    $lblDetect2.Size = New-Object System.Drawing.Size(50, 16)
    $lblDetect2.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#888888')
    $lblDetect2.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $grpCompat.Controls.Add($lblDetect2)

    $lblSanitize2 = New-Object System.Windows.Forms.Label
    $lblSanitize2.Text = 'Sanitize'
    $lblSanitize2.Location = New-Object System.Drawing.Point(280, 18)
    $lblSanitize2.Size = New-Object System.Drawing.Size(55, 16)
    $lblSanitize2.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#888888')
    $lblSanitize2.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $grpCompat.Controls.Add($lblSanitize2)

    $compatControls = @{}
    $row = 0
    foreach ($name in $cfg.compat.Keys) {
        $y = 36 + ($row * 24)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $name
        $lbl.Location = New-Object System.Drawing.Point(12, ($y + 2))
        $lbl.Size = New-Object System.Drawing.Size(200, 20)
        $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
        $grpCompat.Controls.Add($lbl)

        $cbD = New-Object System.Windows.Forms.CheckBox
        $cbD.Location = New-Object System.Drawing.Point(232, $y)
        $cbD.Size = New-Object System.Drawing.Size(20, 20)
        $cbD.Checked = $cfg.compat[$name].detect
        $grpCompat.Controls.Add($cbD)

        $cbS = New-Object System.Windows.Forms.CheckBox
        $cbS.Location = New-Object System.Drawing.Point(295, $y)
        $cbS.Size = New-Object System.Drawing.Size(20, 20)
        $cbS.Checked = $cfg.compat[$name].sanitize
        $grpCompat.Controls.Add($cbS)

        $compatControls[$name] = @{ Detect = $cbD; Sanitize = $cbS }
        $row++
    }

    # --- OK / Cancel ---
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Size = New-Object System.Drawing.Size(90, 30)
    $btnOk.Location = New-Object System.Drawing.Point(380, 585)
    $btnOk.FlatStyle = 'Flat'
    $btnOk.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#0e639c')
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)
    $form.AcceptButton = $btnOk

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(480, 585)
    $btnCancel.FlatStyle = 'Flat'
    $btnCancel.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#3c3c3c')
    $btnCancel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $form.CancelButton = $btnCancel

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $newCfg = @{ edr = [ordered]@{}; compat = [ordered]@{} }
        foreach ($name in $edrControls.Keys) {
            $newCfg.edr[$name] = @{ detect = $edrControls[$name].Detect.Checked; sanitize = $edrControls[$name].Sanitize.Checked }
        }
        foreach ($name in $compatControls.Keys) {
            $newCfg.compat[$name] = @{ detect = $compatControls[$name].Detect.Checked; sanitize = $compatControls[$name].Sanitize.Checked }
        }
        Save-AnalyzeConfig $newCfg $configPath
        Write-Host "Settings saved to $configPath" -ForegroundColor Green
    } else {
        Write-Host "Cancelled." -ForegroundColor Yellow
    }
    $form.Dispose()
    exit 0
}

# ============================================================================
# Mode 2/3: File/Folder analysis
# ============================================================================

$sw = [System.Diagnostics.Stopwatch]::StartNew()

# Collect all xlsm/xlam/xls files
$files = [System.Collections.ArrayList]::new()
$baseDir = $null

foreach ($p in $Paths) {
    $resolved = (Resolve-Path $p -ErrorAction SilentlyContinue).Path
    if (-not $resolved) { Write-VbaError 'Analyze' $p 'Path not found'; continue }

    if (Test-Path $resolved -PathType Container) {
        if (-not $baseDir) { $baseDir = $resolved }
        Get-ChildItem $resolved -Recurse -File -Include '*.xlsm','*.xlam','*.xls' | ForEach-Object {
            [void]$files.Add($_.FullName)
        }
    } else {
        $ext = [IO.Path]::GetExtension($resolved).ToLower()
        if ($ext -in '.xls','.xlsm','.xlam') {
            if (-not $baseDir) { $baseDir = [IO.Path]::GetDirectoryName($resolved) }
            [void]$files.Add($resolved)
        }
    }
}

if ($files.Count -eq 0) {
    Write-Host "No Excel files found." -ForegroundColor Yellow
    exit 0
}

if (-not $baseDir) { $baseDir = [IO.Path]::GetDirectoryName($files[0]) }

# Load config
$cfg = Load-AnalyzeConfig $configPath

# Build active rules from analysis patterns
$dummyAnalysis = Get-VbaAnalysis -Project @{ Modules = [ordered]@{}; Ole2 = $null }
$sanitizeRules = [System.Collections.ArrayList]::new()
$detectRules = [System.Collections.ArrayList]::new()
$anySanitize = $false

foreach ($name in $dummyAnalysis.Patterns.Keys) {
    $edrCfg = $cfg.edr[$name]
    if ($edrCfg -and $edrCfg.sanitize) {
        [void]$sanitizeRules.Add(@{ Name = $name; Pattern = $dummyAnalysis.Patterns[$name].Pattern; Prefix = "' [EDR] "; Category = 'edr' })
        $anySanitize = $true
    }
    if ($edrCfg -and $edrCfg.detect) {
        [void]$detectRules.Add(@{ Name = $name; Pattern = $dummyAnalysis.Patterns[$name].Pattern; Category = 'edr' })
    }
}
foreach ($name in $dummyAnalysis.CompatPatterns.Keys) {
    $cCfg = $cfg.compat[$name]
    if ($cCfg -and $cCfg.sanitize) {
        [void]$sanitizeRules.Add(@{ Name = $name; Pattern = $dummyAnalysis.CompatPatterns[$name].Pattern; Prefix = "' [COMPAT] "; Category = 'compat' })
        $anySanitize = $true
    }
    if ($cCfg -and $cCfg.detect) {
        [void]$detectRules.Add(@{ Name = $name; Pattern = $dummyAnalysis.CompatPatterns[$name].Pattern; Category = 'compat' })
    }
}

# Create output directory
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outputRoot = Join-Path $baseDir 'output'
$outDir = Join-Path $outputRoot "${timestamp}_analyze"
New-Item $outDir -ItemType Directory -Force | Out-Null

# Detect filename collisions for prefix logic
$fileNameCounts = @{}
foreach ($f in $files) {
    $fn = [IO.Path]::GetFileNameWithoutExtension($f)
    if ($fileNameCounts.ContainsKey($fn)) { $fileNameCounts[$fn]++ } else { $fileNameCounts[$fn] = 1 }
}

$replacements = Get-VbaApiReplacements
$csvRows = [System.Collections.ArrayList]::new()
$he = { param($s) [System.Net.WebUtility]::HtmlEncode($s) }
$processed = 0

foreach ($filePath in $files) {
    $processed++
    $fileName = [IO.Path]::GetFileName($filePath)
    $baseName = [IO.Path]::GetFileNameWithoutExtension($filePath)
    $fileExt = [IO.Path]::GetExtension($filePath)

    # Determine output prefix for colliding names
    $outPrefix = $baseName
    if ($fileNameCounts[$baseName] -gt 1) {
        $parentDir = Split-Path (Split-Path $filePath -Parent) -Leaf
        $outPrefix = "${parentDir}_${baseName}"
    }

    # Relative path from base
    $relPath = $filePath
    if ($filePath.StartsWith($baseDir)) {
        $relPath = $filePath.Substring($baseDir.Length).TrimStart('\', '/')
    }

    Write-VbaHeader 'Analyze' $fileName

    $csvRow = [ordered]@{
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        RelativePath = [IO.Path]::GetDirectoryName($relPath)
        FileName = $fileName
        Bas = 0; Cls = 0; Frm = 0; TotalModules = 0; CodeLines = 0
        EdrIssues = 0; CompatIssues = 0; SanitizedLines = 0
        References = ''; Error = ''
    }

    try {
        # Load project
        $project = if ($anySanitize) {
            Get-AllModuleCode $filePath -IncludeRawData
        } else {
            Get-AllModuleCode $filePath -StripAttributes
        }
        if (-not $project) {
            $csvRow.Error = 'No VBA project'
            [void]$csvRows.Add($csvRow)
            Write-VbaError 'Analyze' $fileName 'No vbaProject.bin found'
            continue
        }

        # Module counts
        foreach ($modName in $project.Modules.Keys) {
            $mod = $project.Modules[$modName]
            $csvRow.TotalModules++
            switch ($mod.Ext) { 'bas' { $csvRow.Bas++ } 'cls' { $csvRow.Cls++ } 'frm' { $csvRow.Frm++ } }
            $csvRow.CodeLines += $mod.Lines.Count
        }

        Write-VbaStatus 'Analyze' $fileName "Modules: $($csvRow.TotalModules) ($($csvRow.Bas) bas, $($csvRow.Cls) cls, $($csvRow.Frm) frm)"

        # Analysis
        $analysis = Get-VbaAnalysis -Project $project
        $csvRow.EdrIssues = $analysis.IssueCount
        $csvRow.CompatIssues = $analysis.CompatIssueCount
        $csvRow.References = $analysis.ExternalRefs -join '; '

        Write-VbaStatus 'Analyze' $fileName "EDR issues: $($csvRow.EdrIssues)"
        Write-VbaStatus 'Analyze' $fileName "Compat issues: $($csvRow.CompatIssues)"

        # === Sanitize pass ===
        $totalSanitized = 0
        $sanitizedLineMap = @{}  # "modName:lineIdx" -> rule name

        # Collect declared API names for call-site sanitizing
        $declaredNames = [System.Collections.ArrayList]::new()
        $hasApiDeclareRule = $false
        foreach ($r in $sanitizeRules) { if ($r.Name -eq 'Win32 API (Declare)') { $hasApiDeclareRule = $true; break } }
        if ($hasApiDeclareRule) {
            foreach ($modName in $project.Modules.Keys) {
                $mod = $project.Modules[$modName]
                foreach ($m in [regex]::Matches($mod.Code, '(?im)^\s*(?:Public\s+|Private\s+)?Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+(\w+)')) {
                    $n = $m.Groups[1].Value
                    if ($declaredNames -notcontains $n) { [void]$declaredNames.Add($n) }
                }
            }
        }

        if ($anySanitize -and $sanitizeRules.Count -gt 0) {
            # Copy file to output
            $copyPath = Join-Path $outDir "$outPrefix$fileExt"
            Copy-Item $filePath $copyPath -Force

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
                    foreach ($rule in $sanitizeRules) {
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
                        $sanitizedLineMap["${modName}:$i"] = $ruleName
                    }
                }

                if ($changes -eq 0) { continue }
                $totalSanitized += $changes

                $newCode = $lines -join "`r`n"
                $newCodeBytes = $encoding.GetBytes($newCode)
                $newCompressed = Compress-VBA $newCodeBytes

                $prefixBytes = New-Object byte[] $md.Offset
                [Array]::Copy($md.StreamData, $prefixBytes, $md.Offset)
                $newStream = New-Object byte[] ($prefixBytes.Length + $newCompressed.Length)
                [Array]::Copy($prefixBytes, $newStream, $prefixBytes.Length)
                [Array]::Copy($newCompressed, 0, $newStream, $prefixBytes.Length, $newCompressed.Length)

                Write-Ole2Stream $ole2Bytes $project.Ole2 $md.Entry $newStream

                # Update module lines in project for HTML/report generation
                $md.Lines = $lines
                $md.Code = $newCode
            }

            if ($totalSanitized -gt 0) {
                Save-VbaProjectBytes $copyPath $ole2Bytes $project.IsZip
            } else {
                Remove-Item $copyPath -Force -ErrorAction SilentlyContinue
            }
        }

        $csvRow.SanitizedLines = $totalSanitized
        Write-VbaStatus 'Analyze' $fileName "Sanitized: $totalSanitized lines"

        # === Build line-level highlight data for HTML ===
        # Re-read modules (use sanitized lines if available)
        $allModLines = [ordered]@{}
        foreach ($modName in $project.Modules.Keys) {
            $mod = $project.Modules[$modName]
            $lines = if ($mod.Lines -is [array]) { $mod.Lines } else { @($mod.Lines) }
            # Strip attributes for display
            $displayLines = @($lines | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_' })
            $allModLines[$modName] = @{ Ext = $mod.Ext; Lines = $displayLines }
        }

        # Build highlight map per module: lineIdx -> @{ color, category, patternName }
        $modHighlights = @{}
        foreach ($modName in $allModLines.Keys) {
            $ml = $allModLines[$modName]
            $hlMap = @{}
            for ($i = 0; $i -lt $ml.Lines.Count; $i++) {
                $line = $ml.Lines[$i]

                # Sanitized lines (yellow, highest priority)
                if ($line -match "^\s*'\s*\[EDR\]") {
                    $hlMap[$i] = @{ Color = 'hl-sanitized'; Category = 'edr'; PatternName = 'Sanitized (EDR)' }
                    continue
                }
                if ($line -match "^\s*'\s*\[COMPAT\]") {
                    $hlMap[$i] = @{ Color = 'hl-sanitized'; Category = 'compat'; PatternName = 'Sanitized (Compat)' }
                    continue
                }

                if ($line -match '^\s*''') { continue }

                # Detect EDR (blue)
                $foundRule = $null
                foreach ($rule in $detectRules) {
                    if ($line -match $rule.Pattern) { $foundRule = $rule; break }
                }
                if ($foundRule) {
                    if ($foundRule.Category -eq 'edr') {
                        $hlMap[$i] = @{ Color = 'hl-edr'; Category = 'edr'; PatternName = $foundRule.Name }
                    } else {
                        $hlMap[$i] = @{ Color = 'hl-compat'; Category = 'compat'; PatternName = $foundRule.Name }
                    }
                    continue
                }

                # Check API call sites
                foreach ($apiName in $analysis.ApiCallNames) {
                    if ($line -match "\b$([regex]::Escape($apiName))\b") {
                        $hlMap[$i] = @{ Color = 'hl-edr'; Category = 'edr'; PatternName = "API: $apiName" }
                        break
                    }
                }
            }
            $modHighlights[$modName] = $hlMap
        }

        # === Generate analyze.txt ===
        $txtSb = [System.Text.StringBuilder]::new()
        [void]$txtSb.AppendLine("# VBA Analysis Report")
        [void]$txtSb.AppendLine("# Source: $fileName")
        [void]$txtSb.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        [void]$txtSb.AppendLine("")
        [void]$txtSb.AppendLine("## Modules ($($csvRow.TotalModules))")
        foreach ($modName in $allModLines.Keys) {
            $ml = $allModLines[$modName]
            [void]$txtSb.AppendLine("  $modName.$($ml.Ext) ($($ml.Lines.Count) lines)")
        }
        [void]$txtSb.AppendLine("  Total: $($csvRow.CodeLines) lines")
        [void]$txtSb.AppendLine("")

        if ($analysis.Findings.Count -gt 0) {
            [void]$txtSb.AppendLine("## EDR Risks ($($analysis.IssueCount))")
            foreach ($cat in $analysis.Findings.Keys) {
                $f = $analysis.Findings[$cat]
                [void]$txtSb.AppendLine("  $cat ($($f.Findings.Count))")
                if ($f.Aggregate) {
                    [void]$txtSb.AppendLine("    (aggregated: $($f.Findings.Count) occurrences)")
                } else {
                    foreach ($finding in $f.Findings) { [void]$txtSb.AppendLine("    $finding") }
                }
            }
            [void]$txtSb.AppendLine("")
        }

        if ($analysis.CompatFindings.Count -gt 0) {
            [void]$txtSb.AppendLine("## Compatibility Risks ($($analysis.CompatIssueCount))")
            foreach ($cat in $analysis.CompatFindings.Keys) {
                $f = $analysis.CompatFindings[$cat]
                [void]$txtSb.AppendLine("  $cat ($($f.Findings.Count))")
                foreach ($finding in $f.Findings) { [void]$txtSb.AppendLine("    $finding") }
            }
            [void]$txtSb.AppendLine("")
        }

        if ($analysis.ComBindings.Count -gt 0) {
            [void]$txtSb.AppendLine("## COM Object Usage Details")
            $comByProg = @{}
            foreach ($b in $analysis.ComBindings) {
                if (-not $comByProg.ContainsKey($b.ProgId)) { $comByProg[$b.ProgId] = [System.Collections.ArrayList]::new() }
                [void]$comByProg[$b.ProgId].Add($b)
            }
            foreach ($prog in $comByProg.Keys) {
                [void]$txtSb.AppendLine("  $prog")
                foreach ($b in $comByProg[$prog]) {
                    [void]$txtSb.AppendLine("    $($b.File) L$($b.Line): Set $($b.VarName) = ...")
                }
            }
            [void]$txtSb.AppendLine("")
        }

        if ($analysis.ApiDecls.Count -gt 0) {
            [void]$txtSb.AppendLine("## Win32 API Usage Details")
            foreach ($decl in $analysis.ApiDecls) {
                [void]$txtSb.AppendLine("  $($decl.Name)")
                [void]$txtSb.AppendLine("    $($decl.File) L$($decl.Line): $($decl.Sig)")
                $info = $replacements[$decl.Name]
                if ($info) { [void]$txtSb.AppendLine("    Alternative: $($info.Alt)") }
            }
            [void]$txtSb.AppendLine("")
        }

        if ($analysis.ExternalRefs.Count -gt 0) {
            [void]$txtSb.AppendLine("## External References ($($analysis.ExternalRefs.Count))")
            foreach ($ref in $analysis.ExternalRefs) { [void]$txtSb.AppendLine("  $ref") }
            [void]$txtSb.AppendLine("")
        }

        [void]$txtSb.AppendLine("## Summary")
        [void]$txtSb.AppendLine("  $($csvRow.EdrIssues) EDR issue(s), $($csvRow.CompatIssues) compatibility issue(s), $totalSanitized line(s) sanitized.")

        [IO.File]::WriteAllText((Join-Path $outDir "${outPrefix}_analyze.txt"), $txtSb.ToString(), [System.Text.Encoding]::UTF8)

        # === Generate analyze.html ===
        # Build sidebar
        $sidebarSb = [System.Text.StringBuilder]::new()
        $modIdx = 0; $firstHlIdx = -1
        foreach ($modName in $allModLines.Keys) {
            $ml = $allModLines[$modName]
            $hlCount = 0
            if ($modHighlights[$modName]) { $hlCount = $modHighlights[$modName].Count }
            $cls = if ($hlCount -gt 0) { 'has-hl' } else { 'no-hl' }
            if ($firstHlIdx -eq -1 -and $hlCount -gt 0) { $firstHlIdx = $modIdx }
            $label = "$modName.$($ml.Ext)"
            if ($hlCount -gt 0) { $label += " ($hlCount)" }
            [void]$sidebarSb.Append("<div class=`"item $cls`" onclick=`"showTab($modIdx)`" id=`"tab$modIdx`">$(& $he $label)</div>")
            $modIdx++
        }
        if ($firstHlIdx -eq -1) { $firstHlIdx = 0 }

        # Build content
        $contentSb = [System.Text.StringBuilder]::new()
        $modIdx = 0
        foreach ($modName in $allModLines.Keys) {
            $ml = $allModLines[$modName]
            $hlMap = $modHighlights[$modName]
            [void]$contentSb.Append("<div class=`"module`" id=`"mod$modIdx`"><table class=`"code-table`">")
            for ($i = 0; $i -lt $ml.Lines.Count; $i++) {
                $trClass = ''
                $dataApi = ''
                if ($hlMap -and $hlMap.ContainsKey($i)) {
                    $hl = $hlMap[$i]
                    $trClass = $hl.Color
                    $dataApi = $hl.PatternName
                }
                $ln = $i + 1
                $dataAttr = if ($dataApi) { " data-api=`"$(& $he $dataApi)`"" } else { '' }
                [void]$contentSb.Append("<tr class=`"$trClass`"$dataAttr><td class=`"ln`">$ln</td><td class=`"code`">$(& $he $ml.Lines[$i])</td></tr>")
            }
            [void]$contentSb.Append("</table></div>")
            $modIdx++
        }

        # Build outline items (flat, sorted by module then line number)
        $outlineItems = [System.Collections.ArrayList]::new()
        foreach ($modName in $allModLines.Keys) {
            $hlMap = $modHighlights[$modName]
            if (-not $hlMap) { continue }
            $ext = $allModLines[$modName].Ext
            foreach ($lineIdx in ($hlMap.Keys | Sort-Object { [int]$_ })) {
                $hl = $hlMap[$lineIdx]
                $ln = [int]$lineIdx + 1
                $label = "L$ln $($hl.PatternName)"
                if ($label.Length -gt 50) { $label = $label.Substring(0, 47) + '...' }
                [void]$outlineItems.Add(@{ ModName = "$modName.$ext"; LineNum = $ln; Label = $label; Color = $hl.Color })
            }
        }

        # Build tooltip data for replacements
        $tooltipEntries = [System.Text.StringBuilder]::new()
        [void]$tooltipEntries.Append('{')
        $first = $true
        foreach ($apiName in @($analysis.ApiCallNames) + @($analysis.ApiDecls | ForEach-Object { $_.Name })) {
            $info = $replacements[$apiName]
            if (-not $info) { continue }
            $key = "API: $apiName"
            $altJs = ($info.Alt -replace '\\','\\\\' -replace "'","\'")
            $noteJs = ($info.Note -replace '\\','\\\\' -replace "'","\'")
            $exJs = ((& $he $info.Example) -replace '\\','\\\\' -replace "'","\'" -replace "`r`n",'\n' -replace "`n",'\n')
            $comma = if ($first) { '' } else { ',' }
            $first = $false
            [void]$tooltipEntries.Append("$comma'$(& $he $key)':{alt:'$altJs',note:'$noteJs',ex:'$exJs'}")
        }
        # Also add pattern-level replacements
        foreach ($patName in @($dummyAnalysis.Patterns.Keys) + @($dummyAnalysis.CompatPatterns.Keys)) {
            $info = $replacements[$patName]
            if (-not $info) { continue }
            $altJs = ($info.Alt -replace '\\','\\\\' -replace "'","\'")
            $noteJs = ($info.Note -replace '\\','\\\\' -replace "'","\'")
            $exJs = ((& $he $info.Example) -replace '\\','\\\\' -replace "'","\'" -replace "`r`n",'\n' -replace "`n",'\n')
            $comma = if ($first) { '' } else { ',' }
            $first = $false
            [void]$tooltipEntries.Append("$comma'$(& $he $patName)':{alt:'$altJs',note:'$noteJs',ex:'$exJs'}")
        }
        [void]$tooltipEntries.Append('}')

        # CSS
        $analyzeCss = @"
.sidebar .item.has-hl { color: #e8ab53; }
.sidebar .item.no-hl { color: #606060; }
.code-table { width: 100%; border-collapse: collapse; }
.code-table td { padding: 0 8px; line-height: 20px; vertical-align: top; white-space: pre; overflow: hidden; text-overflow: ellipsis; }
.code-table .ln { width: 50px; min-width: 50px; text-align: right; color: #606060; padding-right: 12px; user-select: none; border-right: 1px solid #3c3c3c; }
.code-table .code { color: #d4d4d4; }
tr.hl-sanitized td.code { background: #4b3a00; color: #f0d080; }
tr.hl-sanitized td.ln { color: #cccccc; }
tr.hl-edr td.code { background: #1b2e4a; color: #a0c4f0; cursor: pointer; }
tr.hl-edr td.ln { color: #cccccc; }
tr.hl-compat td.code { background: #3a1b4a; color: #c4a0f0; cursor: pointer; }
tr.hl-compat td.ln { color: #cccccc; }
.minimap { right: 250px; }
.minimap .mark.m-hl-sanitized { background: #e8ab53; }
.minimap .mark.m-hl-edr { background: #4fc1ff; }
.minimap .mark.m-hl-compat { background: #9a5eff; }
.outline { width: 250px; min-width: 250px; background: #252526; border-left: 1px solid #3c3c3c; overflow-y: auto; padding: 8px 0; }
.outline .ol-header { padding: 6px 12px; font-size: 11px; color: #888; text-transform: uppercase; }
.outline .ol-item { padding: 3px 12px; font-size: 12px; cursor: pointer; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.outline .ol-item:hover { background: #2a2d2e; }
.outline .ol-item.c-sanitized { color: #e8ab53; }
.outline .ol-item.c-edr { color: #4fc1ff; }
.outline .ol-item.c-compat { color: #9a5eff; }
.hover-hint { position: fixed; background: #444; color: #ccc; padding: 2px 8px; border-radius: 3px; font-size: 11px; pointer-events: none; z-index: 50; display: none; }
.tooltip { position: fixed; background: #2d2d2d; border: 1px solid #555; border-radius: 4px; padding: 10px 14px; max-width: 500px; z-index: 100; display: none; font-size: 12px; line-height: 1.5; box-shadow: 0 4px 12px rgba(0,0,0,0.5); user-select: text; }
.tooltip .tt-api { color: #4fc1ff; font-weight: bold; font-size: 14px; }
.tooltip .tt-alt { color: #6a9955; margin-top: 4px; }
.tooltip .tt-note { color: #b0b0b0; font-style: italic; margin-top: 4px; }
.tooltip pre { background: #1e1e1e; border: 1px solid #3c3c3c; border-radius: 3px; padding: 8px; margin-top: 6px; font-size: 11px; line-height: 1.4; max-height: 200px; overflow-y: auto; position: relative; }
.tooltip .tt-copy { position: absolute; top: 6px; right: 6px; background: none; border: none; cursor: pointer; opacity: 0.5; padding: 2px; }
.tooltip .tt-copy:hover { opacity: 1; }
.tooltip .tt-copy svg { width: 14px; height: 14px; fill: #ccc; }
"@

        # Extra HTML
        $extraHtml = @"
<div class="outline" id="outline"></div>
<div class="tooltip" id="tooltip"></div>
<div class="hover-hint" id="hoverHint">Click for details</div>
"@

        # Build outline data as JS array
        $olJsSb = [System.Text.StringBuilder]::new()
        [void]$olJsSb.Append('[')
        $olFirst = $true
        foreach ($item in $outlineItems) {
            $comma = if ($olFirst) { '' } else { ',' }
            $olFirst = $false
            $colorCls = switch ($item.Color) {
                'hl-sanitized' { 'c-sanitized' }
                'hl-edr' { 'c-edr' }
                'hl-compat' { 'c-compat' }
                default { '' }
            }
            $modLabel = & $he $item.ModName
            $olLabel = & $he $item.Label
            [void]$olJsSb.Append("${comma}{mod:'$modLabel',ln:$($item.LineNum),label:'$olLabel',cls:'$colorCls'}")
        }
        [void]$olJsSb.Append(']')

        # JS
        $analyzeJs = @"
const outline = document.getElementById('outline');
const tooltip = document.getElementById('tooltip');
const hoverHint = document.getElementById('hoverHint');
const apiInfo = $($tooltipEntries.ToString());
const outlineData = $($olJsSb.ToString());

var _baseShowTab = showTab;
showTab = function(idx) {
  _baseShowTab(idx);
  updateOutline();
};

function scrollToRow(r) {
  const rRect = r.getBoundingClientRect();
  const cRect = content.getBoundingClientRect();
  const offset = rRect.top - cRect.top + content.scrollTop;
  content.scrollTo({ top: offset - content.clientHeight / 3, behavior: 'smooth' });
}

function updateOutline() {
  outline.innerHTML = '';
  const hdr = document.createElement('div');
  hdr.className = 'ol-header'; hdr.textContent = 'Detected Lines';
  outline.appendChild(hdr);
  const mod = document.querySelector('.module.active');
  if (!mod) return;
  const modIdx = parseInt(mod.id.replace('mod', ''));
  const tabEl = document.getElementById('tab' + modIdx);
  const modName = tabEl ? tabEl.textContent.replace(/ \(\d+\)$/, '') : '';
  const rows = mod.querySelectorAll('tr');
  rows.forEach(r => {
    const cls = r.className;
    if (!cls || (!cls.includes('hl-sanitized') && !cls.includes('hl-edr') && !cls.includes('hl-compat'))) return;
    const ln = r.querySelector('.ln');
    if (!ln) return;
    const lineNum = ln.textContent;
    const api = r.dataset.api || '';
    const label = 'L' + lineNum + ' ' + api;
    const colorCls = cls.includes('hl-sanitized') ? 'c-sanitized' : cls.includes('hl-edr') ? 'c-edr' : 'c-compat';
    const item = document.createElement('div');
    item.className = 'ol-item ' + colorCls;
    item.textContent = label.substring(0, 50);
    item.addEventListener('click', () => scrollToRow(r));
    outline.appendChild(item);
  });
}

let pinnedTooltip = null;
function showTooltipAt(tr) {
  const api = tr.dataset.api;
  if (!api) return;
  const info = apiInfo[api];
  if (!info) return;
  let html = '<div class="tt-api">' + api + '</div>';
  html += '<div class="tt-alt">Alternative: ' + info.alt + '</div>';
  if (info.note) html += '<div class="tt-note">' + info.note + '</div>';
  if (info.ex) html += '<pre><button class="tt-copy" onclick="copyPre(this)" title="Copy"><svg viewBox="0 0 24 24"><path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg></button>' + info.ex.replace(/\\n/g, '\n') + '</pre>';
  tooltip.innerHTML = html;
  tooltip.style.display = 'block';
  const rect = tr.getBoundingClientRect();
  let top = rect.bottom + 4;
  let left = rect.left + 60;
  if (top + tooltip.offsetHeight > window.innerHeight) top = rect.top - tooltip.offsetHeight - 4;
  if (left + tooltip.offsetWidth > window.innerWidth - 270) left = window.innerWidth - 270 - tooltip.offsetWidth - 10;
  tooltip.style.top = top + 'px';
  tooltip.style.left = left + 'px';
  pinnedTooltip = tr;
}
function copyPre(btn) {
  const pre = btn.closest('pre');
  const text = pre.textContent.trim();
  navigator.clipboard.writeText(text).then(() => {
    btn.innerHTML = '<svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z" fill="#6a9955"/></svg>';
    setTimeout(() => { btn.innerHTML = '<svg viewBox="0 0 24 24"><path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg>'; }, 1500);
  });
}

content.addEventListener('mousemove', (e) => {
  const tr = e.target.closest('tr.hl-edr, tr.hl-compat');
  if (tr && tr.dataset.api && !pinnedTooltip) {
    hoverHint.style.display = 'block';
    hoverHint.style.left = (e.clientX + 12) + 'px';
    hoverHint.style.top = (e.clientY - 8) + 'px';
  } else {
    hoverHint.style.display = 'none';
  }
});
content.addEventListener('mouseleave', () => { hoverHint.style.display = 'none'; });
content.addEventListener('click', (e) => {
  hoverHint.style.display = 'none';
  const tr = e.target.closest('tr.hl-edr, tr.hl-compat');
  if (!tr) { tooltip.style.display = 'none'; pinnedTooltip = null; return; }
  if (pinnedTooltip === tr) {
    tooltip.style.display = 'none'; pinnedTooltip = null;
  } else {
    showTooltipAt(tr);
  }
});
"@

        $htmlSubtitle = "$fileName -- $($csvRow.EdrIssues) EDR, $($csvRow.CompatIssues) compat, $totalSanitized sanitized"
        $htmlPath = Join-Path $outDir "${outPrefix}_analyze.html"

        New-HtmlBase -Title "VBA Analysis: $fileName" -Subtitle $htmlSubtitle `
            -ExtraCss $analyzeCss -SidebarHtml $sidebarSb.ToString() -ContentHtml $contentSb.ToString() `
            -ExtraHtml $extraHtml -ExtraJs $analyzeJs `
            -HighlightSelector 'tr.hl-sanitized, tr.hl-edr, tr.hl-compat' `
            -FirstTabIndex $firstHlIdx -OutputPath $htmlPath

    } catch {
        $csvRow.Error = $_.Exception.Message
        Write-VbaError 'Analyze' $fileName $_.Exception.Message
    }

    [void]$csvRows.Add($csvRow)

    $fileSw = $sw.Elapsed.TotalSeconds
    Write-VbaResult 'Analyze' $fileName "$($csvRow.EdrIssues) EDR, $($csvRow.CompatIssues) compat, $totalSanitized sanitized" $outDir $fileSw
    Write-VbaLog 'Analyze' $filePath "$($csvRow.TotalModules) modules, $($csvRow.EdrIssues) EDR, $($csvRow.CompatIssues) compat, $totalSanitized sanitized | -> $outDir"
}

# === Write CSV ===
$csvPath = Join-Path $outDir 'analyze.csv'
$csvSb = [System.Text.StringBuilder]::new()
[void]$csvSb.AppendLine('Timestamp,RelativePath,FileName,Bas,Cls,Frm,TotalModules,CodeLines,EdrIssues,CompatIssues,SanitizedLines,References,Error')

foreach ($row in $csvRows) {
    $fields = @(
        '"' + ($row.Timestamp -replace '"','""') + '"'
        '"' + ($row.RelativePath -replace '"','""') + '"'
        '"' + ($row.FileName -replace '"','""') + '"'
        $row.Bas
        $row.Cls
        $row.Frm
        $row.TotalModules
        $row.CodeLines
        $row.EdrIssues
        $row.CompatIssues
        $row.SanitizedLines
        '"' + ($row.References -replace '"','""') + '"'
        '"' + ($row.Error -replace '"','""') + '"'
    )
    [void]$csvSb.AppendLine($fields -join ',')
}

$utf8Bom = New-Object System.Text.UTF8Encoding $true
[IO.File]::WriteAllText($csvPath, $csvSb.ToString(), $utf8Bom)

$sw.Stop()
if ($files.Count -gt 1) {
    Write-Host "`n  Total: $($files.Count) files analyzed" -ForegroundColor Green
}
Write-Host "  CSV: $csvPath" -ForegroundColor Gray
Write-Host "  Output: $outDir" -ForegroundColor Gray
Write-Host "  Done ($([Math]::Round($sw.Elapsed.TotalSeconds, 1))s)" -ForegroundColor Gray
