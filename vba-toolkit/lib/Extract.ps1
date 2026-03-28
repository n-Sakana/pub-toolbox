param([Parameter(Mandatory)][string[]]$Paths)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

$sw = [System.Diagnostics.Stopwatch]::StartNew()

# Collect all xlsm/xlam/xls files
$files = [System.Collections.ArrayList]::new()
foreach ($p in $Paths) {
    $resolved = (Resolve-Path $p -ErrorAction SilentlyContinue).Path
    if (-not $resolved) { Write-VbaError 'Extract' $p 'Path not found'; continue }

    if (Test-Path $resolved -PathType Container) {
        Get-ChildItem $resolved -Recurse -File -Include '*.xlsm','*.xlam','*.xls' | ForEach-Object {
            [void]$files.Add($_.FullName)
        }
    } else {
        $ext = [IO.Path]::GetExtension($resolved).ToLower()
        if ($ext -in '.xls','.xlsm','.xlam') {
            [void]$files.Add($resolved)
        }
    }
}

if ($files.Count -eq 0) {
    Write-Host "No Excel files found." -ForegroundColor Yellow
    exit 0
}

$totalExtracted = 0
foreach ($FilePath in $files) {
    $fileName = [IO.Path]::GetFileName($FilePath)
    $baseName = [IO.Path]::GetFileNameWithoutExtension($FilePath)

    Write-VbaHeader 'Extract' $fileName
    Write-VbaLog 'Extract' $FilePath 'Started'

    $project = Get-AllModuleCode $FilePath -StripAttributes
    if (-not $project) { Write-VbaError 'Extract' $fileName 'No vbaProject.bin found'; continue }

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

    # Combined source with module index
    $allFiles = Get-ChildItem $modulesDir -File
    $totalLines = 0

    $combined = [System.Text.StringBuilder]::new()
    [void]$combined.AppendLine("=" * 80)
    [void]$combined.AppendLine(" $fileName - VBA Source Code")
    [void]$combined.AppendLine(" Extracted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$combined.AppendLine("=" * 80)
    [void]$combined.AppendLine("")

    # Module index
    [void]$combined.AppendLine("MODULE INDEX")
    [void]$combined.AppendLine("-" * 40)
    [void]$combined.AppendLine("")

    $stdMods = @(); $clsMods = @(); $frmMods = @(); $docMods = @()
    foreach ($f in $allFiles) {
        $ext = $f.Extension.TrimStart('.')
        $lc = (Get-Content $f.FullName -Encoding UTF8).Count
        $totalLines += $lc
        $entry = "    $([IO.Path]::GetFileNameWithoutExtension($f.Name)) ($lc lines)"
        switch ($ext) {
            'bas' { $stdMods += $entry }
            'cls' { $clsMods += $entry }
            'frm' { $frmMods += $entry }
            default { $docMods += $entry }
        }
    }
    if ($stdMods.Count -gt 0) { [void]$combined.AppendLine("  Standard Modules:"); foreach ($e in $stdMods) { [void]$combined.AppendLine($e) }; [void]$combined.AppendLine("") }
    if ($clsMods.Count -gt 0) { [void]$combined.AppendLine("  Class Modules:"); foreach ($e in $clsMods) { [void]$combined.AppendLine($e) }; [void]$combined.AppendLine("") }
    if ($frmMods.Count -gt 0) { [void]$combined.AppendLine("  UserForms:"); foreach ($e in $frmMods) { [void]$combined.AppendLine($e) }; [void]$combined.AppendLine("") }
    if ($docMods.Count -gt 0) { [void]$combined.AppendLine("  Document Modules:"); foreach ($e in $docMods) { [void]$combined.AppendLine($e) }; [void]$combined.AppendLine("") }
    [void]$combined.AppendLine("  Total: $totalLines lines across $($allFiles.Count) modules")
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

    $totalExtracted += $extracted
    Write-VbaResult 'Extract' $fileName "$extracted module(s), $totalLines lines" $outDir $sw.Elapsed.TotalSeconds
    Write-VbaLog 'Extract' $FilePath "$extracted modules, $totalLines lines | -> $outDir"
}

$sw.Stop()
if ($files.Count -gt 1) {
    Write-Host "`n  Total: $($files.Count) files, $totalExtracted modules" -ForegroundColor Green
}
