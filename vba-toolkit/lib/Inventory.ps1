param([Parameter(Mandatory)][string[]]$Paths)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

$sw = [System.Diagnostics.Stopwatch]::StartNew()

# Collect all xlsm/xlam/xls files
$files = [System.Collections.ArrayList]::new()
$baseDir = $null

foreach ($p in $Paths) {
    $resolved = (Resolve-Path $p -ErrorAction SilentlyContinue).Path
    if (-not $resolved) { Write-VbaError 'Inventory' $p 'Path not found'; continue }

    if (Test-Path $resolved -PathType Container) {
        # Folder: recurse
        if (-not $baseDir) { $baseDir = $resolved }
        Get-ChildItem $resolved -Recurse -File -Include '*.xlsm','*.xlam','*.xls' | ForEach-Object {
            [void]$files.Add($_.FullName)
        }
    } else {
        # Single file
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

Write-VbaHeader 'Inventory' "$($files.Count) file(s)"
Write-VbaLog 'Inventory' $baseDir "Started: $($files.Count) files"

# Output
$outDir = New-VbaOutputDir ($files[0]) 'inventory'

# CSV rows
$rows = [System.Collections.ArrayList]::new()

$processed = 0
foreach ($filePath in $files) {
    $processed++
    $fileName = [IO.Path]::GetFileName($filePath)
    $relPath = $filePath
    if ($filePath.StartsWith($baseDir)) {
        $relPath = $filePath.Substring($baseDir.Length).TrimStart('\', '/')
        $relDir = [IO.Path]::GetDirectoryName($relPath)
        if (-not $relDir) { $relDir = '.' }
    } else {
        $relDir = [IO.Path]::GetDirectoryName($filePath)
    }

    Write-VbaStatus 'Inventory' '' "[$processed/$($files.Count)] $fileName"

    $row = [ordered]@{
        RelativePath = $relDir
        FileName = $fileName
        BasModules = 0
        ClsModules = 0
        FrmModules = 0
        TotalModules = 0
        CodeLines = 0
        ApiDeclareCount = 0
        ComObjectCount = 0
        References = ''
        Error = ''
    }

    try {
        $project = Get-AllModuleCode $filePath -StripAttributes
        if (-not $project) {
            $row.Error = 'No VBA project'
            [void]$rows.Add($row)
            continue
        }

        # Module counts
        foreach ($modName in $project.Modules.Keys) {
            $mod = $project.Modules[$modName]
            $row.TotalModules++
            switch ($mod.Ext) {
                'bas' { $row.BasModules++ }
                'cls' { $row.ClsModules++ }
                'frm' { $row.FrmModules++ }
            }
            $row.CodeLines += $mod.Lines.Count
        }

        # Analysis (shared engine)
        $analysis = Get-VbaAnalysis -Project $project
        $row.ApiDeclareCount = $analysis.ApiDecls.Count
        $row.ComObjectCount = $analysis.ComBindings.Count
        $row.References = $analysis.ExternalRefs -join '; '
    } catch {
        $row.Error = $_.Exception.Message
        Write-VbaError 'Inventory' $fileName $_.Exception.Message
    }

    [void]$rows.Add($row)
}

# Write CSV (BOM付き UTF-8)
$csvPath = Join-Path $outDir 'inventory.csv'
$csvSb = [System.Text.StringBuilder]::new()

# Header
[void]$csvSb.AppendLine('RelativePath,FileName,Bas,Cls,Frm,TotalModules,CodeLines,ApiDeclare,ComObjects,References,Error')

# Rows
foreach ($row in $rows) {
    $fields = @(
        '"' + ($row.RelativePath -replace '"','""') + '"'
        '"' + ($row.FileName -replace '"','""') + '"'
        $row.BasModules
        $row.ClsModules
        $row.FrmModules
        $row.TotalModules
        $row.CodeLines
        $row.ApiDeclareCount
        $row.ComObjectCount
        '"' + ($row.References -replace '"','""') + '"'
        '"' + ($row.Error -replace '"','""') + '"'
    )
    [void]$csvSb.AppendLine($fields -join ',')
}

$utf8Bom = New-Object System.Text.UTF8Encoding $true
[IO.File]::WriteAllText($csvPath, $csvSb.ToString(), $utf8Bom)

# Summary
$totalFiles = $rows.Count
$totalModules = 0; $totalLines = 0; $apiFiles = 0; $comFiles = 0; $errorFiles = 0
foreach ($r in $rows) {
    $totalModules += $r.TotalModules
    $totalLines += $r.CodeLines
    if ($r.ApiDeclareCount -gt 0) { $apiFiles++ }
    if ($r.ComObjectCount -gt 0) { $comFiles++ }
    if ($r.Error -ne '') { $errorFiles++ }
}

$sw.Stop()

Write-Host ""
Write-Host "  Files:       $totalFiles" -ForegroundColor Gray
Write-Host "  Modules:     $totalModules" -ForegroundColor Gray
Write-Host "  Code lines:  $totalLines" -ForegroundColor Gray
Write-Host "  With API:    $apiFiles file(s)" -ForegroundColor $(if ($apiFiles -gt 0) { 'Yellow' } else { 'Gray' })
Write-Host "  With COM:    $comFiles file(s)" -ForegroundColor Gray
if ($errorFiles -gt 0) { Write-Host "  Errors:      $errorFiles file(s)" -ForegroundColor Red }

Write-VbaResult 'Inventory' "$($files.Count) files" "CSV: $csvPath" $outDir $sw.Elapsed.TotalSeconds
Write-VbaLog 'Inventory' $baseDir "$totalFiles files, $totalModules modules, $totalLines lines, $apiFiles with API | -> $outDir"
