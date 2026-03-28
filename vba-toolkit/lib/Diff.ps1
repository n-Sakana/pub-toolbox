param(
    [Parameter(Mandatory)][string]$FileA,
    [Parameter(Mandatory)][string]$FileB
)

$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

foreach ($p in @($FileA, $FileB)) {
    if (-not (Test-Path $p)) { Write-Host "Error: file not found: $p" -ForegroundColor Red; exit 1 }
}
$FileA = (Resolve-Path $FileA).Path
$FileB = (Resolve-Path $FileB).Path
$nameA = [IO.Path]::GetFileName($FileA)
$nameB = [IO.Path]::GetFileName($FileB)

Write-Host "Comparing VBA code:"
Write-Host "  A: $nameA"
Write-Host "  B: $nameB"
Write-Host ""

# Extract all module code from a file
function Get-AllModules([string]$path) {
    $proj = Get-VbaProjectBytes $path
    if (-not $proj.Bytes) { return @{} }
    $ole2 = Read-Ole2 $proj.Bytes
    $modules = Get-VbaModuleList $ole2
    $result = [ordered]@{}
    foreach ($mod in $modules) {
        $mc = Get-VbaModuleCode $ole2 $mod.Name
        if ($mc) {
            # Strip Attribute lines for cleaner diff
            $clean = (($mc.Code -split "`r`n|`n") | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_' }) -join "`n"
            $result[$mod.Name] = @{ Code = $clean.TrimStart("`n"); Ext = $mod.Ext }
        }
    }
    return $result
}

$modsA = Get-AllModules $FileA
$modsB = Get-AllModules $FileB

$allNames = [System.Collections.ArrayList]::new()
foreach ($k in $modsA.Keys) { if ($allNames -notcontains $k) { [void]$allNames.Add($k) } }
foreach ($k in $modsB.Keys) { if ($allNames -notcontains $k) { [void]$allNames.Add($k) } }
$allNames = $allNames | Sort-Object

$report = [System.Text.StringBuilder]::new()
[void]$report.AppendLine("# VBA Diff Report")
[void]$report.AppendLine("# A: $nameA")
[void]$report.AppendLine("# B: $nameB")
[void]$report.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$report.AppendLine("")

$added = 0; $removed = 0; $modified = 0; $unchanged = 0

foreach ($name in $allNames) {
    $inA = $modsA.Contains($name)
    $inB = $modsB.Contains($name)

    if ($inA -and -not $inB) {
        $removed++
        $ext = $modsA[$name].Ext
        Write-Host "  - $name.$ext (removed)" -ForegroundColor Red
        [void]$report.AppendLine("## $name.$ext  [REMOVED]")
        [void]$report.AppendLine("")
        $linesA = ($modsA[$name].Code -split "`n").Count
        [void]$report.AppendLine("  $linesA lines removed")
        [void]$report.AppendLine("")
    }
    elseif (-not $inA -and $inB) {
        $added++
        $ext = $modsB[$name].Ext
        Write-Host "  + $name.$ext (added)" -ForegroundColor Green
        [void]$report.AppendLine("## $name.$ext  [ADDED]")
        [void]$report.AppendLine("")
        $linesB = ($modsB[$name].Code -split "`n").Count
        [void]$report.AppendLine("  $linesB lines added")
        [void]$report.AppendLine("")
    }
    else {
        $codeA = $modsA[$name].Code
        $codeB = $modsB[$name].Code
        $ext = $modsA[$name].Ext

        if ($codeA -eq $codeB) {
            $unchanged++
            continue
        }

        $modified++
        Write-Host "  ~ $name.$ext (modified)" -ForegroundColor Yellow

        # Line-by-line diff
        $linesA = $codeA -split "`n"
        $linesB = $codeB -split "`n"

        [void]$report.AppendLine("## $name.$ext  [MODIFIED]")
        [void]$report.AppendLine("")

        # Simple LCS-based diff
        $maxLen = [Math]::Max($linesA.Count, $linesB.Count)
        $hunks = [System.Collections.ArrayList]::new()
        $ia = 0; $ib = 0

        while ($ia -lt $linesA.Count -or $ib -lt $linesB.Count) {
            # Find matching lines
            if ($ia -lt $linesA.Count -and $ib -lt $linesB.Count -and $linesA[$ia] -eq $linesB[$ib]) {
                $ia++; $ib++; continue
            }

            # Diverged - find next sync point
            $bestAi = -1; $bestBi = -1; $bestDist = $maxLen * 2
            $searchA = [Math]::Min($ia + 50, $linesA.Count)
            $searchB = [Math]::Min($ib + 50, $linesB.Count)

            for ($ai = $ia; $ai -lt $searchA; $ai++) {
                for ($bi = $ib; $bi -lt $searchB; $bi++) {
                    if ($linesA[$ai] -eq $linesB[$bi]) {
                        $dist = ($ai - $ia) + ($bi - $ib)
                        if ($dist -lt $bestDist) {
                            $bestDist = $dist; $bestAi = $ai; $bestBi = $bi
                        }
                        break  # found match for this ai
                    }
                }
            }

            if ($bestAi -eq -1) {
                # No sync point found in window - dump rest
                $bestAi = $linesA.Count; $bestBi = $linesB.Count
            }

            # Output hunk
            $hunkLines = [System.Collections.ArrayList]::new()
            $lineNum = $ia + 1
            for ($x = $ia; $x -lt $bestAi; $x++) {
                [void]$hunkLines.Add("  - L$($x+1): $($linesA[$x])")
            }
            for ($x = $ib; $x -lt $bestBi; $x++) {
                [void]$hunkLines.Add("  + L$($x+1): $($linesB[$x])")
            }

            if ($hunkLines.Count -gt 0) {
                [void]$report.AppendLine("  @@ A:L$($ia+1) B:L$($ib+1) @@")
                foreach ($hl in $hunkLines) { [void]$report.AppendLine($hl) }
                [void]$report.AppendLine("")
            }

            $ia = $bestAi; $ib = $bestBi
        }
    }
}

# Summary
Write-Host ""
Write-Host "Summary: " -NoNewline
$parts = @()
if ($added) { $parts += "$added added" }
if ($removed) { $parts += "$removed removed" }
if ($modified) { $parts += "$modified modified" }
if ($unchanged) { $parts += "$unchanged unchanged" }
Write-Host ($parts -join ', ')

[void]$report.AppendLine("## Summary")
[void]$report.AppendLine("")
[void]$report.AppendLine("  Added: $added, Removed: $removed, Modified: $modified, Unchanged: $unchanged")

# Save report
$dirA = [IO.Path]::GetDirectoryName($FileA)
$baseA = [IO.Path]::GetFileNameWithoutExtension($FileA)
$baseB = [IO.Path]::GetFileNameWithoutExtension($FileB)
$reportPath = Join-Path $dirA "${baseA}_vs_${baseB}_diff.txt"
[IO.File]::WriteAllText($reportPath, $report.ToString(), [System.Text.Encoding]::UTF8)
Write-Host "Report: $reportPath" -ForegroundColor Green
