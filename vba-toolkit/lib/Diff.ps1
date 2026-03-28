param(
    [Parameter(Mandatory)][string]$FileA,
    [Parameter(Mandatory)][string]$FileB
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

$FileA = Resolve-VbaFilePath $FileA
$FileB = Resolve-VbaFilePath $FileB
$nameA = [IO.Path]::GetFileName($FileA)
$nameB = [IO.Path]::GetFileName($FileB)
$sw = [System.Diagnostics.Stopwatch]::StartNew()

Write-VbaHeader 'Diff' "$nameA vs $nameB"
Write-VbaLog 'Diff' $FileA "Comparing with $nameB"

function Get-AllModules([string]$path) {
    $project = Get-AllModuleCode $path -StripAttributes
    if (-not $project) { return @{} }
    $result = [ordered]@{}
    foreach ($modName in $project.Modules.Keys) {
        $mod = $project.Modules[$modName]
        $result[$modName] = @{ Code = ($mod.Lines -join "`n"); Ext = $mod.Ext }
    }
    return $result
}

# ============================================================================
# LCS diff algorithm
# ============================================================================

function Get-LcsDiff([string[]]$a, [string[]]$b) {
    $m = $a.Count; $n = $b.Count
    # For large files, use a simpler greedy approach
    $result = [System.Collections.ArrayList]::new()
    $ia = 0; $ib = 0

    while ($ia -lt $m -or $ib -lt $n) {
        if ($ia -lt $m -and $ib -lt $n -and $a[$ia] -eq $b[$ib]) {
            [void]$result.Add(@{ Type = 'equal'; LineA = $ia; LineB = $ib; TextA = $a[$ia]; TextB = $b[$ib] })
            $ia++; $ib++; continue
        }

        $bestAi = -1; $bestBi = -1; $bestDist = ($m + $n) * 2
        $searchA = [Math]::Min($ia + 100, $m)
        $searchB = [Math]::Min($ib + 100, $n)
        for ($ai = $ia; $ai -lt $searchA; $ai++) {
            for ($bi = $ib; $bi -lt $searchB; $bi++) {
                if ($a[$ai] -eq $b[$bi]) {
                    $dist = ($ai - $ia) + ($bi - $ib)
                    if ($dist -lt $bestDist) { $bestDist = $dist; $bestAi = $ai; $bestBi = $bi }
                    break
                }
            }
        }
        if ($bestAi -eq -1) { $bestAi = $m; $bestBi = $n }

        # Pair up removed/added lines where possible
        $remCount = $bestAi - $ia; $addCount = $bestBi - $ib
        $pairCount = [Math]::Min($remCount, $addCount)
        for ($p = 0; $p -lt $pairCount; $p++) {
            [void]$result.Add(@{ Type = 'changed'; LineA = $ia; LineB = $ib; TextA = $a[$ia]; TextB = $b[$ib] })
            $ia++; $ib++
        }
        while ($ia -lt $bestAi) {
            [void]$result.Add(@{ Type = 'removed'; LineA = $ia; LineB = -1; TextA = $a[$ia]; TextB = '' })
            $ia++
        }
        while ($ib -lt $bestBi) {
            [void]$result.Add(@{ Type = 'added'; LineA = -1; LineB = $ib; TextA = ''; TextB = $b[$ib] })
            $ib++
        }
        $ia = $bestAi; $ib = $bestBi
    }
    return ,$result
}

# ============================================================================
# Build module diffs
# ============================================================================

$modsA = Get-AllModules $FileA
$modsB = Get-AllModules $FileB

$allNames = [System.Collections.ArrayList]::new()
foreach ($k in $modsA.Keys) { if ($allNames -notcontains $k) { [void]$allNames.Add($k) } }
foreach ($k in $modsB.Keys) { if ($allNames -notcontains $k) { [void]$allNames.Add($k) } }
$allNames = $allNames | Sort-Object

$moduleDiffs = [System.Collections.ArrayList]::new()
$added = 0; $removed = 0; $modified = 0; $unchanged = 0

foreach ($name in $allNames) {
    $inA = $modsA.Contains($name); $inB = $modsB.Contains($name)
    if ($inA -and -not $inB) {
        $removed++
        $ext = $modsA[$name].Ext
        Write-Host "  - $name.$ext (removed)" -ForegroundColor Red
        $linesA = $modsA[$name].Code -split "`n"
        $diff = [System.Collections.ArrayList]::new()
        for ($i = 0; $i -lt $linesA.Count; $i++) {
            [void]$diff.Add(@{ Type = 'removed'; LineA = $i; LineB = -1; TextA = $linesA[$i]; TextB = '' })
        }
        [void]$moduleDiffs.Add(@{ Name = "$name.$ext"; Status = 'removed'; Diff = $diff })
    }
    elseif (-not $inA -and $inB) {
        $added++
        $ext = $modsB[$name].Ext
        Write-Host "  + $name.$ext (added)" -ForegroundColor Green
        $linesB = $modsB[$name].Code -split "`n"
        $diff = [System.Collections.ArrayList]::new()
        for ($i = 0; $i -lt $linesB.Count; $i++) {
            [void]$diff.Add(@{ Type = 'added'; LineA = -1; LineB = $i; TextA = ''; TextB = $linesB[$i] })
        }
        [void]$moduleDiffs.Add(@{ Name = "$name.$ext"; Status = 'added'; Diff = $diff })
    }
    else {
        $codeA = $modsA[$name].Code; $codeB = $modsB[$name].Code
        $ext = $modsA[$name].Ext
        if ($codeA -eq $codeB) {
            $unchanged++
            $linesA = $codeA -split "`n"
            $diff = [System.Collections.ArrayList]::new()
            for ($i = 0; $i -lt $linesA.Count; $i++) {
                [void]$diff.Add(@{ Type = 'equal'; LineA = $i; LineB = $i; TextA = $linesA[$i]; TextB = $linesA[$i] })
            }
            [void]$moduleDiffs.Add(@{ Name = "$name.$ext"; Status = 'unchanged'; Diff = $diff })
        } else {
            $modified++
            Write-Host "  ~ $name.$ext (modified)" -ForegroundColor Yellow
            $diff = Get-LcsDiff ($codeA -split "`n") ($codeB -split "`n")
            [void]$moduleDiffs.Add(@{ Name = "$name.$ext"; Status = 'modified'; Diff = $diff })
        }
    }
}

Write-Host ""
$parts = @()
if ($added) { $parts += "$added added" }
if ($removed) { $parts += "$removed removed" }
if ($modified) { $parts += "$modified modified" }
if ($unchanged) { $parts += "$unchanged unchanged" }
Write-Host "Summary: $($parts -join ', ')"

# ============================================================================
# Generate HTML
# ============================================================================

function HtmlEncode([string]$s) { return [System.Net.WebUtility]::HtmlEncode($s) }

$html = [System.Text.StringBuilder]::new()
[void]$html.Append(@"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>VBA Diff: $([System.Net.WebUtility]::HtmlEncode($nameA)) vs $([System.Net.WebUtility]::HtmlEncode($nameB))</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: Consolas, 'Courier New', monospace; font-size: 13px; background: #1e1e1e; color: #d4d4d4; }
.header { background: #252526; padding: 10px 20px; border-bottom: 1px solid #3c3c3c; }
.header h1 { font-size: 15px; font-weight: normal; color: #cccccc; }
.header .summary { margin-top: 4px; font-size: 12px; color: #888; }
.main { display: flex; height: calc(100vh - 52px); }
.sidebar { width: 200px; min-width: 200px; background: #252526; border-right: 1px solid #3c3c3c; overflow-y: auto; padding: 8px 0; }
.sidebar .item { padding: 5px 16px; cursor: pointer; color: #888; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sidebar .item:hover { color: #d4d4d4; background: #2a2d2e; }
.sidebar .item.active { color: #ffffff; background: #37373d; border-left: 2px solid #0078d4; }
.sidebar .item.modified { color: #e8ab53; }
.sidebar .item.added { color: #6a9955; }
.sidebar .item.removed { color: #f44747; }
.sidebar .item.unchanged { color: #606060; }
.content { flex: 1; overflow: auto; position: relative; }
.module { display: none; }
.module.active { display: block; }
.minimap { position: fixed; right: 0; top: 52px; width: 14px; bottom: 0; background: #1e1e1e; border-left: 1px solid #3c3c3c; z-index: 20; cursor: pointer; }
.minimap .mark { position: absolute; right: 2px; width: 10px; height: 3px; border-radius: 1px; }
.minimap .mark.m-changed { background: #e8ab53; }
.minimap .mark.m-removed { background: #f44747; }
.minimap .mark.m-added { background: #6a9955; }
.minimap .viewport { position: absolute; right: 0; width: 14px; background: rgba(255,255,255,0.25); border-radius: 2px; pointer-events: none; }
.diff-table { width: 100%; border-collapse: collapse; table-layout: fixed; }
.diff-table td { padding: 0 8px; line-height: 20px; vertical-align: top; white-space: pre; overflow: hidden; text-overflow: ellipsis; }
.diff-table .ln { width: 45px; min-width: 45px; text-align: right; color: #606060; padding-right: 12px; user-select: none; border-right: 1px solid #3c3c3c; }
.diff-table .code { width: calc(50% - 45px); }
.diff-table .sep { width: 1px; background: #3c3c3c; padding: 0; }
tr.equal td.code { color: #d4d4d4; }
tr.changed td.code-a { background: #4b1818; color: #f8d7d7; }
tr.changed td.code-b { background: #1b3a1b; color: #d7f8d7; }
tr.removed td.code-a { background: #4b1818; color: #f8d7d7; }
tr.removed td.code-b { background: #2d2d2d; color: #606060; }
tr.added td.code-a { background: #2d2d2d; color: #606060; }
tr.added td.code-b { background: #1b3a1b; color: #d7f8d7; }
tr.changed td.ln, tr.removed td.ln-a, tr.added td.ln-b { color: #cccccc; }
.unchanged-marker { text-align: center; color: #606060; padding: 4px; background: #252526; font-size: 11px; }
</style>
</head>
<body>
<div class="header">
  <h1>VBA Diff</h1>
  <div class="summary">
    A: $(HtmlEncode $nameA) &nbsp;&mdash;&nbsp; B: $(HtmlEncode $nameB)
    &nbsp;&nbsp;|&nbsp;&nbsp;
    $($parts -join ' &nbsp; ')
  </div>
</div>
<div class="main">
<div class="sidebar" id="sidebar">
"@)

# Sidebar items
$tabIdx = 0
$firstModifiedIdx = -1
foreach ($md in $moduleDiffs) {
    $cls = $md.Status
    if ($firstModifiedIdx -eq -1 -and $md.Status -ne 'unchanged') { $firstModifiedIdx = $tabIdx }
    [void]$html.Append("<div class=`"item $cls`" onclick=`"showTab($tabIdx)`" id=`"tab$tabIdx`">$(HtmlEncode $md.Name)</div>")
    $tabIdx++
}
if ($firstModifiedIdx -eq -1) { $firstModifiedIdx = 0 }

[void]$html.Append("</div><div class=`"content`">")

# Module panels
$tabIdx = 0
foreach ($md in $moduleDiffs) {
    [void]$html.Append("<div class=`"module`" id=`"mod$tabIdx`"><table class=`"diff-table`">")

    $diff = $md.Diff
    foreach ($row in $diff) {
        $type = $row.Type
        $lnA = if ($row.LineA -ge 0) { $row.LineA + 1 } else { '' }
        $lnB = if ($row.LineB -ge 0) { $row.LineB + 1 } else { '' }
        $tA = HtmlEncode $row.TextA
        $tB = HtmlEncode $row.TextB
        [void]$html.Append("<tr class=`"$type`"><td class=`"ln ln-a`">$lnA</td><td class=`"code code-a`">$tA</td><td class=`"sep`"></td><td class=`"ln ln-b`">$lnB</td><td class=`"code code-b`">$tB</td></tr>")
    }

    [void]$html.Append("</table></div>")
    $tabIdx++
}

[void]$html.Append(@"
</div>
<div class="minimap" id="minimap"><div class="viewport" id="viewport"></div></div>
</div>
<script>
const content = document.querySelector('.content');
const minimap = document.getElementById('minimap');
const viewport = document.getElementById('viewport');

function showTab(idx) {
  document.querySelectorAll('.module').forEach(m => m.classList.remove('active'));
  document.querySelectorAll('.item').forEach(t => t.classList.remove('active'));
  document.getElementById('mod' + idx).classList.add('active');
  document.getElementById('tab' + idx).classList.add('active');
  updateMinimap();
}

function updateMinimap() {
  minimap.querySelectorAll('.mark').forEach(m => m.remove());
  const mod = document.querySelector('.module.active');
  if (!mod) return;
  const rows = mod.querySelectorAll('tr.changed, tr.removed, tr.added');
  const allRows = mod.querySelectorAll('tr');
  const totalRows = allRows.length;
  if (totalRows === 0) return;
  const mapH = minimap.clientHeight;
  rows.forEach(r => {
    const idx = Array.from(allRows).indexOf(r);
    const mark = document.createElement('div');
    mark.className = 'mark m-' + r.className;
    mark.style.top = (idx / totalRows * mapH) + 'px';
    mark.addEventListener('click', () => r.scrollIntoView({block:'center'}));
    minimap.appendChild(mark);
  });
  updateViewport();
}

function updateViewport() {
  const sh = content.scrollHeight;
  const ch = content.clientHeight;
  const st = content.scrollTop;
  const mapH = minimap.clientHeight;
  if (sh <= ch) { viewport.style.display = 'none'; return; }
  viewport.style.display = '';
  viewport.style.top = (st / sh * mapH) + 'px';
  viewport.style.height = (ch / sh * mapH) + 'px';
}

content.addEventListener('scroll', updateViewport);
minimap.addEventListener('click', (e) => {
  if (e.target.classList.contains('mark')) return;
  const ratio = e.offsetY / minimap.clientHeight;
  content.scrollTop = ratio * content.scrollHeight - content.clientHeight / 2;
});

showTab($firstModifiedIdx);
</script>
</body></html>
"@)

$outDir = New-VbaOutputDir $FileA 'diff'
$htmlPath = Join-Path $outDir 'diff.html'
[IO.File]::WriteAllText($htmlPath, $html.ToString(), [System.Text.Encoding]::UTF8)

# Text report
$textReport = [System.Text.StringBuilder]::new()
[void]$textReport.AppendLine("# VBA Diff Report")
[void]$textReport.AppendLine("# A: $nameA")
[void]$textReport.AppendLine("# B: $nameB")
[void]$textReport.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$textReport.AppendLine("")
[void]$textReport.AppendLine("Added: $added, Removed: $removed, Modified: $modified, Unchanged: $unchanged")
[IO.File]::WriteAllText((Join-Path $outDir 'diff.txt'), $textReport.ToString(), [System.Text.Encoding]::UTF8)

Start-Process $htmlPath

$sw.Stop()
Write-VbaResult 'Diff' "$nameA vs $nameB" "$($parts -join ', ')" $outDir $sw.Elapsed.TotalSeconds
Write-VbaLog 'Diff' $FileA "vs $nameB | $($parts -join ', ') | -> $outDir"
