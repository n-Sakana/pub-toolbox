param(
    [Parameter(Mandatory)][string]$FilePath
)

$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

if (-not (Test-Path $FilePath)) { Write-Host "Error: file not found: $FilePath" -ForegroundColor Red; exit 1 }
$FilePath = (Resolve-Path $FilePath).Path
$ext = [IO.Path]::GetExtension($FilePath).ToLower()
if ($ext -notin '.xls', '.xlsm', '.xlam') { Write-Host "Error: unsupported format: $ext" -ForegroundColor Red; exit 1 }

# ============================================================================
# Win32 API replacement database
# ============================================================================

$replacements = [ordered]@{
    # --- Timer / Sleep ---
    'GetTickCount' = @{
        Lib = 'kernel32'
        Risk = 'LOW'
        Alt = 'Timer (VBA built-in)'
        Example = @'
' Before:
Dim t As Long: t = GetTickCount()
DoSomething
Debug.Print "Elapsed: " & (GetTickCount() - t) & " ms"

' After:
Dim t As Double: t = Timer
DoSomething
Debug.Print "Elapsed: " & Format((Timer - t) * 1000, "0") & " ms"
'@
        Note = 'Timer returns seconds as Double (midnight reset). For short measurements this is sufficient.'
    }
    'GetTickCount64' = @{
        Lib = 'kernel32'
        Risk = 'LOW'
        Alt = 'Timer (VBA built-in)'
        Example = '(Same as GetTickCount)'
        Note = ''
    }
    'Sleep' = @{
        Lib = 'kernel32'
        Risk = 'LOW'
        Alt = 'Application.Wait or DoEvents loop'
        Example = @'
' Before:
Sleep 1000  ' 1 second

' After (Option A - simple):
Application.Wait Now + TimeSerial(0, 0, 1)

' After (Option B - sub-second, non-blocking):
Dim endTime As Double: endTime = Timer + 0.5  ' 500ms
Do While Timer < endTime: DoEvents: Loop
'@
        Note = 'Application.Wait has 1-second resolution. Use DoEvents loop for sub-second or non-blocking waits.'
    }
    'timeGetTime' = @{
        Lib = 'winmm'
        Risk = 'LOW'
        Alt = 'Timer (VBA built-in)'
        Example = '(Same as GetTickCount)'
        Note = ''
    }
    'QueryPerformanceCounter' = @{
        Lib = 'kernel32'
        Risk = 'MEDIUM'
        Alt = 'Timer (lower precision but usually sufficient)'
        Example = @'
' Before:
QueryPerformanceCounter startCount
DoSomething
QueryPerformanceCounter endCount
elapsed = (endCount - startCount) / freq

' After:
Dim t As Double: t = Timer
DoSomething
Debug.Print "Elapsed: " & Format((Timer - t) * 1000, "0") & " ms"
'@
        Note = 'Timer has ~15ms resolution. If microsecond precision is needed, there is no pure VBA alternative.'
    }
    'QueryPerformanceFrequency' = @{
        Lib = 'kernel32'
        Risk = 'LOW'
        Alt = '(Remove together with QueryPerformanceCounter)'
        Example = ''
        Note = ''
    }

    # --- String / Memory ---
    'CopyMemory' = @{
        Lib = 'kernel32 (RtlMoveMemory)'
        Risk = 'HIGH'
        Alt = 'Array operations or byte-by-byte copy'
        Example = @'
' Before:
CopyMemory ByVal dest, ByVal src, length

' After (for byte arrays):
Dim i As Long
For i = 0 To length - 1
    dest(i) = src(i)
Next i

' Or use mid$ for strings:
Mid$(dest, pos, length) = Mid$(src, 1, length)
'@
        Note = 'CopyMemory is often used for type punning. Refactor to avoid raw memory manipulation.'
    }
    'lstrlen' = @{
        Lib = 'kernel32'
        Risk = 'LOW'
        Alt = 'Len / LenB (VBA built-in)'
        Example = @'
' Before:
length = lstrlen(ByVal ptr)

' After:
length = Len(str)     ' character count
length = LenB(str)    ' byte count
'@
        Note = ''
    }

    # --- User / System info ---
    'GetUserName' = @{
        Lib = 'advapi32'
        Risk = 'LOW'
        Alt = 'Environ$("USERNAME") or Application.UserName'
        Example = @'
' Before:
Dim buf As String: buf = Space(256)
Dim sz As Long: sz = 256
GetUserName buf, sz
userName = Left$(buf, sz - 1)

' After:
userName = Environ$("USERNAME")
' Or:
userName = Application.UserName
'@
        Note = 'Environ$ reads the environment variable. Application.UserName reads the Office setting.'
    }
    'GetComputerName' = @{
        Lib = 'kernel32'
        Risk = 'LOW'
        Alt = 'Environ$("COMPUTERNAME")'
        Example = @'
' Before:
Dim buf As String: buf = Space(256)
Dim sz As Long: sz = 256
GetComputerName buf, sz
compName = Left$(buf, sz - 1)

' After:
compName = Environ$("COMPUTERNAME")
'@
        Note = ''
    }
    'GetTempPath' = @{
        Lib = 'kernel32'
        Risk = 'LOW'
        Alt = 'Environ$("TEMP")'
        Example = @'
' Before:
Dim buf As String: buf = Space(260)
GetTempPath 260, buf
tmpPath = Left$(buf, InStr(buf, vbNullChar) - 1)

' After:
tmpPath = Environ$("TEMP")
'@
        Note = ''
    }
    'GetSystemDirectory' = @{
        Lib = 'kernel32'
        Risk = 'LOW'
        Alt = 'Environ$("WINDIR") & "\System32"'
        Example = ''
        Note = ''
    }
    'GetWindowsDirectory' = @{
        Lib = 'kernel32'
        Risk = 'LOW'
        Alt = 'Environ$("WINDIR")'
        Example = ''
        Note = ''
    }

    # --- Window / UI ---
    'FindWindow' = @{
        Lib = 'user32'
        Risk = 'HIGH'
        Alt = 'Application object properties or AppActivate'
        Example = @'
' Before:
hWnd = FindWindow(vbNullString, "Window Title")

' After (activate by title):
AppActivate "Window Title"

' After (get Excel window handle):
hWnd = Application.hWnd  ' Excel 2010+
'@
        Note = 'Most FindWindow usage in VBA is for getting Excel/form window handles. Use Application.hWnd instead.'
    }
    'SetWindowPos' = @{
        Lib = 'user32'
        Risk = 'HIGH'
        Alt = 'UserForm position properties or Application window properties'
        Example = @'
' Before:
SetWindowPos hWnd, HWND_TOPMOST, x, y, w, h, SWP_NOSIZE

' After (for UserForm):
Me.StartUpPosition = 0
Me.Left = x: Me.Top = y
'@
        Note = 'If used to make a form topmost, there is no pure VBA equivalent. Consider if it is really necessary.'
    }
    'GetSystemMetrics' = @{
        Lib = 'user32'
        Risk = 'MEDIUM'
        Alt = 'Application.Width / Application.Height or hard-coded values'
        Example = @'
' Before:
screenW = GetSystemMetrics(SM_CXSCREEN)
screenH = GetSystemMetrics(SM_CYSCREEN)

' After:
screenW = Application.Width   ' in points
screenH = Application.Height
'@
        Note = 'Units differ: API returns pixels, Application properties return points.'
    }
    'ShowWindow' = @{
        Lib = 'user32'
        Risk = 'HIGH'
        Alt = 'Application.Visible or UserForm.Show/Hide'
        Example = ''
        Note = ''
    }
    'SetForegroundWindow' = @{
        Lib = 'user32'
        Risk = 'MEDIUM'
        Alt = 'AppActivate (VBA built-in)'
        Example = @'
' Before:
SetForegroundWindow hWnd

' After:
AppActivate "Window Title"
' Or:
Application.Visible = True
'@
        Note = ''
    }
    'SendMessage' = @{
        Lib = 'user32'
        Risk = 'HIGH'
        Alt = 'Depends on message type. Often no direct alternative.'
        Example = ''
        Note = 'SendMessage is highly versatile. Review each call site individually. Common uses: scrolling listboxes, setting control properties.'
    }
    'PostMessage' = @{
        Lib = 'user32'
        Risk = 'HIGH'
        Alt = '(Same as SendMessage - review individually)'
        Example = ''
        Note = ''
    }

    # --- File ---
    'SHFileOperation' = @{
        Lib = 'shell32'
        Risk = 'MEDIUM'
        Alt = 'FileSystemObject or VBA Kill/FileCopy/Name'
        Example = @'
' Before:
SHFileOperation fileOp  ' copy/move/delete with recycle bin

' After:
' Copy:  FileCopy src, dst
' Move:  Name src As dst
' Delete (permanent): Kill path
' Delete (recycle bin): no pure VBA equivalent
'@
        Note = 'Recycle bin delete has no VBA equivalent. Use FileSystemObject.DeleteFile for permanent delete.'
    }
    'ShellExecute' = @{
        Lib = 'shell32'
        Risk = 'MEDIUM'
        Alt = 'Shell (VBA built-in) or ThisWorkbook.FollowHyperlink'
        Example = @'
' Before:
ShellExecute 0, "open", path, vbNullString, vbNullString, SW_SHOW

' After (open file with default app):
ThisWorkbook.FollowHyperlink path

' After (run program):
Shell path, vbNormalFocus
'@
        Note = 'Shell function is also flagged by some EDR. FollowHyperlink is safer for opening documents/URLs.'
    }

    # --- Clipboard ---
    'OpenClipboard' = @{
        Lib = 'user32'
        Risk = 'MEDIUM'
        Alt = 'MSForms.DataObject'
        Example = @'
' Before:
OpenClipboard 0
hData = GetClipboardData(CF_TEXT)
' ...
CloseClipboard

' After:
Dim d As New MSForms.DataObject
d.GetFromClipboard
text = d.GetText
'@
        Note = 'Requires reference to Microsoft Forms 2.0 Object Library, or use late binding.'
    }
    'GetClipboardData' = @{
        Lib = 'user32'
        Risk = 'MEDIUM'
        Alt = '(See OpenClipboard)'
        Example = ''
        Note = ''
    }
    'SetClipboardData' = @{
        Lib = 'user32'
        Risk = 'MEDIUM'
        Alt = 'MSForms.DataObject.SetText / PutInClipboard'
        Example = ''
        Note = ''
    }
    'CloseClipboard' = @{
        Lib = 'user32'
        Risk = 'LOW'
        Alt = '(Remove together with OpenClipboard)'
        Example = ''
        Note = ''
    }
}

# ============================================================================
# Scan file for API usage
# ============================================================================

Write-Host "Scanning: $FilePath"

$proj = Get-VbaProjectBytes $FilePath
if (-not $proj.Bytes) { Write-Host "No vbaProject.bin found." -ForegroundColor Yellow; exit 0 }
$ole2 = Read-Ole2 $proj.Bytes
$modules = Get-VbaModuleList $ole2

$found = [ordered]@{}  # apiName -> @{ Decl = @{File;Line;Sig}; Calls = @(@{File;Line;Code}) }

foreach ($mod in $modules) {
    $result = Get-VbaModuleCode $ole2 $mod.Name
    if (-not $result) { continue }
    $lines = ($result.Code -split "`r`n|`n")

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line -match '^\s*''') { continue }
        if ($line -match '(?i)\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)') {
            $apiName = $Matches[3]
            if (-not $found.Contains($apiName)) {
                $found[$apiName] = @{ Decl = $null; Calls = [System.Collections.ArrayList]::new() }
            }
            $found[$apiName].Decl = @{ File = "$($mod.Name).$($mod.Ext)"; Line = $i + 1; Sig = $line.Trim() }
        }
    }
}

# Find call sites
foreach ($apiName in @($found.Keys)) {
    foreach ($mod in $modules) {
        $result = Get-VbaModuleCode $ole2 $mod.Name
        if (-not $result) { continue }
        $lines = ($result.Code -split "`r`n|`n")
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            if ($line -match '^\s*''') { continue }
            if ($line -match '(?i)\bDeclare\s') { continue }
            if ($line -match "\b$([regex]::Escape($apiName))\b") {
                $trimmed = $line.Trim()
                if ($trimmed.Length -gt 100) { $trimmed = $trimmed.Substring(0, 97) + '...' }
                [void]$found[$apiName].Calls.Add(@{ File = "$($mod.Name).$($mod.Ext)"; Line = $i + 1; Code = $trimmed })
            }
        }
    }
}

if ($found.Count -eq 0) {
    Write-Host "No Win32 API declarations found." -ForegroundColor Cyan
    exit 0
}

Write-Host "Found $($found.Count) API(s)"

# ============================================================================
# Generate HTML — left: modules, center: code, right: API outline + hover tips
# ============================================================================

# Build code data per module (strip Attribute lines)
$allModCode = [ordered]@{}
foreach ($mod in $modules) {
    $result = Get-VbaModuleCode $ole2 $mod.Name
    if (-not $result) { continue }
    $lines = ($result.Code -split "`r`n|`n") | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_' }
    $allModCode["$($mod.Name).$($mod.Ext)"] = @($lines)
}

# Build per-module highlight + tooltip data
# For each module, which lines match which API, and what's the tooltip
$allApiNames = @($found.Keys)

$he = { param($s) [System.Net.WebUtility]::HtmlEncode($s) }

# Build tooltip JSON per API
$tooltipData = [ordered]@{}
foreach ($apiName in $allApiNames) {
    $info = $replacements[$apiName]
    $alt = if ($info) { $info.Alt } else { 'Not in database' }
    $note = if ($info -and $info.Note) { $info.Note } else { '' }
    $example = if ($info -and $info.Example) { $info.Example } else { '' }
    $tooltipData[$apiName] = @{ Alt = $alt; Note = $note; Example = $example }
}

$html = [System.Text.StringBuilder]::new()
[void]$html.Append(@"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>Cheatsheet: $([System.Net.WebUtility]::HtmlEncode([IO.Path]::GetFileName($FilePath)))</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: Consolas, 'Courier New', monospace; font-size: 13px; background: #1e1e1e; color: #d4d4d4; }
.header { background: #252526; padding: 10px 20px; border-bottom: 1px solid #3c3c3c; }
.header h1 { font-size: 15px; font-weight: normal; color: #cccccc; }
.header .sub { margin-top: 4px; font-size: 12px; color: #888; }
.main { display: flex; height: calc(100vh - 52px); }
.sidebar { width: 200px; min-width: 200px; background: #252526; border-right: 1px solid #3c3c3c; overflow-y: auto; padding: 8px 0; }
.sidebar .item { padding: 5px 16px; cursor: pointer; color: #888; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sidebar .item:hover { color: #d4d4d4; background: #2a2d2e; }
.sidebar .item.active { color: #ffffff; background: #37373d; border-left: 2px solid #0078d4; }
.sidebar .item.has-hl { color: #4fc1ff; }
.sidebar .item.no-hl { color: #606060; }
.content { flex: 1; overflow: auto; position: relative; }
.module { display: none; }
.module.active { display: block; }
.code-table { width: 100%; border-collapse: collapse; }
.code-table td { padding: 0 8px; line-height: 20px; vertical-align: top; white-space: pre; overflow: hidden; text-overflow: ellipsis; }
.code-table .ln { width: 50px; min-width: 50px; text-align: right; color: #606060; padding-right: 12px; user-select: none; border-right: 1px solid #3c3c3c; }
.code-table .code { color: #d4d4d4; }
tr.hl-api td.code { background: #1b2e4a; color: #a0c4f0; cursor: pointer; }
tr.hl-api td.ln { color: #cccccc; }
.outline { width: 250px; min-width: 250px; background: #252526; border-left: 1px solid #3c3c3c; overflow-y: auto; padding: 8px 0; }
.outline .ol-header { padding: 6px 12px; font-size: 11px; color: #888; text-transform: uppercase; }
.outline .ol-api { padding: 4px 12px 2px; font-size: 13px; color: #4fc1ff; font-weight: bold; }
.outline .ol-alt { display: none; }
.outline .ol-loc { padding: 2px 12px 2px 24px; font-size: 12px; color: #888; cursor: pointer; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.outline .ol-loc:hover { color: #d4d4d4; background: #2a2d2e; }
.outline .ol-sep { border-top: 1px solid #3c3c3c; margin: 6px 12px; }
.tooltip { position: fixed; background: #2d2d2d; border: 1px solid #555; border-radius: 4px; padding: 10px 14px; max-width: 500px; z-index: 100; display: none; font-size: 12px; line-height: 1.5; box-shadow: 0 4px 12px rgba(0,0,0,0.5); }
.tooltip .tt-api { color: #4fc1ff; font-weight: bold; font-size: 14px; }
.tooltip .tt-alt { color: #6a9955; margin-top: 4px; }
.tooltip .tt-note { color: #b0b0b0; font-style: italic; margin-top: 4px; }
.tooltip pre { background: #1e1e1e; border: 1px solid #3c3c3c; border-radius: 3px; padding: 8px; margin-top: 6px; font-size: 11px; line-height: 1.4; max-height: 200px; overflow-y: auto; }
.tooltip pre .cmt { color: #6a9955; }
.minimap { position: fixed; right: 250px; top: 52px; width: 14px; bottom: 0; background: #1e1e1e; border-left: 1px solid #3c3c3c; z-index: 20; cursor: pointer; }
.minimap .mark { position: absolute; right: 2px; width: 10px; height: 3px; border-radius: 1px; background: #4fc1ff; }
.minimap .viewport { position: absolute; right: 0; width: 14px; background: rgba(255,255,255,0.25); border-radius: 2px; pointer-events: none; }
</style>
</head>
<body>
<div class="header">
  <h1>Win32 API Cheatsheet</h1>
  <div class="sub">$([System.Net.WebUtility]::HtmlEncode([IO.Path]::GetFileName($FilePath))) &mdash; $($found.Count) API(s) detected</div>
</div>
<div class="main">
<div class="sidebar">
"@)

# Sidebar: modules
$modIdx = 0
$firstHlIdx = -1
foreach ($modLabel in $allModCode.Keys) {
    $lines = $allModCode[$modLabel]
    $hlCount = 0
    foreach ($line in $lines) {
        if ($line -match '^\s*''') { continue }
        foreach ($apiName in $allApiNames) {
            if ($line -match "\b$([regex]::Escape($apiName))\b") { $hlCount++; break }
        }
    }
    $cls = if ($hlCount -gt 0) { 'has-hl' } else { 'no-hl' }
    if ($firstHlIdx -eq -1 -and $hlCount -gt 0) { $firstHlIdx = $modIdx }
    $label = if ($hlCount -gt 0) { "$modLabel ($hlCount)" } else { $modLabel }
    [void]$html.Append("<div class=`"item $cls`" onclick=`"showTab($modIdx)`" id=`"tab$modIdx`">$(& $he $label)</div>")
    $modIdx++
}
if ($firstHlIdx -eq -1) { $firstHlIdx = 0 }

[void]$html.Append("</div><div class=`"content`">")

# Module panels
$modIdx = 0
foreach ($modLabel in $allModCode.Keys) {
    $lines = $allModCode[$modLabel]
    [void]$html.Append("<div class=`"module`" id=`"mod$modIdx`"><table class=`"code-table`">")
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $trClass = ''
        $dataApi = ''
        if ($line -notmatch '^\s*''') {
            foreach ($apiName in $allApiNames) {
                if ($line -match "\b$([regex]::Escape($apiName))\b") {
                    $trClass = 'hl-api'
                    $dataApi = $apiName
                    break
                }
            }
        }
        $ln = $i + 1
        $dataAttr = if ($dataApi) { " data-api=`"$(& $he $dataApi)`"" } else { '' }
        [void]$html.Append("<tr class=`"$trClass`"$dataAttr><td class=`"ln`">$ln</td><td class=`"code`">$(& $he $line)</td></tr>")
    }
    [void]$html.Append("</table></div>")
    $modIdx++
}

[void]$html.Append(@"
<div class="minimap" id="minimap"><div class="viewport" id="viewport"></div></div>
</div>
<div class="outline" id="outline"></div>
</div>
<div class="tooltip" id="tooltip"></div>
<script>
const content = document.querySelector('.content');
const minimap = document.getElementById('minimap');
const viewport = document.getElementById('viewport');
const outline = document.getElementById('outline');
const tooltip = document.getElementById('tooltip');

const apiInfo = {
"@)

# Emit API info as JS object
$first = $true
foreach ($apiName in $allApiNames) {
    $info = $tooltipData[$apiName]
    $altJs = ($info.Alt -replace '\\','\\\\' -replace "'","\'")
    $noteJs = ($info.Note -replace '\\','\\\\' -replace "'","\'")
    $exJs = ((& $he $info.Example) -replace '\\','\\\\' -replace "'","\'" -replace "`r`n",'\n' -replace "`n",'\n')
    $exJs = $exJs -replace "('.*?)\\n", '<span class="cmt">$1</span>\n'
    $exJs = $exJs -replace "('[^<]*)$", '<span class="cmt">$1</span>'
    $comma = if ($first) { '' } else { ',' }
    $first = $false
    [void]$html.Append("$comma'$(& $he $apiName)':{alt:'$altJs',note:'$noteJs',ex:'$exJs'}")
}

[void]$html.Append(@"
};

function showTab(idx) {
  document.querySelectorAll('.module').forEach(m => m.classList.remove('active'));
  document.querySelectorAll('.item').forEach(t => t.classList.remove('active'));
  document.getElementById('mod' + idx).classList.add('active');
  document.getElementById('tab' + idx).classList.add('active');
  content.scrollTop = 0;
  updateOutline();
  updateMinimap();
}

function updateOutline() {
  outline.innerHTML = '<div class="ol-header">API Usage</div>';
  const mod = document.querySelector('.module.active');
  if (!mod) return;
  const rows = mod.querySelectorAll('tr.hl-api');
  let lastApi = '';
  rows.forEach(r => {
    const api = r.dataset.api;
    if (api !== lastApi) {
      if (lastApi) outline.innerHTML += '<div class="ol-sep"></div>';
      outline.innerHTML += '<div class="ol-api">' + api + '</div>';
      const info = apiInfo[api];
      if (info) outline.innerHTML += '<div class="ol-alt">' + info.alt + '</div>';
      lastApi = api;
    }
    const ln = r.querySelector('.ln').textContent;
    const code = r.querySelector('.code').textContent.trim().substring(0, 50);
    const div = document.createElement('div');
    div.className = 'ol-loc';
    div.textContent = 'L' + ln + ': ' + code;
    div.addEventListener('click', () => r.scrollIntoView({block:'center'}));
    outline.appendChild(div);
  });
}

// Tooltip on hover
document.addEventListener('mouseover', (e) => {
  const tr = e.target.closest('tr.hl-api');
  if (!tr) { tooltip.style.display = 'none'; return; }
  const api = tr.dataset.api;
  const info = apiInfo[api];
  if (!info) return;
  let html = '<div class="tt-api">' + api + '</div>';
  html += '<div class="tt-alt">Alternative: ' + info.alt + '</div>';
  if (info.note) html += '<div class="tt-note">' + info.note + '</div>';
  if (info.ex) html += '<pre>' + info.ex.replace(/\\n/g, '\n') + '</pre>';
  tooltip.innerHTML = html;
  tooltip.style.display = 'block';
  const rect = tr.getBoundingClientRect();
  let top = rect.bottom + 4;
  let left = rect.left + 60;
  if (top + tooltip.offsetHeight > window.innerHeight) top = rect.top - tooltip.offsetHeight - 4;
  if (left + tooltip.offsetWidth > window.innerWidth - 270) left = window.innerWidth - 270 - tooltip.offsetWidth - 10;
  tooltip.style.top = top + 'px';
  tooltip.style.left = left + 'px';
});
document.addEventListener('mouseout', (e) => {
  if (!e.target.closest('tr.hl-api')) tooltip.style.display = 'none';
});

function updateMinimap() {
  minimap.querySelectorAll('.mark').forEach(m => m.remove());
  const mod = document.querySelector('.module.active');
  if (!mod) return;
  const rows = mod.querySelectorAll('tr.hl-api');
  const allRows = mod.querySelectorAll('tr');
  if (allRows.length === 0) return;
  const mapH = minimap.clientHeight;
  rows.forEach(r => {
    const idx = Array.from(allRows).indexOf(r);
    const mark = document.createElement('div');
    mark.className = 'mark';
    mark.style.top = (idx / allRows.length * mapH) + 'px';
    mark.addEventListener('click', () => r.scrollIntoView({block:'center'}));
    minimap.appendChild(mark);
  });
  updateViewport();
}
function updateViewport() {
  const sh = content.scrollHeight, ch = content.clientHeight, st = content.scrollTop;
  const mapH = minimap.clientHeight;
  if (sh <= ch) { viewport.style.display = 'none'; return; }
  viewport.style.display = '';
  viewport.style.top = (st / sh * mapH) + 'px';
  viewport.style.height = (ch / sh * mapH) + 'px';
}
content.addEventListener('scroll', updateViewport);
minimap.addEventListener('click', (e) => {
  if (e.target.classList.contains('mark')) return;
  content.scrollTop = e.offsetY / minimap.clientHeight * content.scrollHeight - content.clientHeight / 2;
});
showTab($firstHlIdx);
</script>
</body></html>
"@)

$baseName = [IO.Path]::GetFileNameWithoutExtension($FilePath)
$htmlPath = Join-Path ([IO.Path]::GetDirectoryName($FilePath)) "${baseName}_cheatsheet.html"
[IO.File]::WriteAllText($htmlPath, $html.ToString(), [System.Text.Encoding]::UTF8)

# Text log
$text = [System.Text.StringBuilder]::new()
[void]$text.AppendLine("# Win32 API Migration Cheatsheet")
[void]$text.AppendLine("# Source: $([IO.Path]::GetFileName($FilePath))")
[void]$text.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$text.AppendLine("")
foreach ($apiName in $found.Keys) {
    $f = $found[$apiName]
    $info = $replacements[$apiName]
    $alt = if ($info) { $info.Alt } else { 'Not in database' }
    [void]$text.AppendLine("## $apiName")
    [void]$text.AppendLine("  Alternative: $alt")
    if ($info -and $info.Note -ne '') { [void]$text.AppendLine("  Note: $($info.Note)") }
    if ($f.Decl) { [void]$text.AppendLine("  Declared: $($f.Decl.File) L$($f.Decl.Line)") }
    foreach ($call in $f.Calls) { [void]$text.AppendLine("  Called:   $($call.File) L$($call.Line): $($call.Code)") }
    [void]$text.AppendLine("")
}
$textPath = Join-Path ([IO.Path]::GetDirectoryName($FilePath)) "${baseName}_cheatsheet.txt"
[IO.File]::WriteAllText($textPath, $text.ToString(), [System.Text.Encoding]::UTF8)

Start-Process $htmlPath
Write-Host "Cheatsheet: $htmlPath" -ForegroundColor Green
Write-Host "Text log:   $textPath"
