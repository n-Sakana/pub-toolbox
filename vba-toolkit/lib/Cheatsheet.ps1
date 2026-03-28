param([Parameter(Mandatory)][string]$FilePath)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force

$FilePath = Resolve-VbaFilePath $FilePath
$fileName = [IO.Path]::GetFileName($FilePath)
$sw = [System.Diagnostics.Stopwatch]::StartNew()

Write-VbaHeader 'Cheatsheet' $fileName
Write-VbaLog 'Cheatsheet' $FilePath 'Started'

# ============================================================================
# Win32 API replacement database
# ============================================================================

$replacements = [ordered]@{
    # --- Timer / Sleep ---
    'GetTickCount' = @{
        Lib = 'kernel32'
        Alt = 'Timer (VBA built-in, Single type)'
        Example = @'
' Before:
Dim t As Long: t = GetTickCount()
DoSomething
Debug.Print "Elapsed: " & (GetTickCount() - t) & " ms"

' After:
Dim t As Single: t = Timer
DoSomething
Dim elapsed As Single: elapsed = Timer - t
If elapsed < 0 Then elapsed = elapsed + 86400  ' midnight rollover
Debug.Print "Elapsed: " & Format(elapsed * 1000, "0") & " ms"
'@
        Note = 'Timer is Single (~15ms resolution), resets at midnight. Add 86400 if elapsed < 0. GetTickCount wraps at ~49.7 days.'
    }
    'GetTickCount64' = @{
        Lib = 'kernel32'
        Alt = 'Timer (VBA built-in)'
        Example = '(Same as GetTickCount)'
        Note = ''
    }
    'Sleep' = @{
        Lib = 'kernel32'
        Alt = 'Application.Wait (Excel only) or DoEvents loop'
        Example = @'
' Before:
Sleep 1000  ' 1 second

' After (Option A - Excel only, 1sec resolution):
Application.Wait Now + TimeSerial(0, 0, 1)

' After (Option B - any host, sub-second, busy-wait):
Dim endTime As Single: endTime = Timer + 0.5  ' 500ms
Do While Timer < endTime: DoEvents: Loop
' Note: DoEvents loop uses 100% CPU on one core

' After (Option C - non-blocking delayed execution):
Application.OnTime Now + TimeSerial(0, 0, 1), "MyCallback"
'@
        Note = 'Application.Wait is Excel-only (not Word/Access/Outlook). DoEvents loop is a busy-wait. Application.OnTime is non-blocking but requires a callback Sub.'
    }
    'timeGetTime' = @{
        Lib = 'winmm'
        Alt = 'Timer (VBA built-in)'
        Example = '(Same as GetTickCount)'
        Note = ''
    }
    'QueryPerformanceCounter' = @{
        Lib = 'kernel32'
        Alt = 'No equivalent for high-resolution timing. Timer (~15ms) for rough measurements.'
        Example = @'
' Before:
QueryPerformanceCounter startCount
DoSomething
QueryPerformanceCounter endCount
elapsed = (endCount - startCount) / freq

' After (rough timing only):
Dim t As Single: t = Timer
DoSomething
Debug.Print "Elapsed: " & Format((Timer - t) * 1000, "0") & " ms"
'@
        Note = 'QPC provides sub-microsecond precision. Timer provides ~15ms at best. No pure VBA equivalent for high-resolution timing.'
    }
    'QueryPerformanceFrequency' = @{
        Lib = 'kernel32'
        Alt = '(Remove together with QueryPerformanceCounter)'
        Example = ''
        Note = ''
    }

    # --- String / Memory ---
    'CopyMemory' = @{
        Lib = 'kernel32 (RtlMoveMemory)'
        Alt = 'LSet (UDT copy), array assignment, or byte-by-byte copy'
        Example = @'
' Before:
CopyMemory ByVal dest, ByVal src, length

' After (byte arrays - direct assignment):
destBytes() = sourceBytes()

' After (byte-by-byte):
Dim i As Long
For i = 0 To length - 1
    dest(i) = src(i)
Next i

' After (UDT to UDT of same size):
LSet destUDT = sourceUDT

' After (in-place string modification):
Mid$(dest, pos, length) = Mid$(src, 1, length)
'@
        Note = 'LSet copies between UDTs of the same size without API. Mid$ as a statement (left-hand side) modifies strings in-place.'
    }
    'lstrlen' = @{
        Lib = 'kernel32'
        Alt = 'Len / LenB (VBA built-in)'
        Example = @'
' Before:
length = lstrlen(ByVal ptr)

' After:
length = Len(str)     ' character count
length = LenB(str)    ' byte count (= Len * 2 in VBA Unicode)
'@
        Note = 'If original code used lstrlen for ANSI buffer sizing, use LenB instead of Len.'
    }

    # --- User / System info ---
    'GetUserName' = @{
        Lib = 'advapi32'
        Alt = 'Environ$("USERNAME") for Windows login name'
        Example = @'
' Before:
Dim buf As String: buf = Space(256)
Dim sz As Long: sz = 256
GetUserName buf, sz
userName = Left$(buf, sz - 1)

' After (Windows login name):
userName = Environ$("USERNAME")

' CAUTION: Application.UserName is the Office display name,
' NOT the Windows login. These are often different in
' corporate environments.
'@
        Note = 'Environ$("USERNAME") = Windows login. Application.UserName = Office display name. These differ in corporate environments.'
    }
    'GetComputerName' = @{
        Lib = 'kernel32'
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
        Note = 'Environ$ returns empty string if variable is not set. Always validate the result.'
    }
    'GetTempPath' = @{
        Lib = 'kernel32'
        Alt = 'Environ$("TEMP")'
        Example = @'
' Before:
Dim buf As String: buf = Space(260)
GetTempPath 260, buf
tmpPath = Left$(buf, InStr(buf, vbNullChar) - 1)

' After:
tmpPath = Environ$("TEMP") & "\"
' Note: API appends trailing "\", Environ$ does not.
'@
        Note = 'GetTempPath appends trailing backslash. Environ$("TEMP") does not. Add "\" when concatenating paths.'
    }
    'GetSystemDirectory' = @{
        Lib = 'kernel32'
        Alt = 'Environ$("WINDIR") & "\System32"'
        Example = ''
        Note = 'Caution: 32-bit Office on 64-bit Windows uses SysWOW64. Environ$ always returns System32. Behavior may differ from the original API call.'
    }
    'GetWindowsDirectory' = @{
        Lib = 'kernel32'
        Alt = 'Environ$("WINDIR")'
        Example = ''
        Note = ''
    }

    # --- Window / UI ---
    'FindWindow' = @{
        Lib = 'user32'
        Alt = 'Application.hWnd (own window only, Excel 2010+) or AppActivate'
        Example = @'
' Before:
hWnd = FindWindow(vbNullString, "Window Title")

' After (get own Excel window handle):
hWnd = Application.hWnd  ' Excel 2010+

' After (activate by title - no handle returned):
AppActivate "Window Title"
'@
        Note = 'Application.hWnd only returns the host app window handle. FindWindow for other app windows has no VBA equivalent. AppActivate does partial title matching - may activate wrong window.'
    }
    'SetWindowPos' = @{
        Lib = 'user32'
        Alt = 'UserForm position properties (positioning only, no topmost)'
        Example = @'
' Before:
SetWindowPos hWnd, HWND_TOPMOST, x, y, w, h, SWP_NOSIZE

' After (UserForm positioning only):
Me.StartUpPosition = 0
Me.Left = x: Me.Top = y

' For "stay visible" behavior:
frm.Show vbModeless
'@
        Note = 'No VBA equivalent for HWND_TOPMOST. Positioning alternative applies to UserForms only, not the application window. vbModeless keeps form visible while user works.'
    }
    'GetSystemMetrics' = @{
        Lib = 'user32'
        Alt = 'Application.UsableWidth/Height (Excel, in points) - no pixel equivalent'
        Example = @'
' Before:
screenW = GetSystemMetrics(SM_CXSCREEN)  ' pixels
screenH = GetSystemMetrics(SM_CYSCREEN)  ' pixels

' After (Excel - workspace size in points, excludes taskbar):
workW = Application.UsableWidth    ' points
workH = Application.UsableHeight   ' points
' Note: 1 point = 1/72 inch. NOT pixels.

' CAUTION: Application.Width/Height is the Excel
' WINDOW size, not the screen size.
'@
        Note = 'Application.UsableWidth/Height = workspace in points (Excel only). For pixels, no pure VBA equivalent. In Access: Screen.Width/Height (twips).'
    }
    'ShowWindow' = @{
        Lib = 'user32'
        Alt = 'Application.Visible, WindowState, or UserForm.Show/Hide'
        Example = @'
' Before:
ShowWindow hWnd, SW_SHOW
ShowWindow hWnd, SW_MINIMIZE
ShowWindow hWnd, SW_MAXIMIZE

' After:
Application.Visible = True          ' show
Application.WindowState = xlMinimized  ' minimize
Application.WindowState = xlMaximized  ' maximize
Application.WindowState = xlNormal     ' restore

' For UserForm:
frm.Show / frm.Hide
'@
        Note = ''
    }
    'SetForegroundWindow' = @{
        Lib = 'user32'
        Alt = 'AppActivate (VBA built-in)'
        Example = @'
' Before:
SetForegroundWindow hWnd

' After:
On Error Resume Next  ' raises error 5 if not found
AppActivate "Window Title"
On Error GoTo 0
'@
        Note = 'AppActivate does partial title matching - may activate the wrong window if titles are similar. Always wrap in error handling.'
    }
    'SendMessage' = @{
        Lib = 'user32'
        Alt = 'Depends on message type. Often no direct alternative.'
        Example = ''
        Note = 'SendMessage is highly versatile. Review each call site individually. Common uses: scrolling listboxes, setting control properties. If used for external app automation, the business process itself may need redesign.'
    }
    'PostMessage' = @{
        Lib = 'user32'
        Alt = '(Same as SendMessage - review individually)'
        Example = ''
        Note = ''
    }

    # --- File ---
    'SHFileOperation' = @{
        Lib = 'shell32'
        Alt = 'FileCopy / Kill / Name / MkDir / RmDir or FileSystemObject'
        Example = @'
' Before:
SHFileOperation fileOp  ' copy/move/delete with recycle bin

' After:
' Copy file:   FileCopy src, dst
' Move/rename: Name src As dst
' Delete:      Kill path  ' permanent, no recycle bin
' Create dir:  MkDir path
' Remove dir:  RmDir path  ' must be empty

' Or use FileSystemObject for folders:
' fso.CopyFolder / fso.DeleteFolder / fso.MoveFolder
' Note: fso.DeleteFile is also permanent (no recycle bin)

' Delete (recycle bin): no pure VBA equivalent
'@
        Note = 'Kill does not support wildcards in the path portion (only filename). Recycle bin delete has no VBA equivalent. FileSystemObject may also be restricted by EDR.'
    }
    'ShellExecute' = @{
        Lib = 'shell32'
        Alt = 'ThisWorkbook.FollowHyperlink (documents) or Shell (executables only)'
        Example = @'
' Before:
ShellExecute 0, "open", path, vbNullString, vbNullString, SW_SHOW

' After (open document/URL with default app - Excel only):
ThisWorkbook.FollowHyperlink path
' Note: may trigger security warnings

' After (run executable only - not documents):
Shell "notepad.exe C:\file.txt", vbNormalFocus
' CAUTION: Shell cannot open .pdf, .xlsx etc.
' by file association. Use FollowHyperlink instead.
'@
        Note = 'Shell only launches executables, not documents by association. FollowHyperlink is Excel-specific and may trigger security prompts.'
    }

    # --- Clipboard ---
    'OpenClipboard' = @{
        Lib = 'user32'
        Alt = 'MSForms.DataObject (text only)'
        Example = @'
' Before:
OpenClipboard 0
hData = GetClipboardData(CF_TEXT)
' ...
CloseClipboard

' After (text only):
Dim d As New MSForms.DataObject
d.GetFromClipboard
text = d.GetText
'@
        Note = 'MSForms.DataObject handles text only (no images/files). Requires Microsoft Forms 2.0 reference. Add error handling for clipboard lock failures.'
    }
    'GetClipboardData' = @{
        Lib = 'user32'
        Alt = '(See OpenClipboard - text only via MSForms.DataObject)'
        Example = ''
        Note = ''
    }
    'SetClipboardData' = @{
        Lib = 'user32'
        Alt = 'MSForms.DataObject.SetText / PutInClipboard (text only)'
        Example = ''
        Note = ''
    }
    'CloseClipboard' = @{
        Lib = 'user32'
        Alt = '(Remove together with OpenClipboard)'
        Example = ''
        Note = ''
    }
}

# ============================================================================
# Scan file for API usage
# ============================================================================

Write-VbaStatus 'Cheatsheet' $fileName "Scanning..."

$project = Get-AllModuleCode $FilePath -StripAttributes
if (-not $project) { Write-VbaError 'Cheatsheet' $fileName 'No vbaProject.bin found'; exit 0 }

$found = [ordered]@{}

# Pass 1: find declarations
foreach ($modName in $project.Modules.Keys) {
    $mod = $project.Modules[$modName]
    for ($i = 0; $i -lt $mod.Lines.Count; $i++) {
        $line = $mod.Lines[$i]
        if ($line -match '^\s*''') { continue }
        if ($line -match '(?i)\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)') {
            $apiName = $Matches[3]
            if (-not $found.Contains($apiName)) {
                $found[$apiName] = @{ Decl = $null; Calls = [System.Collections.ArrayList]::new() }
            }
            $found[$apiName].Decl = @{ File = "$modName.$($mod.Ext)"; Line = $i + 1; Sig = $line.Trim() }
        }
    }
}

# Pass 2: find call sites (single pass over all modules)
foreach ($apiName in @($found.Keys)) {
    foreach ($modName in $project.Modules.Keys) {
        $mod = $project.Modules[$modName]
        $lines = $mod.Lines
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            if ($line -match '^\s*''') { continue }
            if ($line -match '(?i)\bDeclare\s') { continue }
            if ($line -match "\b$([regex]::Escape($apiName))\b") {
                $trimmed = $line.Trim()
                if ($trimmed.Length -gt 100) { $trimmed = $trimmed.Substring(0, 97) + '...' }
                [void]$found[$apiName].Calls.Add(@{ File = "$modName.$($mod.Ext)"; Line = $i + 1; Code = $trimmed })
            }
        }
    }
}

if ($found.Count -eq 0) {
    Write-VbaStatus 'Cheatsheet' $fileName "No Win32 API declarations found"
    Write-VbaLog 'Cheatsheet' $FilePath 'No API found'
    exit 0
}

Write-VbaStatus 'Cheatsheet' $fileName "Found $($found.Count) API(s)"

# ============================================================================
# Generate HTML — left: modules, center: code, right: API outline + hover tips
# ============================================================================

# Build code data per module (already stripped via Get-AllModuleCode)
$allModCode = [ordered]@{}
foreach ($modName in $project.Modules.Keys) {
    $mod = $project.Modules[$modName]
    $allModCode["$modName.$($mod.Ext)"] = @($mod.Lines)
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
.hover-hint { position: fixed; background: #444; color: #ccc; padding: 2px 8px; border-radius: 3px; font-size: 11px; pointer-events: none; z-index: 50; display: none; }
.outline { width: 250px; min-width: 250px; background: #252526; border-left: 1px solid #3c3c3c; overflow-y: auto; padding: 8px 0; }
.outline .ol-header { padding: 6px 12px; font-size: 11px; color: #888; text-transform: uppercase; }
.outline .ol-api { padding: 4px 12px 2px; font-size: 13px; color: #4fc1ff; font-weight: bold; }
.outline .ol-loc { padding: 2px 12px 2px 24px; font-size: 12px; color: #888; cursor: pointer; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.outline .ol-loc:hover { color: #d4d4d4; background: #2a2d2e; }
.outline .ol-sep { border-top: 1px solid #3c3c3c; margin: 6px 12px; }
.tooltip { position: fixed; background: #2d2d2d; border: 1px solid #555; border-radius: 4px; padding: 10px 14px; max-width: 500px; z-index: 100; display: none; font-size: 12px; line-height: 1.5; box-shadow: 0 4px 12px rgba(0,0,0,0.5); user-select: text; }
.tooltip .tt-api { color: #4fc1ff; font-weight: bold; font-size: 14px; }
.tooltip .tt-alt { color: #6a9955; margin-top: 4px; }
.tooltip .tt-note { color: #b0b0b0; font-style: italic; margin-top: 4px; }
.tooltip pre { background: #1e1e1e; border: 1px solid #3c3c3c; border-radius: 3px; padding: 8px; margin-top: 6px; font-size: 11px; line-height: 1.4; max-height: 200px; overflow-y: auto; position: relative; }
.tooltip pre .cmt { color: #6a9955; }
.tooltip .tt-copy { position: absolute; top: 6px; right: 6px; background: none; border: none; cursor: pointer; opacity: 0.5; padding: 2px; }
.tooltip .tt-copy:hover { opacity: 1; }
.tooltip .tt-copy svg { width: 14px; height: 14px; fill: #ccc; }
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
<div class="hover-hint" id="hoverHint">Click for details</div>
<script>
const content = document.querySelector('.content');
const minimap = document.getElementById('minimap');
const viewport = document.getElementById('viewport');
const outline = document.getElementById('outline');
const tooltip = document.getElementById('tooltip');
const hoverHint = document.getElementById('hoverHint');

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

function scrollToRow(r) {
  const rRect = r.getBoundingClientRect();
  const cRect = content.getBoundingClientRect();
  const offset = rRect.top - cRect.top + content.scrollTop;
  content.scrollTo({ top: offset - content.clientHeight / 3, behavior: 'smooth' });
}

function updateOutline() {
  outline.innerHTML = '';
  const hdr = document.createElement('div');
  hdr.className = 'ol-header'; hdr.textContent = 'API Usage';
  outline.appendChild(hdr);
  const mod = document.querySelector('.module.active');
  if (!mod) return;
  const rows = mod.querySelectorAll('tr.hl-api');
  let lastApi = '';
  rows.forEach(r => {
    const api = r.dataset.api;
    if (api !== lastApi) {
      if (lastApi) { const sep = document.createElement('div'); sep.className = 'ol-sep'; outline.appendChild(sep); }
      const apiDiv = document.createElement('div');
      apiDiv.className = 'ol-api'; apiDiv.textContent = api;
      outline.appendChild(apiDiv);
      lastApi = api;
    }
    const ln = r.querySelector('.ln').textContent;
    const code = r.querySelector('.code').textContent.trim().substring(0, 50);
    const loc = document.createElement('div');
    loc.className = 'ol-loc';
    loc.textContent = 'L' + ln + ': ' + code;
    loc.addEventListener('click', () => scrollToRow(r));
    outline.appendChild(loc);
  });
}

// Tooltip: click to pin, click again or X to close
let pinnedTooltip = null;
function showTooltipAt(tr) {
  const api = tr.dataset.api;
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
  const tr = e.target.closest('tr.hl-api');
  if (tr && !pinnedTooltip) {
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
  const tr = e.target.closest('tr.hl-api');
  if (!tr) { tooltip.style.display = 'none'; pinnedTooltip = null; return; }
  if (pinnedTooltip === tr) {
    tooltip.style.display = 'none'; pinnedTooltip = null;
  } else {
    showTooltipAt(tr);
  }
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
    mark.addEventListener('click', () => scrollToRow(r));
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

$outDir = New-VbaOutputDir $FilePath 'cheatsheet'
$htmlPath = Join-Path $outDir 'cheatsheet.html'
[IO.File]::WriteAllText($htmlPath, $html.ToString(), [System.Text.Encoding]::UTF8)

# Text log
$text = [System.Text.StringBuilder]::new()
[void]$text.AppendLine("# Win32 API Migration Cheatsheet")
[void]$text.AppendLine("# Source: $fileName")
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
[IO.File]::WriteAllText((Join-Path $outDir 'cheatsheet.txt'), $text.ToString(), [System.Text.Encoding]::UTF8)

Start-Process $htmlPath

$sw.Stop()
Write-VbaResult 'Cheatsheet' $fileName "$($found.Count) API(s) documented" $outDir $sw.Elapsed.TotalSeconds
Write-VbaLog 'Cheatsheet' $FilePath "$($found.Count) APIs | -> $outDir"
