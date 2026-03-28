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
# Generate HTML cheatsheet
# ============================================================================

$he = { param($s) [System.Net.WebUtility]::HtmlEncode($s) }

$html = [System.Text.StringBuilder]::new()
[void]$html.Append(@"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>Win32 API Cheatsheet: $([IO.Path]::GetFileName($FilePath))</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Segoe UI', Meiryo, sans-serif; font-size: 14px; background: #1e1e1e; color: #d4d4d4; padding: 20px 40px; }
h1 { font-size: 18px; color: #cccccc; margin-bottom: 4px; }
.subtitle { font-size: 12px; color: #888; margin-bottom: 24px; }
.api-card { background: #252526; border: 1px solid #3c3c3c; border-radius: 6px; margin-bottom: 16px; overflow: hidden; }
.api-header { padding: 12px 16px; display: flex; justify-content: space-between; align-items: center; cursor: pointer; }
.api-header:hover { background: #2a2d2e; }
.api-name { font-family: Consolas, monospace; font-size: 16px; font-weight: bold; color: #4fc1ff; }
.api-lib { font-size: 12px; color: #888; }
.risk { padding: 2px 8px; border-radius: 3px; font-size: 11px; font-weight: bold; }
.risk-LOW { background: #1b3a1b; color: #6a9955; }
.risk-MEDIUM { background: #4b3a00; color: #e8ab53; }
.risk-HIGH { background: #4b1818; color: #f44747; }
.risk-UNKNOWN { background: #333; color: #888; }
.api-body { padding: 0 16px 16px; display: none; }
.api-body.open { display: block; }
.section-label { font-size: 11px; color: #888; text-transform: uppercase; margin-top: 12px; margin-bottom: 4px; }
.alt { color: #6a9955; font-weight: bold; }
.location { font-family: Consolas, monospace; font-size: 12px; color: #888; padding: 2px 0; }
.location .file { color: #4fc1ff; }
pre { background: #1e1e1e; border: 1px solid #3c3c3c; border-radius: 4px; padding: 12px; margin-top: 4px; font-family: Consolas, monospace; font-size: 13px; overflow-x: auto; line-height: 1.5; }
pre .comment { color: #6a9955; }
.note { font-size: 12px; color: #b0b0b0; margin-top: 8px; font-style: italic; }
.summary { background: #252526; border: 1px solid #3c3c3c; border-radius: 6px; padding: 16px; margin-bottom: 24px; }
.summary-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin-top: 8px; }
.summary-item { text-align: center; padding: 8px; border-radius: 4px; }
</style>
</head>
<body>
<h1>Win32 API Migration Cheatsheet</h1>
<div class="subtitle">$([IO.Path]::GetFileName($FilePath)) &mdash; $(Get-Date -Format 'yyyy-MM-dd')</div>
"@)

# Summary
$lowCount = 0; $medCount = 0; $highCount = 0; $unknownCount = 0
foreach ($apiName in $found.Keys) {
    $info = $replacements[$apiName]
    if ($info) {
        switch ($info.Risk) { 'LOW' { $lowCount++ } 'MEDIUM' { $medCount++ } 'HIGH' { $highCount++ } }
    } else { $unknownCount++ }
}

[void]$html.Append(@"
<div class="summary">
  <strong>$($found.Count) API(s) detected</strong>
  <div class="summary-grid">
    <div class="summary-item risk-LOW">LOW: $lowCount</div>
    <div class="summary-item risk-MEDIUM">MEDIUM: $medCount</div>
    <div class="summary-item risk-HIGH">HIGH: $highCount</div>
  </div>
  $(if ($unknownCount -gt 0) { "<div style='margin-top:8px;color:#888;'>Unknown (not in database): $unknownCount</div>" })
</div>
"@)

# API cards
foreach ($apiName in $found.Keys) {
    $f = $found[$apiName]
    $info = $replacements[$apiName]
    $risk = if ($info) { $info.Risk } else { 'UNKNOWN' }
    $lib = if ($info) { $info.Lib } else { $f.Decl.Sig -replace '.*Lib\s+(\S+).*','$1' }
    $alt = if ($info) { $info.Alt } else { 'No known alternative in database' }

    [void]$html.Append("<div class=`"api-card`">")
    [void]$html.Append("<div class=`"api-header`" onclick=`"this.nextElementSibling.classList.toggle('open')`">")
    [void]$html.Append("<div><span class=`"api-name`">$(& $he $apiName)</span> <span class=`"api-lib`">$(& $he $lib)</span></div>")
    [void]$html.Append("<span class=`"risk risk-$risk`">$risk</span>")
    [void]$html.Append("</div>")
    [void]$html.Append("<div class=`"api-body`">")

    # Alternative
    [void]$html.Append("<div class=`"section-label`">Alternative</div>")
    [void]$html.Append("<div class=`"alt`">$(& $he $alt)</div>")

    # Usage locations
    [void]$html.Append("<div class=`"section-label`">Usage in this file</div>")
    if ($f.Decl) {
        [void]$html.Append("<div class=`"location`"><span class=`"file`">$($f.Decl.File)</span> L$($f.Decl.Line) (declaration)</div>")
    }
    foreach ($call in $f.Calls) {
        [void]$html.Append("<div class=`"location`"><span class=`"file`">$($call.File)</span> L$($call.Line): $(& $he $call.Code)</div>")
    }

    # Example
    if ($info -and $info.Example -and $info.Example -ne '') {
        [void]$html.Append("<div class=`"section-label`">Migration Example</div>")
        $exHtml = (& $he $info.Example) -replace "('.*)", '<span class="comment">$1</span>'
        [void]$html.Append("<pre>$exHtml</pre>")
    }

    # Note
    if ($info -and $info.Note -and $info.Note -ne '') {
        [void]$html.Append("<div class=`"note`">$(& $he $info.Note)</div>")
    }

    [void]$html.Append("</div></div>")
}

[void]$html.Append(@"
<script>
// Open first card by default
document.querySelector('.api-body').classList.add('open');
</script>
</body></html>
"@)

$baseName = [IO.Path]::GetFileNameWithoutExtension($FilePath)
$htmlPath = Join-Path ([IO.Path]::GetDirectoryName($FilePath)) "${baseName}_cheatsheet.html"
[IO.File]::WriteAllText($htmlPath, $html.ToString(), [System.Text.Encoding]::UTF8)

# Also output text version
$text = [System.Text.StringBuilder]::new()
[void]$text.AppendLine("# Win32 API Migration Cheatsheet")
[void]$text.AppendLine("# Source: $([IO.Path]::GetFileName($FilePath))")
[void]$text.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$text.AppendLine("")
foreach ($apiName in $found.Keys) {
    $f = $found[$apiName]
    $info = $replacements[$apiName]
    $risk = if ($info) { $info.Risk } else { 'UNKNOWN' }
    $alt = if ($info) { $info.Alt } else { 'Not in database' }
    [void]$text.AppendLine("## $apiName [$risk]")
    [void]$text.AppendLine("  Alternative: $alt")
    if ($f.Decl) { [void]$text.AppendLine("  Declared: $($f.Decl.File) L$($f.Decl.Line)") }
    foreach ($call in $f.Calls) { [void]$text.AppendLine("  Called:   $($call.File) L$($call.Line): $($call.Code)") }
    [void]$text.AppendLine("")
}
$textPath = Join-Path ([IO.Path]::GetDirectoryName($FilePath)) "${baseName}_cheatsheet.txt"
[IO.File]::WriteAllText($textPath, $text.ToString(), [System.Text.Encoding]::UTF8)

Start-Process $htmlPath
Write-Host "Cheatsheet: $htmlPath" -ForegroundColor Green
Write-Host "Text log:   $textPath"
