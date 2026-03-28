$ErrorActionPreference = 'Stop'

# ============================================================================
# Environment Probe - Tests EDR/compat patterns via Excel COM
# Creates temporary xlsm files, injects VBA code, tests save/run/result
# ============================================================================

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$results = [System.Collections.ArrayList]::new()
$tempDir = Join-Path ([IO.Path]::GetTempPath()) "vba-probe-$(Get-Date -Format yyyyMMddHHmmss)"
New-Item $tempDir -ItemType Directory -Force | Out-Null

function Add-ProbeResult {
    param([string]$Level, [string]$Category, [string]$Pattern, [string]$Target,
          [string]$Result, [string]$Phase = '', [int]$ErrNum = 0, [string]$ErrMsg = '', [string]$Detail = '')
    [void]$script:results.Add([ordered]@{
        Level = $Level; Category = $Category; Pattern = $Pattern; Target = $Target
        Result = $Result; Phase = $Phase; ErrNum = $ErrNum; ErrMsg = $ErrMsg; Detail = $Detail
    })
    $color = switch ($Result) { 'OK' { 'Green' } 'FAIL' { 'Red' } 'SKIP' { 'DarkGray' } default { 'Yellow' } }
    Write-Host "  [$Result] $Pattern - $Target $(if($Phase){"($Phase) "})$Detail" -ForegroundColor $color
}

# ============================================================================
# Test runner: create xlsm, inject code, save, optionally run, clean up
# ============================================================================

function Test-VbaCode {
    param(
        [string]$Level,
        [string]$Category,
        [string]$Pattern,
        [string]$Target,
        [string]$VbaCode,
        [string]$RunMacro = '',     # macro name to execute after save (empty = save-only test)
        [switch]$ExpectSaveFail     # if true, save failure = expected (e.g. EDR blocks Declare)
    )

    $testFile = Join-Path $tempDir "probe_$([guid]::NewGuid().ToString('N').Substring(0,8)).xlsm"
    $excel = $null
    $wb = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false

        $wb = $excel.Workbooks.Add()

        # Inject VBA code
        try {
            $mod = $wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
            $mod.Name = 'ProbeTest'
            $mod.CodeModule.AddFromString($VbaCode)
        } catch {
            Add-ProbeResult $Level $Category $Pattern $Target 'FAIL' 'Inject' 0 $_.Exception.Message
            return
        }

        # Save
        try {
            $wb.SaveAs($testFile, 52)  # xlOpenXMLWorkbookMacroEnabled
        } catch {
            if ($ExpectSaveFail) {
                Add-ProbeResult $Level $Category $Pattern $Target 'FAIL' 'Save (expected)' 0 $_.Exception.Message 'EDR likely blocked save'
            } else {
                Add-ProbeResult $Level $Category $Pattern $Target 'FAIL' 'Save' 0 $_.Exception.Message
            }
            return
        }

        # If save succeeded but we expected it to fail
        if ($ExpectSaveFail) {
            Add-ProbeResult $Level $Category $Pattern $Target 'OK' 'Save' 0 '' 'Save succeeded (EDR did not block)'
        }

        # Run macro if specified
        if ($RunMacro) {
            try {
                $runResult = $excel.Run($RunMacro)
                Add-ProbeResult $Level $Category $Pattern $Target 'OK' 'Run' 0 '' "Result: $runResult"
            } catch {
                Add-ProbeResult $Level $Category $Pattern $Target 'FAIL' 'Run' 0 $_.Exception.Message
            }
        } elseif (-not $ExpectSaveFail) {
            Add-ProbeResult $Level $Category $Pattern $Target 'OK' 'Save' 0 '' 'Code accepted'
        }

    } catch {
        Add-ProbeResult $Level $Category $Pattern $Target 'FAIL' 'Setup' 0 $_.Exception.Message
    } finally {
        try { if ($wb) { $wb.Close($false) } } catch {}
        try { if ($excel) { $excel.Quit() } } catch {}
        if ($wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
        if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        Remove-Item $testFile -Force -ErrorAction SilentlyContinue
    }
}

# ============================================================================
# Simple COM/function tests (no Excel COM needed)
# ============================================================================

function Test-CreateObject {
    param([string]$Level, [string]$Category, [string]$Pattern, [string]$ProgId)
    try {
        $obj = New-Object -ComObject $ProgId
        Add-ProbeResult $Level $Category $Pattern $ProgId 'OK' 'Create'
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
    } catch {
        Add-ProbeResult $Level $Category $Pattern $ProgId 'FAIL' 'Create' 0 $_.Exception.Message
    }
}

# ============================================================================
# Mode selection
# ============================================================================

Write-Host "=== Environment Probe ===" -ForegroundColor Cyan
Write-Host ""

$mode = Read-Host "Run mode? (B=Basic only, E=Basic+Extended, Q=Quit)"
if ($mode -eq 'Q' -or $mode -eq 'q') { exit 0 }
$runExtended = ($mode -eq 'E' -or $mode -eq 'e')

Write-Host ""
Write-Host "Mode: $(if($runExtended){'Basic + Extended'}else{'Basic only'})" -ForegroundColor Gray
Write-Host ""

# ============================================================================
# System Info
# ============================================================================

Write-Host "--- System Info ---" -ForegroundColor Cyan
Add-ProbeResult 'Aux' 'SystemInfo' 'Computer' 'Environ' 'OK' '' 0 '' $env:COMPUTERNAME
Add-ProbeResult 'Aux' 'SystemInfo' 'User' 'Environ' 'OK' '' 0 '' $env:USERNAME

# Get Office info via COM
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $ver = $xl.Version
    $build = $xl.Build
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Add-ProbeResult 'Aux' 'SystemInfo' 'Office Version' 'Excel.Application' 'OK' '' 0 '' "Version: $ver Build: $build"
} catch {
    Add-ProbeResult 'Aux' 'SystemInfo' 'Office Version' 'Excel.Application' 'FAIL' '' 0 $_.Exception.Message
}

# ============================================================================
# Basic Tests
# ============================================================================

Write-Host ""
Write-Host "--- Basic Tests ---" -ForegroundColor Cyan

# COM CreateObject tests
Test-CreateObject 'Basic' 'EDR' 'COM / CreateObject' 'Scripting.FileSystemObject'
Test-CreateObject 'Basic' 'EDR' 'COM / CreateObject' 'Scripting.Dictionary'
Test-CreateObject 'Basic' 'EDR' 'COM / CreateObject' 'ADODB.Connection'
Test-CreateObject 'Basic' 'EDR' 'COM / CreateObject' 'MSXML2.XMLHTTP.6.0'

# File I/O
$testFilePath = Join-Path $tempDir 'probe_fileio.txt'
try {
    [IO.File]::WriteAllText($testFilePath, 'probe test')
    Remove-Item $testFilePath -Force
    Add-ProbeResult 'Basic' 'EDR' 'File I/O' 'Write+Delete' 'OK' 'Run'
} catch {
    Add-ProbeResult 'Basic' 'EDR' 'File I/O' 'Write+Delete' 'FAIL' 'Run' 0 $_.Exception.Message
}

# Registry (VBA GetSetting equivalent)
try {
    $regPath = 'HKCU:\Software\VB and VBA Program Settings\ProbeTest\TestSection'
    New-Item $regPath -Force | Out-Null
    Set-ItemProperty $regPath -Name 'TestKey' -Value 'probe' -Force
    Remove-Item 'HKCU:\Software\VB and VBA Program Settings\ProbeTest' -Recurse -Force
    Add-ProbeResult 'Basic' 'EDR' 'Registry' 'GetSetting/SaveSetting' 'OK' 'Run'
} catch {
    Add-ProbeResult 'Basic' 'EDR' 'Registry' 'GetSetting/SaveSetting' 'FAIL' 'Run' 0 $_.Exception.Message
}

# Environ
Add-ProbeResult 'Basic' 'EDR' 'Environment' 'Environ' 'OK' '' 0 '' "USERNAME=$env:USERNAME"

# DAO
Test-CreateObject 'Basic' 'Compat' 'Deprecated: DAO' 'DAO.DBEngine.36'

# Legacy Controls
Test-CreateObject 'Basic' 'Compat' 'Deprecated: Legacy Controls' 'MSComDlg.CommonDialog'
Test-CreateObject 'Basic' 'Compat' 'Deprecated: Legacy Controls' 'MSCAL.Calendar'

# IE
Test-CreateObject 'Basic' 'Compat' 'Deprecated: IE Automation' 'InternetExplorer.Application'

# ============================================================================
# VBA injection tests (save/run via Excel COM)
# ============================================================================

Write-Host ""
Write-Host "--- VBA Injection Tests ---" -ForegroundColor Cyan

# Win32 API Declare - test if file with Declare can be saved
Test-VbaCode 'Basic' 'EDR' 'Win32 API (Declare)' 'Declare PtrSafe Function' -ExpectSaveFail -VbaCode @'
Option Explicit
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Public Function TestAPI() As String
    TestAPI = "OK:" & GetTickCount()
End Function
'@

# FileSystemObject via VBA
Test-VbaCode 'Basic' 'EDR' 'FileSystemObject' 'FSO in VBA' -RunMacro 'ProbeTest.TestFSO' -VbaCode @'
Option Explicit
Public Function TestFSO() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    TestFSO = "OK:" & fso.GetTempName()
End Function
'@

# Clipboard via VBA
Test-VbaCode 'Basic' 'EDR' 'Clipboard' 'MSForms.DataObject in VBA' -RunMacro 'ProbeTest.TestClipboard' -VbaCode @'
Option Explicit
Public Function TestClipboard() As String
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If Err.Number <> 0 Then
        TestClipboard = "FAIL:" & Err.Description
    Else
        d.SetText "probe"
        d.PutInClipboard
        TestClipboard = "OK"
    End If
End Function
'@

# ============================================================================
# Extended Tests
# ============================================================================

if ($runExtended) {
    Write-Host ""
    Write-Host "--- Extended Tests ---" -ForegroundColor Cyan

    # Shell via WScript.Shell
    try {
        $wsh = New-Object -ComObject WScript.Shell
        $exitCode = $wsh.Run('cmd /c echo probe', 0, $true)
        Add-ProbeResult 'Extended' 'EDR' 'Shell / process' 'cmd /c echo' 'OK' 'Run' 0 '' "ExitCode: $exitCode"
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsh)
    } catch {
        Add-ProbeResult 'Extended' 'EDR' 'Shell / process' 'cmd /c echo' 'FAIL' 'Run' 0 $_.Exception.Message
    }

    # PowerShell
    try {
        $wsh = New-Object -ComObject WScript.Shell
        $exitCode = $wsh.Run('powershell -Command exit', 0, $true)
        Add-ProbeResult 'Extended' 'EDR' 'PowerShell / WScript' 'powershell -Command exit' 'OK' 'Run' 0 '' "ExitCode: $exitCode"
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsh)
    } catch {
        Add-ProbeResult 'Extended' 'EDR' 'PowerShell / WScript' 'powershell -Command exit' 'FAIL' 'Run' 0 $_.Exception.Message
    }

    # WMI
    try {
        $wmi = [wmiclass]'Win32_Process'
        Add-ProbeResult 'Extended' 'EDR' 'Process / WMI' 'Win32_Process' 'OK' 'Create'
    } catch {
        Add-ProbeResult 'Extended' 'EDR' 'Process / WMI' 'Win32_Process' 'FAIL' 'Create' 0 $_.Exception.Message
    }

    # DDE via VBA
    Test-VbaCode 'Extended' 'Compat' 'Deprecated: DDE' 'DDEInitiate' -RunMacro 'ProbeTest.TestDDE' -VbaCode @'
Option Explicit
Public Function TestDDE() As String
    On Error Resume Next
    Dim ch As Long
    ch = DDEInitiate("Excel", "Sheet1")
    If Err.Number <> 0 Then
        TestDDE = "FAIL:" & Err.Description
    Else
        DDETerminate ch
        TestDDE = "OK"
    End If
End Function
'@

    # SendKeys via VBA
    Test-VbaCode 'Extended' 'EDR' 'SendKeys' 'SendKeys (empty)' -RunMacro 'ProbeTest.TestSendKeys' -VbaCode @'
Option Explicit
Public Function TestSendKeys() As String
    On Error Resume Next
    SendKeys ""
    If Err.Number <> 0 Then
        TestSendKeys = "FAIL:" & Err.Description
    Else
        TestSendKeys = "OK"
    End If
End Function
'@
}

# ============================================================================
# Output
# ============================================================================

Write-Host ""
Write-Host "--- Writing Results ---" -ForegroundColor Cyan

$outDir = Join-Path (Get-Location) 'output'
if (-not (Test-Path $outDir)) { New-Item $outDir -ItemType Directory -Force | Out-Null }
$outPath = Join-Path $outDir "probe_result_${env:COMPUTERNAME}_$(Get-Date -Format yyyyMMdd_HHmmss).txt"

$sb = [System.Text.StringBuilder]::new()
[void]$sb.AppendLine("# Environment Probe Results")
[void]$sb.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$sb.AppendLine("# Computer: $env:COMPUTERNAME")
[void]$sb.AppendLine("# User: $env:USERNAME")
[void]$sb.AppendLine("# Mode: $(if($runExtended){'Basic + Extended'}else{'Basic only'})")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("Level`tCategory`tPattern`tTarget`tResult`tPhase`tErrMsg`tDetail")

foreach ($r in $results) {
    [void]$sb.AppendLine("$($r.Level)`t$($r.Category)`t$($r.Pattern)`t$($r.Target)`t$($r.Result)`t$($r.Phase)`t$($r.ErrMsg)`t$($r.Detail)")
}

$utf8Bom = New-Object System.Text.UTF8Encoding $true
[IO.File]::WriteAllText($outPath, $sb.ToString(), $utf8Bom)

# Cleanup
Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue

# Summary
$sw.Stop()
$okCount = ($results | Where-Object { $_.Result -eq 'OK' }).Count
$failCount = ($results | Where-Object { $_.Result -eq 'FAIL' }).Count
$skipCount = ($results | Where-Object { $_.Result -eq 'SKIP' }).Count

Write-Host ""
Write-Host "=== Results ===" -ForegroundColor Cyan
Write-Host "  OK:   $okCount" -ForegroundColor Green
Write-Host "  FAIL: $failCount" -ForegroundColor $(if($failCount -gt 0){'Red'}else{'Green'})
Write-Host "  SKIP: $skipCount" -ForegroundColor DarkGray
Write-Host "  Time: $([Math]::Round($sw.Elapsed.TotalSeconds, 1))s" -ForegroundColor Gray
Write-Host "  Output: $outPath" -ForegroundColor Gray
