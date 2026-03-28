Attribute VB_Name = "Probe"
Option Explicit

'=============================================================================
' Environment Probe - Tests whether EDR/compat risk patterns actually work
' in the target environment. Single-file, disposable.
'
' Usage:
'   1. Open a blank .xlsm
'   2. Alt+F11 > File > Import > this file
'   3. Alt+F8 > Probe_Run
'   4. Choose Basic or Basic+Extended
'   5. Results written to probe_result_<PC>_<time>.txt
'=============================================================================

Private Type ProbeResult
    TestNo As Long
    Level As String       ' Basic / Extended / Aux
    Category As String    ' EDR / Compat / Reference / SystemInfo
    PatternName As String
    Target As String
    Result As String      ' OK / FAIL / SKIP
    ErrNum As Long
    ErrMsg As String
    Detail As String
End Type

Private m_results() As ProbeResult
Private m_count As Long
Private m_runExtended As Boolean

'=============================================================================
' Result helpers (consistent output format per iris feedback)
'=============================================================================

Private Sub AddOk(level As String, cat As String, pattern As String, target As String, Optional detail As String = "")
    AddResult level, cat, pattern, target, "OK", 0, "", detail
End Sub

Private Sub AddFail(level As String, cat As String, pattern As String, target As String, errNum As Long, errMsg As String, Optional detail As String = "")
    AddResult level, cat, pattern, target, "FAIL", errNum, errMsg, detail
End Sub

Private Sub AddSkip(level As String, cat As String, pattern As String, target As String, Optional detail As String = "")
    AddResult level, cat, pattern, target, "SKIP", 0, "", detail
End Sub

Private Sub AddResult(level As String, cat As String, pattern As String, target As String, result As String, errNum As Long, errMsg As String, detail As String)
    m_count = m_count + 1
    If m_count > UBound(m_results) Then ReDim Preserve m_results(1 To m_count + 10)
    With m_results(m_count)
        .TestNo = m_count
        .Level = level
        .Category = cat
        .PatternName = pattern
        .Target = target
        .Result = result
        .ErrNum = errNum
        .ErrMsg = errMsg
        .Detail = detail
    End With
End Sub

'=============================================================================
' Entry point
'=============================================================================

Public Sub Probe_Run()
    Dim mode As VbMsgBoxResult
    mode = MsgBox("Run Extended tests too?" & vbCrLf & vbCrLf & _
                  "Basic: COM, File I/O, Registry, Environ, etc." & vbCrLf & _
                  "Extended: Win32 API, Shell, PowerShell, DDE, IE, WMI" & vbCrLf & vbCrLf & _
                  "Yes = Basic + Extended" & vbCrLf & _
                  "No = Basic only" & vbCrLf & _
                  "Cancel = Abort", _
                  vbYesNoCancel + vbQuestion, "Environment Probe")

    If mode = vbCancel Then Exit Sub
    m_runExtended = (mode = vbYes)

    ReDim m_results(1 To 40)
    m_count = 0

    ' System info
    RunSystemInfo

    ' Basic tests
    RunBasicTests

    ' Extended tests
    If m_runExtended Then RunExtendedTests

    ' Reference tests (auxiliary)
    RunReferenceTests

    ' Output
    WriteResults

    ' Summary
    Dim okCount As Long, failCount As Long, skipCount As Long
    Dim i As Long
    For i = 1 To m_count
        Select Case m_results(i).Result
            Case "OK": okCount = okCount + 1
            Case "FAIL": failCount = failCount + 1
            Case "SKIP": skipCount = skipCount + 1
        End Select
    Next i

    MsgBox "Probe complete." & vbCrLf & vbCrLf & _
           "OK: " & okCount & vbCrLf & _
           "FAIL: " & failCount & vbCrLf & _
           "SKIP: " & skipCount, _
           vbInformation, "Results"
End Sub

'=============================================================================
' System info (auxiliary)
'=============================================================================

Private Sub RunSystemInfo()
    AddOk "Aux", "SystemInfo", "Office Version", "Application.Version", Application.Version

    #If Win64 Then
        AddOk "Aux", "SystemInfo", "Office Bitness", "#If Win64", "64-bit"
    #Else
        AddOk "Aux", "SystemInfo", "Office Bitness", "#If Win64", "32-bit"
    #End If

    #If VBA7 Then
        AddOk "Aux", "SystemInfo", "VBA Version", "#If VBA7", "VBA7"
    #Else
        AddOk "Aux", "SystemInfo", "VBA Version", "#If VBA7", "VBA6"
    #End If
End Sub

'=============================================================================
' Basic tests
'=============================================================================

Private Sub RunBasicTests()
    TestCOM "Scripting.FileSystemObject", "COM / CreateObject"
    TestCOM "Scripting.Dictionary", "COM / CreateObject"
    TestCOM "ADODB.Connection", "COM / CreateObject"
    TestCOM "ADODB.Recordset", "COM / CreateObject"
    TestCOM "MSXML2.XMLHTTP.6.0", "COM / CreateObject"
    TestCOM "WinHttp.WinHttpRequest.5.1", "COM / CreateObject"
    TestFileIO
    TestFSO
    TestRegistry
    TestEnviron
    TestClipboard
    TestVarPtr
    TestDAO
    TestLegacyControls "MSComDlg.CommonDialog"
    TestLegacyControls "MSCAL.Calendar"
End Sub

Private Sub TestCOM(progId As String, pattern As String)
    On Error Resume Next
    Dim obj As Object
    Set obj = CreateObject(progId)
    If Err.Number <> 0 Then
        AddFail "Basic", "EDR", pattern, progId, Err.Number, Err.Description
    Else
        AddOk "Basic", "EDR", pattern, progId
    End If
    Set obj = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestFileIO()
    Dim tmp As String
    tmp = Environ$("TEMP") & "\probe_test_" & Format(Now, "yyyymmddhhnnss") & ".txt"
    On Error Resume Next

    ' Write
    Dim f As Long: f = FreeFile
    Open tmp For Output As #f
    Print #f, "probe test"
    Close #f
    If Err.Number <> 0 Then
        AddFail "Basic", "EDR", "File I/O", "Open For Output", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    AddOk "Basic", "EDR", "File I/O", "Open For Output"

    ' Delete
    Kill tmp
    If Err.Number <> 0 Then
        AddFail "Basic", "EDR", "File I/O", "Kill", Err.Number, Err.Description
    Else
        AddOk "Basic", "EDR", "File I/O", "Kill"
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestFSO()
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        AddFail "Basic", "EDR", "FileSystemObject", "FileExists", Err.Number, Err.Description
    Else
        Dim exists As Boolean
        exists = fso.FileExists(Environ$("TEMP") & "\nonexistent_probe_file.xyz")
        AddOk "Basic", "EDR", "FileSystemObject", "FileExists", "Returned: " & exists
    End If
    Set fso = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestRegistry()
    On Error Resume Next
    Dim val As String
    val = GetSetting("ProbeTest", "TestSection", "TestKey", "default_value")
    If Err.Number <> 0 Then
        AddFail "Basic", "EDR", "Registry", "GetSetting", Err.Number, Err.Description
    Else
        AddOk "Basic", "EDR", "Registry", "GetSetting", "Returned: " & val
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestEnviron()
    On Error Resume Next
    Dim val As String
    val = Environ$("USERNAME")
    If Err.Number <> 0 Then
        AddFail "Basic", "EDR", "Environment", "Environ$(USERNAME)", Err.Number, Err.Description
    Else
        AddOk "Basic", "EDR", "Environment", "Environ$(USERNAME)", val
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestClipboard()
    ' MSForms.DataObject via GUID (COM class for clipboard text access)
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If Err.Number <> 0 Then
        AddFail "Basic", "EDR", "Clipboard", "MSForms.DataObject", Err.Number, Err.Description
    Else
        d.SetText "probe_test"
        d.PutInClipboard
        If Err.Number <> 0 Then
            AddFail "Basic", "EDR", "Clipboard", "PutInClipboard", Err.Number, Err.Description
        Else
            AddOk "Basic", "EDR", "Clipboard", "MSForms.DataObject"
        End If
    End If
    Set d = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestVarPtr()
    On Error Resume Next
    #If VBA7 Then
        Dim p As LongPtr
        p = VarPtr(p)
        If Err.Number <> 0 Then
            AddFail "Basic", "Compat", "64-bit: VarPtr/ObjPtr/StrPtr", "VarPtr (LongPtr)", Err.Number, Err.Description
        Else
            AddOk "Basic", "Compat", "64-bit: VarPtr/ObjPtr/StrPtr", "VarPtr (LongPtr)", "64-bit safe"
        End If
    #Else
        Dim p As Long
        p = VarPtr(p)
        If Err.Number <> 0 Then
            AddFail "Basic", "Compat", "64-bit: VarPtr/ObjPtr/StrPtr", "VarPtr (Long)", Err.Number, Err.Description
        Else
            AddOk "Basic", "Compat", "64-bit: VarPtr/ObjPtr/StrPtr", "VarPtr (Long)", "32-bit only"
        End If
    #End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestDAO()
    On Error Resume Next
    Dim db As Object
    Set db = CreateObject("DAO.DBEngine.36")
    If Err.Number <> 0 Then
        AddFail "Basic", "Compat", "Deprecated: DAO", "DAO.DBEngine.36", Err.Number, Err.Description
    Else
        AddOk "Basic", "Compat", "Deprecated: DAO", "DAO.DBEngine.36"
    End If
    Set db = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestLegacyControls(progId As String)
    On Error Resume Next
    Dim obj As Object
    Set obj = CreateObject(progId)
    If Err.Number <> 0 Then
        AddFail "Basic", "Compat", "Deprecated: Legacy Controls", progId, Err.Number, Err.Description
    Else
        AddOk "Basic", "Compat", "Deprecated: Legacy Controls", progId
    End If
    Set obj = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

'=============================================================================
' Extended tests (default disabled, requires user opt-in)
'=============================================================================

Private Sub RunExtendedTests()
    TestWin32API
    TestLoadLibrary
    TestGetObjectWMI
    TestShell
    TestPowerShell
    TestSendKeysProbe
    TestDDE
    TestIE
End Sub

Private Sub TestWin32API()
    ' Tests whether Declare PtrSafe Function compiles and runs
    ' Sleep is declared at module level via conditional compilation
    On Error Resume Next
    #If VBA7 Then
        ' Sleep is available if kernel32 Declare works
        Dim startT As Single: startT = Timer
        Sleep 50
        If Err.Number <> 0 Then
            AddFail "Extended", "EDR", "Win32 API (Declare)", "Sleep Lib kernel32", Err.Number, Err.Description
        Else
            Dim elapsed As Single: elapsed = (Timer - startT) * 1000
            AddOk "Extended", "EDR", "Win32 API (Declare)", "Sleep Lib kernel32", "Elapsed: " & Format(elapsed, "0") & "ms"
        End If
    #Else
        AddSkip "Extended", "EDR", "Win32 API (Declare)", "Sleep Lib kernel32", "Not VBA7"
    #End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestLoadLibrary()
    On Error Resume Next
    #If VBA7 Then
        Dim h As LongPtr
        h = LoadLibraryA("kernel32.dll")
        If Err.Number <> 0 Then
            AddFail "Extended", "EDR", "DLL loading", "LoadLibrary kernel32", Err.Number, Err.Description
        ElseIf h = 0 Then
            AddFail "Extended", "EDR", "DLL loading", "LoadLibrary kernel32", 0, "Returned null handle"
        Else
            AddOk "Extended", "EDR", "DLL loading", "LoadLibrary kernel32", "Handle: " & h
            FreeLibraryPtr h
        End If
    #Else
        AddSkip "Extended", "EDR", "DLL loading", "LoadLibrary kernel32", "Not VBA7"
    #End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestGetObjectWMI()
    On Error Resume Next
    Dim wmi As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number <> 0 Then
        AddFail "Extended", "EDR", "COM / GetObject", "winmgmts", Err.Number, Err.Description
    Else
        Dim rs As Object
        Set rs = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE ProcessId = " & Application.Hwnd)
        If Err.Number <> 0 Then
            AddFail "Extended", "EDR", "Process / WMI", "ExecQuery", Err.Number, Err.Description
        Else
            AddOk "Extended", "EDR", "COM / GetObject", "winmgmts"
            AddOk "Extended", "EDR", "Process / WMI", "ExecQuery"
        End If
    End If
    Set rs = Nothing
    Set wmi = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestShell()
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        AddFail "Extended", "EDR", "Shell / process", "WScript.Shell", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    Dim exitCode As Long
    exitCode = wsh.Run("cmd /c echo probe_test", 0, True)
    If Err.Number <> 0 Then
        AddFail "Extended", "EDR", "Shell / process", "cmd /c echo", Err.Number, Err.Description
    Else
        AddOk "Extended", "EDR", "Shell / process", "cmd /c echo", "ExitCode: " & exitCode
    End If
    Set wsh = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestPowerShell()
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        AddFail "Extended", "EDR", "PowerShell / WScript", "WScript.Shell", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    Dim exitCode As Long
    exitCode = wsh.Run("powershell -Command exit", 0, True)
    If Err.Number <> 0 Then
        AddFail "Extended", "EDR", "PowerShell / WScript", "powershell -Command exit", Err.Number, Err.Description
    Else
        AddOk "Extended", "EDR", "PowerShell / WScript", "powershell -Command exit", "ExitCode: " & exitCode
    End If
    Set wsh = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestSendKeysProbe()
    ' Call-check only. Empty string = no side effect.
    ' Limited value as an EDR detection test.
    On Error Resume Next
    SendKeys ""
    If Err.Number <> 0 Then
        AddFail "Extended", "EDR", "SendKeys", "SendKeys (empty)", Err.Number, Err.Description, "Call-check only"
    Else
        AddOk "Extended", "EDR", "SendKeys", "SendKeys (empty)", "Call-check only"
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestDDE()
    On Error Resume Next
    Dim ch As Long
    ch = DDEInitiate("Excel", "Sheet1")
    If Err.Number <> 0 Then
        AddFail "Extended", "Compat", "Deprecated: DDE", "DDEInitiate", Err.Number, Err.Description
    Else
        DDETerminate ch
        AddOk "Extended", "Compat", "Deprecated: DDE", "DDEInitiate"
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestIE()
    On Error Resume Next
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    If Err.Number <> 0 Then
        AddFail "Extended", "Compat", "Deprecated: IE Automation", "InternetExplorer.Application", Err.Number, Err.Description
    Else
        ie.Quit
        AddOk "Extended", "Compat", "Deprecated: IE Automation", "InternetExplorer.Application"
    End If
    Set ie = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

'=============================================================================
' Reference tests (auxiliary, SKIP if VBIDE access denied)
'=============================================================================

Private Sub RunReferenceTests()
    On Error Resume Next
    Dim refs As Object
    Set refs = ThisWorkbook.VBProject.References
    If Err.Number <> 0 Then
        AddSkip "Aux", "Reference", "Reference List", "VBProject.References", "VBIDE access denied (Trust Access not enabled)"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    ' List all references
    Dim refList As String
    Dim ref As Object
    Dim missingCount As Long
    For Each ref In refs
        refList = refList & ref.Name & " (" & ref.Major & "." & ref.Minor & ")"
        If ref.IsBroken Then
            refList = refList & " [MISSING]"
            missingCount = missingCount + 1
        End If
        refList = refList & "; "
    Next ref

    AddOk "Aux", "Reference", "Reference List", "VBProject.References", refList

    If missingCount > 0 Then
        AddFail "Aux", "Reference", "Missing References", "IsBroken", missingCount, missingCount & " missing reference(s)"
    Else
        AddOk "Aux", "Reference", "Missing References", "IsBroken", "None"
    End If

    Err.Clear
    On Error GoTo 0
End Sub

'=============================================================================
' Win32 API declarations (conditional, used by Extended tests)
'=============================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function FreeLibraryPtr Lib "kernel32" Alias "FreeLibrary" (ByVal hLibModule As LongPtr) As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
    Private Declare Function FreeLibraryPtr Lib "kernel32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
#End If

'=============================================================================
' Output - write results to text file
'=============================================================================

Private Sub WriteResults()
    Dim outPath As String
    outPath = ThisWorkbook.Path & "\probe_result_" & Environ$("COMPUTERNAME") & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"

    Dim f As Long: f = FreeFile
    Open outPath For Output As #f

    ' Header
    Print #f, "# Environment Probe Results"
    Print #f, "# Date: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, "# Computer: " & Environ$("COMPUTERNAME")
    Print #f, "# User: " & Environ$("USERNAME")
    Print #f, "# Mode: " & IIf(m_runExtended, "Basic + Extended", "Basic only")
    Print #f, ""
    Print #f, "TestNo" & vbTab & "Level" & vbTab & "Category" & vbTab & "PatternName" & vbTab & "Target" & vbTab & "Result" & vbTab & "ErrNum" & vbTab & "ErrMsg" & vbTab & "Detail"

    Dim i As Long
    For i = 1 To m_count
        With m_results(i)
            Dim errNumStr As String
            errNumStr = IIf(.ErrNum = 0, "", CStr(.ErrNum))
            Print #f, .TestNo & vbTab & .Level & vbTab & .Category & vbTab & .PatternName & vbTab & .Target & vbTab & .Result & vbTab & errNumStr & vbTab & .ErrMsg & vbTab & .Detail
        End With
    Next i

    Close #f

    MsgBox "Results saved to:" & vbCrLf & outPath, vbInformation, "Output"
End Sub
