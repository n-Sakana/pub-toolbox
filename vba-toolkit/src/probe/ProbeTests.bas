Attribute VB_Name = "ProbeTests"
Option Explicit

'=============================================================================
' ProbeTests - All test implementations for the Environment Probe
'=============================================================================

' ---------------------------------------------------------------------------
' Win32 API declarations (used by Extended tests #17, #18)
' These Declare statements are always present but only called when Extended
' tests are enabled.
' ---------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function LoadLibraryA Lib "kernel32" _
        (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32" _
        (ByVal hLibModule As LongPtr) As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function LoadLibraryA Lib "kernel32" _
        (ByVal lpLibFileName As String) As Long
    Private Declare Function FreeLibrary Lib "kernel32" _
        (ByVal hLibModule As Long) As Long
#End If

' ===========================================================================
' Environment info (#28-30) - always run first
' ===========================================================================
Public Sub RunEnvTests()

    ' --- #28 Office version -------------------------------------------------
    Dim ver As String
    On Error Resume Next
    ver = Application.Version
    If Err.Number = 0 Then
        AddResult 28, "Aux", "SystemInfo", "Office Version", _
                  "Application.Version", "OK", Detail:=ver
    Else
        AddResult 28, "Aux", "SystemInfo", "Office Version", _
                  "Application.Version", "FAIL", Err.Number, Err.Description
    End If
    Err.Clear
    On Error GoTo 0

    ' --- #29 Office bitness -------------------------------------------------
    Dim bitness As String
    #If Win64 Then
        bitness = "64-bit"
    #Else
        bitness = "32-bit"
    #End If
    AddResult 29, "Aux", "SystemInfo", "Office Bitness", _
              "#If Win64", "OK", Detail:=bitness

    ' --- #30 VBA version ----------------------------------------------------
    Dim vbaVer As String
    #If VBA7 Then
        vbaVer = "VBA7"
    #Else
        vbaVer = "VBA6"
    #End If
    AddResult 30, "Aux", "SystemInfo", "VBA Version", _
              "#If VBA7", "OK", Detail:=vbaVer

End Sub

' ===========================================================================
' Basic tests (#1-16)
' ===========================================================================
Public Sub RunBasicTests()

    ' --- #1 CreateObject("Scripting.FileSystemObject") ----------------------
    TestCreateObject 1, "Basic", "EDR", "COM / CreateObject", _
                     "Scripting.FileSystemObject"

    ' --- #2 CreateObject("Scripting.Dictionary") ----------------------------
    TestCreateObject 2, "Basic", "EDR", "COM / CreateObject", _
                     "Scripting.Dictionary"

    ' --- #3 CreateObject("ADODB.Connection") --------------------------------
    TestCreateObject 3, "Basic", "EDR", "COM / CreateObject", _
                     "ADODB.Connection"

    ' --- #4 CreateObject("ADODB.Recordset") ---------------------------------
    TestCreateObject 4, "Basic", "EDR", "COM / CreateObject", _
                     "ADODB.Recordset"

    ' --- #5 CreateObject("MSXML2.XMLHTTP.6.0") -----------------------------
    TestCreateObject 5, "Basic", "EDR", "COM / CreateObject", _
                     "MSXML2.XMLHTTP.6.0"

    ' --- #6 CreateObject("WinHttp.WinHttpRequest.5.1") ---------------------
    TestCreateObject 6, "Basic", "EDR", "COM / CreateObject", _
                     "WinHttp.WinHttpRequest.5.1"

    ' --- #7 File I/O -------------------------------------------------------
    TestFileIO

    ' --- #8 FileSystemObject.FileExists ------------------------------------
    TestFSOFileExists

    ' --- #9 Registry -------------------------------------------------------
    TestRegistry

    ' --- #10 Environ -------------------------------------------------------
    TestEnviron

    ' --- #11 Clipboard (MSForms.DataObject) --------------------------------
    TestClipboard

    ' --- #12 VarPtr (64-bit) -----------------------------------------------
    TestVarPtr

    ' --- #13 DAO.DBEngine.36 -----------------------------------------------
    TestCreateObject 13, "Basic", "Compat", "Deprecated: DAO", _
                     "DAO.DBEngine.36"

    ' --- #14 MSComDlg.CommonDialog -----------------------------------------
    TestCreateObject 14, "Basic", "Compat", "Deprecated: Legacy Controls", _
                     "MSComDlg.CommonDialog"

    ' --- #15 MSCAL.Calendar ------------------------------------------------
    TestCreateObject 15, "Basic", "Compat", "Deprecated: Legacy Controls", _
                     "MSCAL.Calendar"

    ' --- #16 HTTP test -----------------------------------------------------
    TestHTTP

End Sub

' ===========================================================================
' Extended tests (#17-25) - only called when RunExtended = TRUE
' ===========================================================================
Public Sub RunExtendedTests()

    ' --- #17 Win32 API Declare (Sleep) -------------------------------------
    TestWin32Sleep

    ' --- #18 LoadLibrary ---------------------------------------------------
    TestLoadLibrary

    ' --- #19 GetObject WMI -------------------------------------------------
    TestGetObjectWMI

    ' --- #20 Shell via WScript.Shell.Run -----------------------------------
    TestShellCmd

    ' --- #21 PowerShell via WScript.Shell.Run ------------------------------
    TestPowerShell

    ' --- #22 WMI ExecQuery -------------------------------------------------
    TestWMIExecQuery

    ' --- #23 SendKeys (call check only) ------------------------------------
    TestSendKeys

    ' --- #24 DDE -----------------------------------------------------------
    TestDDE

    ' --- #25 IE Automation -------------------------------------------------
    TestCreateObject 25, "Extended", "Compat", "Deprecated: IE Automation", _
                     "InternetExplorer.Application"

End Sub

' ===========================================================================
' Reference tests (#26-27) - auxiliary, skip if access denied
' ===========================================================================
Public Sub RunReferenceTests()

    Dim refAccess As Boolean
    Dim refList As String
    Dim hasMissing As Boolean
    Dim ref As Object  ' VBProject.Reference

    On Error Resume Next

    ' Try to access VBProject.References
    Dim refs As Object
    Set refs = ThisWorkbook.VBProject.References
    If Err.Number <> 0 Then
        ' Access denied - SKIP both tests
        AddResult 26, "Aux", "Reference", "VBProject.References", _
                  "VBProject.References enumeration", "SKIP", _
                  Err.Number, Err.Description, _
                  "VBA project access not trusted"
        Err.Clear
        AddResult 27, "Aux", "Reference", "Missing References", _
                  "IsBroken check", "SKIP", _
                  Detail:="VBA project access not trusted"
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear

    ' --- #26 Enumerate references ------------------------------------------
    refList = ""
    hasMissing = False
    For Each ref In refs
        Err.Clear
        Dim refName As String
        refName = ref.Name
        If Err.Number = 0 Then
            If Len(refList) > 0 Then refList = refList & "; "
            refList = refList & refName
        End If
        ' Check IsBroken for #27
        If ref.IsBroken Then hasMissing = True
    Next ref
    Err.Clear

    AddResult 26, "Aux", "Reference", "VBProject.References", _
              "VBProject.References enumeration", "OK", _
              Detail:=refList

    ' --- #27 Missing references --------------------------------------------
    If hasMissing Then
        AddResult 27, "Aux", "Reference", "Missing References", _
                  "IsBroken check", "FAIL", _
                  Detail:="Missing references detected"
    Else
        AddResult 27, "Aux", "Reference", "Missing References", _
                  "IsBroken check", "OK", _
                  Detail:="No missing references"
    End If

    On Error GoTo 0

End Sub

' ===========================================================================
' Helper: generic CreateObject test
' ===========================================================================
Private Sub TestCreateObject(ByVal TestNo As Long, _
                             ByVal Level As String, _
                             ByVal Category As String, _
                             ByVal PatternName As String, _
                             ByVal ProgID As String)
    Dim obj As Object
    On Error Resume Next
    Set obj = CreateObject(ProgID)
    If Err.Number = 0 Then
        AddResult TestNo, Level, Category, PatternName, _
                  "CreateObject(""" & ProgID & """)", "OK"
    Else
        AddResult TestNo, Level, Category, PatternName, _
                  "CreateObject(""" & ProgID & """)", "FAIL", _
                  Err.Number, Err.Description
    End If
    Err.Clear
    Set obj = Nothing
    On Error GoTo 0
End Sub

' ===========================================================================
' #7 File I/O test
' ===========================================================================
Private Sub TestFileIO()
    Dim filePath As String
    Dim fn As Integer

    filePath = g_OutputFolder & g_DummyFileName

    On Error Resume Next

    fn = FreeFile
    Open filePath For Output As #fn
    If Err.Number <> 0 Then
        AddResult 7, "Basic", "EDR", "File I/O", _
                  "Open/Write/Close/Kill", "FAIL", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    Print #fn, "probe test"
    Close #fn

    ' Clean up
    Kill filePath
    Err.Clear
    On Error GoTo 0

    AddResult 7, "Basic", "EDR", "File I/O", _
              "Open/Write/Close/Kill", "OK"
End Sub

' ===========================================================================
' #8 FSO FileExists
' ===========================================================================
Private Sub TestFSOFileExists()
    Dim fso As Object
    Dim result As Boolean

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        AddResult 8, "Basic", "EDR", "FileSystemObject", _
                  "FSO.FileExists", "FAIL", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    result = fso.FileExists(ThisWorkbook.FullName)
    If Err.Number = 0 Then
        AddResult 8, "Basic", "EDR", "FileSystemObject", _
                  "FSO.FileExists", "OK", Detail:=CStr(result)
    Else
        AddResult 8, "Basic", "EDR", "FileSystemObject", _
                  "FSO.FileExists", "FAIL", Err.Number, Err.Description
    End If
    Err.Clear
    Set fso = Nothing
    On Error GoTo 0
End Sub

' ===========================================================================
' #9 Registry (GetSetting / SaveSetting / DeleteSetting)
' ===========================================================================
Private Sub TestRegistry()
    Dim val As String

    On Error Resume Next

    ' Write, read, delete
    SaveSetting "ProbeTest", "Test", "Key", "ProbeValue"
    If Err.Number <> 0 Then
        AddResult 9, "Basic", "EDR", "Registry", _
                  "GetSetting/SaveSetting", "FAIL", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    val = GetSetting("ProbeTest", "Test", "Key", "")

    ' Clean up
    DeleteSetting "ProbeTest", "Test"
    Err.Clear
    On Error GoTo 0

    AddResult 9, "Basic", "EDR", "Registry", _
              "GetSetting/SaveSetting", "OK", Detail:=val
End Sub

' ===========================================================================
' #10 Environ
' ===========================================================================
Private Sub TestEnviron()
    Dim val As String

    On Error Resume Next
    val = Environ$("USERNAME")
    If Err.Number = 0 And Len(val) > 0 Then
        AddResult 10, "Basic", "EDR", "Environment", _
                  "Environ$(""USERNAME"")", "OK", Detail:=val
    Else
        AddResult 10, "Basic", "EDR", "Environment", _
                  "Environ$(""USERNAME"")", "FAIL", Err.Number, Err.Description
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' ===========================================================================
' #11 Clipboard (MSForms.DataObject)
' ===========================================================================
Private Sub TestClipboard()
    Dim d As Object  ' MSForms.DataObject

    On Error Resume Next
    Set d = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If Err.Number <> 0 Then
        Err.Clear
        ' Alternative: try via MSForms UserForm reference
        ' This may also fail if FM20.DLL is not available
        AddResult 11, "Basic", "EDR", "Clipboard", _
                  "MSForms.DataObject", "FAIL", _
                  Detail:="Cannot create MSForms.DataObject"
        On Error GoTo 0
        Exit Sub
    End If

    d.SetText "test"
    d.PutInClipboard

    If Err.Number = 0 Then
        AddResult 11, "Basic", "EDR", "Clipboard", _
                  "MSForms.DataObject", "OK"
    Else
        AddResult 11, "Basic", "EDR", "Clipboard", _
                  "MSForms.DataObject", "FAIL", Err.Number, Err.Description
    End If
    Err.Clear
    Set d = Nothing
    On Error GoTo 0
End Sub

' ===========================================================================
' #12 VarPtr (64-bit pointer)
' ===========================================================================
Private Sub TestVarPtr()

    On Error Resume Next

    #If VBA7 Then
        Dim p As LongPtr
        p = VarPtr(p)
        If Err.Number = 0 Then
            AddResult 12, "Basic", "EDR", "64-bit: VarPtr", _
                      "VarPtr(LongPtr)", "OK", Detail:=CStr(p)
        Else
            AddResult 12, "Basic", "EDR", "64-bit: VarPtr", _
                      "VarPtr(LongPtr)", "FAIL", Err.Number, Err.Description
        End If
    #Else
        Dim p As Long
        p = VarPtr(p)
        If Err.Number = 0 Then
            AddResult 12, "Basic", "EDR", "64-bit: VarPtr", _
                      "VarPtr(Long)", "OK", Detail:=CStr(p)
        Else
            AddResult 12, "Basic", "EDR", "64-bit: VarPtr", _
                      "VarPtr(Long)", "FAIL", Err.Number, Err.Description
        End If
    #End If

    Err.Clear
    On Error GoTo 0
End Sub

' ===========================================================================
' #16 HTTP test (uses TestURL from settings)
' ===========================================================================
Private Sub TestHTTP()
    If Len(g_TestURL) = 0 Then
        AddResult 16, "Basic", "EDR", "Network / HTTP", _
                  "XMLHTTP GET", "SKIP", Detail:="TestURL is empty"
        Exit Sub
    End If

    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    If Err.Number <> 0 Then
        AddResult 16, "Basic", "EDR", "Network / HTTP", _
                  "XMLHTTP GET", "FAIL", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    http.Open "GET", g_TestURL, False
    http.Send

    If Err.Number = 0 Then
        AddResult 16, "Basic", "EDR", "Network / HTTP", _
                  "XMLHTTP GET " & g_TestURL, "OK", _
                  Detail:="Status=" & http.Status
    Else
        AddResult 16, "Basic", "EDR", "Network / HTTP", _
                  "XMLHTTP GET " & g_TestURL, "FAIL", Err.Number, Err.Description
    End If
    Err.Clear
    Set http = Nothing
    On Error GoTo 0
End Sub

' ===========================================================================
' #17 Win32 API Declare (Sleep)
' ===========================================================================
Private Sub TestWin32Sleep()
    On Error Resume Next
    Sleep 1  ' sleep 1 ms - minimal impact
    If Err.Number = 0 Then
        AddResult 17, "Extended", "EDR", "Win32 API (Declare)", _
                  "Sleep Lib ""kernel32""", "OK"
    Else
        AddResult 17, "Extended", "EDR", "Win32 API (Declare)", _
                  "Sleep Lib ""kernel32""", "FAIL", Err.Number, Err.Description
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' ===========================================================================
' #18 LoadLibrary
' ===========================================================================
Private Sub TestLoadLibrary()
    On Error Resume Next

    #If VBA7 Then
        Dim hLib As LongPtr
    #Else
        Dim hLib As Long
    #End If

    hLib = LoadLibraryA("kernel32.dll")
    If Err.Number = 0 And hLib <> 0 Then
        FreeLibrary hLib
        AddResult 18, "Extended", "EDR", "DLL loading", _
                  "LoadLibrary ""kernel32.dll""", "OK"
    Else
        AddResult 18, "Extended", "EDR", "DLL loading", _
                  "LoadLibrary ""kernel32.dll""", "FAIL", Err.Number, Err.Description
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' ===========================================================================
' #19 GetObject WMI
' ===========================================================================
Private Sub TestGetObjectWMI()
    Dim wmi As Object
    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number = 0 Then
        AddResult 19, "Extended", "EDR", "COM / GetObject", _
                  "GetObject(""winmgmts:\\.\root\cimv2"")", "OK"
    Else
        AddResult 19, "Extended", "EDR", "COM / GetObject", _
                  "GetObject(""winmgmts:\\.\root\cimv2"")", "FAIL", _
                  Err.Number, Err.Description
    End If
    Err.Clear
    Set wmi = Nothing
    On Error GoTo 0
End Sub

' ===========================================================================
' #20 Shell via WScript.Shell.Run (cmd)
' ===========================================================================
Private Sub TestShellCmd()
    Dim wsh As Object
    Dim rc As Long

    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        AddResult 20, "Extended", "EDR", "Shell / process", _
                  "WScript.Shell.Run cmd", "FAIL", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    rc = wsh.Run("cmd /c echo test", 0, True)
    If Err.Number = 0 Then
        AddResult 20, "Extended", "EDR", "Shell / process", _
                  "WScript.Shell.Run ""cmd /c echo test""", "OK", _
                  Detail:="ExitCode=" & rc
    Else
        AddResult 20, "Extended", "EDR", "Shell / process", _
                  "WScript.Shell.Run ""cmd /c echo test""", "FAIL", _
                  Err.Number, Err.Description
    End If
    Err.Clear
    Set wsh = Nothing
    On Error GoTo 0
End Sub

' ===========================================================================
' #21 PowerShell via WScript.Shell.Run
' ===========================================================================
Private Sub TestPowerShell()
    Dim wsh As Object
    Dim rc As Long

    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        AddResult 21, "Extended", "EDR", "PowerShell / WScript", _
                  "WScript.Shell.Run powershell", "FAIL", _
                  Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    rc = wsh.Run("powershell -Command exit", 0, True)
    If Err.Number = 0 Then
        AddResult 21, "Extended", "EDR", "PowerShell / WScript", _
                  "WScript.Shell.Run ""powershell -Command exit""", "OK", _
                  Detail:="ExitCode=" & rc
    Else
        AddResult 21, "Extended", "EDR", "PowerShell / WScript", _
                  "WScript.Shell.Run ""powershell -Command exit""", "FAIL", _
                  Err.Number, Err.Description
    End If
    Err.Clear
    Set wsh = Nothing
    On Error GoTo 0
End Sub

' ===========================================================================
' #22 WMI ExecQuery
' ===========================================================================
Private Sub TestWMIExecQuery()
    Dim wmi As Object
    Dim results As Object

    On Error Resume Next
    Set wmi = GetObject("winmgmts:")
    If Err.Number <> 0 Then
        AddResult 22, "Extended", "EDR", "Process / WMI", _
                  "WMI ExecQuery", "FAIL", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    Set results = wmi.ExecQuery("SELECT Name FROM Win32_OperatingSystem")
    If Err.Number = 0 Then
        Dim osName As String
        Dim item As Object
        For Each item In results
            osName = item.Name
            Exit For
        Next item
        AddResult 22, "Extended", "EDR", "Process / WMI", _
                  "WMI ExecQuery Win32_OperatingSystem", "OK", _
                  Detail:=Left$(osName, 100)
    Else
        AddResult 22, "Extended", "EDR", "Process / WMI", _
                  "WMI ExecQuery", "FAIL", Err.Number, Err.Description
    End If
    Err.Clear
    Set results = Nothing
    Set wmi = Nothing
    On Error GoTo 0
End Sub

' ===========================================================================
' #23 SendKeys (call check only - empty string)
' ===========================================================================
Private Sub TestSendKeys()
    On Error Resume Next
    SendKeys ""
    If Err.Number = 0 Then
        AddResult 23, "Extended", "EDR", "SendKeys", _
                  "SendKeys """"", "OK", _
                  Detail:="Call check only"
    Else
        AddResult 23, "Extended", "EDR", "SendKeys", _
                  "SendKeys """"", "FAIL", Err.Number, Err.Description
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' ===========================================================================
' #24 DDE (DDEInitiate - expected to fail)
' ===========================================================================
Private Sub TestDDE()
    Dim chan As Long

    On Error Resume Next
    chan = Application.DDEInitiate("Excel", "System")
    If Err.Number = 0 Then
        ' Unexpectedly succeeded - close the channel
        Application.DDETerminate chan
        AddResult 24, "Extended", "Compat", "Deprecated: DDE", _
                  "DDEInitiate", "OK", Detail:="DDE available"
    Else
        ' Expected failure - distinguish between "blocked" and "no server"
        AddResult 24, "Extended", "Compat", "Deprecated: DDE", _
                  "DDEInitiate", "FAIL", Err.Number, Err.Description
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' ===========================================================================
' (#25 IE Automation is handled by TestCreateObject in RunExtendedTests)
' ===========================================================================
