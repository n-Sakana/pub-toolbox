Attribute VB_Name = "ProbeMain"
Option Explicit

'=============================================================================
' ProbeMain - Entry point for the Environment Probe macro
'
' Usage: Import all Probe*.bas into a blank .xlsm, create _probe_config
'        and _probe_result sheets, then run Probe_Run from the macro dialog.
'=============================================================================

' ---------------------------------------------------------------------------
' Settings read from _probe_config sheet
' ---------------------------------------------------------------------------
Public g_RunExtended    As Boolean
Public g_TestURL        As String
Public g_OutputFolder   As String
Public g_DummyFileName  As String

' ---------------------------------------------------------------------------
' Result storage
' ---------------------------------------------------------------------------
Public Type ProbeResult
    TestNo       As Long
    Level        As String   ' Basic / Extended / Aux
    Category     As String   ' EDR / Compat / Reference / SystemInfo
    PatternName  As String
    Target       As String
    Result       As String   ' OK / FAIL / SKIP
    ErrorNumber  As Long
    ErrorMessage As String
    Detail       As String
End Type

Public g_Results()  As ProbeResult
Public g_Count      As Long

' ---------------------------------------------------------------------------
' Public: initialise result array
' ---------------------------------------------------------------------------
Public Sub InitResults()
    g_Count = 0
    ReDim g_Results(1 To 30)
End Sub

' ---------------------------------------------------------------------------
' Public: add one result row
' ---------------------------------------------------------------------------
Public Sub AddResult(ByVal TestNo As Long, _
                     ByVal Level As String, _
                     ByVal Category As String, _
                     ByVal PatternName As String, _
                     ByVal Target As String, _
                     ByVal Result As String, _
                     Optional ByVal ErrNo As Long = 0, _
                     Optional ByVal ErrMsg As String = "", _
                     Optional ByVal Detail As String = "")

    g_Count = g_Count + 1
    If g_Count > UBound(g_Results) Then
        ReDim Preserve g_Results(1 To g_Count + 10)
    End If

    With g_Results(g_Count)
        .TestNo = TestNo
        .Level = Level
        .Category = Category
        .PatternName = PatternName
        .Target = Target
        .Result = Result
        .ErrorNumber = ErrNo
        .ErrorMessage = ErrMsg
        .Detail = Detail
    End With
End Sub

' ---------------------------------------------------------------------------
' Main entry point
' ---------------------------------------------------------------------------
Public Sub Probe_Run()

    Dim wsConfig As Worksheet
    Dim okCnt As Long, failCnt As Long, skipCnt As Long
    Dim i As Long

    ' --- Read settings -------------------------------------------------------
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("_probe_config")
    On Error GoTo 0

    If wsConfig Is Nothing Then
        MsgBox "_probe_config sheet not found. Please create it first.", _
               vbCritical, "Probe"
        Exit Sub
    End If

    ReadSettings wsConfig

    ' --- Initialise ----------------------------------------------------------
    InitResults

    ' --- Environment info (#28-30) ------------------------------------------
    RunEnvTests

    ' --- Basic tests (#1-16) ------------------------------------------------
    RunBasicTests

    ' --- Extended tests (#17-25) - only when enabled -------------------------
    If g_RunExtended Then
        RunExtendedTests
    End If

    ' --- Reference tests (#26-27) - auxiliary --------------------------------
    RunReferenceTests

    ' --- Output results ------------------------------------------------------
    WriteCSV
    WriteSheet

    ' --- Summary message -----------------------------------------------------
    okCnt = 0: failCnt = 0: skipCnt = 0
    For i = 1 To g_Count
        Select Case g_Results(i).Result
            Case "OK":   okCnt = okCnt + 1
            Case "FAIL": failCnt = failCnt + 1
            Case "SKIP": skipCnt = skipCnt + 1
        End Select
    Next i

    MsgBox "Probe complete." & vbCrLf & vbCrLf & _
           "OK:   " & okCnt & vbCrLf & _
           "FAIL: " & failCnt & vbCrLf & _
           "SKIP: " & skipCnt, _
           vbInformation, "Probe"

End Sub

' ---------------------------------------------------------------------------
' Private: read _probe_config key/value pairs (A=key, B=value)
' ---------------------------------------------------------------------------
Private Sub ReadSettings(ws As Worksheet)

    Dim r As Long
    Dim key As String

    ' Defaults
    g_RunExtended = False
    g_TestURL = ""
    g_OutputFolder = ""
    g_DummyFileName = "_probe_test.txt"

    For r = 1 To ws.UsedRange.Rows.Count
        key = Trim$(CStr(ws.Cells(r, 1).Value))
        Select Case key
            Case "RunExtended"
                g_RunExtended = (UCase$(Trim$(CStr(ws.Cells(r, 2).Value))) = "TRUE")
            Case "TestURL"
                g_TestURL = Trim$(CStr(ws.Cells(r, 2).Value))
            Case "OutputFolder"
                g_OutputFolder = Trim$(CStr(ws.Cells(r, 2).Value))
            Case "DummyFileName"
                Dim tmp As String
                tmp = Trim$(CStr(ws.Cells(r, 2).Value))
                If Len(tmp) > 0 Then g_DummyFileName = tmp
        End Select
    Next r

    ' Default output folder = same as workbook
    If Len(g_OutputFolder) = 0 Then
        g_OutputFolder = ThisWorkbook.Path
    End If
    ' Ensure trailing backslash
    If Right$(g_OutputFolder, 1) <> "\" Then
        g_OutputFolder = g_OutputFolder & "\"
    End If

End Sub
