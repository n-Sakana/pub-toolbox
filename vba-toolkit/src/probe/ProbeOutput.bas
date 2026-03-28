Attribute VB_Name = "ProbeOutput"
Option Explicit

'=============================================================================
' ProbeOutput - CSV and sheet output for the Environment Probe
'=============================================================================

' ---------------------------------------------------------------------------
' CSV output (BOM UTF-8)
' ---------------------------------------------------------------------------
Public Sub WriteCSV()

    Dim filePath As String
    Dim fn As Integer
    Dim i As Long
    Dim computerName As String

    On Error Resume Next
    computerName = Environ$("COMPUTERNAME")
    If Len(computerName) = 0 Then computerName = "UNKNOWN"
    Err.Clear
    On Error GoTo 0

    filePath = g_OutputFolder & "probe_result_" & computerName & _
               "_" & Format$(Now, "yyyymmdd_hhnnss") & ".csv"

    fn = FreeFile

    On Error Resume Next
    Open filePath For Binary As #fn
    If Err.Number <> 0 Then
        MsgBox "Cannot create CSV: " & filePath & vbCrLf & Err.Description, _
               vbExclamation, "Probe"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Write BOM
    Dim bom(0 To 2) As Byte
    bom(0) = &HEF: bom(1) = &HBB: bom(2) = &HBF
    Put #fn, , bom

    ' Header
    Dim line As String
    line = "TestNo,Level,Category,PatternName,Target,Result," & _
           "ErrorNumber,ErrorMessage,Detail" & vbCrLf
    PutUTF8 fn, line

    ' Data rows
    For i = 1 To g_Count
        With g_Results(i)
            line = CStr(.TestNo) & "," & _
                   EscCSV(.Level) & "," & _
                   EscCSV(.Category) & "," & _
                   EscCSV(.PatternName) & "," & _
                   EscCSV(.Target) & "," & _
                   .Result & "," & _
                   IIf(.ErrorNumber <> 0, CStr(.ErrorNumber), "") & "," & _
                   EscCSV(.ErrorMessage) & "," & _
                   EscCSV(.Detail) & vbCrLf
            PutUTF8 fn, line
        End With
    Next i

    Close #fn

End Sub

' ---------------------------------------------------------------------------
' Write a VBA string (UTF-16) as UTF-8 bytes to an open binary file
' This is a simplified conversion that handles ASCII and common characters.
' For full Unicode support, use ADODB.Stream instead.
' ---------------------------------------------------------------------------
Private Sub PutUTF8(ByVal fileNum As Integer, ByVal text As String)
    ' Use ADODB.Stream for proper UTF-8 conversion
    Dim stm As Object
    Dim buf() As Byte

    On Error Resume Next
    Set stm = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        ' Fallback: write as ANSI
        Err.Clear
        Dim ansi() As Byte
        ansi = StrConv(text, vbFromUnicode)
        Put #fileNum, , ansi
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    stm.Type = 2  ' adTypeText
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText text

    ' Read back as binary (skip the BOM that ADODB adds)
    stm.Position = 0
    stm.Type = 1  ' adTypeBinary
    ' Skip BOM (3 bytes)
    stm.Position = 3
    buf = stm.Read
    stm.Close
    Set stm = Nothing

    Put #fileNum, , buf
End Sub

' ---------------------------------------------------------------------------
' Escape a value for CSV (RFC 4180)
' ---------------------------------------------------------------------------
Private Function EscCSV(ByVal val As String) As String
    If InStr(val, ",") > 0 Or InStr(val, """") > 0 Or _
       InStr(val, vbCr) > 0 Or InStr(val, vbLf) > 0 Then
        EscCSV = """" & Replace(val, """", """""") & """"
    Else
        EscCSV = val
    End If
End Function

' ---------------------------------------------------------------------------
' Sheet output - write to _probe_result with color formatting
' ---------------------------------------------------------------------------
Public Sub WriteSheet()

    Dim ws As Worksheet
    Dim i As Long
    Dim r As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_probe_result")
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create the sheet if it does not exist
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "_probe_result"
    End If

    ' Clear existing content
    ws.Cells.Clear

    ' Header row
    r = 1
    ws.Cells(r, 1).Value = "TestNo"
    ws.Cells(r, 2).Value = "Level"
    ws.Cells(r, 3).Value = "Category"
    ws.Cells(r, 4).Value = "PatternName"
    ws.Cells(r, 5).Value = "Target"
    ws.Cells(r, 6).Value = "Result"
    ws.Cells(r, 7).Value = "ErrorNumber"
    ws.Cells(r, 8).Value = "ErrorMessage"
    ws.Cells(r, 9).Value = "Detail"

    ' Bold header
    ws.Range("A1:I1").Font.Bold = True

    ' Data rows
    For i = 1 To g_Count
        r = i + 1
        With g_Results(i)
            ws.Cells(r, 1).Value = .TestNo
            ws.Cells(r, 2).Value = .Level
            ws.Cells(r, 3).Value = .Category
            ws.Cells(r, 4).Value = .PatternName
            ws.Cells(r, 5).Value = .Target
            ws.Cells(r, 6).Value = .Result
            If .ErrorNumber <> 0 Then
                ws.Cells(r, 7).Value = .ErrorNumber
            End If
            ws.Cells(r, 8).Value = .ErrorMessage
            ws.Cells(r, 9).Value = .Detail
        End With

        ' Color formatting: OK=green, FAIL=red, SKIP=gray
        Select Case g_Results(i).Result
            Case "OK"
                ws.Range("A" & r & ":I" & r).Interior.Color = RGB(198, 239, 206)  ' light green
                ws.Cells(r, 6).Font.Color = RGB(0, 128, 0)
            Case "FAIL"
                ws.Range("A" & r & ":I" & r).Interior.Color = RGB(255, 199, 206)  ' light red
                ws.Cells(r, 6).Font.Color = RGB(192, 0, 0)
            Case "SKIP"
                ws.Range("A" & r & ":I" & r).Interior.Color = RGB(217, 217, 217)  ' light gray
                ws.Cells(r, 6).Font.Color = RGB(128, 128, 128)
        End Select
    Next i

    ' Auto-fit columns
    ws.Columns("A:I").AutoFit

End Sub
