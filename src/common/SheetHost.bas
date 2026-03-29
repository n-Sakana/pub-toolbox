Attribute VB_Name = "SheetHost"
Option Explicit

Public Function EnsureWorksheet(ByVal wb As Workbook, ByVal SheetName As String, Optional ByVal ClearFirst As Boolean = False) As Worksheet
    Dim ws As Worksheet
    Dim shp As Shape

    On Error Resume Next
    Set ws = wb.Worksheets(SheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SheetName
    ElseIf ClearFirst Then
        ws.Cells.Clear
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    End If

    Set EnsureWorksheet = ws
End Function

Public Sub AddActionButton(ByVal ws As Worksheet, ByVal Name As String, ByVal Caption As String, ByVal LeftPos As Double, ByVal TopPos As Double, ByVal MacroName As String)
    Dim shp As Shape

    On Error Resume Next
    ws.Shapes(Name).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(1, LeftPos, TopPos, 140, 28)
    shp.Name = Name
    shp.TextFrame.Characters.Text = Caption
    shp.OnAction = MacroName
End Sub

Public Function EscapeSheetName(ByVal SheetName As String) As String
    EscapeSheetName = Replace(SheetName, "'", "''")
End Function

Public Function NormalizeHeaderName(ByVal Value As String, ByVal Index As Long) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    s = LCase$(Trim$(Value))
    If Len(s) = 0 Then s = "column_" & CStr(Index)

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i

    Do While InStr(result, "__") > 0
        result = Replace(result, "__", "_")
    Loop
    If Left$(result, 1) = "_" Then result = Mid$(result, 2)
    If Right$(result, 1) = "_" Then result = Left$(result, Len(result) - 1)
    If Len(result) = 0 Then result = "column_" & CStr(Index)

    NormalizeHeaderName = "src_" & result
End Function

