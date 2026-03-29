Attribute VB_Name = "OutlookDraftSheet"
Option Explicit

Private Const COL_FROM As String = "from"
Private Const COL_TO As String = "to"
Private Const COL_CC As String = "cc"
Private Const COL_BCC As String = "bcc"
Private Const COL_SUBJECT As String = "subject"
Private Const COL_BODY As String = "body"
Private Const COL_ATTACHMENTS As String = "attachments"
Private Const COL_STATUS As String = "status"
Private Const COL_ERROR As String = "error"

Public Sub CreateDraftTemplate(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim headers As Variant
    Dim headerRow As Long
    Dim endCol As Long
    Dim i As Long
    Dim tableRange As Range
    Dim lo As ListObject

    Set ws = SheetHost.EnsureWorksheet(wb, ToolRegistry.SHEET_OUTLOOK_DRAFT, True)

    headers = Array( _
        COL_FROM, _
        COL_TO, _
        COL_CC, _
        COL_BCC, _
        COL_SUBJECT, _
        COL_BODY, _
        COL_ATTACHMENTS, _
        COL_STATUS, _
        COL_ERROR)

    headerRow = 1
    For i = LBound(headers) To UBound(headers)
        ws.Cells(headerRow, i + 1).Value2 = CStr(headers(i))
    Next i

    endCol = UBound(headers) + 1
    Set tableRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, endCol))
    Set lo = ws.ListObjects.Add(SourceType:=1, Source:=tableRange, XlListObjectHasHeaders:=1)
    lo.Name = ToolRegistry.TABLE_OUTLOOK_DRAFT
    lo.TableStyle = "TableStyleMedium2"

    ws.Columns.AutoFit
    ws.Columns(1).ColumnWidth = 28
    ws.Columns(2).ColumnWidth = 32
    ws.Columns(3).ColumnWidth = 28
    ws.Columns(4).ColumnWidth = 28
    ws.Columns(5).ColumnWidth = 32
    ws.Columns(6).ColumnWidth = 48
    ws.Columns(7).ColumnWidth = 48
    ws.Columns(8).ColumnWidth = 16
    ws.Columns(9).ColumnWidth = 24
    ws.Activate
    ws.Range("A1").Select
End Sub
