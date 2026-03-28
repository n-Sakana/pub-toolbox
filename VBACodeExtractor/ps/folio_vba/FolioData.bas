Attribute VB_Name = "FolioData"
Option Explicit

' ============================================================================
' FE-side cache (populated from hidden sheets written by FolioWorker)
' FE detects changes via Workbook_SheetChange on _folio_signal.
' ============================================================================

Private m_feMailRecords As Object    ' Dict: entry_id -> record Dict
Private m_feMailIndex As Object      ' Dict: normalized_key -> Dict(entry_id -> True)
Private m_feCaseNames As Object      ' Dict: folder_name -> True
Private m_feCaseFiles As Object      ' Dict: case_id -> Dict(file_path -> record Dict)

' ============================================================================
' Table Operations (FE: reads/writes the source Excel file directly)
' ============================================================================

Public Function GetWorkbookTableNames(wb As Workbook) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "GetWorkbookTableNames"
    On Error GoTo ErrHandler
    Set GetWorkbookTableNames = New Collection
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then GoTo NextSheet
        Dim tbl As ListObject
        For Each tbl In ws.ListObjects
            GetWorkbookTableNames.Add tbl.Name
        Next tbl
NextSheet:
    Next ws
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Function FindTable(wb As Workbook, tableName As String) As ListObject
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "FindTable"
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim tbl As ListObject
        For Each tbl In ws.ListObjects
            If LCase$(tbl.Name) = LCase$(tableName) Then
                Set FindTable = tbl: eh.OK: Exit Function
            End If
        Next tbl
    Next ws
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Function ReadTableRecords(tbl As ListObject) As Object
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ReadTableRecords"
    On Error GoTo ErrHandler
    Set ReadTableRecords = FolioLib.NewDict()
    If tbl.DataBodyRange Is Nothing Then eh.OK: Exit Function
    Dim data As Variant: data = tbl.DataBodyRange.value
    Dim nCols As Long: nCols = tbl.ListColumns.count
    Dim colNames() As String: ReDim colNames(1 To nCols)
    Dim c As Long
    For c = 1 To nCols
        colNames(c) = tbl.ListColumns(c).Name
    Next c
    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim rec As Object: Set rec = FolioLib.NewDict()
        rec.Add "_row_index", r
        For c = 1 To nCols
            rec.Add colNames(c), data(r, c)
        Next c
        ReadTableRecords.Add CStr(r), rec
    Next r
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Sub WriteTableCell(tbl As ListObject, rowIndex As Long, colName As String, val As Variant)
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "WriteTableCell"
    On Error GoTo ErrHandler
    Dim col As ListColumn: Set col = tbl.ListColumns(colName)
    tbl.DataBodyRange.Cells(rowIndex, col.Index).value = val
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Function GetTableColumnNames(tbl As ListObject) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "GetTableColumnNames"
    On Error GoTo ErrHandler
    Set GetTableColumnNames = New Collection
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If Not (col.Name Like "_*") Then GetTableColumnNames.Add col.Name
    Next col
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' ============================================================================
' FE: Mail/Case counts 窶・read from FE-side Dictionary cache
' ============================================================================

Public Function GetMailCount() As Long
    GetMailCount = 0
    If Not m_feMailRecords Is Nothing Then GetMailCount = m_feMailRecords.count
End Function

' FE: Find mail records matching keyValue via FE-side Dictionary cache
Public Function FindMailRecords(keyValue As String, matchField As String, matchMode As String) As Object
    Dim result As Object: Set result = FolioLib.NewDict()
    Set FindMailRecords = result
    If Len(keyValue) = 0 Then Exit Function
    If m_feMailIndex Is Nothing Then Exit Function
    If m_feMailRecords Is Nothing Then Exit Function

    ' Build lookup keys (split ";" separated, normalize)
    Dim keyParts() As String: keyParts = Split(keyValue, ";")
    Dim kp As Long
    For kp = 0 To UBound(keyParts)
        Dim normKey As String: normKey = LCase$(Trim$(keyParts(kp)))
        If matchMode = "domain" Then normKey = LCase$(GetDomain(normKey))
        If Len(normKey) = 0 Then GoTo NextKey

        ' O(1) lookup in index
        If m_feMailIndex.Exists(normKey) Then
            Dim inner As Object: Set inner = m_feMailIndex(normKey)
            Dim eids As Variant: eids = inner.keys
            Dim j As Long
            For j = 0 To UBound(eids)
                Dim eid As String: eid = CStr(eids(j))
                If Not result.Exists(eid) And m_feMailRecords.Exists(eid) Then
                    Set result(eid) = m_feMailRecords(eid)
                End If
            Next j
        End If
NextKey:
    Next kp
    Set FindMailRecords = result
End Function

Public Function GetCaseCount() As Long
    GetCaseCount = 0
    If Not m_feCaseNames Is Nothing Then GetCaseCount = m_feCaseNames.count
End Function


Public Sub CreateCaseFolder(rootPath As String, caseId As String, displayName As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "CreateCaseFolder"
    On Error GoTo ErrHandler
    If Len(rootPath) = 0 Or Len(caseId) = 0 Then eh.OK: Exit Sub
    Dim folderName As String
    folderName = FolioLib.SafeName(caseId)
    If Len(displayName) > 0 Then folderName = folderName & "_" & FolioLib.SafeName(displayName)
    FolioLib.EnsureFolder rootPath & "\" & folderName
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Function GetDomain(email As String) As String
    Dim pos As Long: pos = InStr(email, "@")
    If pos > 0 Then GetDomain = Mid$(email, pos + 1) Else GetDomain = email
End Function

' Load from FE's own local sheets (no cross-process)
Public Sub LoadFromLocalSheets(wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet

    Set ws = wb.Worksheets("_folio_mail")
    If Not ws Is Nothing Then LoadMailFromLocalSheet wb

    Set ws = wb.Worksheets("_folio_mail_idx")
    If Not ws Is Nothing Then LoadMailIndexFromLocalSheet wb

    Set ws = wb.Worksheets("_folio_cases")
    If Not ws Is Nothing Then LoadCasesFromLocalSheet wb

    Set ws = wb.Worksheets("_folio_files")
    If Not ws Is Nothing Then LoadCaseFilesFromLocalSheet wb

    On Error GoTo 0
End Sub

Private Sub LoadMailFromLocalSheet(wb As Workbook)
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_folio_mail")
    If ws.Range("A1").value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.value
    If IsEmpty(data) Then Exit Sub
    Dim newRecs As Object: Set newRecs = FolioLib.NewDict()
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim eid As String: eid = CStr(data(i, 1))
        If Len(eid) = 0 Then GoTo NextLMail
        Dim rec As Object: Set rec = FolioLib.NewDict()
        rec.Add "entry_id", eid
        rec.Add "sender_email", CStr(data(i, 2))
        rec.Add "sender_name", CStr(data(i, 3))
        rec.Add "subject", CStr(data(i, 4))
        rec.Add "received_at", CStr(data(i, 5))
        rec.Add "folder_path", CStr(data(i, 6))
        rec.Add "body_path", CStr(data(i, 7))
        rec.Add "msg_path", CStr(data(i, 8))
        Dim attDict As Object: Set attDict = FolioLib.NewDict()
        Dim attStr As String: attStr = CStr(data(i, 9))
        If Len(attStr) > 0 Then
            Dim attParts() As String: attParts = Split(attStr, "|")
            Dim a As Long
            For a = 0 To UBound(attParts)
                If Len(attParts(a)) > 0 Then
                    Dim fn As String: fn = Mid$(attParts(a), InStrRev(attParts(a), "\") + 1)
                    attDict.Add attParts(a), fn
                End If
            Next a
        End If
        rec.Add "attachment_paths", attDict
        rec.Add "_mail_folder", CStr(data(i, 10))
        Set newRecs(eid) = rec
NextLMail:
    Next i
    Set m_feMailRecords = newRecs
    Exit Sub
ErrOut:
End Sub

Private Sub LoadMailIndexFromLocalSheet(wb As Workbook)
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_folio_mail_idx")
    If ws.Range("A1").value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.value
    If IsEmpty(data) Then Exit Sub
    Dim newIdx As Object: Set newIdx = FolioLib.NewDict()
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim key As String: key = CStr(data(i, 1))
        If Len(key) = 0 Then GoTo NextLIdx
        If Not newIdx.Exists(key) Then newIdx.Add key, FolioLib.NewDict()
        Dim inner As Object: Set inner = newIdx(key)
        inner(CStr(data(i, 2))) = True
NextLIdx:
    Next i
    Set m_feMailIndex = newIdx
    Exit Sub
ErrOut:
End Sub

Private Sub LoadCasesFromLocalSheet(wb As Workbook)
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_folio_cases")
    If ws.Range("A1").value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.value
    If IsEmpty(data) Then Exit Sub
    Dim newNames As Object: Set newNames = FolioLib.NewDict()
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim nm As String: nm = CStr(data(i, 1))
        If Len(nm) > 0 Then newNames(nm) = True
    Next i
    Set m_feCaseNames = newNames
    Exit Sub
ErrOut:
End Sub

' Load ALL case files into Dict indexed by case_id (from _folio_files sheet)
Private Sub LoadCaseFilesFromLocalSheet(wb As Workbook)
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_folio_files")
    If ws.Range("A1").value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.value
    If IsEmpty(data) Then Exit Sub
    If UBound(data, 2) < 7 Then Exit Sub
    Dim newFiles As Object: Set newFiles = FolioLib.NewDict()
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim cid As String: cid = CStr(data(i, 1))
        If Len(cid) = 0 Then GoTo NextFile
        If Not newFiles.Exists(cid) Then newFiles.Add cid, FolioLib.NewDict()
        Dim inner As Object: Set inner = newFiles(cid)
        Dim rec As Object: Set rec = FolioLib.NewDict()
        rec.Add "case_id", cid
        rec.Add "file_name", CStr(data(i, 2))
        rec.Add "file_path", CStr(data(i, 3))
        rec.Add "folder_path", CStr(data(i, 4))
        rec.Add "relative_path", CStr(data(i, 5))
        rec.Add "file_size", CStr(data(i, 6))
        rec.Add "modified_at", CStr(data(i, 7))
        Set inner(CStr(data(i, 3))) = rec
NextFile:
    Next i
    Set m_feCaseFiles = newFiles
    Exit Sub
ErrOut:
End Sub

' O(1) lookup: get all files for a specific case ID
Public Function FindCaseFiles(caseId As String) As Object
    Set FindCaseFiles = FolioLib.NewDict()
    If m_feCaseFiles Is Nothing Then Exit Function
    If Len(caseId) = 0 Then Exit Function
    ' Prefix match: case folder may be "R06-001" or "R06-001_Name"
    Dim keys As Variant
    If m_feCaseFiles.count = 0 Then Exit Function
    keys = m_feCaseFiles.keys
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim k As String: k = CStr(keys(i))
        Dim baseName As String: baseName = k
        Dim usPos As Long: usPos = InStr(baseName, "_")
        If usPos > 0 Then baseName = Left$(baseName, usPos - 1)
        If LCase$(baseName) = LCase$(caseId) Then
            Set FindCaseFiles = m_feCaseFiles(k)
            Exit Function
        End If
    Next i
End Function
