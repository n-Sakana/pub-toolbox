Attribute VB_Name = "OutlookDraftRun"
Option Explicit

Public Sub RunFromWorksheet(ByVal ws As Worksheet)
    Dim lo As ListObject
    Dim olApp As Object
    Dim draftItem As Object
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim draftedCount As Long
    Dim errorCount As Long
    Dim colFrom As Long
    Dim colTo As Long
    Dim colCc As Long
    Dim colBcc As Long
    Dim colSubject As Long
    Dim colBody As Long
    Dim colAttachments As Long
    Dim colStatus As Long
    Dim colError As Long
    Dim fromValue As String
    Dim toValue As String

    On Error GoTo FailFast

    Set lo = GetDraftTable(ws)
    If lo Is Nothing Then
        ErrorHandler.ShowToolError "Outlook Draft", "Draft table was not found on the active sheet."
        Exit Sub
    End If
    If lo.DataBodyRange Is Nothing Then
        ErrorHandler.ShowToolError "Outlook Draft", "Draft table has no rows."
        Exit Sub
    End If

    colFrom = lo.ListColumns("from").Index
    colTo = lo.ListColumns("to").Index
    colCc = lo.ListColumns("cc").Index
    colBcc = lo.ListColumns("bcc").Index
    colSubject = lo.ListColumns("subject").Index
    colBody = lo.ListColumns("body").Index
    colAttachments = lo.ListColumns("attachments").Index
    colStatus = lo.ListColumns("status").Index
    colError = lo.ListColumns("error").Index

    Set olApp = GetOutlookApp()
    rowCount = lo.DataBodyRange.Rows.Count
    SetStatusBar "Outlook Draft: starting"

    For rowIndex = 1 To rowCount
        lo.DataBodyRange.Cells(rowIndex, colStatus).Value2 = "Running"
        lo.DataBodyRange.Cells(rowIndex, colError).Value2 = vbNullString
        SetStatusBar "Outlook Draft: " & CStr(rowIndex) & "/" & CStr(rowCount) & " drafting..."
        DoEvents

        toValue = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, colTo).Value2))
        If Len(toValue) = 0 Then
            errorCount = errorCount + 1
            lo.DataBodyRange.Cells(rowIndex, colStatus).Value2 = "Error"
            lo.DataBodyRange.Cells(rowIndex, colError).Value2 = "To is empty."
            GoTo NextRow
        End If

        fromValue = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, colFrom).Value2))

        On Error GoTo RowFail
        Set draftItem = olApp.CreateItem(0)
        ApplyFromAccount olApp, draftItem, fromValue
        draftItem.To = toValue
        draftItem.CC = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, colCc).Value2))
        draftItem.BCC = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, colBcc).Value2))
        draftItem.Subject = CStr(lo.DataBodyRange.Cells(rowIndex, colSubject).Value2)
        draftItem.Body = CStr(lo.DataBodyRange.Cells(rowIndex, colBody).Value2)
        AddAttachments draftItem, CStr(lo.DataBodyRange.Cells(rowIndex, colAttachments).Value2)
        draftItem.Save
        lo.DataBodyRange.Cells(rowIndex, colStatus).Value2 = "Drafted"
        draftedCount = draftedCount + 1
        Set draftItem = Nothing
        On Error GoTo FailFast

NextRow:
        DoEvents
    Next rowIndex

    SetStatusBar "Outlook Draft: done"
    MsgBox "Drafted=" & CStr(draftedCount) & ", Error=" & CStr(errorCount), vbInformation, "Outlook Draft"
    ClearStatusBar
    Exit Sub

RowFail:
    errorCount = errorCount + 1
    lo.DataBodyRange.Cells(rowIndex, colStatus).Value2 = "Error"
    lo.DataBodyRange.Cells(rowIndex, colError).Value2 = Err.Description
    Err.Clear
    Set draftItem = Nothing
    Resume NextRow

FailFast:
    ClearStatusBar
    ErrorHandler.ShowToolError "Outlook Draft", Err.Description
End Sub

Private Function GetDraftTable(ByVal ws As Worksheet) As ListObject
    On Error Resume Next
    Set GetDraftTable = ws.ListObjects(ToolRegistry.TABLE_OUTLOOK_DRAFT)
    On Error GoTo 0
End Function

Private Function GetOutlookApp() As Object
    On Error Resume Next
    Set GetOutlookApp = GetObject(, "Outlook.Application")
    If GetOutlookApp Is Nothing Then
        Set GetOutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    If GetOutlookApp Is Nothing Then
        Err.Raise vbObjectError + 6000, "OutlookDraftRun", "Outlook could not be started."
    End If
End Function

Private Sub ApplyFromAccount(ByVal olApp As Object, ByVal draftItem As Object, ByVal FromAddress As String)
    Dim accounts As Object
    Dim account As Object
    Dim accountIndex As Long
    Dim targetAddress As String
    Dim accountAddress As String

    FromAddress = Trim$(FromAddress)
    If Len(FromAddress) = 0 Then Exit Sub

    targetAddress = LCase$(FromAddress)
    Set accounts = olApp.Session.Accounts

    For accountIndex = 1 To accounts.Count
        Set account = accounts.Item(accountIndex)
        accountAddress = vbNullString
        On Error Resume Next
        accountAddress = LCase$(Trim$(CStr(account.SmtpAddress)))
        On Error GoTo 0

        If Len(accountAddress) > 0 And accountAddress = targetAddress Then
            Set draftItem.SendUsingAccount = account
            Exit Sub
        End If
    Next accountIndex

    Err.Raise vbObjectError + 6002, "OutlookDraftRun", "From account not found: " & FromAddress
End Sub

Private Sub AddAttachments(ByVal draftItem As Object, ByVal AttachmentsText As String)
    Dim parts() As String
    Dim i As Long
    Dim filePath As String

    AttachmentsText = Trim$(AttachmentsText)
    If Len(AttachmentsText) = 0 Then Exit Sub

    parts = Split(AttachmentsText, ";")
    For i = LBound(parts) To UBound(parts)
        filePath = Trim$(parts(i))
        If Len(filePath) = 0 Then GoTo NextPart
        If Dir$(filePath) = vbNullString Then
            Err.Raise vbObjectError + 6001, "OutlookDraftRun", "Attachment not found: " & filePath
        End If
        draftItem.Attachments.Add filePath
NextPart:
    Next i
End Sub

Private Sub SetStatusBar(ByVal Message As String)
    Application.StatusBar = Message
End Sub

Private Sub ClearStatusBar()
    Application.StatusBar = False
End Sub
