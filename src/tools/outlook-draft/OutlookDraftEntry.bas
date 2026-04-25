Attribute VB_Name = "OutlookDraftEntry"
Option Explicit

Public Sub CreateDraftSheetFromActiveWorkbook()
    If ActiveWorkbook Is Nothing Then
        ErrorHandler.ShowToolError "vba-toolbox", "No active workbook."
        Exit Sub
    End If
    If ActiveWorkbook Is ThisWorkbook Then
        ErrorHandler.ShowToolError "vba-toolbox", "Switch to the target workbook first."
        Exit Sub
    End If

    OutlookDraftSheet.CreateDraftTemplate ActiveWorkbook
End Sub

Public Sub RunFromActiveSheet()
    If ActiveWorkbook Is Nothing Then Exit Sub
    OutlookDraftRun.RunFromWorksheet ActiveSheet
End Sub
