Attribute VB_Name = "ToolboxMain"
Option Explicit

Private g_initialized As Boolean

Public Sub InitAddin()
    g_initialized = True
End Sub

Public Sub ShutdownAddin()
    g_initialized = False
End Sub

Public Sub Ribbon_CreateDraftSheet(control As IRibbonControl)
    On Error GoTo ErrHandler
    CreateOutlookDraftSheet
    Exit Sub
ErrHandler:
    ErrorHandler.ShowToolError "vba-toolbox", Err.Description
End Sub

Public Sub Ribbon_RunDrafts(control As IRibbonControl)
    On Error GoTo ErrHandler
    RunOutlookDrafts
    Exit Sub
ErrHandler:
    ErrorHandler.ShowToolError "vba-toolbox", Err.Description
End Sub

Public Sub CreateOutlookDraftSheet()
    OutlookDraftEntry.CreateDraftSheetFromActiveWorkbook
End Sub

Public Sub RunOutlookDrafts()
    OutlookDraftEntry.RunFromActiveSheet
End Sub
