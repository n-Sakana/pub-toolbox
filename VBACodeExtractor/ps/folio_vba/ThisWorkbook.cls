Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    FolioMain.BeforeWorkbookClose
    Me.Saved = True
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    Dim sn As String: sn = Sh.Name
    If Left$(sn, 6) <> "_folio" Then Exit Sub
    Application.ScreenUpdating = False
    If FolioMain.g_formLoaded Then frmFolio.OnFolioSheetChange sn
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub
