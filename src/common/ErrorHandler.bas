Attribute VB_Name = "ErrorHandler"
Option Explicit

Public Sub ShowToolError(ByVal Title As String, ByVal Message As String)
    MsgBox Message, vbExclamation, Title
End Sub

