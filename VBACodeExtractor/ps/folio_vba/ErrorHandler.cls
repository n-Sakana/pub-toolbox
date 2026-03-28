Attribute VB_Name = "ErrorHandler"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Centralized error handler with log accumulation.
' All output goes to Immediate window (Ctrl+G).
'
' Usage:
'   Dim eh As New ErrorHandler
'   eh.Enter "ModuleName", "ProcName"
'   On Error GoTo ErrHandler
'   eh.Log "processing " & count & " items"
'   ...
'   eh.OK: Exit Sub
'   ErrHandler: eh.Catch

Private m_module As String
Private m_proc As String
Private m_logs As Collection
Private m_t0 As Single
Private m_verbose As Boolean

Public Sub Enter(moduleName As String, procName As String)
    m_module = moduleName
    m_proc = procName
    Set m_logs = New Collection
    m_t0 = Timer
End Sub

Public Sub Log(msg As String)
    If m_logs Is Nothing Then Set m_logs = New Collection
    m_logs.Add Format$(Timer - m_t0, "0.000") & "s " & msg
End Sub

Public Sub Catch()
    Dim prefix As String: prefix = m_module & "." & m_proc
    Debug.Print "!! " & prefix & " | Err " & Err.Number & ": " & Err.Description
    ' Dump accumulated logs for context
    If Not m_logs Is Nothing Then
        If m_logs.count > 0 Then
            Debug.Print "   Log trace (" & m_logs.count & " entries):"
            Dim i As Long
            For i = 1 To m_logs.count
                Debug.Print "     " & m_logs(i)
            Next i
        End If
    End If
    Debug.Print "   Elapsed: " & Format$(Timer - m_t0, "0.000") & "s"
End Sub

Public Sub Trace(detail As String)
    Debug.Print "-- " & m_module & "." & m_proc & " | " & detail
End Sub

Public Sub OK()
    Set m_logs = Nothing
End Sub

Public Property Get Elapsed() As Single
    Elapsed = Timer - m_t0
End Property
