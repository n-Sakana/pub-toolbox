Attribute VB_Name = "FieldEditor"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private WithEvents m_txt As MSForms.TextBox
Attribute m_txt.VB_VarHelpID = -1
Private m_fieldName As String
Private m_fieldType As String
Private m_form As Object
Private m_isReadOnly As Boolean
Private m_prevValue As String
Private m_suppressing As Boolean
Private m_committedValue As String  ' value at last focus-enter / external set

Public Sub Init(txt As MSForms.TextBox, FieldName As String, frm As Object, IsReadOnly As Boolean, Optional fieldType As String = "text")
    Set m_txt = txt
    m_fieldName = FieldName
    m_fieldType = fieldType
    Set m_form = frm
    m_isReadOnly = IsReadOnly
    m_prevValue = ""
    m_suppressing = False
End Sub

Public Property Get FieldName() As String
    FieldName = m_fieldName
End Property

Public Property Get Text() As String
    Text = m_txt.Text
End Property

Public Property Get PrevValue() As String
    PrevValue = m_prevValue
End Property

Public Property Get TextBox() As MSForms.TextBox
    Set TextBox = m_txt
End Property

Public Property Get IsReadOnly() As Boolean
    IsReadOnly = m_isReadOnly
End Property

Public Sub SetValue(val As String)
    m_suppressing = True
    m_prevValue = val
    m_committedValue = val
    m_txt.Text = FormatDisplay(val)
    m_txt.SelStart = 0
    m_suppressing = False
End Sub

Public Sub RefreshIfChanged(val As String)
    If val <> m_prevValue And m_txt.Text = FormatDisplay(m_prevValue) Then
        Dim oldVal As String: oldVal = m_prevValue
        m_suppressing = True
        m_prevValue = val
        m_committedValue = val
        m_txt.Text = FormatDisplay(val)
        m_suppressing = False
        ' Log external change (SheetWatcher detected table edit)
        If m_form Is Nothing Then Exit Sub
        m_form.OnFieldChanged m_fieldName, oldVal, val, "external"
    End If
End Sub

Private Function FormatDisplay(val As String) As String
    FormatDisplay = val
    If m_fieldType <> "currency" Then Exit Function
    If Len(val) = 0 Then Exit Function
    If Not IsNumeric(val) Then Exit Function
    On Error Resume Next
    FormatDisplay = Format$(CDbl(val), "#,##0")
    ' Keep decimals if present
    If InStr(val, ".") > 0 Then FormatDisplay = Format$(CDbl(val), "#,##0.##")
    On Error GoTo 0
End Function

Private Function ParseRawValue() As String
    ' Strip commas from displayed currency to get raw value
    Dim txt As String: txt = m_txt.Text
    If m_fieldType = "currency" Then
        txt = Replace(txt, ",", "")
    End If
    ParseRawValue = txt
End Function

Private Sub m_txt_Change()
    If m_suppressing Then Exit Sub
    If m_form Is Nothing Then Exit Sub
    Dim newVal As String: newVal = ParseRawValue()
    If newVal <> m_prevValue Then
        m_prevValue = newVal
        ' Write to table immediately but don't log yet (log on Exit)
        If Not m_isReadOnly Then
            m_form.OnFieldEdited m_fieldName, newVal
        End If
    End If
End Sub

Public Sub CommitEdit()
    ' Flush pending local edit as a single log entry
    If m_prevValue <> m_committedValue Then
        Dim oldVal As String: oldVal = m_committedValue
        m_committedValue = m_prevValue
        If Not m_form Is Nothing Then
            m_form.OnFieldChanged m_fieldName, oldVal, m_prevValue, "local"
        End If
    End If
End Sub

Private Sub Class_Terminate()
    Set m_txt = Nothing
    Set m_form = Nothing
End Sub
