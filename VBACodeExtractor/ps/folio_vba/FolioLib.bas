Attribute VB_Name = "FolioLib"
Option Explicit

' ============================================================================
' FolioLib: Merged from FolioHelpers + FolioConfig + FolioChangeLog
' ============================================================================

' ##########################################################################
' # SECTION: General Utilities (ex-FolioHelpers)
' ##########################################################################

' --- Helpers state ---
Private m_readStm As Object  ' Reusable ADODB.Stream for ReadTextFile
Private m_writeStm As Object ' Reusable ADODB.Stream for WriteTextFile

' --- Config state ---
Private Const SH_CONFIG As String = "_folio_config"
Private Const SH_SOURCES As String = "_folio_sources"
Private Const SH_FIELDS As String = "_folio_fields"
Private m_cfg As Object       ' Dict: key -> value
Private m_sources As Object   ' Dict: source_name -> Dict(col -> value)
Private m_fields As Object    ' Dict: "source|field" -> Dict(col -> value)
Private m_loaded As Boolean
Private m_dirty As Boolean

' --- ChangeLog state ---
Private Const LOG_SHEET As String = "_folio_log"
Private Const LOG_TABLE As String = "FolioLog"
Private Const MAX_LOG_ROWS As Long = 5000

' ============================================================================
' JSON Parser
' ============================================================================

Public Function ParseJson(ByVal Text As String) As Object
    Dim p As Long: p = 1
    SkipWS Text, p
    If p > Len(Text) Then Set ParseJson = NewDict(): Exit Function
    Dim ch As String: ch = Mid$(Text, p, 1)
    If ch = "{" Then
        Set ParseJson = ParseObj(Text, p)
    ElseIf ch = "[" Then
        Set ParseJson = ParseArr(Text, p)
    Else
        Set ParseJson = NewDict()
    End If
End Function

Private Function ParseObj(ByRef s As String, ByRef p As Long) As Object
    Dim d As Object: Set d = NewDict()
    p = p + 1
    SkipWS s, p
    If p <= Len(s) Then
        If Mid$(s, p, 1) = "}" Then p = p + 1: Set ParseObj = d: Exit Function
    End If
    Do
        SkipWS s, p
        Dim k As String: k = ParseStr(s, p)
        SkipWS s, p
        If p <= Len(s) Then p = p + 1 ' skip :
        Dim v As Variant: ParseVal s, p, v
        d.Add k, v
        SkipWS s, p
        If p > Len(s) Then Exit Do
        If Mid$(s, p, 1) = "}" Then p = p + 1: Exit Do
        p = p + 1 ' skip ,
    Loop
    Set ParseObj = d
End Function

Private Function ParseArr(ByRef s As String, ByRef p As Long) As Object
    Dim c As New Collection
    p = p + 1
    SkipWS s, p
    If p <= Len(s) Then
        If Mid$(s, p, 1) = "]" Then p = p + 1: Set ParseArr = c: Exit Function
    End If
    Do
        SkipWS s, p
        Dim v As Variant: ParseVal s, p, v
        c.Add v
        SkipWS s, p
        If p > Len(s) Then Exit Do
        If Mid$(s, p, 1) = "]" Then p = p + 1: Exit Do
        p = p + 1 ' skip ,
    Loop
    Set ParseArr = c
End Function

Private Sub ParseVal(ByRef s As String, ByRef p As Long, ByRef result As Variant)
    SkipWS s, p
    If p > Len(s) Then result = Null: Exit Sub
    Dim ch As String: ch = Mid$(s, p, 1)
    Select Case ch
        Case "{":  Set result = ParseObj(s, p)
        Case "[":  Set result = ParseArr(s, p)
        Case """": result = ParseStr(s, p)
        Case "t":  result = True: p = p + 4
        Case "f":  result = False: p = p + 5
        Case "n":  result = Null: p = p + 4
        Case Else: result = ParseNum(s, p)
    End Select
End Sub

Private Function ParseStr(ByRef s As String, ByRef p As Long) As String
    p = p + 1
    Dim buf As String, start As Long: start = p
    Do While p <= Len(s)
        Dim ch As String: ch = Mid$(s, p, 1)
        If ch = """" Then
            buf = buf & Mid$(s, start, p - start)
            p = p + 1
            ParseStr = buf: Exit Function
        ElseIf ch = "\" Then
            buf = buf & Mid$(s, start, p - start)
            p = p + 1
            If p <= Len(s) Then
                Dim esc As String: esc = Mid$(s, p, 1)
                Select Case esc
                    Case """", "\", "/": buf = buf & esc
                    Case "n": buf = buf & vbLf
                    Case "r": buf = buf & vbCr
                    Case "t": buf = buf & vbTab
                    Case "u"
                        If p + 4 <= Len(s) Then
                            On Error Resume Next
                            buf = buf & ChrW$(CLng("&H" & Mid$(s, p + 1, 4)))
                            On Error GoTo 0
                            p = p + 4
                        End If
                End Select
                p = p + 1: start = p
            End If
        Else
            p = p + 1
        End If
    Loop
    ParseStr = buf & Mid$(s, start, p - start)
End Function

Private Function ParseNum(ByRef s As String, ByRef p As Long) As Double
    Dim start As Long: start = p
    If p <= Len(s) Then If Mid$(s, p, 1) = "-" Then p = p + 1
    Do While p <= Len(s)
        Dim ch As String: ch = Mid$(s, p, 1)
        If ch Like "[0-9.eE+-]" Then p = p + 1 Else Exit Do
    Loop
    On Error Resume Next
    ParseNum = CDbl(Mid$(s, start, p - start))
    On Error GoTo 0
End Function

Private Sub SkipWS(ByRef s As String, ByRef p As Long)
    Do While p <= Len(s)
        Select Case Mid$(s, p, 1)
            Case " ", vbTab, vbLf, vbCr: p = p + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

' ============================================================================
' JSON Serializer
' ============================================================================

Public Function ToJson(ByVal v As Variant, Optional ind As Long = -1) As String
    If IsObject(v) Then
        If v Is Nothing Then ToJson = "null": Exit Function
        Dim obj As Object: Set obj = v
        If TypeName(obj) = "Dictionary" Then ToJson = DictToJson(obj, ind): Exit Function
        If TypeName(obj) = "Collection" Then ToJson = CollToJson(obj, ind): Exit Function
        ToJson = "null"
    ElseIf IsNull(v) Or IsEmpty(v) Then
        ToJson = "null"
    ElseIf VarType(v) = vbString Then
        ToJson = """" & JsonEscape(CStr(v)) & """"
    ElseIf VarType(v) = vbBoolean Then
        ToJson = IIf(v, "true", "false")
    ElseIf IsNumeric(v) Then
        ToJson = CStr(v)
    Else
        ToJson = """" & JsonEscape(CStr(v)) & """"
    End If
End Function

Private Function DictToJson(d As Object, ind As Long) As String
    If d.count = 0 Then DictToJson = "{}": Exit Function
    Dim keys() As Variant: keys = d.keys
    Dim nl As String, sp As String, ind2 As Long, csp As String
    If ind >= 0 Then nl = vbCrLf: sp = String$(ind + 2, " "): ind2 = ind + 2: csp = String$(ind, " ") Else ind2 = -1
    Dim parts() As String: ReDim parts(d.count - 1)
    Dim i As Long
    For i = 0 To d.count - 1
        Dim val As Variant
        If IsObject(d(keys(i))) Then Set val = d(keys(i)) Else val = d(keys(i))
        parts(i) = sp & """" & JsonEscape(CStr(keys(i))) & """:" & IIf(ind >= 0, " ", "") & ToJson(val, ind2)
    Next i
    DictToJson = "{" & nl & Join(parts, "," & nl) & nl & csp & "}"
End Function

Private Function CollToJson(c As Object, ind As Long) As String
    If c.count = 0 Then CollToJson = "[]": Exit Function
    Dim nl As String, sp As String, ind2 As Long, csp As String
    If ind >= 0 Then nl = vbCrLf: sp = String$(ind + 2, " "): ind2 = ind + 2: csp = String$(ind, " ") Else ind2 = -1
    Dim parts() As String: ReDim parts(c.count - 1)
    Dim i As Long
    For i = 1 To c.count
        Dim val As Variant
        If IsObject(c(i)) Then Set val = c(i) Else val = c(i)
        parts(i - 1) = sp & ToJson(val, ind2)
    Next i
    CollToJson = "[" & nl & Join(parts, "," & nl) & nl & csp & "]"
End Function

Public Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

' ============================================================================
' Fast mail meta.json parser (InStr-based, no char-by-char parsing)
' Expects fixed structure: flat object with "attachments" array of {"path":"..."} objects
' ============================================================================

Public Function ParseMailMeta(ByVal json As String) As Object
    Dim d As Object: Set d = NewDict()
    Set ParseMailMeta = d
    If Len(json) < 5 Then Exit Function

    ' Extract simple string fields via InStr
    ExtractField json, d, "mail_id"
    ExtractField json, d, "entry_id"
    ExtractField json, d, "mailbox_address"
    ExtractField json, d, "folder_path"
    ExtractField json, d, "received_at"
    ExtractField json, d, "sender_name"
    ExtractField json, d, "sender_email"
    ExtractField json, d, "subject"
    ExtractField json, d, "body_path"
    ExtractField json, d, "msg_path"

    ' Extract attachments: just collect path strings into a Collection (no Dict per attachment)
    Dim attStart As Long: attStart = InStr(1, json, """attachments""")
    If attStart > 0 Then
        Dim attCol As New Collection
        Dim searchPos As Long: searchPos = attStart
        Do
            Dim pathKey As Long: pathKey = InStr(searchPos, json, """path""")
            If pathKey = 0 Then Exit Do
            Dim pathVal As String: pathVal = ExtractValueAt(json, pathKey + 6)
            If Len(pathVal) > 0 Then attCol.Add pathVal
            searchPos = pathKey + 6
        Loop
        d.Add "attachments", attCol
    End If
End Function

Private Sub ExtractField(ByRef json As String, d As Object, FieldName As String)
    Dim key As String: key = """" & FieldName & """"
    Dim pos As Long: pos = InStr(1, json, key)
    If pos = 0 Then Exit Sub
    ' Skip past key, colon, optional spaces, opening quote
    Dim valStart As Long: valStart = InStr(pos + Len(key), json, """")
    If valStart = 0 Then Exit Sub
    valStart = valStart + 1
    ' Find closing quote (handle \" escapes)
    Dim valEnd As Long: valEnd = valStart
    Do
        valEnd = InStr(valEnd, json, """")
        If valEnd = 0 Then Exit Sub
        ' Check for escape
        If Mid$(json, valEnd - 1, 1) <> "\" Then Exit Do
        valEnd = valEnd + 1
    Loop
    Dim val As String: val = Mid$(json, valStart, valEnd - valStart)
    ' Unescape common sequences
    If InStr(1, val, "\") > 0 Then
        val = Replace(val, "\""", """")
        val = Replace(val, "\\", "\")
        val = Replace(val, "\n", vbLf)
        val = Replace(val, "\t", vbTab)
    End If
    d.Add FieldName, val
End Sub

Private Function ExtractValueAt(ByRef json As String, startPos As Long) As String
    Dim valStart As Long: valStart = InStr(startPos, json, """")
    If valStart = 0 Then Exit Function
    valStart = valStart + 1
    Dim valEnd As Long: valEnd = InStr(valStart, json, """")
    If valEnd = 0 Then Exit Function
    ExtractValueAt = Mid$(json, valStart, valEnd - valStart)
End Function

' ============================================================================
' Dictionary Helpers
' ============================================================================

Public Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Public Function DictStr(d As Object, key As String, Optional def As String = "") As String
    DictStr = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Or IsNull(d(key)) Then Exit Function
    DictStr = CStr(d(key))
End Function

Public Function DictObj(d As Object, key As String) As Object
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If Not IsObject(d(key)) Then Exit Function
    Set DictObj = d(key)
End Function

Public Function DictBool(d As Object, key As String, Optional def As Boolean = False) As Boolean
    DictBool = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Or IsNull(d(key)) Then Exit Function
    DictBool = CBool(d(key))
End Function

Public Function DictLng(d As Object, key As String, Optional def As Long = 0) As Long
    DictLng = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Or IsNull(d(key)) Then Exit Function
    DictLng = CLng(d(key))
End Function

Public Sub DictPut(d As Object, key As String, val As Variant)
    If d.Exists(key) Then d.Remove key
    d.Add key, val
End Sub

' ============================================================================
' File System
' ============================================================================

Public Function ReadTextFile(path As String) As String
    On Error GoTo ErrOut
    ReadTextFile = ""
    If Len(Dir$(path)) = 0 Then Exit Function

    ' Read raw bytes with shared access (avoids file lock)
    Dim f As Long: f = FreeFile
    Open path For Binary Access Read Shared As #f
    Dim size As Long: size = LOF(f)
    If size = 0 Then Close #f: Exit Function
    Dim buf() As Byte: ReDim buf(0 To size - 1)
    Get #f, , buf
    Close #f

    ' Convert UTF-8 bytes to VBA string (reuse stream object)
    If m_readStm Is Nothing Then Set m_readStm = CreateObject("ADODB.Stream")
    m_readStm.Type = 1: m_readStm.Open: m_readStm.Write buf
    m_readStm.Position = 0: m_readStm.Type = 2: m_readStm.Charset = "UTF-8"
    ReadTextFile = m_readStm.ReadText
    m_readStm.Close
    Exit Function
ErrOut:
    ReadTextFile = ""
    On Error Resume Next: If Not m_readStm Is Nothing Then m_readStm.Close: On Error GoTo 0
End Function

Public Sub WriteTextFile(path As String, content As String)
    On Error GoTo ErrOut
    If m_writeStm Is Nothing Then Set m_writeStm = CreateObject("ADODB.Stream")
    m_writeStm.Type = 2: m_writeStm.Charset = "UTF-8"
    m_writeStm.Open: m_writeStm.WriteText content
    m_writeStm.Position = 0: m_writeStm.Type = 1: m_writeStm.Position = 3
    Dim out As Object: Set out = CreateObject("ADODB.Stream")
    out.Type = 1: out.Open
    m_writeStm.CopyTo out
    out.SaveToFile path, 2
    out.Close: m_writeStm.Close
    Exit Sub
ErrOut:
End Sub

Public Sub EnsureFolder(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    Dim parent As String: parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then EnsureFolder parent
    End If
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Public Function FileExists(path As String) As Boolean
    FileExists = Len(Dir$(path)) > 0
End Function

Public Function FolderExists(path As String) As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(path)
End Function

' ============================================================================
' String Utilities
' ============================================================================

Public Function SafeName(ByVal Text As String) As String
    Text = Trim$(Text)
    If Len(Text) = 0 Then Text = "blank"
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        Text = Replace(Text, CStr(bad), "_")
    Next bad
    If Len(Text) > 80 Then Text = Left$(Text, 80)
    SafeName = Text
End Function

Public Function FormatFieldValue(val As Variant, Optional fieldType As String = "text") As String
    If IsNull(val) Or IsEmpty(val) Then FormatFieldValue = "": Exit Function
    If IsObject(val) Then
        If TypeName(val) = "Collection" Then
            Dim parts() As String: ReDim parts(val.count - 1)
            Dim i As Long
            For i = 1 To val.count: parts(i - 1) = CStr(val(i)): Next i
            FormatFieldValue = Join(parts, "; "): Exit Function
        End If
        FormatFieldValue = "": Exit Function
    End If
    Dim s As String: s = CStr(val)
    Select Case fieldType
        Case "date"
            If VarType(val) = vbDate Then
                FormatFieldValue = Format$(CDate(val), "yyyy/mm/dd"): Exit Function
            End If
        Case "currency"
            If IsNumeric(s) Then
                Dim n As Double: n = CDbl(s)
                If n = Int(n) Then
                    FormatFieldValue = Format$(CLng(n), "#,0")
                Else
                    FormatFieldValue = Format$(n, "#,0.##")
                End If
                Exit Function
            End If
    End Select
    FormatFieldValue = s
End Function

Public Function GetFieldLabel(FieldName As String) As String
    GetFieldLabel = Replace(FieldName, "_", " ")
End Function

Public Function GetFieldGroup(FieldName As String) As String
    Dim pos As Long: pos = InStr(FieldName, "_")
    If pos > 1 And pos < Len(FieldName) Then GetFieldGroup = Left$(FieldName, pos - 1)
End Function

Public Function GetFieldShortName(FieldName As String) As String
    Dim pos As Long: pos = InStr(FieldName, "_")
    If pos > 1 And pos < Len(FieldName) Then
        GetFieldShortName = Mid$(FieldName, pos + 1)
    Else
        GetFieldShortName = FieldName
    End If
End Function

Public Function CountFieldGroups(fields As Collection) As Long
    Dim groups As Object: Set groups = NewDict()
    Dim i As Long
    For i = 1 To fields.count
        Dim g As String: g = GetFieldGroup(CStr(fields(i)))
        If Len(g) > 0 And Not groups.Exists(g) Then groups.Add g, True
    Next i
    CountFieldGroups = groups.count
End Function

' ##########################################################################
' # SECTION: Config Management (ex-FolioConfig)
' ##########################################################################

' ============================================================================
' Init / Save
' ============================================================================

Public Sub EnsureConfigSheets()
    EnsureSheet SH_CONFIG, Array("key", "value")
    EnsureSheet SH_SOURCES, Array("source_name", "key_column", "display_name_column", "mail_link_column", "folder_link_column", "mail_match_mode")
    EnsureSheet SH_FIELDS, Array("source_name", "field_name", "type", "in_list", "editable", "multiline")
    If Not m_loaded Then LoadFromSheets
End Sub

Public Sub SaveToSheets()
    If Not m_loaded Then Exit Sub
    If Not m_dirty Then Exit Sub
    On Error Resume Next
    SaveConfigSheet
    SaveSourcesSheet
    SaveFieldsSheet
    m_dirty = False
    On Error GoTo 0
End Sub

Private Sub LoadFromSheets()
    On Error Resume Next
    Set m_cfg = CreateObject("Scripting.Dictionary")
    Set m_sources = CreateObject("Scripting.Dictionary")
    Set m_fields = CreateObject("Scripting.Dictionary")

    ' Load config KV
    Dim wsCfg As Worksheet: Set wsCfg = ThisWorkbook.Worksheets(SH_CONFIG)
    If Not wsCfg Is Nothing Then
        Dim r As Long
        For r = 2 To wsCfg.Cells(wsCfg.Rows.count, 1).End(xlUp).Row
            Dim k As String: k = CStr(wsCfg.Cells(r, 1).value)
            If Len(k) > 0 Then m_cfg(k) = CStr(wsCfg.Cells(r, 2).value)
        Next r
    End If

    ' Load sources
    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets(SH_SOURCES)
    If Not wsSrc Is Nothing Then
        Dim srcCols As Object: Set srcCols = ReadHeaderMap(wsSrc)
        For r = 2 To wsSrc.Cells(wsSrc.Rows.count, 1).End(xlUp).Row
            Dim sn As String: sn = CStr(wsSrc.Cells(r, 1).value)
            If Len(sn) > 0 Then
                Dim sd As Object: Set sd = CreateObject("Scripting.Dictionary")
                Dim ck As Variant
                For Each ck In srcCols.keys
                    sd(CStr(ck)) = CStr(wsSrc.Cells(r, CLng(srcCols(ck))).value)
                Next ck
                Set m_sources(sn) = sd
            End If
        Next r
    End If

    ' Load fields
    Dim wsFld As Worksheet: Set wsFld = ThisWorkbook.Worksheets(SH_FIELDS)
    If Not wsFld Is Nothing Then
        Dim fldCols As Object: Set fldCols = ReadHeaderMap(wsFld)
        For r = 2 To wsFld.Cells(wsFld.Rows.count, 1).End(xlUp).Row
            Dim fs As String: fs = CStr(wsFld.Cells(r, 1).value)
            Dim ff As String: ff = CStr(wsFld.Cells(r, 2).value)
            If Len(fs) > 0 And Len(ff) > 0 Then
                Dim fk As String: fk = LCase$(fs) & "|" & LCase$(ff)
                Dim fd As Object: Set fd = CreateObject("Scripting.Dictionary")
                fd("source_name") = fs
                fd("field_name") = ff
                For Each ck In fldCols.keys
                    fd(CStr(ck)) = CStr(wsFld.Cells(r, CLng(fldCols(ck))).value)
                Next ck
                Set m_fields(fk) = fd
            End If
        Next r
    End If

    m_loaded = True
    m_dirty = False
    On Error GoTo 0
End Sub

Private Function ReadHeaderMap(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        Dim h As String: h = CStr(ws.Cells(1, c).value)
        If Len(h) > 0 Then d(h) = c
    Next c
    Set ReadHeaderMap = d
End Function

' ============================================================================
' Config (key-value)
' ============================================================================

Public Function GetStr(key As String, Optional def As String = "") As String
    If Not m_loaded Then EnsureConfigSheets
    GetStr = def
    If m_cfg.Exists(key) Then
        Dim v As String: v = CStr(m_cfg(key))
        If Len(v) > 0 Then GetStr = v
    End If
End Function

Public Sub SetStr(key As String, value As String)
    If Not m_loaded Then EnsureConfigSheets
    m_cfg(key) = value
    m_dirty = True
End Sub

Public Function GetLng(key As String, Optional def As Long = 0) As Long
    GetLng = def
    Dim s As String: s = GetStr(key)
    If Len(s) > 0 And IsNumeric(s) Then GetLng = CLng(s)
End Function

Public Sub SetLng(key As String, value As Long)
    SetStr key, CStr(value)
End Sub

' ============================================================================
' Sources
' ============================================================================

Public Function GetSourceNames() As Collection
    If Not m_loaded Then EnsureConfigSheets
    Set GetSourceNames = New Collection
    Dim k As Variant
    For Each k In m_sources.keys
        GetSourceNames.Add CStr(k)
    Next k
End Function

Public Function GetSourceStr(src As String, col As String, Optional def As String = "") As String
    If Not m_loaded Then EnsureConfigSheets
    GetSourceStr = def
    If Not m_sources.Exists(src) Then Exit Function
    Dim sd As Object: Set sd = m_sources(src)
    If sd.Exists(col) Then
        Dim v As String: v = CStr(sd(col))
        If Len(v) > 0 Then GetSourceStr = v
    End If
End Function

Public Sub SetSourceStr(src As String, col As String, value As String)
    If Not m_loaded Then EnsureConfigSheets
    If Not m_sources.Exists(src) Then Set m_sources(src) = CreateObject("Scripting.Dictionary")
    Dim sd As Object: Set sd = m_sources(src)
    sd(col) = value
    m_dirty = True
End Sub

Public Sub EnsureSource(src As String)
    If Not m_loaded Then EnsureConfigSheets
    If Not m_sources.Exists(src) Then
        Set m_sources(src) = CreateObject("Scripting.Dictionary")
        m_sources(src)("source_name") = src
        m_dirty = True
    End If
End Sub

' ============================================================================
' Fields
' ============================================================================

Public Function GetFieldNames(src As String) As Collection
    If Not m_loaded Then EnsureConfigSheets
    Set GetFieldNames = New Collection
    Dim k As Variant
    For Each k In m_fields.keys
        If Left$(CStr(k), Len(src) + 1) = LCase$(src) & "|" Then
            Dim fd As Object: Set fd = m_fields(k)
            GetFieldNames.Add CStr(fd("field_name"))
        End If
    Next k
End Function

Public Function GetFieldStr(src As String, fld As String, col As String, Optional def As String = "") As String
    If Not m_loaded Then EnsureConfigSheets
    GetFieldStr = def
    Dim fk As String: fk = LCase$(src) & "|" & LCase$(fld)
    If Not m_fields.Exists(fk) Then Exit Function
    Dim fd As Object: Set fd = m_fields(fk)
    If fd.Exists(col) Then
        Dim v As String: v = CStr(fd(col))
        If Len(v) > 0 Then GetFieldStr = v
    End If
End Function

Public Function GetFieldBool(src As String, fld As String, col As String, Optional def As Boolean = False) As Boolean
    GetFieldBool = def
    Dim v As String: v = GetFieldStr(src, fld, col)
    If Len(v) > 0 Then GetFieldBool = CBool(v)
End Function

Public Sub SetFieldStr(src As String, fld As String, col As String, value As String)
    If Not m_loaded Then EnsureConfigSheets
    Dim fk As String: fk = LCase$(src) & "|" & LCase$(fld)
    If Not m_fields.Exists(fk) Then
        Dim fd As Object: Set fd = CreateObject("Scripting.Dictionary")
        fd("source_name") = src
        fd("field_name") = fld
        Set m_fields(fk) = fd
    End If
    Dim d As Object: Set d = m_fields(fk)
    d(col) = value
    m_dirty = True
End Sub

Public Sub SetFieldBool(src As String, fld As String, col As String, value As Boolean)
    SetFieldStr src, fld, col, CStr(value)
End Sub

Public Sub EnsureField(src As String, fld As String)
    Dim fk As String: fk = LCase$(src) & "|" & LCase$(fld)
    If Not m_fields.Exists(fk) Then
        SetFieldStr src, fld, "type", "text"
        SetFieldStr src, fld, "in_list", CStr(False)
        SetFieldStr src, fld, "editable", CStr(True)
        SetFieldStr src, fld, "multiline", CStr(False)
    End If
End Sub

' ============================================================================
' Field Settings Auto-Init
' ============================================================================

Public Sub InitFieldSettingsFromTable(src As String, tbl As ListObject)
    If Not m_loaded Then EnsureConfigSheets
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name Like "_*" Then GoTo NextCol
        Dim fk As String: fk = LCase$(src) & "|" & LCase$(col.Name)
        If Not m_fields.Exists(fk) Then
            EnsureField src, col.Name
            SetFieldStr src, col.Name, "type", GuessFieldType(col)
            SetFieldStr src, col.Name, "multiline", CStr(GuessMultiline(col))
        End If
NextCol:
    Next col
End Sub

Private Function GuessFieldType(col As ListColumn) As String
    GuessFieldType = "text"
    On Error Resume Next
    If col.DataBodyRange Is Nothing Then Exit Function
    If col.DataBodyRange.Rows.count = 0 Then Exit Function
    Dim fmt As String: fmt = CStr(col.DataBodyRange.Cells(1, 1).NumberFormat)
    If fmt Like "*yy*" Or fmt Like "*mm*dd*" Then GuessFieldType = "date": Exit Function
    If fmt Like "*#,##0*" Or fmt Like "*" & ChrW$(165) & "*" Or fmt Like "*$*" Then GuessFieldType = "currency": Exit Function
    If fmt Like "#*" Or fmt Like "0*" Or fmt Like "*%*" Then GuessFieldType = "number": Exit Function
    Dim r As Long
    For r = 1 To Application.Min(10, col.DataBodyRange.Rows.count)
        Dim v As Variant: v = col.DataBodyRange.Cells(r, 1).value
        If Not IsEmpty(v) And Not IsNull(v) Then
            If VarType(v) = vbDate Then GuessFieldType = "date": Exit Function
            If VarType(v) = vbCurrency Then GuessFieldType = "currency": Exit Function
            If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or _
               VarType(v) = vbSingle Then GuessFieldType = "number": Exit Function
            Exit Function
        End If
    Next r
    On Error GoTo 0
End Function

Private Function GuessMultiline(col As ListColumn) As Boolean
    GuessMultiline = False
    On Error Resume Next
    If col.DataBodyRange Is Nothing Then Exit Function
    Dim r As Long
    For r = 1 To Application.Min(10, col.DataBodyRange.Rows.count)
        Dim v As Variant: v = col.DataBodyRange.Cells(r, 1).value
        If Not IsEmpty(v) And Not IsNull(v) Then
            Dim s As String: s = CStr(v)
            If InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Or Len(s) > 30 Then
                GuessMultiline = True: Exit Function
            End If
        End If
    Next r
    On Error GoTo 0
End Function

' ============================================================================
' Sheet Persistence (private)
' ============================================================================

Private Sub EnsureSheet(shName As String, headers As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = shName
        ws.Visible = xlSheetVeryHidden
        Dim i As Long
        For i = 0 To UBound(headers)
            ws.Cells(1, i + 1).value = headers(i)
        Next i
    End If
End Sub

Private Sub SaveConfigSheet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    ' Clear existing data
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Delete
    If m_cfg.count = 0 Then Exit Sub
    Dim keys As Variant: keys = m_cfg.keys
    Dim data() As Variant: ReDim data(1 To m_cfg.count, 1 To 2)
    Dim i As Long
    For i = 0 To UBound(keys)
        data(i + 1, 1) = CStr(keys(i))
        data(i + 1, 2) = CStr(m_cfg(keys(i)))
    Next i
    ws.Cells(2, 1).Resize(m_cfg.count, 2).value = data
End Sub

Private Sub SaveSourcesSheet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_SOURCES)
    Dim hdr As Object: Set hdr = ReadHeaderMap(ws)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Delete
    If m_sources.count = 0 Then Exit Sub
    Dim srcKeys As Variant: srcKeys = m_sources.keys
    Dim nCols As Long: nCols = hdr.count
    Dim data() As Variant: ReDim data(1 To m_sources.count, 1 To nCols)
    Dim i As Long
    For i = 0 To UBound(srcKeys)
        Dim sd As Object: Set sd = m_sources(srcKeys(i))
        Dim hk As Variant
        For Each hk In hdr.keys
            Dim c As Long: c = hdr(hk)
            If sd.Exists(CStr(hk)) Then data(i + 1, c) = CStr(sd(CStr(hk)))
        Next hk
        data(i + 1, 1) = CStr(srcKeys(i))
    Next i
    ws.Cells(2, 1).Resize(m_sources.count, nCols).value = data
End Sub

Private Sub SaveFieldsSheet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_FIELDS)
    Dim hdr As Object: Set hdr = ReadHeaderMap(ws)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Delete
    If m_fields.count = 0 Then Exit Sub
    Dim fKeys As Variant: fKeys = m_fields.keys
    Dim nCols As Long: nCols = hdr.count
    Dim data() As Variant: ReDim data(1 To m_fields.count, 1 To nCols)
    Dim i As Long
    For i = 0 To UBound(fKeys)
        Dim fd As Object: Set fd = m_fields(fKeys(i))
        Dim hk As Variant
        For Each hk In hdr.keys
            Dim ci As Long: ci = hdr(hk)
            If fd.Exists(CStr(hk)) Then data(i + 1, ci) = CStr(fd(CStr(hk)))
        Next hk
    Next i
    ws.Cells(2, 1).Resize(m_fields.count, nCols).value = data
End Sub

' ##########################################################################
' # SECTION: Change Log (ex-FolioChangeLog)
' ##########################################################################

Private Function GetLogTable() As ListObject
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    If ws Is Nothing Then Exit Function
    Set GetLogTable = ws.ListObjects(LOG_TABLE)
    On Error GoTo 0
End Function

Public Sub EnsureLogSheet()
    Dim eh As New ErrorHandler: eh.Enter "FolioLib", "EnsureLogSheet"
    On Error GoTo ErrHandler
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(LOG_SHEET)
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = LOG_SHEET
        ws.Visible = xlSheetVeryHidden
    End If
    ' Ensure ListObject exists
    If ws.ListObjects.count = 0 Then
        Dim headers As Variant: headers = Array("timestamp", "source", "key", "field", "old_value", "new_value", "origin")
        Dim c As Long
        For c = 0 To 6: ws.Cells(1, c + 1).value = headers(c): Next c
        ws.ListObjects.Add(xlSrcRange, ws.Range("A1:G1"), , xlYes).Name = LOG_TABLE
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Sub AddLogEntry(src As String, key As String, field As String, _
                       oldVal As String, newVal As String, origin As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioLib", "AddLogEntry"
    On Error GoTo ErrHandler
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then EnsureLogSheet: Set tbl = GetLogTable()
    If tbl Is Nothing Then Exit Sub

    RotateIfNeeded tbl, 1
    Dim lr As ListRow: Set lr = tbl.ListRows.Add
    lr.Range(1, 1).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    lr.Range(1, 2).value = src
    lr.Range(1, 3).value = key
    lr.Range(1, 4).value = field
    lr.Range(1, 5).value = oldVal
    lr.Range(1, 6).value = newVal
    lr.Range(1, 7).value = origin
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Sub AddLogEntries(entries As Collection)
    If entries Is Nothing Then Exit Sub
    If entries.count = 0 Then Exit Sub
    On Error GoTo ErrHandler
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then EnsureLogSheet: Set tbl = GetLogTable()
    If tbl Is Nothing Then Exit Sub

    RotateIfNeeded tbl, entries.count

    Dim n As Long: n = entries.count
    Dim data() As Variant: ReDim data(1 To n, 1 To 7)
    Dim ts As String: ts = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Dim i As Long
    For i = 1 To n
        Dim e As Object: Set e = entries(i)
        data(i, 1) = ts
        data(i, 2) = DictStr(e, "type")
        data(i, 3) = DictStr(e, "id")
        Dim act As String: act = DictStr(e, "action")
        If act = "added" Then data(i, 4) = "+" & DictStr(e, "type") _
        Else data(i, 4) = "-" & DictStr(e, "type")
        data(i, 5) = ""
        data(i, 6) = DictStr(e, "description")
        data(i, 7) = "external"
    Next i

    ' Add rows and batch write
    Dim startRow As Long
    If tbl.DataBodyRange Is Nothing Then
        tbl.ListRows.Add
        startRow = 1
    Else
        startRow = tbl.ListRows.count + 1
    End If
    ' Add remaining rows
    For i = 2 To n: tbl.ListRows.Add: Next i
    tbl.DataBodyRange.Rows(startRow).Resize(n, 7).value = data
    Exit Sub
ErrHandler:
End Sub

Private Sub RotateIfNeeded(tbl As ListObject, addCount As Long)
    On Error Resume Next
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim total As Long: total = tbl.ListRows.count + addCount
    If total <= MAX_LOG_ROWS Then Exit Sub
    Dim delCount As Long: delCount = total - MAX_LOG_ROWS
    Dim i As Long
    For i = 1 To delCount: tbl.ListRows(1).Delete: Next i
    On Error GoTo 0
End Sub

Public Function GetRecentEntries(Optional count As Long = 200) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioLib", "GetRecentEntries"
    On Error GoTo ErrHandler
    Set GetRecentEntries = New Collection
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then eh.OK: Exit Function
    If tbl.DataBodyRange Is Nothing Then eh.OK: Exit Function

    Dim rowCount As Long: rowCount = tbl.ListRows.count
    If rowCount = 0 Then eh.OK: Exit Function
    Dim startIdx As Long: startIdx = rowCount - count + 1
    If startIdx < 1 Then startIdx = 1

    Dim r As Long
    For r = rowCount To startIdx Step -1
        Dim rng As Range: Set rng = tbl.ListRows(r).Range
        Dim entry As Object: Set entry = NewDict()
        entry.Add "ts", CStr(rng(1, 1).value)
        entry.Add "src", CStr(rng(1, 2).value)
        entry.Add "key", CStr(rng(1, 3).value)
        entry.Add "field", CStr(rng(1, 4).value)
        entry.Add "old", CStr(rng(1, 5).value)
        entry.Add "new", CStr(rng(1, 6).value)
        entry.Add "origin", CStr(rng(1, 7).value)
        GetRecentEntries.Add entry
    Next r
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Sub ClearLog()
    Dim eh As New ErrorHandler: eh.Enter "FolioLib", "ClearLog"
    On Error GoTo ErrHandler
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Function FormatLogLine(entry As Object) As String
    On Error Resume Next
    Dim ts As String: ts = DictStr(entry, "ts")
    If IsDate(ts) Then ts = Format$(CDate(ts), "hh:nn:ss")

    Dim origin As String: origin = DictStr(entry, "origin")
    Dim key As String: key = DictStr(entry, "key")
    Dim nm As String: nm = DictStr(entry, "name")
    Dim field As String: field = DictStr(entry, "field")
    Dim oldV As String: oldV = DictStr(entry, "old")
    Dim newV As String: newV = DictStr(entry, "new")

    Dim change As String
    If Len(field) > 0 Then change = field & ": "
    If Len(oldV) > 0 Or Len(newV) > 0 Then change = change & oldV & " -> " & newV

    Dim id As String: id = key
    If Len(nm) > 0 And nm <> key Then id = id & " " & nm

    FormatLogLine = ts & "  " & origin & "  " & id & "  " & change
    On Error GoTo 0
End Function
