'===============================================================================
' ExcelDataExtractor - ADO版: 複数Excelファイルからセルデータを一括抽出
'===============================================================================
' 使い方:
'   1. 「設定」シートの A2:B列 にフィールド名とセル番地を入力
'      例) A2="会社名", B2="B3"
'           A3="売上",   B3="D10"
'   2. ExtractData マクロを実行
'   3. フォルダ選択ダイアログで対象フォルダを選択
'   4. 「結果」シートに抽出結果が出力される
'
' 特徴:
'   - ADO経由でファイルを読むため Workbooks.Open 不要 → 高速
'   - 参照設定不要 (CreateObject による遅延バインディング)
'
' インポート方法:
'   VBE (Alt+F11) > ファイル > ファイルのインポート > このファイルを選択
'===============================================================================
Option Explicit

' --- 定数 ---
Private Const SETTING_SHEET As String = "設定"
Private Const RESULT_SHEET  As String = "結果"
Private Const FIXED_COLS    As Long = 3  ' ディレクトリ / ファイル名 / シート名

' --- エントリポイント ---
Public Sub ExtractData()
    Dim settings()  As Variant
    Dim settingCnt  As Long
    Dim folderPath  As String
    Dim files       As Collection
    Dim allRows     As Collection  ' 各要素は Variant配列(1 To totalCols)
    Dim errFiles    As Collection
    Dim t           As Double

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    t = Timer

    ' --- 設定読み込み ---
    If Not ReadSettings(settings, settingCnt) Then GoTo Cleanup

    ' --- フォルダ選択 ---
    folderPath = PickFolder()
    If folderPath = "" Then GoTo Cleanup

    ' --- ファイル列挙 (サブフォルダ含む) ---
    Set files = New Collection
    CollectExcelFiles folderPath, files
    If files.Count = 0 Then
        MsgBox "対象のExcelファイルが見つかりませんでした。", vbExclamation
        GoTo Cleanup
    End If

    ' --- ADOで各ファイルからデータ抽出 ---
    Set allRows = New Collection
    Set errFiles = New Collection

    Dim i As Long
    For i = 1 To files.Count
        Application.StatusBar = "処理中... (" & i & "/" & files.Count & ") " & Mid(files(i), InStrRev(files(i), "\") + 1)
        DoEvents
        ExtractFromFileADO files(i), settings, settingCnt, allRows, errFiles
    Next i

    ' --- 結果出力 ---
    WriteResults allRows, settings, settingCnt

    Application.StatusBar = False

    Dim msg As String
    msg = "完了: " & allRows.Count & " 件抽出 (" & Format(Timer - t, "0.0") & " 秒)" & vbCrLf & _
          "対象ファイル: " & files.Count & " 件"
    If errFiles.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & "読取失敗: " & errFiles.Count & " 件"
        Dim e As Long
        For e = 1 To Application.WorksheetFunction.Min(errFiles.Count, 10)
            msg = msg & vbCrLf & "  " & errFiles(e)
        Next e
        If errFiles.Count > 10 Then msg = msg & vbCrLf & "  ... 他 " & (errFiles.Count - 10) & " 件"
    End If
    MsgBox msg, vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub

' --- 設定シートからフィールド名・セル番地を配列に読み込む ---
Private Function ReadSettings(ByRef settings() As Variant, ByRef cnt As Long) As Boolean
    Dim ws As Worksheet
    ReadSettings = False

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTING_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "「" & SETTING_SHEET & "」シートが見つかりません。" & vbCrLf & _
               "A列にフィールド名、B列にセル番地を入力してください。", vbCritical
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "設定が空です。A2以降にフィールド名とセル番地を入力してください。", vbExclamation
        Exit Function
    End If

    cnt = lastRow - 1
    Dim raw As Variant
    raw = ws.Range("A2:B" & lastRow).Value

    ' 1行だけの場合は2次元配列にならないので補正
    If cnt = 1 Then
        Dim tmp As Variant
        tmp = raw
        ReDim raw(1 To 1, 1 To 2)
        raw(1, 1) = tmp(1, 1)
        raw(1, 2) = tmp(1, 2)
    End If

    ReDim settings(1 To cnt, 1 To 2)
    Dim i As Long
    For i = 1 To cnt
        If Trim(CStr(raw(i, 1))) = "" Or Trim(CStr(raw(i, 2))) = "" Then
            MsgBox "設定の " & (i + 1) & " 行目が不完全です。" & vbCrLf & _
                   "フィールド名とセル番地の両方を入力してください。", vbExclamation
            Exit Function
        End If
        If Not IsValidCellAddress(Trim(CStr(raw(i, 2)))) Then
            MsgBox "設定の " & (i + 1) & " 行目のセル番地「" & raw(i, 2) & "」が無効です。", vbExclamation
            Exit Function
        End If
        settings(i, 1) = Trim(CStr(raw(i, 1)))
        settings(i, 2) = Trim(CStr(raw(i, 2)))
    Next i

    ReadSettings = True
End Function

' --- セル番地の簡易バリデーション ---
Private Function IsValidCellAddress(ByVal addr As String) As Boolean
    On Error Resume Next
    Dim rng As Range
    Set rng = Range(addr)
    IsValidCellAddress = (Not rng Is Nothing)
    On Error GoTo 0
End Function

' --- フォルダ選択ダイアログ ---
Private Function PickFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "データ抽出対象のフォルダを選択"
        .ButtonName = "選択"
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
            If Right(PickFolder, 1) <> "\" Then PickFolder = PickFolder & "\"
        Else
            PickFolder = ""
        End If
    End With
End Function

' --- サブフォルダを含むExcelファイル列挙 (再帰) ---
Private Sub CollectExcelFiles(ByVal folderPath As String, ByRef files As Collection)
    Dim fso  As Object
    Dim fld  As Object
    Dim sub_ As Object
    Dim f    As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set fld = fso.GetFolder(folderPath)

    For Each f In fld.files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(f.Name))
        If (ext = "xlsx" Or ext = "xlsm" Or ext = "xls" Or ext = "xlsb") Then
            If Left(f.Name, 1) <> "~" And f.Path <> ThisWorkbook.FullName Then
                files.Add f.Path
            End If
        End If
    Next f

    For Each sub_ In fld.SubFolders
        CollectExcelFiles sub_.Path & "\", files
    Next sub_
End Sub

' --- 拡張子に応じたADO接続文字列を返す ---
Private Function GetConnectionString(ByVal filePath As String) As String
    Dim ext As String
    ext = LCase(Mid(filePath, InStrRev(filePath, ".") + 1))

    Select Case ext
        Case "xlsx", "xlsm", "xlsb"
            GetConnectionString = _
                "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & filePath & ";" & _
                "Extended Properties=""Excel 12.0;HDR=NO;IMEX=1;"";"
        Case "xls"
            GetConnectionString = _
                "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & filePath & ";" & _
                "Extended Properties=""Excel 8.0;HDR=NO;IMEX=1;"";"
    End Select
End Function

' --- ADO OpenSchemaからシート名を取得し、クリーンな名前を返す ---
Private Function GetSheetNames(ByRef conn As Object) As Collection
    Dim schema As Object
    Dim col    As New Collection

    Set schema = conn.OpenSchema(20) ' adSchemaTables = 20
    Do While Not schema.EOF
        Dim tblName As String
        tblName = schema.Fields("TABLE_NAME").Value

        ' シート名の形式:
        '   通常:      Sheet1$
        '   空白含む:  'Sheet 1$'
        '   引用符含む: 'Sheet''s$'
        Dim cleanName As String
        cleanName = ""

        If Right(tblName, 1) = "$" Then
            ' Sheet1$ → Sheet1
            cleanName = Left(tblName, Len(tblName) - 1)
        ElseIf Right(tblName, 2) = "$'" Then
            ' 'Sheet 1$' → Sheet 1
            cleanName = Mid(tblName, 2, Len(tblName) - 3)
            cleanName = Replace(cleanName, "''", "'")  ' エスケープ解除
        End If

        If cleanName <> "" Then col.Add cleanName
        schema.MoveNext
    Loop
    schema.Close

    Set GetSheetNames = col
End Function

' --- シート名をSQL用にエスケープ ---
Private Function EscapeSheetName(ByVal name As String) As String
    ' シート名に ' が含まれる場合は '' にエスケープ
    EscapeSheetName = Replace(name, "'", "''")
End Function

' --- ADO版: 1ファイルからデータ抽出 ---
Private Sub ExtractFromFileADO( _
    ByVal filePath As String, _
    ByRef settings() As Variant, _
    ByVal settingCnt As Long, _
    ByRef allRows As Collection, _
    ByRef errFiles As Collection)

    Dim conn        As Object
    Dim rs          As Object
    Dim sheetNames  As Collection
    Dim fso         As Object
    Dim dirName     As String
    Dim fileName    As String
    Dim totalCols   As Long

    totalCols = FIXED_COLS + settingCnt

    Set fso = CreateObject("Scripting.FileSystemObject")
    dirName = fso.GetParentFolderName(filePath)
    fileName = fso.GetFileName(filePath)

    ' --- ADO接続 ---
    Set conn = CreateObject("ADODB.Connection")
    conn.CursorLocation = 3  ' adUseClient
    On Error Resume Next
    conn.Open GetConnectionString(filePath)
    If Err.Number <> 0 Then
        errFiles.Add fileName & " | 接続失敗: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' --- シート名一覧を取得 ---
    On Error Resume Next
    Set sheetNames = GetSheetNames(conn)
    If Err.Number <> 0 Then
        errFiles.Add fileName & " | シート取得失敗: " & Err.Description
        Err.Clear
        conn.Close
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If sheetNames.Count = 0 Then
        errFiles.Add fileName & " | シート0件"
        conn.Close
        Exit Sub
    End If

    ' --- 各シート × 各セルをクエリで取得 ---
    Dim sheetName As Variant
    Dim j As Long
    For Each sheetName In sheetNames
        Dim row() As Variant
        ReDim row(1 To totalCols)

        row(1) = dirName
        row(2) = fileName
        row(3) = CStr(sheetName)

        For j = 1 To settingCnt
            Dim cellAddr As String
            Dim sql      As String

            cellAddr = settings(j, 2)
            ' [Sheet1$B3:B3] or ['Sheet Name'$B3:B3]
            sql = "SELECT F1 FROM [" & EscapeSheetName(CStr(sheetName)) & "$" & cellAddr & ":" & cellAddr & "]"

            On Error Resume Next
            Set rs = CreateObject("ADODB.Recordset")
            rs.Open sql, conn, 0, 1  ' adOpenForwardOnly, adLockReadOnly
            If Err.Number <> 0 Then
                row(FIXED_COLS + j) = "#ERROR: " & Err.Description
                Err.Clear
                Set rs = Nothing
            Else
                If Not rs.EOF Then
                    Dim v As Variant
                    v = rs.Fields(0).Value
                    If IsNull(v) Then
                        row(FIXED_COLS + j) = Empty
                    Else
                        row(FIXED_COLS + j) = v
                    End If
                Else
                    row(FIXED_COLS + j) = Empty
                End If
                rs.Close
                Set rs = Nothing
            End If
            On Error GoTo 0
        Next j

        allRows.Add row
    Next sheetName

    conn.Close
    Set conn = Nothing
End Sub

' --- 結果を「結果」シートに出力 ---
Private Sub WriteResults( _
    ByRef allRows As Collection, _
    ByRef settings() As Variant, _
    ByVal settingCnt As Long)

    Dim ws As Worksheet
    Dim totalCols As Long
    Dim rowCount As Long
    totalCols = FIXED_COLS + settingCnt
    rowCount = allRows.Count

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RESULT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = RESULT_SHEET
    Else
        ws.Cells.Clear
    End If

    ' --- ヘッダー ---
    Dim headers() As Variant
    ReDim headers(1 To 1, 1 To totalCols)
    headers(1, 1) = "ディレクトリ"
    headers(1, 2) = "ファイル名"
    headers(1, 3) = "シート名"
    Dim k As Long
    For k = 1 To settingCnt
        headers(1, FIXED_COLS + k) = settings(k, 1) & " (" & settings(k, 2) & ")"
    Next k
    ws.Range("A1").Resize(1, totalCols).Value = headers

    With ws.Range("A1").Resize(1, totalCols)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' --- データ (Collection → 2次元配列 → 一括書き込み) ---
    If rowCount > 0 Then
        Dim output() As Variant
        ReDim output(1 To rowCount, 1 To totalCols)
        Dim r As Long, c As Long
        For r = 1 To rowCount
            Dim rowData As Variant
            rowData = allRows(r)
            For c = 1 To totalCols
                output(r, c) = rowData(c)
            Next c
        Next r
        ws.Range("A2").Resize(rowCount, totalCols).Value = output
    End If

    ' --- テーブル化 ---
    Dim dataRows As Long
    dataRows = rowCount
    If dataRows = 0 Then dataRows = 1  ' テーブルは最低1データ行必要
    Dim tblRange As Range
    Set tblRange = ws.Range("A1").Resize(dataRows + 1, totalCols)

    On Error Resume Next
    Dim lo As Object
    For Each lo In ws.ListObjects
        lo.Delete
    Next lo
    On Error GoTo 0

    ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes).Name = "抽出結果"

    ws.Columns.AutoFit
    ws.Activate
End Sub
