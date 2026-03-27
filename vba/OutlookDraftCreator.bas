'===============================================================================
' OutlookDraftCreator - Excelシートから一括でOutlook下書きメールを作成
'===============================================================================
' 使い方:
'   1. 「入力」シートに以下のヘッダー付きでデータを入力 (2行目からデータ)
'      A列: 送信先メールアドレス (必須)
'      B列: CC
'      C列: BCC
'      D列: メール件名
'      E列: メール本文
'      F列: 添付物パス (ファイルまたはフォルダ。セミコロン区切りで複数指定可)
'      G列: 送信元メールアドレス
'      H列: 結果 (マクロが書き込む)
'   2. CreateDrafts マクロを実行
'   3. Outlookの下書きフォルダにメールが作成される
'   4. H列に処理結果 (OK / エラー内容) が出力される
'
' 特徴:
'   - 参照設定不要 (CreateObject による遅延バインディング)
'   - 秒数ベースの DoEvents でフリーズ対策
'   - 1行あたり30秒のタイムアウトで停滞防止
'   - フォルダ指定で中身のファイルを一括添付
'
' インポート方法:
'   VBE (Alt+F11) > ファイル > ファイルのインポート > このファイルを選択
'===============================================================================
Option Explicit

' --- 定数 ---
Private Const INPUT_SHEET       As String = "入力"
Private Const COL_TO            As Long = 1   ' A列
Private Const COL_CC            As Long = 2   ' B列
Private Const COL_BCC           As Long = 3   ' C列
Private Const COL_SUBJECT       As Long = 4   ' D列
Private Const COL_BODY          As Long = 5   ' E列
Private Const COL_ATTACHMENT    As Long = 6   ' F列
Private Const COL_FROM          As Long = 7   ' G列
Private Const COL_RESULT        As Long = 8   ' H列
Private Const DOEVENTS_SEC      As Double = 0.5  ' DoEvents実行間隔 (秒)
Private Const TIMEOUT_SEC       As Long = 30     ' 1行あたりタイムアウト (秒)

' --- エントリポイント ---
Public Sub CreateDrafts()
    Dim ws          As Worksheet
    Dim lastRow     As Long
    Dim rowCount    As Long
    Dim data()      As Variant
    Dim results()   As Variant
    Dim olApp       As Object
    Dim t           As Double
    Dim lastDoEvent As Double
    Dim i           As Long
    Dim successCnt  As Long
    Dim failCnt     As Long
    Dim skipCnt     As Long
    Dim timeoutCnt  As Long

    ' --- 入力シート取得 ---
    Set ws = GetInputSheet()
    If ws Is Nothing Then Exit Sub

    ' --- データ範囲検出 ---
    lastRow = GetLastDataRow(ws)
    If lastRow < 2 Then
        MsgBox "「" & INPUT_SHEET & "」シートにデータがありません。" & vbCrLf & _
               "2行目以降にデータを入力してください。", vbExclamation
        Exit Sub
    End If

    rowCount = lastRow - 1

    ' --- データ一括読込 ---
    data = ws.Range("A2").Resize(rowCount, COL_RESULT).Value
    ReDim results(1 To rowCount, 1 To 1)

    ' --- Outlook起動 ---
    On Error Resume Next
    Set olApp = CreateObject("Outlook.Application")
    If Err.Number <> 0 Then
        MsgBox "Outlookの起動に失敗しました。" & vbCrLf & _
               "Outlookがインストールされているか確認してください。" & vbCrLf & _
               Err.Description, vbCritical
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' --- パフォーマンス設定 ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    t = Timer
    lastDoEvent = Timer
    successCnt = 0
    failCnt = 0
    skipCnt = 0
    timeoutCnt = 0

    ' --- メインループ ---
    For i = 1 To rowCount
        Application.StatusBar = "下書き作成中... (" & i & "/" & rowCount & ")"

        ' 秒数ベース DoEvents
        If Timer - lastDoEvent >= DOEVENTS_SEC Then
            DoEvents
            lastDoEvent = Timer
        End If

        ' 空行スキップ
        If IsBlankRow(data, i) Then
            results(i, 1) = "スキップ (空行)"
            skipCnt = skipCnt + 1
        Else
            results(i, 1) = ProcessRow(olApp, data, i)
            Select Case results(i, 1)
                Case "OK"
                    successCnt = successCnt + 1
                Case "タイムアウト"
                    timeoutCnt = timeoutCnt + 1
                Case Else
                    failCnt = failCnt + 1
            End Select
        End If
    Next i

    ' --- 結果一括書込 ---
    ws.Range("H2").Resize(rowCount, 1).Value = results

    ' --- 後処理 ---
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False

    ShowSummary rowCount, successCnt, failCnt, skipCnt, timeoutCnt, Timer - t

    Set olApp = Nothing
End Sub

' --- 入力シートの取得・検証 ---
Private Function GetInputSheet() As Worksheet
    On Error Resume Next
    Set GetInputSheet = ThisWorkbook.Worksheets(INPUT_SHEET)
    On Error GoTo 0

    If GetInputSheet Is Nothing Then
        MsgBox "「" & INPUT_SHEET & "」シートが見つかりません。" & vbCrLf & _
               "A列に送信先メールアドレス等のヘッダーを持つシートを作成してください。", vbCritical
    End If
End Function

' --- データ最終行の検出 ---
Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    GetLastDataRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
End Function

' --- 1行分の下書き作成 ---
Private Function ProcessRow( _
    ByRef olApp As Object, _
    ByRef data() As Variant, _
    ByVal rowIdx As Long) As String

    Dim mail    As Object
    Dim toAddr  As String
    Dim rowStart As Double

    rowStart = Timer

    ' --- 送信先アドレス取得・検証 ---
    toAddr = Trim(CStr(data(rowIdx, COL_TO)))
    If Not IsValidEmail(toAddr) Then
        ProcessRow = "エラー: 送信先アドレスが無効 (" & toAddr & ")"
        Exit Function
    End If

    ' --- CC/BCC検証 ---
    Dim ccAddr As String, bccAddr As String
    ccAddr = Trim(CStr(data(rowIdx, COL_CC) & ""))
    bccAddr = Trim(CStr(data(rowIdx, COL_BCC) & ""))

    If ccAddr <> "" And Not IsValidEmail(ccAddr) Then
        ProcessRow = "エラー: CCアドレスが無効 (" & ccAddr & ")"
        Exit Function
    End If
    If bccAddr <> "" And Not IsValidEmail(bccAddr) Then
        ProcessRow = "エラー: BCCアドレスが無効 (" & bccAddr & ")"
        Exit Function
    End If

    ' --- タイムアウトチェック ---
    If Timer - rowStart > TIMEOUT_SEC Then
        ProcessRow = "タイムアウト"
        Exit Function
    End If

    ' --- MailItem作成 ---
    On Error Resume Next
    Set mail = olApp.CreateItem(0)  ' olMailItem = 0
    If Err.Number <> 0 Then
        ProcessRow = "エラー: メール作成失敗 - " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' --- プロパティ設定 ---
    On Error Resume Next
    mail.To = toAddr
    If ccAddr <> "" Then mail.CC = ccAddr
    If bccAddr <> "" Then mail.BCC = bccAddr
    mail.Subject = CStr(data(rowIdx, COL_SUBJECT) & "")
    mail.Body = CStr(data(rowIdx, COL_BODY) & "")

    ' 送信元アドレス
    Dim fromAddr As String
    fromAddr = Trim(CStr(data(rowIdx, COL_FROM) & ""))
    If fromAddr <> "" Then
        mail.SentOnBehalfOfName = fromAddr
    End If

    If Err.Number <> 0 Then
        ProcessRow = "エラー: プロパティ設定失敗 - " & Err.Description
        Err.Clear
        On Error GoTo 0
        Set mail = Nothing
        Exit Function
    End If
    On Error GoTo 0

    ' --- タイムアウトチェック ---
    If Timer - rowStart > TIMEOUT_SEC Then
        Set mail = Nothing
        ProcessRow = "タイムアウト"
        Exit Function
    End If

    ' --- 添付物処理 ---
    Dim attachPath As String
    attachPath = Trim(CStr(data(rowIdx, COL_ATTACHMENT) & ""))
    If attachPath <> "" Then
        Dim attachResult As String
        attachResult = AddAttachments(mail, attachPath, rowStart)
        If attachResult <> "" Then
            Set mail = Nothing
            ProcessRow = attachResult
            Exit Function
        End If
    End If

    ' --- 下書き保存 ---
    On Error Resume Next
    mail.Save
    If Err.Number <> 0 Then
        ProcessRow = "エラー: 下書き保存失敗 - " & Err.Description
        Err.Clear
        On Error GoTo 0
        Set mail = Nothing
        Exit Function
    End If
    On Error GoTo 0

    Set mail = Nothing
    ProcessRow = "OK"
End Function

' --- メールアドレス簡易バリデーション ---
' セミコロン区切りの複数アドレスにも対応
Private Function IsValidEmail(ByVal addrList As String) As Boolean
    Dim parts()  As String
    Dim addr     As String
    Dim atPos    As Long
    Dim dotPos   As Long
    Dim i        As Long

    IsValidEmail = False

    If Trim(addrList) = "" Then Exit Function

    parts = Split(addrList, ";")
    For i = LBound(parts) To UBound(parts)
        addr = Trim(parts(i))
        If addr = "" Then GoTo NextAddr

        ' スペースを含む場合は無効
        If InStr(addr, " ") > 0 Then Exit Function

        ' @ が1つだけ含まれること
        atPos = InStr(addr, "@")
        If atPos = 0 Then Exit Function
        If InStr(atPos + 1, addr, "@") > 0 Then Exit Function

        ' @ の前後に文字があること
        If atPos <= 1 Then Exit Function
        If atPos >= Len(addr) Then Exit Function

        ' @ の後にドットがあること
        dotPos = InStr(atPos + 1, addr, ".")
        If dotPos = 0 Then Exit Function
        If dotPos >= Len(addr) Then Exit Function
NextAddr:
    Next i

    IsValidEmail = True
End Function

' --- 添付物処理 ---
' セミコロン区切りで複数パス対応、フォルダ指定時は中身のファイルを全添付
' 戻り値: 空文字=成功、それ以外=エラーメッセージ
Private Function AddAttachments( _
    ByRef mail As Object, _
    ByVal attachPath As String, _
    ByVal rowStart As Double) As String

    Dim fso     As Object
    Dim paths() As String
    Dim p       As String
    Dim i       As Long
    Dim f       As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    paths = Split(attachPath, ";")

    For i = LBound(paths) To UBound(paths)
        p = Trim(paths(i))
        If p = "" Then GoTo NextPath

        ' タイムアウトチェック
        If Timer - rowStart > TIMEOUT_SEC Then
            AddAttachments = "タイムアウト"
            Exit Function
        End If

        If fso.FileExists(p) Then
            ' ファイル添付
            On Error Resume Next
            mail.Attachments.Add p
            If Err.Number <> 0 Then
                AddAttachments = "エラー: 添付失敗 (" & p & ") - " & Err.Description
                Err.Clear
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0

        ElseIf fso.FolderExists(p) Then
            ' フォルダ内の全ファイルを添付
            For Each f In fso.GetFolder(p).files
                ' タイムアウトチェック
                If Timer - rowStart > TIMEOUT_SEC Then
                    AddAttachments = "タイムアウト"
                    Exit Function
                End If

                On Error Resume Next
                mail.Attachments.Add f.Path
                If Err.Number <> 0 Then
                    AddAttachments = "エラー: 添付失敗 (" & f.Path & ") - " & Err.Description
                    Err.Clear
                    On Error GoTo 0
                    Exit Function
                End If
                On Error GoTo 0
            Next f
        Else
            AddAttachments = "エラー: パスが存在しません (" & p & ")"
            Exit Function
        End If
NextPath:
    Next i

    AddAttachments = ""
End Function

' --- 完了サマリー表示 ---
Private Sub ShowSummary( _
    ByVal totalRows As Long, _
    ByVal successCnt As Long, _
    ByVal failCnt As Long, _
    ByVal skipCnt As Long, _
    ByVal timeoutCnt As Long, _
    ByVal elapsed As Double)

    Dim msg As String
    msg = "完了: " & successCnt & " 件の下書きを作成 (" & Format(elapsed, "0.0") & " 秒)" & vbCrLf & _
          "対象行数: " & totalRows & " 件" & vbCrLf & _
          "成功: " & successCnt & " 件"

    If failCnt > 0 Then msg = msg & vbCrLf & "失敗: " & failCnt & " 件"
    If timeoutCnt > 0 Then msg = msg & vbCrLf & "タイムアウト: " & timeoutCnt & " 件"
    If skipCnt > 0 Then msg = msg & vbCrLf & "スキップ (空行): " & skipCnt & " 件"

    If failCnt > 0 Or timeoutCnt > 0 Then
        msg = msg & vbCrLf & vbCrLf & "※ H列で詳細を確認してください。"
    End If

    MsgBox msg, vbInformation
End Sub

' --- 空行判定 ---
Private Function IsBlankRow(ByRef data() As Variant, ByVal rowIdx As Long) As Boolean
    IsBlankRow = (Trim(CStr(data(rowIdx, COL_TO) & "")) = "")
End Function
