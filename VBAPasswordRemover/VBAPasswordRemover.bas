'===============================================================================
' VBAPasswordRemover - VBAプロジェクトのパスワード保護を解除
'===============================================================================
' 使い方:
'   1. RemoveVBAPassword マクロを実行
'   2. ファイル選択ダイアログで対象の Excel ファイルを選択
'   3. パスワード保護キーが無効化された状態でファイルが保存される
'   4. 対象ファイルを開き、VBE (Alt+F11) > ツール > VBAProject のプロパティ
'      > 保護タブ で既存パスワードを空欄にして OK
'
' 対応形式:
'   .xls  (OLE2 Compound Document)
'   .xlsm / .xlam (ZIP ベースの OOXML)
'
' 注意:
'   - 自分が管理するファイルのパスワード回復用途を想定しています
'   - 元ファイルのバックアップが同じフォルダに自動作成されます
'
' インポート方法:
'   VBE (Alt+F11) > ファイル > ファイルのインポート > このファイルを選択
'===============================================================================
Option Explicit

' --- エントリポイント ---
Public Sub RemoveVBAPassword()
    Dim filePath As String
    filePath = PickFile()
    If filePath = "" Then Exit Sub

    Dim ext As String
    ext = LCase(Mid(filePath, InStrRev(filePath, ".") + 1))

    ' バックアップ作成
    If Not CreateBackup(filePath) Then
        MsgBox "バックアップの作成に失敗しました。処理を中止します。", vbCritical
        Exit Sub
    End If

    Dim result As Boolean
    Select Case ext
        Case "xls"
            result = RemovePasswordXls(filePath)
        Case "xlsm", "xlam"
            result = RemovePasswordXlsm(filePath)
        Case Else
            MsgBox "対応していない形式です: ." & ext & vbCrLf & _
                   ".xls / .xlsm / .xlam に対応しています。", vbExclamation
            Exit Sub
    End Select

    If result Then
        MsgBox "パスワード保護を無効化しました。" & vbCrLf & vbCrLf & _
               "次の手順で完全に解除してください:" & vbCrLf & _
               "  1. 対象ファイルを開く" & vbCrLf & _
               "  2. VBE (Alt+F11) を開く" & vbCrLf & _
               "  3. ツール > VBAProject のプロパティ > 保護タブ" & vbCrLf & _
               "  4. パスワード欄を空にして OK" & vbCrLf & _
               "  5. ファイルを保存" & vbCrLf & vbCrLf & _
               "バックアップ: " & filePath & ".bak", vbInformation
    Else
        MsgBox "VBAプロジェクトのパスワード情報が見つかりませんでした。" & vbCrLf & _
               "ファイルにVBAプロジェクトが含まれていない可能性があります。", vbExclamation
    End If
End Sub

' --- ファイル選択ダイアログ ---
Private Function PickFile() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "VBAパスワードを解除するファイルを選択"
        .ButtonName = "選択"
        .Filters.Clear
        .Filters.Add "Excel マクロ有効ファイル", "*.xls;*.xlsm;*.xlam"
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFile = .SelectedItems(1)
        Else
            PickFile = ""
        End If
    End With
End Function

' --- バックアップ作成 ---
Private Function CreateBackup(ByVal filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fso.CopyFile filePath, filePath & ".bak", True
    CreateBackup = (Err.Number = 0)
    On Error GoTo 0
End Function

' --- .xls (OLE2) のパスワード解除 ---
'     バイナリ内の "DPB=" を "DPx=" に書き換えることで
'     パスワードハッシュを無効化する
Private Function RemovePasswordXls(ByVal filePath As String) As Boolean
    Dim fileNum As Integer
    Dim fileData() As Byte
    Dim fileLen As Long

    RemovePasswordXls = False

    ' ファイル全体をバイト配列に読み込み
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    fileLen = LOF(fileNum)
    If fileLen = 0 Then
        Close #fileNum
        Exit Function
    End If
    ReDim fileData(0 To fileLen - 1)
    Get #fileNum, , fileData
    Close #fileNum

    ' "DPB=" のバイト列を検索 (ASCII)
    Dim targetBytes() As Byte
    targetBytes = StrConv("DPB=", vbFromUnicode)

    Dim pos As Long
    pos = FindBytes(fileData, targetBytes)

    If pos = -1 Then Exit Function

    ' "DPB=" → "DPx=" に書き換え (3バイト目の 'B' を 'x' に変更)
    ' B=0x42, x=0x78
    fileData(pos + 2) = &H78

    ' 書き戻し
    fileNum = FreeFile
    Open filePath For Binary Access Write As #fileNum
    Put #fileNum, , fileData
    Close #fileNum

    RemovePasswordXls = True
End Function

' --- .xlsm / .xlam (OOXML ZIP) のパスワード解除 ---
'     Excel 自身で一時的に .xls (OLE2) 形式に変換し、
'     バイナリパッチを適用してから .xlsm に戻す
'     外部プロセスを使わないため EDR にブロックされない
Private Function RemovePasswordXlsm(ByVal filePath As String) As Boolean
    Dim fso As Object
    Dim tempXlsPath As String
    Dim wb As Workbook
    Dim origAlerts As Boolean
    Dim origEvents As Boolean
    Dim ext As String

    RemovePasswordXlsm = False

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 一時 .xls ファイルのパス
    tempXlsPath = fso.GetSpecialFolder(2).Path & "\" & _
                  "VBAPwdRemover_" & Format(Now, "yyyymmddhhnnss") & ".xls"

    ' ダイアログ・イベント抑制
    origAlerts = Application.DisplayAlerts
    origEvents = Application.EnableEvents
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrHandler

    ' 対象ファイルを開いて .xls 形式で保存
    Set wb = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=False)
    wb.SaveAs tempXlsPath, xlExcel8  ' .xls (OLE2) 形式
    wb.Close SaveChanges:=False

    ' .xls のバイナリを直接パッチ (OLE2 なので DPB= が平文で存在)
    If Not RemovePasswordXls(tempXlsPath) Then
        ' パッチ失敗 → 一時ファイル削除して終了
        On Error Resume Next
        fso.DeleteFile tempXlsPath, True
        On Error GoTo 0
        GoTo Cleanup
    End If

    ' パッチ済み .xls を開いて元の形式で保存し直す
    Set wb = Workbooks.Open(tempXlsPath, UpdateLinks:=0, ReadOnly:=False)

    ext = LCase(fso.GetExtensionName(filePath))
    If ext = "xlam" Then
        wb.SaveAs filePath, xlOpenXMLAddIn  ' .xlam
    Else
        wb.SaveAs filePath, xlOpenXMLWorkbookMacroEnabled  ' .xlsm
    End If
    wb.Close SaveChanges:=False

    ' 一時ファイル削除
    On Error Resume Next
    fso.DeleteFile tempXlsPath, True
    On Error GoTo 0

    RemovePasswordXlsm = True
    GoTo Cleanup

ErrHandler:
    ' エラー時: ブックが開いていたら閉じる
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    fso.DeleteFile tempXlsPath, True
    On Error GoTo 0

Cleanup:
    Application.DisplayAlerts = origAlerts
    Application.EnableEvents = origEvents
End Function

' --- バイト配列内でパターンを検索 ---
Private Function FindBytes(ByRef data() As Byte, ByRef pattern() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim patLen As Long
    Dim dataLen As Long

    FindBytes = -1
    patLen = UBound(pattern) - LBound(pattern) + 1
    dataLen = UBound(data) - LBound(data) + 1

    If patLen > dataLen Then Exit Function

    For i = 0 To dataLen - patLen
        Dim matched As Boolean
        matched = True
        For j = 0 To patLen - 1
            If data(i + j) <> pattern(j) Then
                matched = False
                Exit For
            End If
        Next j
        If matched Then
            FindBytes = i
            Exit Function
        End If
    Next i
End Function
