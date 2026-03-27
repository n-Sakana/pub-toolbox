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
'     ZIP内の vbaProject.bin を展開し、同様に DPB= を書き換える
Private Function RemovePasswordXlsm(ByVal filePath As String) As Boolean
    Dim fso As Object
    Dim shellApp As Object
    Dim tempFolder As String
    Dim zipPath As String
    Dim vbaProjPath As String

    RemovePasswordXlsm = False

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shellApp = CreateObject("Shell.Application")

    ' 一時フォルダ作成
    tempFolder = fso.GetSpecialFolder(2).Path & "\" & _
                 "VBAPwdRemover_" & Format(Now, "yyyymmddhhnnss")
    If Not fso.FolderExists(tempFolder) Then fso.CreateFolder tempFolder

    ' ZIP としてコピー
    zipPath = tempFolder & "\temp.zip"
    fso.CopyFile filePath, zipPath

    ' ZIP を展開
    Dim extractFolder As String
    extractFolder = tempFolder & "\extracted"
    fso.CreateFolder extractFolder

    Dim zipItems As Object
    Set zipItems = shellApp.Namespace(zipPath).Items

    shellApp.Namespace(extractFolder).CopyHere zipItems, 16 + 4  ' 上書き + 非表示

    ' Shell.Application の CopyHere は非同期のため完了を待機
    WaitForExtraction extractFolder, zipItems.Count

    ' vbaProject.bin を検索
    vbaProjPath = FindVbaProjectBin(fso, extractFolder)
    If vbaProjPath = "" Then
        CleanupTempFolder fso, tempFolder
        Exit Function
    End If

    ' vbaProject.bin 内の DPB= を書き換え
    If Not PatchDPBInFile(vbaProjPath) Then
        CleanupTempFolder fso, tempFolder
        Exit Function
    End If

    ' 再ZIP化: 展開フォルダの中身を新しいZIPにまとめる
    Dim newZipPath As String
    newZipPath = tempFolder & "\patched.zip"
    CreateEmptyZip newZipPath

    Dim srcFolder As Object
    Set srcFolder = shellApp.Namespace(extractFolder)

    shellApp.Namespace(newZipPath).CopyHere srcFolder.Items, 16 + 4

    ' ZIP化完了を待機
    WaitForZip newZipPath, fso

    ' 元ファイルを差し替え
    fso.CopyFile newZipPath, filePath, True

    ' 一時フォルダ削除
    CleanupTempFolder fso, tempFolder

    RemovePasswordXlsm = True
End Function

' --- vbaProject.bin をフォルダ内から再帰検索 ---
Private Function FindVbaProjectBin(ByRef fso As Object, ByVal folderPath As String) As String
    Dim fld As Object
    Dim f As Object
    Dim sub_ As Object

    FindVbaProjectBin = ""

    Set fld = fso.GetFolder(folderPath)

    For Each f In fld.Files
        If LCase(f.Name) = "vbaproject.bin" Then
            FindVbaProjectBin = f.Path
            Exit Function
        End If
    Next f

    For Each sub_ In fld.SubFolders
        FindVbaProjectBin = FindVbaProjectBin(fso, sub_.Path)
        If FindVbaProjectBin <> "" Then Exit Function
    Next sub_
End Function

' --- ファイル内の DPB= を DPx= に書き換え ---
Private Function PatchDPBInFile(ByVal filePath As String) As Boolean
    Dim fileNum As Integer
    Dim fileData() As Byte
    Dim fileLen As Long

    PatchDPBInFile = False

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

    Dim targetBytes() As Byte
    targetBytes = StrConv("DPB=", vbFromUnicode)

    Dim pos As Long
    pos = FindBytes(fileData, targetBytes)

    If pos = -1 Then Exit Function

    ' DPB= → DPx=
    fileData(pos + 2) = &H78

    fileNum = FreeFile
    Open filePath For Binary Access Write As #fileNum
    Put #fileNum, , fileData
    Close #fileNum

    PatchDPBInFile = True
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

' --- 空のZIPファイルを作成 ---
Private Sub CreateEmptyZip(ByVal zipPath As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open zipPath For Output As #fileNum
    ' ZIP ファイルの最小ヘッダー (End of Central Directory Record)
    Print #fileNum, Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, Chr(0))
    Close #fileNum
End Sub

' --- Shell.Application の非同期展開を待機 ---
Private Sub WaitForExtraction(ByVal folderPath As String, ByVal expectedCount As Long)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim elapsed As Long
    elapsed = 0
    Do
        Application.Wait Now + TimeSerial(0, 0, 1)
        elapsed = elapsed + 1
        If elapsed > 60 Then Exit Do  ' 最大60秒
        If fso.GetFolder(folderPath).Files.Count + _
           fso.GetFolder(folderPath).SubFolders.Count >= expectedCount Then Exit Do
    Loop
End Sub

' --- ZIP化完了を待機 ---
Private Sub WaitForZip(ByVal zipPath As String, ByRef fso As Object)
    Dim prevSize As Long
    Dim currSize As Long
    Dim stableCount As Long

    prevSize = 0
    stableCount = 0

    Do
        Application.Wait Now + TimeSerial(0, 0, 1)
        currSize = fso.GetFile(zipPath).Size
        If currSize = prevSize And currSize > 22 Then
            stableCount = stableCount + 1
            If stableCount >= 3 Then Exit Do
        Else
            stableCount = 0
        End If
        prevSize = currSize
        If stableCount > 30 Then Exit Do  ' 最大30秒
    Loop
End Sub

' --- 一時フォルダ削除 ---
Private Sub CleanupTempFolder(ByRef fso As Object, ByVal folderPath As String)
    On Error Resume Next
    If fso.FolderExists(folderPath) Then
        fso.DeleteFolder folderPath, True
    End If
    On Error GoTo 0
End Sub
