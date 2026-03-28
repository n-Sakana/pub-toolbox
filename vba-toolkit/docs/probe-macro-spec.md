# 調査マクロ仕様書 (Environment Probe)

## 概要

新環境の基盤能力を実測する VBA マクロ。対象ブックを直接操作するのではなく、「新環境で何が使えるか」「どこで落ちるか」を先に測る。

## 基本方針

- **副作用なし**: 作って消す、下書きまでで送信しない、読み取り中心、接続しても更新しない
- **CSV 出力**: 縦持ちで固定列。項目追加で構造が崩れない
- **外部設定対応**: 確認先 URL、共有パス、プリンタ名等は設定シートで指定
- **単一 xlsm**: マクロ本体 + 設定シート + 結果シートを1ファイルに収める

## 出力形式

### CSV（主出力）

| 列 | 内容 |
|---|---|
| Category | テストカテゴリ |
| TestName | テスト名 |
| Target | テスト対象（ProgID、パス、URL 等） |
| Result | OK / FAIL / SKIP / WARN |
| ErrorNumber | エラー番号（正常時は空） |
| ErrorMessage | エラーメッセージ（正常時は空） |
| ElapsedMs | 所要時間（ミリ秒） |
| Detail | 補足情報（バージョン、パス、値等） |

### 結果シート

CSV と同じ内容を Excel シートにも出力。色分け: OK=緑, FAIL=赤, WARN=黄, SKIP=灰

## テスト項目

### 1. 端末情報 (SystemInfo)

| テスト名 | 取得方法 | Detail に出力 |
|---|---|---|
| ComputerName | Environ$("COMPUTERNAME") | 端末名 |
| UserName | Environ$("USERNAME") | ユーザー名 |
| OS | Environ$("OS") | OS 名 |
| OfficeVersion | Application.Version | バージョン番号 |
| OfficeBuild | Application.Build | ビルド番号 |
| OfficeBitness | #If Win64 → "64-bit" Else "32-bit" | ビット数 |
| VBAVersion | #If VBA7 → "VBA7" Else "VBA6" | VBA バージョン |
| ExcelPath | Application.Path | Excel インストールパス |
| TempFolder | Environ$("TEMP") | 一時フォルダパス |
| UserProfile | Environ$("USERPROFILE") | ユーザープロファイルパス |

### 2. VBA 実行基盤 (VBARuntime)

| テスト名 | テスト内容 | 判定 |
|---|---|---|
| StandardModule | 標準モジュールの Sub 実行 | OK/FAIL |
| ClassCreate | クラスインスタンス生成 | OK/FAIL |
| CollectionUse | Collection 操作 | OK/FAIL |
| DictionaryCreate | CreateObject("Scripting.Dictionary") | OK/FAIL |
| ErrorHandling | On Error GoTo の動作 | OK/FAIL |
| UserFormLoad | UserForm の Load（表示せず） | OK/FAIL |
| TimerFunction | Timer 関数の動作 | OK/FAIL |

### 3. VBIDE アクセス (VBIDEAccess)

| テスト名 | テスト内容 | 判定 |
|---|---|---|
| VBProjectAccess | ThisWorkbook.VBProject にアクセス | OK/FAIL |
| VBComponentEnum | VBComponents の列挙 | OK/FAIL |
| TrustAccessSetting | 上記が動くかで判定 | OK(信頼する=ON) / FAIL(OFF) |

### 4. 参照ライブラリ (References)

| テスト名 | テスト内容 | 判定 |
|---|---|---|
| ReferenceList | 全参照の列挙 | OK (Detail に一覧) |
| MissingReferences | IsBroken=True の参照 | OK(なし) / FAIL(あり) |
| Scripting | Microsoft Scripting Runtime | OK/FAIL |
| DAO | Microsoft DAO | OK/FAIL/WARN |
| ADO | Microsoft ActiveX Data Objects | OK/FAIL |
| MSForms | Microsoft Forms 2.0 | OK/FAIL |
| Outlook | Microsoft Outlook Object Library | OK/FAIL |
| Word | Microsoft Word Object Library | OK/FAIL |

### 5. COM 生成テスト (COMCreate)

各 ProgID に対して `CreateObject` を実行。成功=OK、失敗=FAIL + エラーメッセージ。

| Target (ProgID) |
|---|
| Scripting.FileSystemObject |
| Scripting.Dictionary |
| ADODB.Connection |
| ADODB.Recordset |
| Outlook.Application |
| Word.Application |
| WScript.Shell |
| MSXML2.XMLHTTP.6.0 |
| WinHttp.WinHttpRequest.5.1 |
| Shell.Application |
| MSXML2.DOMDocument.6.0 |

### 6. ActiveX / フォーム (ActiveXTest)

| テスト名 | テスト内容 |
|---|---|
| MSFormsUserForm | UserForm 生成 |
| TreeView | MSComctlLib.TreeView 生成 |
| ListView | MSComctlLib.ListView 生成 |
| Calendar | MSCAL.Calendar 生成 |
| WebBrowser | WebBrowser コントロール |
| CommonDialog | MSComDlg.CommonDialog 生成 |

### 7. ファイルアクセス (FileAccess)

| テスト名 | テスト内容 |
|---|---|
| TempWrite | 一時フォルダにテキスト書き込み |
| TempDelete | 書き込んだファイルの削除 |
| TempOverwrite | 上書き |
| UTF8Write | ADODB.Stream で UTF-8 テキスト出力 |
| CSVWrite | CSV ファイル出力 |
| SharedFolderExists | 設定シートの共有パスの存在確認（Dir） |
| SharedFolderRead | 共有フォルダからの読み取り |
| SharedFolderWrite | 共有フォルダへの書き込み |
| OneDriveSyncPath | OneDrive 同期フォルダの場所確認 |
| OneDriveReadWrite | 同期フォルダへの読み書き |
| ThisWorkbookPath | ThisWorkbook.Path の値を取得 |
| ExternalFileOpen | 同じフォルダのダミーファイルを開けるか |
| DesktopAppCheck | Application.OperatingSystem で判定（ブラウザ起点でないか） |

### 8. 通信 (NetworkTest)

設定シートの URL リストを使用。

| テスト名 | テスト内容 |
|---|---|
| HTTP_GET | HTTP GET 応答確認 |
| HTTPS_GET | HTTPS GET 応答確認 |
| DNS_Resolve | ホスト名の名前解決 |
| ProxyDetect | プロキシ設定の有無 |

### 9. Office 連携 (OfficeInterop)

| テスト名 | テスト内容 |
|---|---|
| OutlookDraft | Outlook で下書き作成（送信しない）→ 削除 |
| WordLaunch | Word.Application 起動 → 終了 |
| ExcelCOM | 別 Excel インスタンス起動 → 終了 |

### 10. 外部実行 (ExternalExec)

| テスト名 | テスト内容 |
|---|---|
| PowerShell | Shell "powershell -Command echo test" |
| CMD | Shell "cmd /c echo test" |
| WScript | CreateObject("WScript.Shell").Run |

### 11. 印刷 / PDF (PrintTest)

| テスト名 | テスト内容 |
|---|---|
| DefaultPrinter | Application.ActivePrinter の取得 |
| PrinterList | Win32_Printer の列挙（WMI が使えれば） |
| PDFExport | ExportAsFixedFormat で PDF 出力 → 削除 |

### 12. DB 接続 (DBConnect)

| テスト名 | テスト内容 |
|---|---|
| DAOEngine | DBEngine オブジェクト生成 |
| ADOConnection | ADODB.Connection オブジェクト生成 |
| ACEProvider | Microsoft.ACE.OLEDB.12.0 の存在確認 |
| JetProvider | Microsoft.Jet.OLEDB.4.0 の存在確認 |

### 13. セキュリティ制約 (SecurityCheck)

| テスト名 | テスト内容 |
|---|---|
| MacroSecurity | Application.AutomationSecurity の値 |
| ProtectedView | Application.ProtectedViewWindows の挙動 |
| ActiveXSecurity | ActiveX 制限の有無 |

## 設定シート (_probe_config)

| A列 | B列 |
|---|---|
| SharedFolder1 | \\\\server\\share |
| SharedFolder2 | |
| TestURL1 | https://www.google.com |
| TestURL2 | |
| OutputFolder | (空=ThisWorkbook と同じフォルダ) |
| DummyFileName | _probe_test.txt |

## 実行フロー

```
1. 設定シート読み込み
2. 結果シートクリア
3. 各テストカテゴリを順番に実行
   - 各テストは On Error Resume Next で保護
   - Timer で所要時間計測
   - 結果を結果シートに即時出力
4. CSV ファイルに結果を書き出し
5. 結果シートを整形（色分け）
6. 完了メッセージ
```

## ファイル構成

```
probe.xlsm
├── Module: ProbeMain       エントリポイント (Probe_Run)
├── Module: ProbeTests      各テストカテゴリの実装
├── Module: ProbeOutput     CSV/シート出力
├── Sheet: _probe_config    設定シート
└── Sheet: _probe_result    結果シート
```

## ビルド方法

vba-toolkit の Build-Probe.ps1 (新規) で .bas ファイルから probe.xlsm を生成。
または手動で VBE にインポート。

## 優先実装順

iris のメモに基づく優先度:

1. 端末情報 (SystemInfo) — 必須
2. Office ビット数確認 — 必須
3. 参照ライブラリ (References) — 必須
4. COM 生成テスト (COMCreate) — 必須
5. ファイルアクセス (FileAccess) — 必須
6. 共有フォルダ確認 — 必須
7. 印刷 / PDF (PrintTest) — 必須
8. VBIDE アクセス — 必須
9. VBA 実行基盤 — 次点
10. ActiveX テスト — 次点
11. 通信テスト — 案件次第
12. Office 連携 — 案件次第
13. 外部実行 — 案件次第
14. DB 接続 — 案件次第
15. セキュリティ制約 — 次点
