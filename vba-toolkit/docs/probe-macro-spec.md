# 調査マクロ仕様書 (Environment Probe)

## 概要

toolkit が検知する EDR リスク・互換性リスクのパターンが、新環境で実際に動くか動かないかを確認するマクロ。使い捨て。

## 基本方針

- toolkit の検知パターンと 1:1 対応するテスト項目
- 各パターンを実際に実行してみて OK/FAIL を記録するだけ
- 副作用なし（作って消す、下書きまでで送信しない）
- 結果は CSV + シート出力
- 単一 xlsm（手動で VBE にインポートして実行）

## 二層構成: Basic / Extended

配布先の温度感に合わせて使い分ける。

### Basic（安全プローブ）

COM 生成、ファイル I/O、レジストリ、参照確認、環境情報など。支社配布向け。

### Extended（強いプローブ）

Win32 API 呼び出し、Shell、WScript、PowerShell、DDE、IE、WMI など。実行するとEDR のアラートや管理部門の警戒を招く可能性がある。必要時のみ、事前了解のうえで実行。

実装は分けなくてもよい。設定シートで Basic / Extended を切り替える。デフォルトは Basic のみ。

## テスト項目

### Basic（安全プローブ）

| # | toolkit パターン | テスト内容 |
|---|---|---|
| 1 | COM / CreateObject | `CreateObject("Scripting.FileSystemObject")` |
| 2 | COM / CreateObject | `CreateObject("Scripting.Dictionary")` |
| 3 | COM / CreateObject | `CreateObject("ADODB.Connection")` |
| 4 | COM / CreateObject | `CreateObject("ADODB.Recordset")` |
| 5 | COM / CreateObject | `CreateObject("MSXML2.XMLHTTP.6.0")` |
| 6 | COM / CreateObject | `CreateObject("WinHttp.WinHttpRequest.5.1")` |
| 7 | File I/O | `Open ... For Output As #1` → 書き込み → `Close` → `Kill` |
| 8 | FileSystemObject | `CreateObject("Scripting.FileSystemObject").FileExists(...)` |
| 9 | Registry | `GetSetting "ProbeTest", "Test", "Key", ""` |
| 10 | Environment | `Environ$("USERNAME")` |
| 11 | Clipboard | `Dim d As New MSForms.DataObject: d.SetText "test": d.PutInClipboard` |
| 12 | 64-bit: VarPtr | `Dim p As LongPtr: p = VarPtr(p)` |
| 13 | Deprecated: DAO | `CreateObject("DAO.DBEngine.36")` |
| 14 | Deprecated: Legacy Controls | `CreateObject("MSComDlg.CommonDialog")` |
| 15 | Deprecated: Legacy Controls | `CreateObject("MSCAL.Calendar")` |
| 16 | Network / HTTP | 設定シートの URL に対して `XMLHTTP.Open "GET"` → `.Send` |

### Extended（強いプローブ） ※デフォルト無効

| # | toolkit パターン | テスト内容 |
|---|---|---|
| 17 | Win32 API (Declare) | `Declare PtrSafe Function Sleep Lib "kernel32" ...` を宣言して呼ぶ |
| 18 | DLL loading | `Declare PtrSafe Function LoadLibrary ...` を宣言して呼ぶ |
| 19 | COM / GetObject | `GetObject("winmgmts:\\.\root\cimv2")` |
| 20 | Shell / process | `CreateObject("WScript.Shell").Run "cmd /c echo test", 0, True` |
| 21 | PowerShell / WScript | `CreateObject("WScript.Shell").Run "powershell -Command exit", 0, True` |
| 22 | Process / WMI | `GetObject("winmgmts:").ExecQuery(...)` |
| 23 | SendKeys | `SendKeys ""` |
| 24 | Deprecated: DDE | `DDEInitiate` を呼んでみる（失敗前提） |
| 25 | Deprecated: IE Automation | `CreateObject("InternetExplorer.Application")` |

注意:
- SendKeys（#23）は呼び出し可否の確認のみ。空文字列なので副作用はないが、環境によるブロック有無の実測としては限定的。
- Shell/PowerShell テストは `WScript.Shell.Run` の `waitOnReturn:=True` で実行結果（終了コード）を取得する。VBA の `Shell` 関数は終了コードを返せないため使用しない。

### 参照ライブラリ確認（補助）

VBProject にアクセスできる場合のみ実行。できない場合は SKIP（「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」が OFF だと動かない）。必須扱いにしない。

| # | テスト内容 |
|---|---|
| 26 | 全参照の列挙（VBProject.References）|
| 27 | Missing 参照の有無（IsBroken = True） |

### 環境情報（補助）

| # | 取得項目 | 方法 |
|---|---|---|
| 28 | Office バージョン | Application.Version |
| 29 | Office ビット数 | #If Win64 |
| 30 | VBA バージョン | #If VBA7 |

## テスト不要なパターン

以下は構文やコーディングスタイルの問題であり、環境に依存しないためテスト不要:

- Auto-execution（マクロが動いている時点で確認済み）
- Encoding / obfuscation（Chr$ は VBA 標準関数）
- Legacy: DefType（構文の問題。ただし DefType が原因で暗黙に Long になった変数が 64bit API に渡されて壊れるケースは、API テスト #17 でカバーされる）
- Legacy: GoSub（構文の問題）
- Legacy: While/Wend（構文の問題）

## 設定シート (_probe_config)

| A列 | B列 | 説明 |
|---|---|---|
| RunExtended | FALSE | Extended テストを実行するか |
| TestURL | (空) | HTTP テストの接続先 URL。空の場合は HTTP テストを SKIP |
| OutputFolder | (空) | CSV 出力先。空の場合は ThisWorkbook と同じフォルダ |
| DummyFileName | _probe_test.txt | ファイル I/O テスト用の一時ファイル名 |

## 出力

### CSV

| 列 | 内容 |
|---|---|
| TestNo | テスト番号 |
| Level | Basic / Extended / Aux |
| Category | EDR / Compat / Reference / SystemInfo |
| PatternName | toolkit のパターン名と対応 |
| Target | 実行した内容（ProgID、API 名等） |
| Result | OK / FAIL / SKIP |
| ErrorNumber | VBA エラー番号（正常時は空） |
| ErrorMessage | エラーメッセージ（正常時は空） |
| Detail | 補足（バージョン、値等） |

ファイル名: `probe_result_<端末名>_<日時>.csv`
BOM 付き UTF-8。

### 結果シート

CSV と同内容。OK=緑, FAIL=赤, SKIP=灰。

## 実行フロー

```
1. 設定シート読み込み
2. 結果シートクリア
3. 環境情報取得（#28-30）
4. Basic テスト実行（#1-16）
5. RunExtended = TRUE の場合のみ Extended テスト実行（#17-25）
6. VBProject にアクセスできれば参照確認（#26-27）
7. CSV 書き出し
8. 結果シート整形（色分け）
9. 完了メッセージ（OK/FAIL/SKIP の件数サマリ）
```

## ファイル構成

```
probe.xlsm
├── Module: ProbeMain       エントリポイント (Probe_Run)
├── Module: ProbeTests      各テストの実装
├── Module: ProbeOutput     CSV/シート出力
├── Sheet: _probe_config    設定シート
└── Sheet: _probe_result    結果シート
```

手動で VBE にインポートして `Probe_Run` を実行。Build スクリプトは不要。

## toolkit との対応

調査マクロの結果と toolkit の Analyze 結果を突き合わせることで:

- toolkit が「危ない」と言ったものが、新環境で**本当に動かないのか**を確認
- 「FAIL だったパターンを使っているファイル」を Analyze の CSV から抽出 → 修正対象の確定

例:
```
probe: CreateObject("InternetExplorer.Application") → FAIL
analyze.csv: FileA.xlsm に IE Automation 検知あり
→ FileA.xlsm は IE 関連コードの修正が必要
```
