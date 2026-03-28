# 調査マクロ仕様書 (Environment Probe)

## 概要

toolkit が検知する EDR リスク・互換性リスクのパターンが、新環境で実際に動くか動かないかを確認するマクロ。使い捨て。

## 基本方針

- toolkit の検知パターンと 1:1 対応するテスト項目
- 各パターンを実際に実行してみて OK/FAIL を記録するだけ
- 副作用なし（作って消す、下書きまでで送信しない）
- 結果は CSV + シート出力
- 単一 xlsm（手動で VBE にインポートして実行）

## テスト項目

### EDR リスク（toolkit の検知パターンに対応）

| # | toolkit パターン | テスト内容 |
|---|---|---|
| 1 | Win32 API (Declare) | `Declare PtrSafe Function Sleep Lib "kernel32" ...` を宣言して呼ぶ |
| 2 | DLL loading | `LoadLibrary` を宣言して呼ぶ |
| 3 | COM / CreateObject | `CreateObject("Scripting.FileSystemObject")` |
| 4 | COM / CreateObject | `CreateObject("Scripting.Dictionary")` |
| 5 | COM / CreateObject | `CreateObject("ADODB.Connection")` |
| 6 | COM / CreateObject | `CreateObject("ADODB.Recordset")` |
| 7 | COM / CreateObject | `CreateObject("MSXML2.XMLHTTP.6.0")` |
| 8 | COM / CreateObject | `CreateObject("WinHttp.WinHttpRequest.5.1")` |
| 9 | COM / GetObject | `GetObject("winmgmts:\\.\root\cimv2")` |
| 10 | Shell / process | `Shell "cmd /c echo test"` |
| 11 | Shell / process | `CreateObject("WScript.Shell").Run "cmd /c echo test", 0, True` |
| 12 | File I/O | `Open ... For Output As #1` → 書き込み → `Close` → `Kill` |
| 13 | FileSystemObject | `CreateObject("Scripting.FileSystemObject").FileExists(...)` |
| 14 | Registry | `GetSetting "ProbeTest", "Test", "Key", ""` |
| 15 | SendKeys | `SendKeys ""` (空文字で副作用なし) |
| 16 | Network / HTTP | `CreateObject("MSXML2.XMLHTTP.6.0").Open "GET", "https://www.google.com", False` → `.Send` |
| 17 | PowerShell / WScript | `CreateObject("WScript.Shell").Run "powershell -Command exit", 0, True` |
| 18 | Process / WMI | `GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & ... )` |
| 19 | Clipboard | `Dim d As New MSForms.DataObject: d.SetText "test": d.PutInClipboard` |
| 20 | Environment | `Environ$("USERNAME")` |
| 21 | Auto-execution | (テスト不要 — マクロ実行自体が動いている時点で OK) |
| 22 | Encoding / obfuscation | `Chr$(65)` (テスト不要 — VBA 標準関数) |

### 互換性リスク（toolkit の検知パターンに対応）

| # | toolkit パターン | テスト内容 |
|---|---|---|
| 23 | 64-bit: Missing PtrSafe | `Declare PtrSafe Function` が通るか（#1 で兼用） |
| 24 | 64-bit: Long for handles | 64bit 環境かどうか確認: `#If Win64` → Detail に "64-bit" / "32-bit" |
| 25 | 64-bit: VarPtr/ObjPtr/StrPtr | `Dim p As LongPtr: p = VarPtr(p)` |
| 26 | Deprecated: DDE | `DDEInitiate` を呼んでみる（失敗前提） |
| 27 | Deprecated: IE Automation | `CreateObject("InternetExplorer.Application")` |
| 28 | Deprecated: Legacy Controls | `CreateObject("MSComDlg.CommonDialog")` |
| 29 | Deprecated: Legacy Controls | `CreateObject("MSCAL.Calendar")` |
| 30 | Deprecated: DAO | `CreateObject("DAO.DBEngine.36")` |
| 31 | Legacy: DefType | (テスト不要 — 構文の問題であり環境依存ではない) |
| 32 | Legacy: GoSub | (テスト不要 — 同上) |
| 33 | Legacy: While/Wend | (テスト不要 — 同上) |

### 参照ライブラリ確認

| # | テスト内容 |
|---|---|
| 34 | 全参照の列挙（VBProject.References）— VBIDE アクセスが必要 |
| 35 | Missing 参照の有無（IsBroken = True） |

※ VBProject にアクセスできない場合は SKIP

### 環境情報（補助）

| # | 取得項目 | 方法 |
|---|---|---|
| 36 | Office バージョン | Application.Version |
| 37 | Office ビット数 | #If Win64 |
| 38 | VBA バージョン | #If VBA7 |

## 出力

### CSV

| 列 | 内容 |
|---|---|
| TestNo | テスト番号 |
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
1. 結果シートクリア
2. 環境情報取得（#36-38）
3. 各テスト実行
   - On Error Resume Next で保護
   - 結果をシートに即時出力
4. CSV 書き出し
5. 完了メッセージ（OK/FAIL の件数サマリ）
```

## ファイル構成

```
probe.xlsm
├── Module: ProbeMain       エントリポイント (Probe_Run)
├── Module: ProbeTests      各テストの実装
├── Module: ProbeOutput     CSV/シート出力
└── Sheet: _probe_result    結果シート
```

手動で VBE にインポートして `Probe_Run` を実行。Build スクリプトは不要。

## テスト不要なパターン

以下は構文やコーディングスタイルの問題であり、環境に依存しないためテスト不要:

- Auto-execution (マクロが動いている時点で確認済み)
- Encoding / obfuscation (Chr$ は VBA 標準関数)
- Legacy: DefType (構文の問題)
- Legacy: GoSub (構文の問題)
- Legacy: While/Wend (構文の問題)

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
