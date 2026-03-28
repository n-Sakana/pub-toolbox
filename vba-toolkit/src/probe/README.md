# Probe Macro - Environment Testing Tool

調査マクロ: toolkit が検知するパターンが新環境で動くかどうかを確認するツール。

## 使い方

### 1. 準備

1. 空の `.xlsm` ファイルを作成する
2. Alt+F11 で VBE（Visual Basic Editor）を開く
3. 以下の `.bas` ファイルをインポートする（ファイル → ファイルのインポート）:
   - `ProbeMain.bas`
   - `ProbeTests.bas`
   - `ProbeOutput.bas`

### 2. 設定シートの作成

`_probe_config` という名前のシートを作成し、以下のように入力する:

| A列 | B列 | 説明 |
|---|---|---|
| RunExtended | FALSE | Extended テストを実行するか（TRUE/FALSE） |
| TestURL | | HTTP テストの接続先 URL（空の場合はスキップ） |
| OutputFolder | | CSV 出力先フォルダ（空の場合はブックと同じフォルダ） |
| DummyFileName | _probe_test.txt | ファイル I/O テスト用の一時ファイル名 |

### 3. 結果シートの作成

`_probe_result` という名前のシートを作成する（内容は空でよい。存在しない場合は自動作成される）。

### 4. 実行

1. Alt+F8 でマクロダイアログを開く
2. `Probe_Run` を選択して実行
3. 完了後、結果サマリがメッセージボックスで表示される

### 5. 結果の確認

- **シート**: `_probe_result` シートに色分けされた結果が出力される
  - 緑: OK（正常動作）
  - 赤: FAIL（エラーまたはブロック）
  - 灰: SKIP（テスト条件未達でスキップ）
- **CSV**: `probe_result_<端末名>_<日時>.csv` が出力フォルダに生成される

## テスト項目

### Basic（#1-16）: 安全プローブ
COM生成、ファイルI/O、レジストリ、環境変数、クリップボード、VarPtr、DAO、レガシーコントロール、HTTP

### Extended（#17-25）: 強いプローブ（デフォルト無効）
Win32 API、LoadLibrary、WMI GetObject、Shell、PowerShell、WMI ExecQuery、SendKeys、DDE、IE

### 補助（#26-30）
参照ライブラリ確認（#26-27）、環境情報（#28-30）

## 注意事項

- Extended テストは EDR アラートを引き起こす可能性があります。事前了解のうえ実行してください。
- 参照テスト（#26-27）は「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が ON の場合のみ動作します。OFF の場合は SKIP になります。
- すべてのテストは副作用なし（作成したファイルは削除、送信は行わない）です。
