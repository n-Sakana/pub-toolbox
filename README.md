# toolbox

VBA / PowerShell 単発マクロ・スクリプト集。

## VBA マクロ

| マクロ | 説明 |
|--------|------|
| [ExcelDataExtractor](vba/ExcelDataExtractor/) | ADO 経由で複数 Excel ファイルからセルデータを一括抽出 |
| [OutlookDraftCreator](vba/OutlookDraftCreator/) | Excel シートから一括で Outlook 下書きメールを作成 |
| [VBAPasswordRemover](vba/VBAPasswordRemover/) | VBA プロジェクトのパスワード保護を解除 (.xls / .xlsm / .xlam) |

## PowerShell スクリプト

| スクリプト | 説明 |
|-----------|------|
| [VBAPasswordRemover](ps/VBAPasswordRemover/) | VBA プロジェクトのパスワード保護を解除 (BAT にドラッグ＆ドロップ) |

## 使い方

- **VBA マクロ**: 各フォルダ内の `.bas` ファイルを Excel VBE (Alt+F11) にインポートして実行
- **PowerShell**: `.bat` ファイルに対象ファイルをドラッグ＆ドロップ、またはコマンドラインから実行

詳細は各フォルダの README を参照。
