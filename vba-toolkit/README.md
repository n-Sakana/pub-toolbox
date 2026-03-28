# vba-toolkit

Excel を開かずに VBA プロジェクトをバイナリレベルで操作するツール集。

## ツール

| BAT | 説明 |
|-----|------|
| `Extract.bat` | VBA コードをテキスト抽出 + EDR リスク分析 + 統合ソース出力 |
| `Diff.bat` | 2つの Excel ファイルの VBA コードを差分比較 |
| `Sanitize.bat` | Win32 API 宣言と呼び出しをコメントアウト |
| `RemovePassword.bat` | VBA プロジェクトのパスワード保護を解除 |

## 使い方

対象の `.xls` / `.xlsm` / `.xlam` ファイルを BAT にドラッグ＆ドロップ。

### Extract

- `<ファイル名>_vba/` に個別モジュール (.bas/.cls/.frm) を出力
- `<ファイル名>_vba/_analysis.txt` に EDR リスク分析レポートを出力
- `<ファイル名>_combined.txt` に統合ソース（アーキテクチャヘッダ付き）を出力

### Sanitize

- 元ファイルのバックアップ (`.bak`) を作成
- `Declare PtrSafe Function` 等の行を `' [SANITIZED] ` 付きでコメントアウト
- 宣言された API の呼び出し箇所も自動検出してコメントアウト
- `<ファイル名>_sanitized.txt` にレポートを出力

ルールは `lib/Sanitize.ps1` 冒頭の `$rules` で変更可能。

### Diff

```
Diff.bat file1.xlsm file2.xlsm
```

- 両ファイルの VBA コードをバイナリレベルで抽出・比較
- モジュールの追加/削除/変更を検出
- 変更箇所を行単位で表示
- `<fileA>_vs_<fileB>_diff.txt` にレポートを出力

### RemovePassword

- Excel COM で `.xls` 変換→バイナリパッチ→元形式で保存
- 処理中に「不正なキー 'DPx'」ダイアログが出たら「はい」をクリック
- 処理後に VBE で保護タブからパスワードを空にして保存

## 構成

```
vba-toolkit/
├── Extract.bat
├── Diff.bat
├── Sanitize.bat
├── RemovePassword.bat
└── lib/
    ├── VBAToolkit.psm1    共通モジュール (OLE2, VBA圧縮/展開)
    ├── Extract.ps1
    ├── Diff.ps1
    ├── Sanitize.ps1
    └── RemovePassword.ps1
```

## 仕組み

OLE2 Compound Document と MS-OVBA 圧縮形式を PowerShell で直接解析。Excel COM は RemovePassword のみ使用（Extract/Sanitize は完全バイナリ操作）。
