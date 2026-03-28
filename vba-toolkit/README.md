# vba-toolkit

Excel を開かずに VBA プロジェクトをバイナリレベルで操作するツール集。

## ツール

| BAT | 説明 |
|-----|------|
| `Extract.bat` | VBA コードをテキスト抽出 + EDR リスク分析 + HTML ビューア |
| `Diff.bat` | 2つの Excel ファイルの VBA コードを差分比較 |
| `Sanitize.bat` | Win32 API 宣言と呼び出しをコメントアウト（非破壊） |
| `Cheatsheet.bat` | Win32 API の移行ガイド（代替手段 + コード例） |
| `Unlock.bat` | VBA プロジェクトのパスワード保護を解除（非破壊） |

## 使い方

対象の `.xls` / `.xlsm` / `.xlam` ファイルを BAT にドラッグ＆ドロップ。
複数ファイルの同時ドロップに対応（Extract, Sanitize, Cheatsheet, Unlock）。

**元ファイルは一切変更されません。** 全ての出力は `output/` フォルダに集約されます。

## 出力

入力ファイルと同じフォルダに `output/` が作成され、実行ごとにタイムスタンプ付きサブフォルダが生成されます。

```
output/
├── 20260328_120000_extract/
│   ├── modules/           .bas / .cls / .frm
│   ├── analysis.txt       EDR リスク分析
│   ├── combined.txt       統合ソース
│   └── extract.html       HTML ビューア
├── 20260328_120500_sanitize/
│   ├── sample.xlsm        サニタイズ済みコピー
│   ├── sanitize.txt       レポート
│   └── sanitize.html      HTML ビューア
├── 20260328_121000_cheatsheet/
│   ├── cheatsheet.txt     移行ガイド
│   └── cheatsheet.html    HTML ビューア
├── 20260328_121500_diff/
│   ├── diff.txt           差分レポート
│   └── diff.html          HTML ビューア
└── 20260328_122000_unlock/
    └── sample.xlsm        パスワード解除済みコピー
```

実行ログは `vba-toolkit/vba-toolkit.log` に追記されます。

## 構成

```
vba-toolkit/
├── Extract.bat
├── Diff.bat
├── Sanitize.bat
├── Cheatsheet.bat
├── Unlock.bat
├── vba-toolkit.log        実行ログ (自動生成)
└── lib/
    ├── VBAToolkit.psm1    共通モジュール (OLE2, VBA圧縮/展開, HTML テンプレート)
    ├── Extract.ps1
    ├── Diff.ps1
    ├── Sanitize.ps1
    ├── Cheatsheet.ps1
    └── Unlock.ps1
```

## 仕組み

OLE2 Compound Document と MS-OVBA 圧縮形式を PowerShell + C# (Add-Type) で解析。Excel COM は Unlock のみ使用。
