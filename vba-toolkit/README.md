# vba-toolkit

Excel を開かずに VBA プロジェクトをバイナリレベルで操作するツール集。

## ツール

| BAT | 説明 |
|-----|------|
| `Extract.bat` | VBA コードをテキスト抽出（モジュール個別 + combined.txt） |
| `Analyze.bat` | 分析 + サニタイズ + 移行ガイド + CSV ログ |
| `Diff.bat` | 2つの Excel ファイルの VBA コードを差分比較 |
| `Unlock.bat` | VBA プロジェクトのパスワード保護を解除（非破壊） |

## 使い方

対象の `.xls` / `.xlsm` / `.xlam` ファイルまたはフォルダを BAT にドラッグ＆ドロップ。

- **Extract**: ファイルまたはフォルダをドロップ → VBA ソースコードをテキスト抽出
- **Analyze**: 引数なしで設定 GUI、ファイル/フォルダで分析実行
- **Diff**: 2ファイルをドロップ → サイドバイサイド差分比較
- **Unlock**: ファイルをドロップ → パスワード保護解除

**元ファイルは一切変更されません。** 全ての出力は `output/` フォルダに集約されます。

## Analyze の3モード

1. **設定 GUI**（引数なし）: ダブルクリックで WinForms ダイアログ。パターンごとに Detect / Sanitize を設定。
2. **ファイル分析**: ファイルをドロップ → EDR リスク + 互換性リスク検出 + サニタイズ + HTML ビューア
3. **フォルダ分析**: フォルダをドロップ → 再帰走査で全 xlsm/xlam/xls を一括分析

## 出力

```
output/
├── 20260328_120000_extract/
│   ├── modules/           .bas / .cls / .frm
│   └── combined.txt       統合ソース
├── 20260328_120500_analyze/
│   ├── analyze.csv        分析結果一覧（BOM UTF-8）
│   ├── sample_analyze.txt テキストレポート
│   ├── sample_analyze.html HTML ビューア（3カラム）
│   └── sample.xlsm        サニタイズ済みコピー（該当時のみ）
├── 20260328_121000_diff/
│   ├── diff.txt           差分レポート
│   └── diff.html          HTML ビューア
└── 20260328_121500_unlock/
    └── sample.xlsm        パスワード解除済みコピー
```

実行ログは `vba-toolkit/vba-toolkit.log` に追記されます。

## 構成

```
vba-toolkit/
├── Extract.bat
├── Analyze.bat
├── Diff.bat
├── Unlock.bat
├── vba-toolkit.log          実行ログ (自動生成)
├── config/
│   └── analyze.json         分析・サニタイズ設定
├── lib/
│   ├── VBAToolkit.psm1      共通モジュール (OLE2, VBA圧縮/展開, 分析エンジン, API代替DB, HTML)
│   ├── Extract.ps1          モジュール抽出
│   ├── Analyze.ps1          分析 + サニタイズ + 移行ガイド + CSV
│   ├── Diff.ps1             差分比較
│   └── Unlock.ps1           パスワード解除
├── test/
│   ├── test_sample.xlsm
│   ├── test_protected.xlsm
│   ├── test_large.xlsm
│   ├── test_replacements.ps1
│   └── csharp-check/
│       ├── Check.bat
│       └── Check.ps1
└── docs/
    └── consolidation-spec.md
```

## 仕組み

OLE2 Compound Document と MS-OVBA 圧縮形式を PowerShell + C# (Add-Type) で解析。Excel COM は Unlock のみ使用。
