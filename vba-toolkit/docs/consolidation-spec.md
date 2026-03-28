# vba-toolkit 統合仕様書

## 概要

6ツール (Extract, Sanitize, Cheatsheet, Inventory, Diff, Unlock) を4ツールに統合する。

## ツール構成

| BAT | 用途 | 入力 |
|-----|------|------|
| Extract.bat | VBA ソースコード抽出 | ファイル / フォルダ |
| Analyze.bat | 分析 + サニタイズ + 移行ガイド + CSV ログ | なし / ファイル / フォルダ |
| Diff.bat | 差分比較 | 2ファイル |
| Unlock.bat | パスワード解除 | ファイル |

廃止: Sanitize.bat, Cheatsheet.bat, Inventory.bat

---

## 1. Extract

### 機能
VBA モジュールをテキストファイルとして抽出する。分析やハイライトは行わない。

### 入力
- ファイル（1つ以上の .xls / .xlsm / .xlam）
- フォルダ（再帰的に xlsm を収集）

### 出力
```
output/<timestamp>_extract/
├── modules/
│   ├── Module1.bas
│   ├── Class1.cls
│   └── Form1.frm
└── combined.txt
```

### combined.txt の構造
```
================================================================================
 sample.xlsm - VBA Source Code
 Extracted: 2026-03-28 15:00:00
================================================================================

MODULE INDEX
----------------------------------------

  Standard Modules:
    Module1 (120 lines)
    Module2 (85 lines)

  Class Modules:
    Class1 (45 lines)

  UserForms:
    Form1 (200 lines)

  Document Modules:
    ThisWorkbook (10 lines)
    Sheet1 (5 lines)

  Total: 465 lines across 6 modules

================================================================================
 Module1.bas
================================================================================

Option Explicit
...
```

### ログ
`vba-toolkit.log` に追記。CSV への追記はなし。

---

## 2. Analyze

### 機能
VBA コードの分析、サニタイズ、移行ガイドを1つのツールで行う。

### 3つのモード

#### モード 1: 設定 GUI（引数なし）
BAT をダブルクリックすると WinForms ダイアログを表示。

**レイアウト:**
```
┌─ Analyze Settings ──────────────────────────┐
│                                              │
│  EDR Risks              Detect  Sanitize     │
│  ─────────────────────────────────────────   │
│  Win32 API (Declare)      ☑        ☑         │
│  DLL loading              ☑        ☑         │
│  COM / CreateObject       ☑        ☐         │
│  Shell / process          ☑        ☐         │
│  ...                                         │
│                                              │
│  Compatibility Risks    Detect  Sanitize     │
│  ─────────────────────────────────────────   │
│  64-bit: Missing PtrSafe  ☑        ☐         │
│  64-bit: Long for handles ☑        ☐         │
│  Deprecated: DAO          ☑        ☐         │
│  ...                                         │
│                                              │
│                        [ OK ]  [ Cancel ]    │
└──────────────────────────────────────────────┘
```

- ダークテーマ（#252526 背景）
- OK → `config/analyze.json` に保存
- Cancel → 変更なし

#### モード 2: ファイル分析（ファイルをドロップ）
#### モード 3: フォルダ分析（フォルダをドロップ）

モード2と3は同じ処理。フォルダなら再帰走査で xlsm を収集し、各ファイルに対して同一の処理を実行。

### 処理フロー

```
1. 入力を解決（ファイル → そのまま、フォルダ → 再帰的に xlsm を収集）
2. 出力フォルダ output/<timestamp>_analyze/ を作成
3. 各ファイルに対して:
   a. Get-AllModuleCode でモジュール読み込み
      （サニタイズ設定があれば -IncludeRawData 付き）
   b. Get-VbaAnalysis で EDR + 互換性パターン検出
   c. API 宣言を Get-VbaApiReplacements で代替情報と照合
   d. サニタイズ対象があればコピーを生成してコメントアウト
4. 出力生成:
   a. analyze.csv — 全ファイルの分析結果一覧（1ファイル=1行）
   b. analyze.txt — テキストレポート（全ファイル分）
   c. analyze.html — コードビューア
   d. サニタイズ済みコピー（該当ファイルのみ）
```

### 出力（1回の実行につき1フォルダ）
```
output/<timestamp>_analyze/
├── analyze.csv            分析結果一覧（ファイル数分のレコード）
├── sample1_analyze.txt    テキストレポート（ファイルごと）
├── sample1_analyze.html   HTML ビューア（ファイルごと）
├── sample1.xlsm           サニタイズ済みコピー（該当時のみ）
├── sample2_analyze.txt
├── sample2_analyze.html
└── sample2.xlsm
```

- analyze.csv は実行ごとに新規作成。入力ファイルの数だけレコードが入る。BOM 付き UTF-8。
- HTML / txt はファイルごとに1つ生成（`<baseName>_analyze.html`）。
- サニタイズ済みコピーはファイルごとに1つ（該当時のみ）。
- ファイル名衝突時（異なるサブフォルダに同名ファイル）: 親フォルダ名をプレフィクス（例: `sub1_sample_analyze.html`）。

### バイナリ書き戻し（サニタイズ）
サニタイズ済みコピーの生成は既存の OLE2 バイナリ操作を使用:
1. 元ファイルを出力フォルダにコピー
2. `Get-AllModuleCode -IncludeRawData` でモジュールの生バイトデータを取得
3. コード行を修正（コメントアウト）
4. `Compress-VBA` で再圧縮
5. `Write-Ole2Stream` でコピーの vbaProject.bin に書き戻し
6. `Save-VbaProjectBytes` で ZIP（xlsm）を更新

### 対応拡張子
全ツール共通: `.xls` / `.xlsm` / `.xlam`

| カラム | 内容 |
|--------|------|
| Timestamp | 実行日時 (yyyy-MM-dd HH:mm:ss) |
| RelativePath | ベースフォルダからの相対パス |
| FileName | ファイル名 |
| Bas | 標準モジュール数 |
| Cls | クラスモジュール数 |
| Frm | フォーム数 |
| TotalModules | 合計モジュール数 |
| CodeLines | コード行数（Attribute 行除外） |
| EdrIssues | EDR 検知件数 |
| CompatIssues | 互換性リスク件数 |
| SanitizedLines | サニタイズ済み行数 |
| References | 参照ライブラリ一覧（セミコロン区切り） |
| Error | エラーメッセージ（正常時は空） |

BOM 付き UTF-8。ヘッダー行 + データ行（ファイル数分）。実行ごとに新規作成。

### analyze.txt の構成
```
# VBA Analysis Report
# Source: sample.xlsm
# Date: 2026-03-28 15:00:00

## Modules (6)
  Module1.bas (120 lines)
  ...
  Total: 465 lines

## EDR Risks (3)
  Win32 API (Declare) (3)
    Module1.bas: GetTickCount (PtrSafe)
    ...

## Compatibility Risks (2)
  64-bit: Long for handles (2)
    Module1.bas: As Long in PtrSafe Declare -- review for LongPtr

## COM Object Usage Details
  Scripting.FileSystemObject
    Module1.bas L45: Set fso = ...
    Module1.bas L46: .FolderExists  -- fso.FolderExists(path)

## Win32 API Usage Details
  GetTickCount
    Module1.bas L3: Private Declare PtrSafe Function GetTickCount Lib kernel32 () As Long
    Module1.bas L10: t = GetTickCount()
    Alternative: Timer (VBA built-in, Single type)

## External References (2)
  Microsoft Scripting Runtime
  Microsoft DAO 3.6 Object Library

## Summary
  3 EDR issue(s), 2 compatibility issue(s), 7 line(s) sanitized.
```

### analyze.html の構成

3カラムレイアウト:

```
┌─────────┬──────────────────────────┬──────────────┬──┐
│ Modules │ Code                     │ Outline      │MM│
│         │                          │              │  │
│ Mod1.bas│ Option Explicit          │ L3  GetTick  │  │ ← 青
│ Mod2.bas│ Private Declare PtrSafe  │ L5  Sleep    │  │ ← 青
│ Class1  │   Function GetTickCount  │ L10 [EDR]    │  │ ← 黄
│         │   ...                    │ L20 Declare  │  │ ← 紫
│         │ ' [EDR] t = GetTickCo..  │ L25 [COMPAT] │  │ ← 黄
│         │ ...                      │              │  │
│         │ Declare Function Foo     │              │  │
│         │   ' (PtrSafe 欠落)       │              │  │
│         │                          │              │  │
└─────────┴──────────────────────────┴──────────────┴──┘
```

- **左**: モジュール一覧（検知件数付き）
- **中央**: 全コード（サニタイズ後のコードを表示）。3色ハイライト:
  - サニタイズ済み = 黄 `#4b3a00` — `[EDR]` / `[COMPAT]` プレフィクスでコメントアウトされた行（最優先）
  - 検知（EDR）= 青 `#1b2e4a` — detect=true のパターンにマッチした未サニタイズ行
  - 検知（互換性）= 紫 `#3a1b4a` — detect=true のパターンにマッチした未サニタイズ行
  - 優先度: 黄（サニタイズ済み） > 青（EDR検知） > 紫（互換性検知）
  - 代替情報のある検知行はクリックでツールチップ（代替情報 + 移行例 + コピーボタン）
- **右**: アウトライン（行番号順にフラット表示）
  - 検知行・サニタイズ済み行を行番号順に並べる
  - 表示形式: `L{行番号} {パターン名}` （最大50文字で切り詰め）
  - クリックで該当行にジャンプ
  - 色で種別を区別（青/紫/黄）
- **右端**: ミニマップ（青 + 紫 + 黄マーク）

### config/analyze.json の構造
```json
{
  "edr": {
    "Win32 API (Declare)": { "detect": true, "sanitize": true },
    "DLL loading": { "detect": true, "sanitize": true },
    "COM / CreateObject": { "detect": true, "sanitize": false },
    "Shell / process": { "detect": true, "sanitize": false },
    "File I/O": { "detect": true, "sanitize": false },
    "FileSystemObject": { "detect": true, "sanitize": false },
    "Registry": { "detect": true, "sanitize": false },
    "SendKeys": { "detect": true, "sanitize": false },
    "Network / HTTP": { "detect": true, "sanitize": false },
    "PowerShell / WScript": { "detect": true, "sanitize": false },
    "Process / WMI": { "detect": true, "sanitize": false },
    "Clipboard": { "detect": true, "sanitize": false },
    "Environment": { "detect": true, "sanitize": false },
    "Auto-execution": { "detect": true, "sanitize": false },
    "Encoding / obfuscation": { "detect": true, "sanitize": false }
  },
  "compat": {
    "64-bit: Missing PtrSafe": { "detect": true, "sanitize": false },
    "64-bit: Long for handles": { "detect": true, "sanitize": false },
    "64-bit: VarPtr/ObjPtr/StrPtr": { "detect": true, "sanitize": false },
    "Deprecated: DDE": { "detect": true, "sanitize": false },
    "Deprecated: IE Automation": { "detect": true, "sanitize": false },
    "Deprecated: Legacy Controls": { "detect": true, "sanitize": false },
    "Deprecated: DAO": { "detect": true, "sanitize": false },
    "Legacy: DefType": { "detect": true, "sanitize": false },
    "Legacy: GoSub": { "detect": true, "sanitize": false },
    "Legacy: While/Wend": { "detect": true, "sanitize": false }
  }
}
```

旧 `config/sanitize.json` が存在する場合、初回実行時に自動移行:
- 旧形式 `{ "edr": { "pattern": true } }` → 新形式 `{ "edr": { "pattern": { "detect": true, "sanitize": true } } }`

### サニタイズのコメントプレフィクス
- EDR パターン: `' [EDR] `
- 互換性パターン: `' [COMPAT] `

### ターミナル表示
```
[Analyze] sample.xlsm
  Modules: 6 (4 bas, 1 cls, 1 frm)
  EDR issues: 3
  Compat issues: 2
  Sanitized: 7 lines
  Output: C:\...\output\20260328_150000_analyze\
  Done (1.2s)
```

---

## 3. Diff（変更なし）

2ファイルの VBA コード差分比較。サイドバイサイド HTML ビューア。

---

## 4. Unlock（変更なし）

VBA プロジェクトのパスワード保護を解除。非破壊（コピーを生成）。

---

## プロジェクト構成（統合後）

```
vba-toolkit/
├── Extract.bat
├── Analyze.bat
├── Diff.bat
├── Unlock.bat
├── vba-toolkit.log              実行ログ（自動生成）
├── config/
│   └── analyze.json             分析設定
├── lib/
│   ├── VBAToolkit.psm1          共通モジュール
│   │   ├── C# インライン        OLE2, VBA 圧縮/展開, バイト検索
│   │   ├── OLE2 パーサー        Read-Ole2, Read-Ole2Stream, Write-Ole2Stream
│   │   ├── ヘルパー             Resolve-VbaFilePath, Get-AllModuleCode, Get-VbaCodepage
│   │   ├── 分析エンジン          Get-VbaAnalysis (EDR + 互換性パターン)
│   │   ├── API 代替 DB          Get-VbaApiReplacements ($replacements)
│   │   ├── ログ・表示           Write-VbaLog, Write-VbaStatus, Write-VbaResult
│   │   ├── 出力管理             New-VbaOutputDir
│   │   └── HTML テンプレート     New-HtmlBase, New-HtmlCodeView, New-HtmlAnalyzeView
│   ├── Extract.ps1              モジュール抽出のみ
│   ├── Analyze.ps1              分析 + サニタイズ + チートシート + CSV
│   ├── Diff.ps1                 差分比較
│   └── Unlock.ps1               パスワード解除
├── test/
│   ├── test_sample.xlsm
│   ├── test_protected.xlsm
│   ├── test_large.xlsm
│   ├── test_replacements.ps1
│   └── csharp-check/
│       ├── Check.bat
│       └── Check.ps1
└── docs/
    └── consolidation-spec.md    この文書
```

## VBAToolkit.psm1 の変更

### 追加する関数

| 関数 | 用途 |
|------|------|
| `Get-VbaApiReplacements` | チートシートの API 代替 DB を返す。戻り値: `[ordered]@{ ApiName -> @{ Lib; Alt; Example; Note } }` |
| `New-HtmlAnalyzeView` | 分析用 HTML 生成。`New-HtmlBase` を内部で呼び出す。`New-HtmlCodeView` は Extract/Sanitize 廃止により不要になるが、Diff が使用する可能性があるため残す |

### 移動するデータ

`$replacements`（~40 API エントリの代替情報 DB）を Cheatsheet.ps1 から VBAToolkit.psm1 に移動。`$script:VbaApiReplacements` として格納。

### 設定ファイル移行

旧 `config/sanitize.json` が存在し `config/analyze.json` が存在しない場合:
- 旧形式 `{ "edr": { "pattern": true } }` → 新形式 `{ "edr": { "pattern": { "detect": true, "sanitize": true } } }` に変換
- `config/analyze.json` として保存
- `config/sanitize.json` は削除しない（バックアップとして残す）
- 不明なキーは無視（警告なし）

---

## 出力フォルダ構成（全体）

```
output/
├── 20260328_150000_extract/
│   ├── modules/
│   └── combined.txt
├── 20260328_150100_analyze/
│   ├── analyze.csv
│   ├── sample_analyze.txt
│   ├── sample_analyze.html
│   └── sample.xlsm                    サニタイズ済みコピー（該当時のみ）
├── 20260328_150200_diff/
│   ├── diff.txt
│   └── diff.html
└── 20260328_150300_unlock/
    └── sample.xlsm                    パスワード解除済みコピー
```
