# vba-toolkit プロジェクト概要

## 目的

レガシー環境の VBA マクロ資産を新環境に移行するための支援ツール群。

新環境の制約:
- **EDR（Endpoint Detection & Response）**: `Declare PtrSafe Function` 等を含むファイルを開くとファイルが破損する
- **64bit Office**: 32bit 前提のコード（Long 型ハンドル、PtrSafe 未対応）が動かない
- **OneDrive / SharePoint**: ファイルパスの前提が変わる（ローカルパス → URL、同期フォルダ）
- **セキュリティポリシー強化**: Shell, PowerShell, WScript 等の外部プロセス実行がブロックされる

既知の確定制約:
- **Win32 API 使用不可**: EDR により Declare 文を含むファイルが開けない（コードを見ようとした時点で破損）
- **DAO 使用不可**: 64bit 環境で DAO 3.6 の参照が解決できない
- **共有サーバ → SharePoint + OneDrive 移行**: 固定パス・UNC パス・相対パス前提の処理が崩れる。ブラウザ起動ではマクロ実行不可。OneDrive 同期パスはユーザーごとに異なる

## 3層アプローチ

```
┌─────────────────────────────────────────────────┐
│ 1. toolkit（静的解析）                            │
│    Excel を開かずにバイナリレベルで VBA コードを解析  │
│    → 何が使われているか、何が危ないかを洗い出す      │
├─────────────────────────────────────────────────┤
│ 2. 調査マクロ（環境実測）                          │
│    新環境で実際に各パターンを実行してみる             │
│    → 何が通って何が通らないかを実測する              │
├─────────────────────────────────────────────────┤
│ 3. 突き合わせ → 修正対象の確定                      │
│    toolkit の検知結果 × 調査マクロの実測結果          │
│    → 本当に修正が必要なファイル・箇所を特定           │
└─────────────────────────────────────────────────┘
```

### 1. toolkit（静的解析）

Excel を開かずに、OLE2 + MS-OVBA 圧縮形式をバイナリレベルで解析する PowerShell ツール群。

| ツール | 役割 |
|--------|------|
| **Extract** | VBA ソースコードをテキスト抽出（modules/ + combined.txt） |
| **Analyze** | EDR/互換性リスク検知 + サニタイズ + 移行ガイド + CSV |
| **Diff** | 2ファイル間の VBA コード差分比較 |
| **Unlock** | VBA プロジェクトのパスワード保護解除 |

特徴:
- Excel を一切開かない（Unlock の .xls 変換を除く）
- パスワード保護されたプロジェクトも解析可能
- BAT にドラッグ＆ドロップで動作
- コア処理は C#（Add-Type）で高速化

### 2. 調査マクロ（Environment Probe）

toolkit が検知するパターン（Declare, CreateObject, Shell 等）を、新環境で実際に実行して通るか通らないかを確認する使い捨てマクロ。

- toolkit の検知パターンと 1:1 対応
- 副作用なし
- CSV で結果出力
- 新環境に持ち込んで1回実行するだけ

### 3. 突き合わせ

```
probe: CreateObject("InternetExplorer.Application") → FAIL
analyze.csv: FileA.xlsm に IE Automation 検知あり
→ FileA.xlsm は IE 関連コードの修正が必要
```

toolkit の Analyze CSV と調査マクロの結果 CSV を突き合わせることで、500+ ファイルの中から本当に修正が必要なものを特定する。

## 検知カテゴリ

4大分類 + 情報提供:

| カテゴリ | 内容 | ハイライト色 |
|----------|------|-------------|
| EDR / セキュリティ制約 | Win32 API, Shell, COM, WMI 等 | 青 |
| 64bit / 互換性 | PtrSafe, LongPtr, DAO, DefType 等 | 紫 |
| 環境依存 | 固定パス, UNC, プリンタ, 接続文字列 等 | 緑（予定） |
| 業務依存 | Outlook, Word, Access, 印刷, PDF 等 | 橙（予定） |
| 情報提供 (Info) | ThisWorkbook.Path 等（リスクではないが確認推奨） | なし |

## ワークフロー

```
1. Analyze で全 xlsm を一括スキャン
   → analyze.csv（全ファイルの検知結果一覧）
   → 各ファイルの HTML レポート

2. 調査マクロを新環境で実行
   → probe_result.csv（各パターンの OK/FAIL）

3. 2つの CSV を突き合わせ
   → FAIL なパターンを使っているファイルが修正対象

4. 修正対象ファイルに対して:
   - Analyze のサニタイズ機能で自動コメントアウト
   - HTML レポートのツールチップで代替手段を確認
   - Extract でコードを抽出して手動修正

5. 修正後に Diff で変更確認
```

## プロジェクト構成

```
vba-toolkit/
├── Extract.bat          モジュール抽出
├── Analyze.bat          分析 + サニタイズ + 移行ガイド
├── Diff.bat             差分比較
├── Unlock.bat           パスワード解除
├── config/
│   └── analyze.json     分析・サニタイズ設定
├── lib/
│   ├── VBAToolkit.psm1  共通モジュール
│   ├── Extract.ps1
│   ├── Analyze.ps1
│   ├── Diff.ps1
│   └── Unlock.ps1
├── test/
│   ├── test_sample.xlsm
│   ├── test_protected.xlsm
│   ├── test_large.xlsm
│   └── csharp-check/
└── docs/
    ├── overview.md              この文書
    ├── env-patterns-spec.md     環境依存パターン拡張仕様（未実装）
    └── probe-macro-spec.md      調査マクロ仕様（未実装）
```

## 技術スタック

| レイヤー | 技術 |
|----------|------|
| エントリポイント | BAT（ドラッグ＆ドロップ） |
| オーケストレーション | PowerShell |
| バイナリ処理 | C#（Add-Type でインラインコンパイル） |
| OLE2 解析 | 自前実装（セクタチェーン、FAT、ディレクトリ） |
| VBA 圧縮/展開 | MS-OVBA 2.4.1 準拠の自前実装 |
| ZIP 操作 | System.IO.Compression |
| HTML ビューア | インライン生成の静的 HTML（ダークテーマ、ミニマップ付き） |
| 設定 GUI | WinForms（Analyze の設定モード） |
| パスワード解除 | Excel COM（.xls 変換経由、Unlock のみ） |
