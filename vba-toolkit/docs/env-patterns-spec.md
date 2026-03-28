# toolkit 拡張仕様: 環境依存パターン + 判定列

## 概要

現在の Analyze は EDR リスク（16パターン）と互換性リスク（10パターン）を検知する。ここに「環境依存リスク」と「業務依存リスク」を追加し、判定列を強化する。

## 追加パターン

### カテゴリ3: 環境依存リスク（緑ハイライト `#1b3a2a`）

| パターン名 | 検知対象 | 正規表現 |
|---|---|---|
| 固定ドライブレター | `C:\`, `D:\` 等のハードコード | `(?mi)^[^'\r\n]*"[A-Z]:\\"` |
| UNC パス | `\\server\share` | `(?mi)^[^'\r\n]*"\\\\[^"]+\\` |
| ユーザーフォルダ | `C:\Users\` | `(?mi)^[^'\r\n]*C:\\Users\\` |
| Desktop / Documents | 固定パスでのデスクトップ/ドキュメント参照 | `(?mi)^[^'\r\n]*\\(Desktop|Documents|ドキュメント|デスクトップ)\\` |
| AppData | AppData パス | `(?mi)^[^'\r\n]*\\AppData\\` |
| Program Files | Program Files パス | `(?mi)^[^'\r\n]*\\Program Files` |
| 固定プリンタ名 | ActivePrinter への文字列設定 | `(?mi)^[^'\r\n]*\.ActivePrinter\s*=\s*"` |
| 固定端末名/IP | ハードコードされたホスト名やIP | `(?mi)^[^'\r\n]*"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"` |
| localhost | localhost 参照 | `(?mi)^[^'\r\n]*\blocalhost\b` |
| 接続文字列 | Provider= や DSN= | `(?mi)^[^'\r\n]*(Provider\s*=|DSN\s*=|Data\s+Source\s*=)` |
| ThisWorkbook.Path | パス依存処理 | `(?mi)^[^'\r\n]*\bThisWorkbook\.Path\b` |
| ActiveWorkbook.Path | パス依存処理 | `(?mi)^[^'\r\n]*\bActiveWorkbook\.Path\b` |
| 外部ブック参照 | Workbooks.Open | `(?mi)^[^'\r\n]*\bWorkbooks\.Open\b` |

### カテゴリ4: 業務依存リスク（橙ハイライト `#4a3a1b`）

| パターン名 | 検知対象 | 正規表現 |
|---|---|---|
| Outlook 連携 | Outlook COM | `(?mi)^[^'\r\n]*\bOutlook\.Application\b` |
| Word 連携 | Word COM | `(?mi)^[^'\r\n]*\bWord\.Application\b` |
| Access / DB 連携 | Access COM / DAO / ADO | `(?mi)^[^'\r\n]*\b(Access\.Application|CurrentDb|DoCmd)\b` |
| PDF 出力 | ExportAsFixedFormat | `(?mi)^[^'\r\n]*\.ExportAsFixedFormat\b` |
| 印刷 | PrintOut / PrintPreview | `(?mi)^[^'\r\n]*\.(PrintOut|PrintPreview)\b` |
| 外部 EXE 起動 | Shell で exe 起動 | `(?mi)^[^'\r\n]*\bShell\s.*\.exe` |

## 判定列（CSV 出力に追加）

### RiskLevel (自動判定)

| 値 | 条件 |
|---|---|
| High | GUI操作系API (FindWindow, SendMessage, PostMessage, keybd_event, mouse_event) あり、または PowerShell/WScript あり |
| Medium | Win32 API あり（GUI操作系以外）、または DAO あり、または 固定パス多数（5件以上） |
| Low | 上記に該当しない |

### MigrationClass (自動判定)

| 値 | 条件 |
|---|---|
| そのまま可 | EDR/互換性/環境依存リスクがすべて 0 |
| 軽微修正 | PtrSafe 追加のみ、または Environ$ 置換のみ |
| 要代替実装 | Win32 API（GUI操作系以外）、または DAO → ADO 移行 |
| 再構築必要 | GUI操作系 API 依存、または Shell/PowerShell 依存 |
| 要保存先見直し | 固定パス/UNC パス/共有フォルダ依存 |

複数該当する場合は最も重い分類を採用。

### PrimaryConcern (自動判定)

検出されたパターンの最頻カテゴリ: `API / COM / DB / Process / Network / File / Mail / Mixed / StorageMigration`

### NeedsReviewBy (自動判定)

| 値 | 条件 |
|---|---|
| Security | EDR リスクあり |
| Infra | 環境依存（パス、プリンタ、接続文字列）あり |
| DB | DAO / ADO / 接続文字列あり |
| BusinessOwner | 業務依存（Outlook、Word、印刷）あり |

複数該当する場合はセミコロン区切り。

## CSV カラム（更新後）

既存: Timestamp, RelativePath, FileName, Bas, Cls, Frm, TotalModules, CodeLines, EdrIssues, CompatIssues, SanitizedLines, References, Error

追加: EnvIssues, BizIssues, RiskLevel, MigrationClass, PrimaryConcern, NeedsReviewBy, TopApiNames, TopComProgIds, SampleEvidence

### TopApiNames
検出された API 宣言名の先頭3件をセミコロン区切り。例: `Sleep; FindWindow; GetTickCount`

### TopComProgIds
検出された COM ProgID の先頭3件をセミコロン区切り。例: `Outlook.Application; Scripting.FileSystemObject`

### SampleEvidence
代表的な検出行を1件。例: `Module1.bas:L42 CreateObject("Outlook.Application")`

## HTML ハイライト色（5色）

| カテゴリ | 色 | 背景 | テキスト |
|---|---|---|---|
| サニタイズ済み | 黄 | #4b3a00 | #f0d870 |
| EDR リスク | 青 | #1b2e4a | #a0c4f0 |
| 互換性リスク | 紫 | #3a1b4a | #c4a0f0 |
| 環境依存リスク | 緑 | #1b3a2a | #a0f0c4 |
| 業務依存リスク | 橙 | #4a3a1b | #f0c4a0 |

優先度: 黄 > 青 > 紫 > 緑 > 橙

## 代替情報（ツールチップ）

各パターンにツールチップ情報を追加:

| パターン | Alt | Note |
|---|---|---|
| 固定ドライブレター | Environ$("TEMP") や ThisWorkbook.Path で動的に解決 | OneDrive 移行でパスが変わる |
| UNC パス | 設定ファイルや CustomDocumentProperties で外部化 | サーバー移行でパスが変わる |
| ユーザーフォルダ | Environ$("USERPROFILE") で動的に取得 | ユーザー名はハードコードしない |
| Desktop / Documents | WScript.Shell.SpecialFolders または Environ$ で取得 | 日英環境でフォルダ名が異なる |
| ThisWorkbook.Path | OneDrive 同期時に URL 形式になる場合がある | SharePoint/OneDrive 環境で要検証 |
| Workbooks.Open | 相対パスやThisWorkbook.Path ベースに変更 | 固定パスでの外部参照は移行で壊れる |
| 固定プリンタ名 | ActivePrinter を動的取得するか設定ファイル化 | 新環境にプリンタが存在しない可能性 |
| 接続文字列 | 設定ファイルに外部化。Provider バージョンを確認 | ACE/Jet のバージョン違いに注意 |
| Outlook 連携 | 動作自体は可能だが CreateObject のテストが必要 | 新環境の Outlook バージョンを確認 |
| Word 連携 | 同上 | |
| Access / DB 連携 | DAO → ADO 移行を検討。ACE Provider の存在確認 | 64bit 環境で Provider が異なる |
| PDF 出力 | ExportAsFixedFormat は通常動作する | プリンタドライバ依存の場合あり |
| 印刷 | プリンタ名のハードコードを確認 | 新環境のプリンタ構成を確認 |
| 外部 EXE 起動 | EDR でブロックされる可能性が高い | Shell と同等のリスク |
