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
| 固定端末名/IP | ハードコードされたIP | `(?mi)^[^'\r\n]*"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"` |
| localhost | localhost 参照 | `(?mi)^[^'\r\n]*\blocalhost\b` |
| 接続文字列 | Provider= や DSN= | `(?mi)^[^'\r\n]*(Provider\s*=|DSN\s*=|Data\s+Source\s*=)` |
| 外部ブック参照 (リテラル) | Workbooks.Open に文字列リテラル引数 | `(?mi)^[^'\r\n]*\bWorkbooks\.Open\s*\(\s*"` |

**注: 以下はリスクではなく情報提供（Info）として検知する:**

| パターン名 | 検知対象 | 正規表現 | 理由 |
|---|---|---|---|
| ThisWorkbook.Path | パス依存処理 | `(?mi)^[^'\r\n]*\bThisWorkbook\.Path\b` | 推奨されるパス取得方法だが、SharePoint/OneDrive 環境ではローカルパスにならない可能性がある。確認を促す目的 |
| ActiveWorkbook.Path | パス依存処理 | `(?mi)^[^'\r\n]*\bActiveWorkbook\.Path\b` | 同上 |

Info 項目はハイライトせず、テキストレポートにのみ出力。CSV には `InfoCount` 列で件数を記録。

### カテゴリ4: 業務依存リスク（橙ハイライト `#4a3a1b`）

| パターン名 | 検知対象 | 正規表現 |
|---|---|---|
| Outlook 連携 | Outlook COM | `(?mi)^[^'\r\n]*\bOutlook\.Application\b` |
| Word 連携 | Word COM | `(?mi)^[^'\r\n]*\bWord\.Application\b` |
| Access / DB 連携 | Access COM / DAO / ADO | `(?mi)^[^'\r\n]*\b(Access\.Application|CurrentDb|DoCmd)\b` |
| PDF 出力 | ExportAsFixedFormat | `(?mi)^[^'\r\n]*\.ExportAsFixedFormat\b` |
| 印刷 | PrintOut / PrintPreview | `(?mi)^[^'\r\n]*\.(PrintOut|PrintPreview)\b` |
| 外部 EXE 起動 | Shell で exe 起動 | `(?mi)^[^'\r\n]*\bShell\s.*\.exe` |

### Environ$ の扱い

`Environ$` は既存の EDR パターンで検知されるが、環境依存パターンの推奨修正方法でもある（例: 固定ドライブレター → `Environ$("TEMP")`）。

この矛盾に対する方針:
- EDR カテゴリの `Environ` パターンは **detect=true, sanitize=false** をデフォルトとする
- ツールチップに「Environ$ は VBA 標準関数で通常 EDR にブロックされない。検知は情報提供目的」と明記
- `Environ$` への置き換えを推奨する他パターンのツールチップにも「EDR 検知に表示されるが問題ない」旨を記載

## 判定列（CSV 出力に追加）

### RiskLevel（技術危険度）

技術的にどれだけ危険か。「壊れるかどうか」の指標。

| 値 | 条件 |
|---|---|
| High | GUI操作系API (FindWindow, SendMessage, PostMessage, keybd_event, mouse_event) あり、または PowerShell/WScript あり |
| Medium | Win32 API あり（GUI操作系以外）、または DAO あり |
| Low | 上記に該当しない |

### MigrationClass（対応方針）

移行にあたって何をすべきか。RiskLevel とは異なる軸。**複数該当する場合はセミコロン区切りで全て記載。**

| 値 | 条件 |
|---|---|
| そのまま可 | 全カテゴリのリスクがすべて 0 |
| 軽微修正 | 互換性リスクのみ（PtrSafe 追加、While/Wend 等） |
| 要代替実装 | Win32 API（GUI操作系以外）、または DAO → ADO 移行 |
| 再構築必要 | GUI操作系 API 依存、または Shell/PowerShell 依存 |
| 要保存先見直し | 固定パス/UNC パス/共有フォルダ依存 |

例: `要代替実装; 要保存先見直し` — API の修正と保存先の見直しの両方が必要。

### PrimaryConcern（主要懸念）

重み付きで判定。件数ではなく、以下の優先順で最初にマッチしたものを採用:

1. GUI操作系 API → `GUI`
2. Shell / PowerShell → `Process`
3. 保存先移行関連 → `StorageMigration`
4. DB 連携 → `DB`
5. COM / 外部連携 → `COM`
6. ネットワーク → `Network`
7. メール → `Mail`
8. ファイル I/O → `File`
9. その他 → `Other`

複数に該当しても、最も重いもの1つを採用。

### NeedsReviewBy（確認担当）

**複数付与。** セミコロン区切り。

| 値 | 条件 |
|---|---|
| Security | EDR リスクあり |
| Infra | 環境依存（パス、プリンタ、接続文字列）あり |
| DB | DAO / ADO / 接続文字列あり |
| BusinessOwner | 業務依存（Outlook、Word、印刷）あり |
| Developer | 互換性リスクのみ（PtrSafe, DefType 等の純粋なコード修正） |

## CSV カラム（更新後）

既存: Timestamp, RelativePath, FileName, Bas, Cls, Frm, TotalModules, CodeLines, EdrIssues, CompatIssues, SanitizedLines, References, Error

追加: EnvIssues, BizIssues, InfoCount, RiskLevel, MigrationClass, PrimaryConcern, NeedsReviewBy, TopApiNames, TopComProgIds, SampleEvidence

### TopApiNames
検出された API 宣言名を重みの重い順に先頭3件。セミコロン区切り。
重み: GUI操作系 > その他 API。同重みなら出現順。

### TopComProgIds
検出された COM ProgID を先頭3件（出現順）。セミコロン区切り。

### SampleEvidence
最も重い検出の代表行を1件。重み順で最初のもの。
例: `Module1.bas:L42 CreateObject("Outlook.Application")`

## HTML ハイライト色（5色 + Info）

| カテゴリ | 色 | 背景 | テキスト |
|---|---|---|---|
| サニタイズ済み | 黄 | #4b3a00 | #f0d870 |
| EDR リスク | 青 | #1b2e4a | #a0c4f0 |
| 互換性リスク | 紫 | #3a1b4a | #c4a0f0 |
| 環境依存リスク | 緑 | #1b3a2a | #a0f0c4 |
| 業務依存リスク | 橙 | #4a3a1b | #f0c4a0 |
| Info（参考） | ハイライトなし | — | — |

優先度: 黄 > 青 > 紫 > 緑 > 橙

## 代替情報（ツールチップ）

| パターン | Alt | Note |
|---|---|---|
| 固定ドライブレター | Environ$("TEMP") や ThisWorkbook.Path で動的に解決 | OneDrive 移行でパスが変わる。Environ$ は EDR 検知に表示されるが通常ブロックされない |
| UNC パス | 設定ファイルや CustomDocumentProperties で外部化 | サーバー移行でパスが変わる |
| ユーザーフォルダ | Environ$("USERPROFILE") で動的に取得 | ユーザー名はハードコードしない |
| Desktop / Documents | Environ$("USERPROFILE") & "\Desktop" で取得 | 日英環境でフォルダ名が異なる。WScript.Shell.SpecialFolders は EDR リスクあり |
| ThisWorkbook.Path (Info) | 通常は問題ないが確認推奨 | SharePoint/OneDrive 環境では期待したローカルパスにならない可能性がある |
| ActiveWorkbook.Path (Info) | 同上 | 同上 |
| 外部ブック参照 (リテラル) | 相対パスや ThisWorkbook.Path ベースに変更 | 固定パスでの外部参照は移行で壊れる |
| 固定プリンタ名 | ActivePrinter を動的取得するか設定ファイル化 | 新環境にプリンタが存在しない可能性 |
| 固定端末名/IP | 設定ファイルに外部化 | 環境移行で接続先が変わる |
| 接続文字列 | 設定ファイルに外部化。Provider バージョンを確認 | ACE/Jet のバージョン違いに注意 |
| Outlook 連携 | 動作自体は可能だが新環境でテストが必要 | Outlook バージョンを確認 |
| Word 連携 | 同上 | |
| Access / DB 連携 | DAO → ADO 移行を検討。ACE Provider の存在確認 | 64bit 環境で Provider が異なる |
| PDF 出力 | ExportAsFixedFormat は通常動作する | プリンタドライバ依存の場合あり |
| 印刷 | プリンタ名のハードコードを確認 | 新環境のプリンタ構成を確認 |
| 外部 EXE 起動 | EDR でブロックされる可能性が高い | Shell と同等のリスク |
