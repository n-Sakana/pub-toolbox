# VBA移行調査・検証方針まとめ

## 位置づけ

今回の整理では、作業を三層に分ける。

- toolkit側: 対象ブックの静的解析と資産棚卸し
- 調査マクロ側: 新環境の基盤能力チェック
- 必要に応じた後続工程: 対象ブックの適合性検証

調査マクロは「業務マクロの互換性確認」ではなく、「新環境で何が使えるか」「どこで落ちるか」を別冊で先に測るためのものとする。対象ブックを直接いじらず、環境側の前提条件を先に切り分ける。

## toolkit側で見るもの

### 既存の主軸

EDRリスク:
- Win32 API宣言
- Shell / cmd / PowerShell / WScript
- CreateObject / GetObject
- WMI
- HTTP通信
- ファイル入出力
- レジストリ
- SendKeys
- クリップボード
- 自動実行
- 難読化っぽい記述

互換性リスク:
- PtrSafe欠落
- LongPtr見直し候補
- VarPtr / ObjPtr / StrPtr
- IE Automation
- 古い ActiveX / OCX
- DAO
- GoSub
- While/Wend

### 追加で厚く見るもの

環境依存:
- UNCパス
- 固定ドライブレター
- `C:\Users\`
- Desktop
- Documents
- AppData
- Program Files
- 固定プリンタ名
- 固定端末名
- 社内URL
- localhost
- 固定IP
- ポート番号
- 接続文字列
- 固定メールアドレス
- 共有フォルダ名

保存先・共有方式変更:
- 共有フォルダ依存
- ドライブレター依存
- UNCパス依存
- OneDrive / SharePoint 運用注意
- `ThisWorkbook.Path`
- `ActiveWorkbook.Path`
- `Dir`
- `FileDialog`
- `Workbooks.Open`
- `Open ... For Input/Output`
- 外部ブック参照

業務依存:
- Outlook連携
- Word連携
- Access / DB連携
- CSV / テキスト大量処理
- 外部EXE起動
- PDF出力
- 印刷
- ネットワーク共有
- フォーム / ActiveX UI

### Win32 APIの扱い

Win32 APIは一括りにせず、重みを分ける。

- 軽微代替可能
- 64bit修正で通る可能性あり
- 外部連携あり要調査
- GUI操作依存で再構築寄り

特に以下のようなGUI操作系は、単なる「要修正」ではなく、移行不能寄りとして扱う。

- `FindWindow`
- `SendMessage`
- `PostMessage`
- `GetWindowText`
- `SetForegroundWindow`
- `keybd_event`
- `mouse_event`

この種のものは「VBAマクロとしての単純移行は不可。代替方式での再構築が必要」と表現する。

出力上も、単なる件数だけでなく以下のような判定列を持たせると会議で使いやすい。

- `RiskLevel = Low / Medium / High`
- `PrimaryConcern = API / COM / DB / Process / Network / File / Mail / Mixed / StorageMigration`
- `MigrationClass = そのまま可 / 軽微修正 / 要代替実装 / 再構築必要 / 要保存先見直し`
- `NeedsReviewBy = Infra / DB / Security / BusinessOwner`

SharePoint / OneDrive への移行は、単なるファイルパス問題ではなく、業務マクロの置き場と開き方の前提変更として扱う。特に以下は独立観点として重く見る。

- ブラウザで開くとマクロが実行できない
- OneDrive 同期先パスがユーザーごとに異なる
- 同期状態によりファイルの存在前提が揺れる
- 共有サーバ前提のテンプレート参照や外部ブック参照が不安定になる
- `ThisWorkbook.Path` や相対パス前提の処理が崩れる

大分類としては、少なくとも以下の4本柱に分けると実務上わかりやすい。

- EDR / セキュリティ制約
- 64bit / 互換性
- 外部GUI操作依存
- 保存先・共有方式変更

## 調査マクロ側で見るもの

目的は、対象ブックが動くかではなく、新環境の前提能力があるかを測ること。

### 必須に近い項目

端末情報:
- 端末名
- ユーザー名
- OS
- Officeバージョン
- Excelバージョン
- 32/64bit
- VBAバージョン
- 実行日時

VBA実行基盤:
- 標準モジュール実行可否
- クラス生成可否
- UserForm表示または生成可否

VBIDE周辺:
- 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」有効可否
- VBProjectに触れるか

参照ライブラリ:
- 参照一覧の列挙
- Missing参照の有無
- 主要ライブラリの利用可否
  - Scripting
  - DAO
  - ADO
  - MSForms
  - Outlook
  - Word

COM生成テスト:
- Outlook.Application
- Word.Application
- Scripting.FileSystemObject
- ADODB.Connection
- ADODB.Recordset
- WScript.Shell
- MSXML2.XMLHTTP
- WinHttp.WinHttpRequest.5.1

ActiveX / フォーム系:
- MSForms.UserForm
- TreeView / ListView系
- Calendar系
- WebBrowser系
- CommonDialog系

ファイルアクセス:
- ローカル一時フォルダへの書き込み
- 削除
- 上書き
- 文字コード付きテキスト出力
- CSV出力
- 共有フォルダの存在確認
- 共有フォルダへの読み書き
- OneDrive 同期フォルダの場所確認
- 同期済みファイルへの読み書き可否
- `ThisWorkbook.Path` がどう見えるか
- 外部ファイルを相対的に開けるか
- ブラウザ起点ではなくデスクトップアプリで開いているか

通信:
- HTTP GET
- HTTPS GET
- プロキシ影響の有無
- 名前解決
- 固定URLへの疎通
- 必要ならポート疎通

Office連携:
- Outlook下書き作成可否
- Word起動可否
- Excelから他Office COMを触れるか

外部実行:
- PowerShell起動可否
- cmd起動可否
- WScript / CScript起動可否

印刷・PDF:
- 既定プリンタ取得可否
- PDF出力可否
- プリンタ一覧取得可否

DB接続:
- DAO / ADO オブジェクト生成可否
- DSN / OLEDB Provider の存在確認

セキュリティ制約:
- マクロ実行ポリシーの影響
- ActiveX制限
- Protected View影響
- OneDrive / ネットワーク保護の影響

### 今回の案件で優先度を落としてよいもの

実態としてHTTP通信やメール送信が無いなら、そこは最初から本気で回さなくてよい。今の案件では、以下を優先すれば十分。

- Officeビット数
- 参照ライブラリ
- COM生成可否
- ファイルアクセス
- 共有フォルダ
- 印刷
- VBIDE可否

## 出力方針

### 調査マクロ

人間向けの長文より、まずCSVを優先する。固定列で縦持ちにする。

例:
- 端末情報
- テスト名
- 対象
- 結果
- エラー番号
- エラーメッセージ
- 所要時間
- 備考

横持ちで一端末一行にすると見やすいが、項目追加で崩れやすい。最初は縦持ち、必要なら集計側で横に倒す。

### toolkit側

件数だけでなく、代表的な証拠行や代表カテゴリも出すとよい。

例:
- `TopApiNames = Sleep; FindWindow`
- `TopComProgIds = Outlook.Application; Scripting.FileSystemObject`
- `SampleEvidence = Module1.bas:L42 CreateObject("Outlook.Application")`

## マクロファイル以外に必要なもの

最低限あると便利なもの:
- 書き込み確認用の一時フォルダ
- 存在確認用の共有フォルダ
- 疎通確認用の安全なURL
- 必要なら社内プロキシ前提のURL
- 結果出力先

案件によっては、以下を外部設定ファイルに持つ。
- 代表的な ActiveX の存在確認リスト
- 主要 COM ProgID 一覧
- 確認対象の共有パス一覧
- 確認対象URL一覧
- 確認対象プリンタ名一覧

構成としては、以下の4点くらいあると回しやすい。
- 調査マクロ本体
- 設定ファイル
- 結果出力先
- 疎通確認先や共有パスなどの調査対象一覧

## 副作用なしの原則

調査マクロは以下を守る。

- 作成してすぐ消せる
- 下書きまでで送信しない
- 読み取り中心
- 接続しても更新しない
- 外部起動しても無害

ここが崩れると、調査マクロ自体が危険物になる。

## 実行テストについての整理

「新環境で動くか」の確認は価値が大きいが、いきなり対象ブックを全面実行するのは重くて危ない。

段階としては以下が自然。

1. 静的解析
   - 新環境で失敗しやすい点を予測する
2. 適合性検証
   - 開けるか
   - 参照がMissingにならないか
   - コンパイルできるか
   - UserFormロードで死なないか
3. 限定実行
   - 副作用を抑えたうえで、必要な入口だけ試す

ただし今回の整理では、まずは調査マクロによる基盤能力チェックを先に置く。

## 現時点の結論

- toolkitは単なるコード抽出器ではなく、静的解析と資産棚卸しの基盤としてかなり筋が良い
- 次に効くのは検知対象の追加そのものより、判定列と会議向けの整理
- 調査マクロは、対象ブックを直接いじる前に新環境の基盤能力を測るための別冊として有効
- HTTPやメールが実態として無いなら優先度を落とし、Officeビット数、参照、COM、ファイル、共有、印刷、VBIDEを優先する
- SharePoint / OneDrive への保存先移行は独立の大分類として扱い、固定パスや相対参照の影響を重点確認する
- GUI操作系の Win32 API 依存マクロは、単純移行ではなく再構築前提で扱う

要するに、
- toolkit側で、ブックの静的リスクと依存を洗う
- 調査マクロ側で、新環境の基盤能力を実測する
- 外部設定と確認先で、案件ごとの差分を吸収する

この三段構えにしておくと、「コードが危ないのか」「端末が足りないのか」「接続先がないのか」を切り分けられる。そこまで分かれば、移行会議はかなり進む。