# vba-toolkit 既存ツールレビュー メモ

## 総評

vba-toolkit はかなり筋が良い。特に以下が強い。

- Excel を開かずに VBA を読むという軸が一貫している
- `Analyze` を主軸に据えて、検知・サニタイズ・移行ガイドを一本化している
- `VBAToolkit.psm1` に OLE2 / 圧縮展開 / 分析エンジン / 代替DB を集約しており、知識と処理の共通化が進んでいる
- HTML 出力が実務のレビューにかなり向いている

端的に言えば、個別スクリプト集ではなく、もう小さな解析基盤になっている。

ただし、良い分だけ、いくつか構造的な危うさが見える。特に `Analyze.ps1` と `Extract.ps1` は、今後の拡張を考えると少し繊細。

## 1. Extract.ps1

### 良い点

- ファイル/フォルダ両対応
- `output/` を自動生成し、実行導線が単純
- モジュールインデックス付き `combined.txt` はレビュー用途に向いている
- `output/` 配下を再帰対象から除外しているのは実務的

### 気になる点

#### 1-1. 複数ファイル時に modules/ が衝突する

現在は実行全体で単一の `outDir` を作り、その中の `modules/` に全ファイルのモジュールを書いている。

そのため、複数ファイルに同名モジュールがあると上書きされる。

例:
- A.xlsm の `Module1.bas`
- B.xlsm の `Module1.bas`

後から処理したほうが残る。これはかなり危ない。

対処案:
- ファイルごとに `modules/<baseName>/` を切る
- あるいは出力ファイル名を `<baseName>__<moduleName>.<ext>` にする

前者のほうが自然。

#### 1-2. combined.txt も複数ファイル時に単一ファイルへ上書きされる

`combined.txt` が実行ごとに1つしかないため、複数ファイル処理時は最後のファイルの内容だけが残る。

対処案:
- `${baseName}_combined.txt` にする
- あるいはファイルごとサブフォルダを切る

#### 1-3. allFiles の取得元が modulesDir 全体になっている

`$allFiles = Get-ChildItem $modulesDir -File` だと、複数ファイル実行時に前のファイルのモジュールも混ざる。

結果として、
- module index
- total lines
- combined source

がファイル単位で正しくなくなる。

対処案:
- そのファイルで出力したパス一覧を都度保持する
- `Get-ChildItem` に頼らず、書き出したモジュールの配列をそのまま使う

#### 1-4. 経過時間表示がファイル単位で累積になっている

`Write-VbaResult` に渡している秒数が全体 Stopwatch の経過時間なので、各ファイルの所要時間ではなく累積時間になる。

対処案:
- ファイル単位の Stopwatch を別に持つ

### まとめ

Extract は単体利用ではかなり良いが、複数ファイル処理で出力の整合性が崩れる。ここは早めに直したほうがよい。

## 2. Analyze.ps1

### 良い点

- 現在の toolkit の中心としてかなり良い
- 検知、サニタイズ、HTML、CSV が一本の流れにまとまっている
- HTML の 3 カラム構成、アウトライン、ツールチップはレビュー用途に強い
- API 宣言だけでなく呼び出し側も追っているのは実務的
- サニタイズ済みコードをそのまま表示に反映しているのも分かりやすい

### 気になる点

#### 2-1. Analyze.ps1 が太り始めている

現状でも、
- 検知
- サニタイズ
- レポート生成
- HTML 構築
- CSV 出力

を1ファイルで抱えている。今後、環境依存や業務依存、判定列、Info 項目まで入ると、かなり重くなる。

対処案:
- 内部関数に分ける
  - `Invoke-SanitizePass`
  - `Build-AnalyzeTextReport`
  - `Build-AnalyzeHtml`
  - `Build-AnalyzeCsvRow`
- 表向きは一本の Analyze のままでよいが、中を分けたほうが保守しやすい

#### 2-2. highlight 構築が現在の3カテゴリ前提に寄っている

今は `hl-sanitized / hl-edr / hl-compat` 前提でかなり素直に書かれている。

だが今後、
- 環境依存（緑）
- 業務依存（橙）
- Info（ハイライトなし）

が入ると、
- 優先順位
- 色クラス
- アウトライン色
- ツールチップ対象

の分岐が一気に増える。

対処案:
- highlight 情報を `@{ Color; Category; PatternName; Priority; TooltipKey }` のようにデータ化する
- CSS クラスの直書き分岐を減らす

#### 2-3. tooltipEntries の構築が少し危うい

`$analysis.ApiCallNames + $analysis.ApiDecls.Name + pattern-level replacements` を都度 JS 化しているが、
- 重複
- エスケープ漏れ
- 今後のカテゴリ追加

で壊れやすい。

対処案:
- tooltip データをいったん hashtable でユニーク化してから JS 化する
- `PatternName` と `TooltipKey` を分ける

#### 2-4. CSV がまだ旧列構成のまま

仕様書では今後、
- EnvIssues
- BizIssues
- InfoCount
- RiskLevel
- MigrationClass
- PrimaryConcern
- NeedsReviewBy
- TopApiNames
- TopComProgIds
- SampleEvidence

が入る予定。

今の CSV 出力部はまだ旧構成なので、拡張時にかなり大きく手を入れることになる。

対処案:
- CSV row を最初から ordered hashtable で組み立てる
- ヘッダ文字列を手書き連結するのではなく、列定義から生成する

#### 2-5. サニタイズ対象の意味づけに注意

実装としては良いが、出力上で「sanitized = 修正済み」と誤読される危険はある。

対処案:
- README / report 上で「危険箇所のコメントアウトであり、移行完了ではない」と明記する

### まとめ

Analyze はかなり強い。ただし今後の拡張を支えるには、内部関数化とデータ駆動化を少し進めたほうがよい。

## 3. Diff.ps1

### 良い点

- モジュール単位の added / removed / modified / unchanged の整理が分かりやすい
- サイドバイサイド HTML はかなり見やすい
- 大きいファイル向けに簡易 greedy diff に寄せているのは現実的
- 変更のあるモジュールを先に見せる導線も良い

### 気になる点

#### 3-1. LCS ではなく greedy なので、差分品質はケース依存

仕様上そうしているのは理解できるが、行のズレが大きいケースでは changed / added / removed の見え方が少し荒れる可能性がある。

これは欠点というよりトレードオフ。

対処案:
- README に「完全な LCS ではなく、大きい VBA 向けの軽量差分」と一言あると親切

#### 3-2. diff.txt がかなり薄い

今は要約しか出していない。

HTML が本体なのは分かるが、テキスト側にも
- 変更モジュール一覧
- 追加/削除/変更件数

くらいは欲しい。

対処案:
- modified module 名を列挙する
- added / removed module も出す

### まとめ

Diff は用途が明確で、今のままでも十分使える。大きな問題はない。改善余地はテキストレポートの薄さくらい。

## 4. Unlock.ps1

### 良い点

- 非破壊でコピーを作る
- xls と xlsm/xlam を分けて扱っている
- 一時 xls を経由して再保存する流れは実務的
- 最後に手動で完全解除する手順を出すのも親切

### 気になる点

#### 4-1. Excel COM の後始末が少し弱い

`$wb` の ReleaseComObject が無い。`$excel` は解放しているが、Workbook 側の参照が残る可能性がある。

対処案:
- `$wb` も finally で閉じて解放する
- 変数を `$null` に落とす

#### 4-2. tempXls のファイル名衝突可能性

秒単位タイムスタンプなので、同時実行で衝突余地がある。

対処案:
- GUID を足す

#### 4-3. README 上の位置づけはやや鋭い

実装の問題ではないが、`Extract / Analyze / Diff` と並ぶ主線ツールとしては少し温度差がある。

対処案:
- README 上では「補助ツール」扱いに寄せる

### まとめ

Unlock は用途が鋭いが、実装自体は筋が通っている。COM 解放だけ少し丁寧にしたい。

## 5. VBAToolkit.psm1

### 良い点

- 共通基盤としてかなり良い
- OLE2 / VBA 圧縮展開 / コードページ解析 / モジュール抽出 / 分析エンジンまで一通り揃っている
- `Get-VbaAnalysis` に検知ロジックを集約しているのは正解
- API 代替 DB をここに置いたのも良い

### 気になる点

#### 5-1. 役割が増えてきている

今はこの1モジュールに、
- OLE2 低レベル処理
- HTML テンプレート
- 分析エンジン
- API 代替 DB
- ログ出力

が全部入っている。

これは小さな基盤としては便利だが、今後さらに
- 環境依存パターン
- 業務依存パターン
- 判定列ロジック
- Info 項目

が入ると、少し重い。

対処案:
- 少なくとも概念上は分ける
  - `Core`（OLE2 / 圧縮展開）
  - `Analysis`（Get-VbaAnalysis）
  - `Html`（テンプレート）
  - `Replacements`（代替 DB）
- 物理分割まで急がなくてもよいが、関数配置の意識は持ちたい

#### 5-2. 代替 DB がかなり大きく、今後の保守点になる

今は便利だが、パターン名や API 名が増えるほど、
- 重複キー
- 命名揺れ
- パターン名変更時の不整合

が起きやすい。

対処案:
- pattern-level key と API-level key の命名規則を揃える
- 将来的には JSON / psd1 外出しも検討余地あり

### まとめ

VBAToolkit.psm1 は今の toolkit の心臓部としてかなり良い。ただし、今後の拡張に備えて“どこまで抱えるか”は少し意識したほうがよい。

## 優先度つき改善案

### 優先度: 高

1. Extract の複数ファイル出力衝突を直す
   - modules/ の衝突
   - combined.txt の上書き
   - allFiles 集計の混線

2. Analyze の内部関数化を始める
   - サニタイズ
   - text report
   - html generation
   - csv row

3. Analyze の highlight / tooltip をデータ駆動に寄せる
   - env/biz/info 追加前にやると楽

### 優先度: 中

4. CSV 出力を列定義ベースにする
   - 今後の判定列追加に備える

5. Diff の diff.txt を少し厚くする

6. Unlock の COM 解放を丁寧にする

### 優先度: 低

7. VBAToolkit.psm1 の概念分割を意識する

8. README 上で Unlock を補助ツール寄りに見せる

## ひとことで言うと

- Extract は単体では良いが、複数ファイル時に危ない
- Analyze はかなり強いが、拡張前に少し分けたい
- Diff は安定
- Unlock は鋭いが筋は通っている
- 共通モジュールは優秀だが、今後は抱え込みすぎに注意

今の vba-toolkit はかなり良い。すでに道具箱ではなく、小さな解析ポットです。
ただし茶葉を増やす前に、注ぎ口の漏れだけは先に塞いだほうがいい。
