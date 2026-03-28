# VBAPasswordRemover

VBA プロジェクトのパスワード保護を解除する。

## 対応形式

- `.xls` (OLE2 Compound Document)
- `.xlsm` / `.xlam` (ZIP ベースの OOXML)

## 方法

### PowerShell版 (推奨) — `ps/`

対象の Excel ファイルを `ps/VBAPasswordRemover.bat` にドラッグ＆ドロップ。

処理中に「不正なキー 'DPx' を含んでいます」というダイアログが表示されるので「はい」をクリック。

### VBA版 — `VBAPasswordRemover.bas`

1. VBE (Alt+F11) で `VBAPasswordRemover.bas` をインポート
2. `RemoveVBAPassword` マクロを実行
3. ファイル選択ダイアログで対象ファイルを選択

## 処理後の手順

1. 対象ファイルを開く
2. VBE (Alt+F11) を開く
3. ツール > VBAProject のプロパティ > 保護タブ
4. パスワード欄を空にして OK
5. ファイルを保存

## 仕組み

Excel で `.xlsm` / `.xlam` を一時的に `.xls` (OLE2) に変換し、バイナリ内のパスワードハッシュ (`DPB=`) を無効化してから元の形式で保存し直す。元ファイルのバックアップ (`.bak`) が自動作成される。

## 注意

自分が管理するファイルのパスワード回復用途を想定しています。
