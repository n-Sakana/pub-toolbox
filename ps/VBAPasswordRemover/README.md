# VBAPasswordRemover (PowerShell版)

VBA プロジェクトのパスワード保護を解除する BAT + PowerShell スクリプト。

## 使い方

対象の Excel ファイルを `VBAPasswordRemover.bat` にドラッグ＆ドロップするだけ。

処理完了後、対象ファイルを開くと「不正なキー 'DPx' を含んでいます」と表示されるので「はい」をクリック。その後:
1. VBE (Alt+F11) を開く
2. ツール > VBAProject のプロパティ > 保護タブ
3. パスワード欄を空にして OK
4. ファイルを保存

## 対応形式

- `.xls` (OLE2 Compound Document)
- `.xlsm` / `.xlam` (ZIP ベースの OOXML)

## 仕組み

Excel COM で `.xlsm` / `.xlam` を一時的に `.xls` (OLE2) に変換し、バイナリ内のパスワードハッシュ (`DPB=`) を無効化してから元の形式で保存し直す。元ファイルのバックアップ (`.bak`) が自動作成される。

## 注意

自分が管理するファイルのパスワード回復用途を想定しています。
