# VBAPasswordRemover (PowerShell版)

VBA プロジェクトのパスワード保護を解除する BAT + PowerShell スクリプト。

## 使い方

対象の Excel ファイルを `VBAPasswordRemover.bat` にドラッグ＆ドロップするだけ。

処理完了後、対象ファイルを開いて以下の手順で完全解除:
1. VBE (Alt+F11) を開く
2. ツール > VBAProject のプロパティ > 保護タブ
3. パスワード欄を空にして OK
4. ファイルを保存

## 対応形式

- `.xls` (OLE2 Compound Document)
- `.xlsm` / `.xlam` (ZIP ベースの OOXML)

## 仕組み

VBA プロジェクト内のパスワードハッシュ (`DPB=`) を無効化する。元ファイルのバックアップ (`.bak`) が自動作成される。

## 注意

自分が管理するファイルのパスワード回復用途を想定しています。
