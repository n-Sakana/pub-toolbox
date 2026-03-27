# VBAPasswordRemover

VBA プロジェクトのパスワード保護を解除する VBA マクロ。

## 対応形式

- `.xls` (OLE2 Compound Document)
- `.xlsm` / `.xlam` (ZIP ベースの OOXML)

## 仕組み

VBA プロジェクト内のパスワードハッシュ (`DPB=`) を無効化することで、保護を解除する。元ファイルのバックアップ (`.bak`) が自動作成される。

## 使い方

1. VBE (Alt+F11) で `VBAPasswordRemover.bas` をインポート
2. `RemoveVBAPassword` マクロを実行
3. ファイル選択ダイアログで対象ファイルを選択
4. 処理完了後、対象ファイルを開いて以下の手順で完全解除:
   - VBE (Alt+F11) を開く
   - ツール > VBAProject のプロパティ > 保護タブ
   - パスワード欄を空にして OK
   - ファイルを保存

## 注意

自分が管理するファイルのパスワード回復用途を想定しています。
