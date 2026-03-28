# VBACodeExtractor

Excel を開かずに VBA ソースコードをテキストファイルとして抽出する PowerShell スクリプト。

## 用途

- EDR が厳しい環境で VBA コードを確認・改修する前段階
- Win32 API (`Declare PtrSafe Function`) や Shell 関数の使用箇所を特定
- マクロ資産の棚卸し・移行前の事前調査

## 使い方

対象の Excel ファイルを `ps/VBACodeExtractor.bat` にドラッグ＆ドロップ。

同じフォルダに `<ファイル名>_vba/` ディレクトリが作成され、各モジュールが `.bas` / `.cls` / `.frm` ファイルとして出力される。

### Win32 API の使用チェック

抽出後に grep で確認:

```bat
findstr /i "Declare" *_vba\*.bas *_vba\*.cls
findstr /i "Shell" *_vba\*.bas *_vba\*.cls
```

## 対応形式

- `.xlsm` / `.xlam` (ZIP ベースの OOXML)
- `.xls` (OLE2 Compound Document)

## 仕組み

Excel COM を使わず、バイナリレベルで処理:
1. `.xlsm` → ZIP から `vbaProject.bin` を抽出
2. OLE2 Compound Document を解析
3. `PROJECT` ストリームからモジュール一覧を取得
4. 各モジュールの圧縮ソースコードを MS-OVBA 仕様に基づき展開
