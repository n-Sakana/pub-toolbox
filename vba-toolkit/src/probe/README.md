# Environment Probe

toolkit が検知するパターンが新環境で実際に動くかを確認する使い捨てマクロ。

## 使い方

1. 新しい空の `.xlsm` を作成
2. VBE (Alt+F11) で `Probe.bas` をインポート
3. `Alt+F8` → `Probe_Run` を実行
4. ダイアログで Basic / Basic+Extended を選択
5. 結果が `probe_result_<PC名>_<日時>.txt` に出力される

## Basic テスト

COM 生成、ファイル I/O、レジストリ、Environ、クリップボード、VarPtr、DAO、レガシーコントロール等。安全。

## Extended テスト

Win32 API、Shell、PowerShell、DDE、IE 等。EDR アラートを招く可能性あり。デフォルト無効。

## SendKeys について

呼び出し可否の確認のみ。空文字列で副作用なし。環境によるブロック有無の実測としては限定的。
