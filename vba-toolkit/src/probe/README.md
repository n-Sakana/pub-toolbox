# Environment Probe

toolkit が検知するパターンが新環境で実際に動くかを PowerShell + Excel COM でテストする。

## 使い方

`Probe.bat` をダブルクリック。B (Basic) または E (Extended) を選択。

## 仕組み

VBA ファイルではなく PowerShell スクリプト。Excel COM で一時的な xlsm を生成し、テスト対象のコードを動的に注入して:

1. **保存できるか** — EDR が Declare 文をブロックするかを検出
2. **実行できるか** — マクロを Run して結果を取得
3. **結果が正しいか** — 戻り値を検証

テストごとに独立した xlsm を作成・破棄するため、1つの失敗が他に波及しない。

## Basic / Extended

- **Basic**: COM 生成、ファイル I/O、レジストリ、VBA注入テスト。安全。
- **Extended**: Shell、PowerShell、WMI、DDE。EDR アラートの可能性あり。
