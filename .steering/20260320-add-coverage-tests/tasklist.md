# タスクリスト

## 🚨 タスク完全完了の原則

**このファイルの全タスクが完了するまで作業を継続すること**

---

## フェーズ1: test_cli.py の作成

- [x] `TestParseArgs`: parse_args() の各引数をテスト
  - [x] デフォルト値の確認
  - [x] `--output`, `--sheet`, `--base-font-size`, `--debug` オプション
- [x] `TestValidateInput`: 存在しないファイル・非対応拡張子
- [x] `TestResolveOutputPath`: output指定あり/なし
- [x] `TestSelectSheets`: 全シート・名前指定・インデックス指定・未存在エラー
- [x] `TestRunSuccess`: run() 正常系（単一シート・複数シート・debugモード・空シートスキップ）
- [x] `TestRunErrors`: run() 異常系（ファイル未存在・非対応形式・シート未存在）
- [x] `TestWriteOutput`: 書き込み成功・OSError時のPermissionError変換

## フェーズ2: merge_resolver の to_inline_runs テスト追加

- [x] 非リッチテキスト値（str）→ 空リストを返すこと
- [x] CellRichText の文字列パーツ → InlineRun 変換
- [x] CellRichText の TextBlock パーツ（bold/italic/strike/underline）→ 書式付き InlineRun

## フェーズ3: 品質チェック

- [x] `pytest` 実行 → 全テストパス確認（164 passed）
- [x] カバレッジ80%以上を確認
  - [x] cli.py 89%
  - [x] merge_resolver.py 100%
  - [x] 全体 93%

---

## 実装後の振り返り

### 実装完了日
2026-03-20

### 計画と実績の差分

**計画と異なった点**:
-

**新たに必要になったタスク**:
-

### 学んだこと

-
