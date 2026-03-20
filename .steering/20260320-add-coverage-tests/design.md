# 設計書

## アーキテクチャ概要

既存テストと同じパターン（pytest クラスベース）で新しいテストファイルを追加する。
実際のファイルI/Oが必要な箇所は `tmp_path` フィクスチャ（pytest組み込み）を使用する。

## コンポーネント設計

### 1. tests/test_cli.py（新規）

**責務**:
- `parse_args()` の引数パース検証
- `run()` の正常系・異常系（各エラーパス）
- 内部ヘルパー関数（`_validate_input`, `_resolve_output_path`, `_select_sheets`, etc.）

**実装の要点**:
- ファイルI/Oは `tmp_path` フィクスチャで一時ディレクトリを使用
- 実際の xlsx ファイルは openpyxl でオンメモリ生成して保存
- `run()` は exit code（int）を返すので、直接呼び出してアサート

**テストケース設計**:

| テストクラス | カバーするコード |
|---|---|
| `TestParseArgs` | `parse_args()` - 各引数の組み合わせ |
| `TestValidateInput` | `_validate_input()` - 存在しないファイル、非対応拡張子 |
| `TestResolveOutputPath` | `_resolve_output_path()` - output指定あり/なし |
| `TestSelectSheets` | `_select_sheets()` - 全シート、名前指定、インデックス指定、未存在 |
| `TestRunSuccess` | `run()` 正常系 - 単一シート、複数シート、debug mode |
| `TestRunErrors` | `run()` 異常系 - ファイル未存在、非対応形式、シート未存在、空シート |
| `TestWriteOutput` | `_write_output()` - 書き込み成功・失敗 |

### 2. tests/test_merge_resolver.py への追記

**責務**:
- `to_inline_runs()` の各パスをカバー

**テストケース**:
- 非リッチテキスト（通常文字列）→ 空リスト返却
- `CellRichText` 型の文字列パーツ → `InlineRun(text=...)` 変換
- `CellRichText` 型の `TextBlock` パーツ（bold/italic/strike/underline）→ 書式付き `InlineRun`

## テスト戦略

- ファイルI/O: `tmp_path` (pytest fixture)
- openpyxl xlsx 生成: `conftest.py` の `save_and_reload` パターンを踏襲
- モック: 使わない（実ファイルで統合的にテスト）

## ディレクトリ構造

```
tests/
  test_cli.py          # 新規
  test_merge_resolver.py  # to_inline_runs テスト追加
```

## 実装の順序

1. `tests/test_cli.py` 新規作成
2. `tests/test_merge_resolver.py` に `to_inline_runs` テスト追加
3. pytest でカバレッジ確認
