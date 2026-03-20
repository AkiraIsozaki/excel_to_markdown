# タスクリスト

## 🚨 タスク完全完了の原則

**このファイルの全タスクが完了するまで作業を継続すること**

---

## フェーズ1: データモデル拡張（hyperlink）

- [x] models.py: `RawCell` に `hyperlink: str | None` フィールド追加（デフォルト None、末尾）
- [x] models.py: `TextBlock` に `hyperlink: str | None` フィールド追加（デフォルト None）
- [x] models.py: `DocElement` に `hyperlink: str | None` フィールド追加（デフォルト None）

## フェーズ2: ハイパーリンク変換パイプライン

- [x] xlsx_reader.py: `read_sheet` で `cell.hyperlink.target` を `RawCell.hyperlink` に設定
- [x] merge_resolver.py: `resolve` で `cell.hyperlink` を `TextBlock.hyperlink` に転送
- [x] structure_detector.py: `DocElement` 生成時に `block.hyperlink` を `DocElement.hyperlink` に転送
- [x] markdown_renderer.py: `render_element` で `el.hyperlink` があれば `[text](url)` 形式に変換

## フェーズ3: アトミック書き込み

- [x] cli.py: `_write_output` を tmpファイル → rename のアトミック書き込みに変更

## フェーズ4: .xls 対応

- [x] reader/xls_reader.py: 新規作成（xlrd 2.x でシートから list[RawCell] を抽出）
  - [x] セル値の文字列化
  - [x] 結合セルの処理（merged_cells）
  - [x] フォントプロパティ（bold/italic/strike/underline/size）
  - [x] 背景色（best effort）
- [x] cli.py: `.xls` ファイルを xls_reader 経由でパイプラインに通す分岐を追加

## フェーズ5: バッチ変換

- [x] cli.py: `run()` にディレクトリ入力モードを追加
  - [x] `input_path.is_dir()` の場合、配下の xlsx/xls を収集
  - [x] 各ファイルを順番に変換（エラー時は stderr に出力して継続）

## フェーズ6: テスト追加・修正

- [x] test_integration.py: `test_performance_1000rows` 追加（30秒以内）
- [x] test_cli.py: アトミック書き込みのテスト更新（既存の PermissionError テストを確認）
- [x] test_cli.py: バッチ変換テスト追加
- [x] 全テスト実行 → 全パス・カバレッジ80%以上を確認

---

## 実装後の振り返り

### 実装完了日
2026-03-20

### 計画と実績の差分

**計画と異なった点**:
- `xlrd.open_workbook` に `formatting_info=True` が必要だった（フォント情報取得のため）
- テスト用 xls ファイル生成に xlwt が必要で、環境にインストールして対応
- xls テスト用に `test_xls_reader.py` を別ファイルとして新規作成（設計書では想定外）

**新たに必要になったタスク**:
- `test_xls_reader.py` の新規作成（xls_reader.py の直接ユニットテスト）

### 学んだこと

- xlrd 2.x は `.xls` 専用で、フォント・背景色情報の取得に `formatting_info=True` が必須
- xlwt はフォーマット付き xls ファイルを生成できる唯一の選択肢（Python 3 環境）
