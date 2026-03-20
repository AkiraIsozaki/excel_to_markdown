# タスクリスト: 変換パイプライン基盤

## セットアップ

- [x] T01: pyproject.toml を作成する
- [x] T02: requirements.txt / requirements-dev.txt を作成する
- [x] T03: .gitignore を作成する
- [x] T04: パッケージディレクトリと __init__.py を作成する（excel_to_markdown/, reader/, parser/, renderer/, tests/）

## データモデル

- [x] T05: excel_to_markdown/models.py を実装する（RawCell, TextBlock, InlineRun, DocElement, TableElement, TableCell, ElementType）

## 読み込みレイヤー

- [x] T06: excel_to_markdown/reader/xlsx_reader.py を実装する（read_sheet, extract_font_props, extract_bg_color, get_print_area）

## 空間解析レイヤー

- [x] T07: excel_to_markdown/parser/cell_grid.py を実装する（CellGrid dataclass: baseline_col, col_unit, modal_row_height, is_empty_row）
- [x] T08: excel_to_markdown/parser/merge_resolver.py を実装する（resolve, to_inline_runs）

## テスト

- [x] T09: tests/conftest.py を作成する（xlsx_builder フィクスチャ）
- [x] T10: tests/test_xlsx_reader.py を実装する（結合セル/フォント/印刷領域/非表示行列）
- [x] T11: tests/test_cell_grid.py を実装する（baseline_col, col_unit, modal_row_height）
- [x] T12: tests/test_merge_resolver.py を実装する（空白スキップ/InlineRun変換/ソート）

## 申し送り事項

### 実装完了日
2026-03-20

### 計画と実績の差分

- 計画通り全12タスクを完了
- `test_hidden_col_excluded` が1件失敗 → `ColumnDimension` に `.column` 属性がなく、`column_index_from_string(cd)` で修正。openpyxl の `column_dimensions` のキーは列文字 (str) であることに注意

### 学んだこと

- openpyxl の `column_dimensions` のキーは列文字 ("A", "B"…)。列番号への変換には `column_index_from_string()` を使う
- `merge_resolver.py` の `to_inline_runs()` は openpyxl のリッチテキスト型に依存するが、`xlsx_reader.py` が既に文字列化しているため現状では未使用パス。次フェーズでリッチテキスト対応を強化する場合は reader 側の変換処理を見直す必要あり
- `merge_resolver.py` の `to_inline_runs` のカバレッジが 50% と低い。次フェーズで openpyxl リッチテキストフィクスチャを使ったテストを追加する

### 次フェーズへの引き継ぎ

- 次回: `/add-feature 構造検出` — `parser/structure_detector.py` + `parser/table_detector.py` の実装
- `CellGrid` は `cli.py` または reader 呼び出し後に `col_widths` / `row_heights` を openpyxl から取得して渡す設計。実装例:
  ```python
  col_widths = {ws.column_dimensions[l].column: ws.column_dimensions[l].width or 8.0 for l in ws.column_dimensions}
  row_heights = {r: ws.row_dimensions[r].height or 15.0 for r in ws.row_dimensions}
  ```
  ※ `column_dimensions` キーは列文字であることに注意し、`column_index_from_string()` で変換すること
- `merge_resolver.py` の `to_inline_runs` はリッチテキスト対応時に reader との連携を再設計する
