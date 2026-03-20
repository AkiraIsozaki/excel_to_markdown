# 設計: 変換パイプライン基盤

## 実装アプローチ

### ファイル作成順序

依存関係の上流から順に実装する:

```
pyproject.toml / 設定ファイル
    ↓
excel_to_markdown/__init__.py（バージョン定義）
    ↓
excel_to_markdown/models.py（全データモデル）
    ↓
excel_to_markdown/reader/xlsx_reader.py（RawCell生成）
    ↓
excel_to_markdown/parser/cell_grid.py（空間統計）
    ↓
excel_to_markdown/parser/merge_resolver.py（TextBlock生成）
    ↓
tests/（各モジュールのユニットテスト）
```

### 設計判断

#### models.py
- `RawCell`: `frozen=True` の dataclass。値の不変性を保証
- `TextBlock`: mutable dataclass（indent_level が後処理で更新されるため）
- `InlineRun`: `frozen=True`（値オブジェクト）
- `DocElement`, `TableElement`, `TableCell`: mutable dataclass
- `ElementType`: `enum.Enum`

#### xlsx_reader.py
- openpyxl の `ws.merged_cells.ranges` で結合情報を先に収集する
- 結合の非起点セルは `is_merge_origin=False`, `value=None` として記録
- `extract_font_props` / `extract_bg_color` を独立関数として分離し、テスト容易性を高める
- テーマ色（`cell.font.color.type == "theme"`）はNoneを返す

#### cell_grid.py
- `col_unit` は `statistics.median()` で列幅の中央値を算出
- `modal_row_height` は `statistics.mode()` で最頻値を算出（空の場合はデフォルト15.0）
- `baseline_col` は `min(cell.col for cell in cells if cell.value)` で算出

#### merge_resolver.py
- openpyxl のリッチテキスト (`_CellRichText`) を `to_inline_runs()` で `list[InlineRun]` に変換
- `indent_level` は 0 で初期化（structure_detector が後で更新する）
- セル内改行 (`\n`) はそのまま `text` に保持する

### テスト方針

- openpyxl でプログラム的にワークブックを生成（バイナリ .xlsx をリポジトリにコミットしない）
- `conftest.py` に `xlsx_builder` フィクスチャを定義
- 各テストはモジュール単位で独立させ、パイプライン下流に依存しない
