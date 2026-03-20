# 要求仕様: 変換パイプライン基盤

## 概要

Excel方眼紙→Markdownコンバーターのパイプライン基盤層を実装する。
具体的には以下のコンポーネントが対象:

1. **プロジェクトセットアップ** — pyproject.toml / requirements.txt / .gitignore など
2. **models.py** — パイプライン全体で使用するデータモデル (RawCell, TextBlock, InlineRun, DocElement, TableElement, TableCell, ElementType)
3. **reader/xlsx_reader.py** — openpyxl を使って .xlsx シートから `list[RawCell]` を生成
4. **parser/cell_grid.py** — `CellGrid` データクラス（col_unit・modal_row_height・baseline_col・is_empty_row）
5. **parser/merge_resolver.py** — `list[RawCell]` → `list[TextBlock]` 変換（`resolve` / `to_inline_runs`）

## 機能要件

- `models.py` は標準ライブラリのみに依存し、循環依存なし
- `xlsx_reader.py` は印刷領域フィルタ・非表示行列除外・結合セル処理を行う
- `CellGrid` は col_unit（列幅の中央値）・modal_row_height（行高さの最頻値）を正しく算出する
- `merge_resolver.py` は値なし/空白セルをスキップし、InlineRun (リッチテキスト) を保持する
- 全モジュールに mypy strict モード適合の型ヒントを付与する

## 非機能要件

- Python 3.12+
- ruff (line-length=100) でフォーマット済みであること
- 各モジュールに pytest ユニットテストを実装し、カバレッジ 80% 以上を目指す
