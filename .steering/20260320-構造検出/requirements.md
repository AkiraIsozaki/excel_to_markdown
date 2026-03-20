# 要求仕様: 構造検出

## 概要

変換パイプライン基盤の上に、TextBlock → DocElement 変換を担う構造検出レイヤーを実装する。

1. **parser/table_detector.py** — TextBlockリストからGFMテーブルに変換すべきグリッド領域を検出する
2. **parser/structure_detector.py** — TextBlockリストをDocElementリストに変換する（見出し/段落/リスト/ラベル:値）

## 機能要件

### table_detector.py
- `find_tables(blocks, grid) -> (list[TableElement], list[TextBlock])` を実装
- 2行以上 × 2列以上の矩形グリッドを検出
- 全行の列境界 (left_col) が一致する場合のみ表として検出（保守的）
- 1行目が bold かつ他行が非bold の場合 is_header=True、それ以外は全行 is_header=False

### structure_detector.py
- `detect(blocks, grid, base_font_size) -> list[DocElement]` を実装
- `compute_indent_tiers(blocks, grid) -> dict[int, int]` を実装
- 見出し判定: `classify_heading(block, base_font_size) -> int | None`（優先順位1〜6）
- ラベル:値パターン認識: `is_label_value_pair(left, right) -> bool`（20文字以下）
- 行グループ処理: 同一行2ブロック→ラベル:値、3ブロック以上→スペース結合段落
- 空行挿入: 行ギャップ・背景色変化によるBLANK要素挿入

## 非機能要件

- 全関数に mypy strict 適合の型ヒント
- ruff (line-length=100) 準拠
- pytest ユニットテストを実装（各コンポーネント）
