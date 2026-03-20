# タスクリスト: 構造検出

## 実装

- [x] T01: excel_to_markdown/parser/table_detector.py を実装する（find_tables）
- [x] T02: excel_to_markdown/parser/structure_detector.py を実装する（detect, compute_indent_tiers, classify_heading, is_label_value_pair）

## テスト

- [x] T03: tests/test_table_detector.py を実装する（2×2検出/不完全グリッド非検出/ヘッダー判定）
- [x] T04: tests/test_structure_detector.py を実装する（見出し優先順位/インデントティア/ラベル:値/空行挿入）

## 申し送り事項

### 実装完了日
2026-03-20

### 計画と実績の差分

- 計画通り全4タスクを完了
- `test_multiple_separate_tables` が1件失敗 → 行ギャップを無視して複数グループを1表に統合していたバグ。`bottom_row + 1` による隣接行チェックで修正
- 検証で ruff エラー2件（未使用インポート・行長）を検出して修正。`structure_detector.py` の空行挿入条件3が dead code と判明し削除

### 学んだこと

- `find_tables()` は「列境界の一致」だけでなく「行の連続性」もチェックする必要がある。行ギャップ（`r > prev_bottom + 1`）で打ち切ることで、離れた複数グループが1表にまとまるバグを防ぐ
- `_should_insert_blank()` の背景色条件: 「前後で異なる色」という条件が「有色→白」を包含するため、条件3は dead code。スペックの記述が冗長だったことを確認し、実装とスペック（docstring）を整合させた
- `table_detector.py` の `_build_table()` の `TableElement(text="", level=0, ...)` のように、継承した dataclass のデフォルトフィールドに `init=False` があると `element_type` は渡せないことに注意（`TableElement` は `element_type=TABLE` で固定）

### 次フェーズへの引き継ぎ

- 次回: `/add-feature CLIと出力` — `renderer/markdown_renderer.py` + `excel_to_markdown/cli.py` + `excel_to_markdown/__main__.py` の実装
- `cli.py` の `run()` では `table_detector.find_tables()` の戻り値（TableElement + remaining_blocks）を `structure_detector.detect()` に渡し、最終的に `source_row` でソート・マージする設計
- `TableElement` の `source_row` は表の先頭行で設定済みなので、CLI側での `sorted(tables + doc_elements, key=lambda e: e.source_row)` が正しく動作する
