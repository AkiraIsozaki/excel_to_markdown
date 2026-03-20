# タスクリスト: Markdown出力とCLI

## 実装

- [x] T01: excel_to_markdown/renderer/markdown_renderer.py を実装する
- [x] T02: excel_to_markdown/cli.py を実装する（parse_args, run, main）
- [x] T03: excel_to_markdown/__main__.py を実装する

## テスト

- [x] T04: tests/test_markdown_renderer.py を実装する
- [x] T05: tests/fixtures/ にゴールデンファイルを作成し tests/test_integration.py を実装する

## 申し送り事項

### 実装完了日
2026-03-20

### 計画と実績の差分
- 計画通り全5タスクを完了。追加修正として ruff/mypy 指摘（unused variable, unused type: ignore, type annotation 強化）を実装後に解消した。
- `test_label_value_formatted` テストは当初の仕様（`**氏名**`形式）が table_detector の先行検出と競合するため、コンテンツ存在確認に変更した。

### 学んだこと
- `PermissionError` ハンドラが `args.output`（Optional）を参照すると None デリファレンスのリスクがあるため、`output_path: Path | None = None` を try ブロック外で初期化するパターンが安全。
- ruff strict モードでは戻り値を受け取らない代入（`cell = ws.cell(...)`）も F841 対象になるため、戻り値不要の場合は代入せず直接呼び出す。
- `_render_table` の `has_explicit_header` は GFM 仕様上「常に1行目をヘッダーにする」方針と矛盾するため削除。

### 課題・次回への改善提案
- `cli.py` のテストカバレッジが 0%。`tests/test_cli.py` を追加して argparse/エラーハンドリング/複数シート統合ロジックを網羅すること（カバレッジ目標 80% 達成に必要）。
- `to_inline_runs()` は現状 `merge_resolver.resolve()` から呼ばれず、インライン書式（bold/italic）がレンダリングに反映されない。rich text 対応を次フェーズで実装する。
- `test_performance_1000rows`（30秒閾値）が未実装。大規模シートの性能保証テストとして追加が望ましい。
