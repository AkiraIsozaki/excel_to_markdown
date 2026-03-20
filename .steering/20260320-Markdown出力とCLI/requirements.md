# 要求仕様: Markdown出力とCLI

## 概要

変換パイプラインの最終段（出力レイヤー）とエントリーポイントを実装する。
これによりパイプライン全体が完成し、`python -m excel_to_markdown input.xlsx` が動作する。

1. **renderer/markdown_renderer.py** — DocElement リスト → Markdown 文字列
2. **excel_to_markdown/cli.py** — CLI引数パース・パイプライン全体の実行・エラーハンドリング
3. **excel_to_markdown/__main__.py** — `python -m excel_to_markdown` のエントリーポイント
4. **tests/test_markdown_renderer.py** — renderer のユニットテスト
5. **tests/test_integration.py** — パイプライン全体の統合テスト（ゴールデンファイル比較）
6. **tests/fixtures/** — 統合テスト用ゴールデン Markdown ファイル

## 機能要件

### markdown_renderer.py
- `render(elements, footnotes) -> str`
- `render_element(el, footnote_counter) -> tuple[str, int]`
- セル内改行 (`\n`) → Markdown ハードブレーク（行末スペース2つ + `\n`）
- GFM テーブル形式での出力
- 脚注 `[^1]` 形式での連番出力
- `collapse_blank_lines()` で3行以上の連続空行を2行に圧縮

### cli.py
- `parse_args(argv) -> argparse.Namespace`
- `run(args) -> int` — パイプライン全体を実行し exit code を返す
- 複数シート統合: シート名 H1 + `---` で連結
- エラーハンドリング: functional-design.md のエラーテーブルに準拠
- `--debug`: TextBlock リストを JSON で stderr 出力

### __main__.py
- `python -m excel_to_markdown` で `cli.main()` を呼び出す

## 非機能要件
- 全関数に mypy strict 適合の型ヒント
- ruff (line-length=100) 準拠
- pytest ユニットテスト + 統合テスト
