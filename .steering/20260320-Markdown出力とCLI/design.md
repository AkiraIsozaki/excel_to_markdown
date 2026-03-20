# 設計: Markdown出力とCLI

## 実装アプローチ

```
renderer/markdown_renderer.py  (DocElement → str)
    ↓
excel_to_markdown/cli.py       (パイプライン全体 + エラーハンドリング)
    ↓
excel_to_markdown/__main__.py  (エントリーポイント)
    ↓
tests/test_markdown_renderer.py
tests/test_integration.py + tests/fixtures/
```

## markdown_renderer.py

**ElementType → Markdown 変換テーブル**:
| ElementType | 出力形式 |
|------------|--------|
| HEADING n  | `"#"*n + " " + text + "\n\n"` |
| PARAGRAPH  | `text + "\n\n"` |
| LIST_ITEM level | `"  "*(level-1) + "- " + text + "\n"` |
| TABLE      | GFM テーブル |
| BLANK      | `"\n"` |

**GFM テーブル構築**:
```
| セル1 | セル2 |
| --- | --- |
| 値1 | 値2 |
```
ヘッダー行 (`is_header=True`) がある場合は2行目にセパレータを挿入。
なければ1行目をヘッダーとして使い、2行目にセパレータを挿入（GFM はヘッダー必須）。

**脚注処理**:
- `render()` の `footnotes` リストに comment_text が含まれる
- `render_element()` 内で `[^N]` を text に付記し、counter を返す
- `render()` 末尾に `[^N]: 内容` を一括出力

## cli.py

**パイプライン実行フロー** (`run()`):
1. ファイルパスの検証 (`Path.resolve()` → 存在確認 → 拡張子確認)
2. `openpyxl.load_workbook(path, data_only=True)` で workbook 読み込み
3. 対象シートのリストを決定 (`--sheet` 指定 or 全シート)
4. 各シートに対してパイプライン実行:
   - `read_sheet(ws)` → `list[RawCell]`
   - `CellGrid(cells, col_widths, row_heights)` 構築
   - `resolve(cells)` → `list[TextBlock]`
   - `find_tables(blocks, grid)` → `(tables, remaining)`
   - `detect(remaining, grid, base_font_size)` → `list[DocElement]`
   - `sorted(tables + doc_elements, key=lambda e: e.source_row)` でマージ
   - `footnotes` 収集 → `render(elements, footnotes)`
5. 複数シートは `# シート名\n\n---\n\n` で連結
6. output ファイルに書き出し

**col_widths / row_heights 取得**:
```python
from openpyxl.utils import column_index_from_string
col_widths = {
    column_index_from_string(col): ws.column_dimensions[col].width or 8.0
    for col in ws.column_dimensions
}
row_heights = {r: ws.row_dimensions[r].height or 15.0 for r in ws.row_dimensions}
```

**シート選択**:
`--sheet` が数値文字列 → 0-based インデックス、文字列 → シート名で検索

**エラーハンドリング**: 例外は `run()` の末尾でキャッチ。
- `FileNotFoundError` / `ValueError` → exit code 1, stderr メッセージ
- `Exception` → exit code 2, stderr にメッセージ
