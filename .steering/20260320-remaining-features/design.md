# 設計書

## 1. .xls 対応

### xls_reader.py の設計

xlrd 2.x は `.xls` 専用。0-based インデックスを 1-based に変換して `RawCell` を生成する。

**xlrd → RawCell マッピング**:
- `cell_value()` → `value`（文字列化）
- `xf_list[cell_xf_index].font_index` → font から `bold`, `italic`, `strike`, `underline`, `size`
- `book.colour_map[xf.background.pattern_colour_index]` → `bg_color`（ARGB hex に変換）
- `sheet.merged_cells` → `merge_origins`, `merge_non_origins`
- コメント: xlrd 2.x はサポート外のため `has_comment=False`, `comment_text=None`
- ハイパーリンク: xlrd 2.x はサポート外のため `hyperlink=None`（フィールド追加後）

**フォントサイズ**: xlrd では `font.height` が 1/20 pt 単位 → `height / 20` で pt 値

### cli.py の変更

`_open_workbook` を廃止し、`_load_raw_cells_from_sheet` を各リーダーに委譲する形に変更する。

実際には `run()` 内で拡張子判定を行い:
- `.xlsx` → `openpyxl.load_workbook` + `xlsx_reader.read_sheet`
- `.xls` → `xlrd.open_workbook` + `xls_reader.read_sheet_xls`

## 2. アトミック書き込み

`_write_output` を変更:
1. `tmp_path = path.with_suffix(path.suffix + ".tmp")` に書き込む
2. 成功したら `tmp_path.replace(path)` でアトミックにリネーム
3. 失敗したら `tmp_path` を削除してから例外を送出

## 3. ハイパーリンク変換

### models.py の変更
`RawCell` と `TextBlock` に `hyperlink: str | None = None` を追加。

`RawCell` は frozen dataclass のため、フィールドをデフォルト値なしで追加する場合は全コンストラクタ呼び出しを更新する必要がある。**既存コードへの影響を最小化するため、デフォルト値 `None` 付きで追加する**。ただし `frozen=True` の dataclass でデフォルト値なしフィールドの後にデフォルト値ありフィールドは追加できない→ `hyperlink` を最後に追加することで解決。

### xlsx_reader.py の変更
`RawCell` 生成時に `cell.hyperlink.target if cell.hyperlink else None` を渡す。

### TextBlock の変更
`TextBlock` にも `hyperlink: str | None = None` を追加（`merge_resolver.py` で転送）。

### markdown_renderer.py の変更
`render_element` の text 生成時に `block.hyperlink` があれば `[text](url)` 形式に変換。
`TextBlock` → `DocElement` の変換時に hyperlink 情報を渡す設計が必要。

実装方針: `DocElement` にも `hyperlink: str | None = None` を追加し、`structure_detector` が `TextBlock.hyperlink` を `DocElement.hyperlink` に転送する。`render_element` で `el.hyperlink` があれば `text = f"[{text}]({el.hyperlink})"` に変換する。

## 4. バッチ変換

### cli.py の変更
`run()` の冒頭で `input_path.is_dir()` を判定:
- ディレクトリの場合: `glob("**/*.xlsx") + glob("**/*.xls")` で対象ファイルを収集
- 各ファイルに対して既存の単一ファイル変換ロジックを適用

## ディレクトリ変更サマリー

```
excel_to_markdown/
  models.py                     # RawCell, TextBlock, DocElement に hyperlink 追加
  cli.py                        # .xls分岐・アトミック書き込み・バッチモード
  reader/
    xlsx_reader.py              # hyperlink 抽出
    xls_reader.py               # 新規作成
  parser/
    merge_resolver.py           # hyperlink を TextBlock に転送
    structure_detector.py       # hyperlink を DocElement に転送
  renderer/
    markdown_renderer.py        # [text](url) 変換
tests/
  test_integration.py           # 1000行パフォーマンステスト追加
  test_cli.py                   # バッチ変換・アトミック書き込みテスト追加
```

## 実装の順序

1. models.py に hyperlink フィールド追加（全体の基盤）
2. xlsx_reader.py で hyperlink 抽出
3. merge_resolver.py で TextBlock に hyperlink 転送
4. structure_detector.py で DocElement に hyperlink 転送
5. markdown_renderer.py で [text](url) 変換
6. xls_reader.py 新規作成
7. cli.py を .xls 分岐・アトミック書き込み・バッチモードに対応
8. test_integration.py に 1000行テスト追加
9. test_cli.py にバッチ・アトミック書き込みテスト追加
10. 全テスト実行・確認
