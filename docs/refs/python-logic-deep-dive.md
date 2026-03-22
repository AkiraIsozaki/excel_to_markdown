# ロジック詳細解説（Python 版）

> 対象読者：`python-project-overview.md` を読んだ後、もう一段深くコードを理解したい方
>
> **Java は知っているが Python はあまり知らない方** を想定しています。Python 固有の書き方には随所で Java との対比を示します。

---

## Python 早見表（Java 開発者向け）

コードを読む前に、よく出てくる Python 特有の書き方を確認しておきます。

### dataclass（Java の record/POJO に相当）

```python
# Python
@dataclass
class RawCell:
    row: int
    value: str | None   # Java: String（null 許容は Optional<String>）
    font_bold: bool     # Java: boolean
```

```java
// Java 相当（Java 16+ の record）
record RawCell(int row, String value, boolean fontBold) {}
```

`@dataclass` は `__init__`・`__repr__`・`__eq__` を自動生成します。`frozen=True` をつけると `final` 相当になり、フィールドへの代入が実行時エラーになります。

### 型ヒント（Java の型宣言に相当）

```python
list[RawCell]         # Java: List<RawCell>
dict[int, float]      # Java: Map<Integer, Float>
str | None            # Java: Optional<String>
tuple[int, int, int]  # Java: 値の組（専用クラスがない場合は配列や record で代替）
```

型ヒントは実行時に検証されません（あくまでドキュメントと静的解析ツール用）。

### リスト内包表記（Java の stream に相当）

```python
# Python
hidden_rows = {r for r, rd in ws.row_dimensions.items() if rd.hidden}

# Java 相当
Set<Integer> hiddenRows = ws.getRowDimensions().entrySet().stream()
    .filter(e -> e.getValue().isHidden())
    .map(Map.Entry::getKey)
    .collect(Collectors.toSet());
```

`{...}` は `set`、`[...]` は `list`、`{k: v for ...}` は `dict` を生成します。

### プロパティ（Java の getter に相当）

```python
# Python
@property
def col_unit(self) -> float:
    return statistics.median(...)

# 呼び出し側: grid.col_unit（括弧なし！）
```

```java
// Java 相当
public double getColUnit() { return ...; }
// 呼び出し: grid.getColUnit()
```

### `statistics` モジュール

```python
import statistics
statistics.median([1, 2, 3])   # 中央値 = 2
statistics.mode([1, 1, 2])     # 最頻値 = 1
```

Java に標準の統計ライブラリはなく、Apache Commons Math などを使います。Python では標準ライブラリに含まれています。

### `Path`（Java の `Path` に相当）

```python
from pathlib import Path
path = Path("input.xlsx")
path.suffix         # → ".xlsx"（Java: path.getExtension()相当は自前実装）
path.with_suffix(".md")  # → Path("input.md")
path.write_text(content, encoding="utf-8")
path.read_text(encoding="utf-8")
path.replace(other)  # アトミックなリネーム（Java: Files.move()）
```

### `id()` 関数

```python
id(obj)  # オブジェクトのメモリアドレス（Java: System.identityHashCode(obj) に相当）
```

`==` はオブジェクトの「値の等価性」、`id()` は「同一インスタンスか」を確認します。Java の `==` と `equals()` の区別と同じ考え方です。

---

## 目次（クイック）

| セクション | 内容 |
|-----------|------|
| [0. E2E 通し例](#0-エンドツーエンド通し例) | 小さな Excel がどう変換されるかを一気に確認 |
| [1. CLI レイヤー](#1-cliレイヤー) | parse_args・run・_run_pipeline・アトミック書き出し |
| [2. 読み込みレイヤー](#2-読み込みレイヤー-readerxlsx_readerpy) | xlsx_reader・印刷領域・結合セル・フォント抽出 |
| [3. 空間解析レイヤー](#3-空間解析レイヤー-parsercell_gridpy--parsermerge_resolverpy) | CellGrid・col_unit・modal_row_height・resolve() |
| [4. テーブル検出](#4-テーブル検出-parsertable_detectorpy) | グリッド検出アルゴリズム・ヘッダー判定 |
| [5. 構造検出](#5-構造検出-parserstructure_detectorpy) | インデントティア・見出し判定・ラベル:値・空行挿入 |
| [6. Markdown 生成](#6-markdown生成-renderermarkdown_rendererpy) | render()・テーブルレンダリング・脚注・インライン書式 |
| [7. Web UI レイヤー](#7-web-uiレイヤー-webapppy) | FastAPI・tempfile・単一/ZIP レスポンス |
| [補足 A: frozen dataclass の理由](#補足a-frozen-dataclass-はなぜ使うのか) | Java の record との対比 |
| [補足 B: テスト戦略](#補足b-テスト戦略とゴールデンファイル) | ゴールデンファイル・フィクスチャ生成 |

---

## この文書の地図

```
0. エンドツーエンド通し例（先にここを読むと理解しやすい）
1. CLI レイヤー（cli.py）
2. 読み込みレイヤー（reader/xlsx_reader.py）
   2-1. 全体の流れ
   2-2. 印刷領域・非表示行列フィルタ
   2-3. 結合セルの処理
   2-4. フォントプロパティとフォントサイズ単位
3. 空間解析レイヤー（cell_grid.py / merge_resolver.py）
   3-1. CellGrid: col_unit と modal_row_height
   3-2. resolve(): RawCell → TextBlock
4. テーブル検出（table_detector.py）
   4-1. 検出アルゴリズム（行コルマップ → 貪欲矩形探索）
   4-2. 部分一致許容とヘッダー判定
5. 構造検出（structure_detector.py）
   5-1. compute_indent_tiers(): 列幅ギャップによるティア計算
   5-2. classify_heading(): 6段階見出し判定
   5-3. _process_row_group(): 行グループの分類
   5-4. _should_insert_blank(): 空行挿入判定
6. Markdown 生成（markdown_renderer.py）
   6-1. render() の全体構造
   6-2. render_element(): 要素別変換
   6-3. _render_table(): GFM テーブル出力
   6-4. インライン書式・セル内改行・脚注
7. Web UI レイヤー（web/app.py）
補足 A: frozen dataclass はなぜ使うのか
補足 B: テスト戦略とゴールデンファイル
```

---

## 0. エンドツーエンド通し例

各セクションを読む前に、小さな Excel がパイプライン全体を通してどう変化するかを一気に確認する。

### 入力：Excel シートのイメージ

```
        A列(col=1)            C列(col=3)
row=1   [18pt] "仕様書"
row=3   [14pt bold] "機能一覧"
row=5   [11pt] "機能名"       [11pt] "説明"
row=6   [11pt] "変換"         [11pt] "ExcelをMarkdownに変換する"
row=7   [11pt] "DL"           [11pt] "変換結果をダウンロードする"
row=9   [11pt, col=3] "注意事項"   ← C列起点のためインデント相当
```

### Step 1 後：xlsx_reader → merge_resolver の出力（TextBlock リスト）

すべて `indent_level=0` のまま（後で `structure_detector` が更新する）。

```
TextBlock(top_row=1, left_col=1, text="仕様書",   font_size=18, font_bold=False, indent_level=0)
TextBlock(top_row=3, left_col=1, text="機能一覧", font_size=14, font_bold=True,  indent_level=0)
TextBlock(top_row=5, left_col=1, text="機能名",   font_size=11, font_bold=False, indent_level=0)
TextBlock(top_row=5, left_col=3, text="説明",     font_size=11, font_bold=False, indent_level=0)
TextBlock(top_row=6, left_col=1, text="変換",     font_size=11, font_bold=False, indent_level=0)
TextBlock(top_row=6, left_col=3, text="ExcelをMarkdownに変換する", indent_level=0)
TextBlock(top_row=7, left_col=1, text="DL",       font_size=11, font_bold=False, indent_level=0)
TextBlock(top_row=7, left_col=3, text="変換結果をダウンロードする", indent_level=0)
TextBlock(top_row=9, left_col=3, text="注意事項", font_size=11, font_bold=False, indent_level=0)
```

### Step 2a 後：table_detector の出力

```
row=5 の left_col セット: {1, 3}  → 2列
row=6 の left_col セット: {1, 3}  → {1,3} は {1,3} のサブセット かつ row=5の直後 → 表の候補
row=7 の left_col セット: {1, 3}  → 同様に表を継続

→ row=5〜7 を TableElement として検出（3行×2列）
→ 残余: [TextBlock(row=1,"仕様書"), TextBlock(row=3,"機能一覧"), TextBlock(row=9,"注意事項")]
```

### Step 2b 後：structure_detector の出力（indent_level・ElementType 確定）

```
left_col の収集: {1, 3}
col_unit の推定: 列幅の中央値（仮に 3.0）
threshold = 3.0 × 1.5 = 4.5
sorted_cols = [1, 3]
  col=1: tier 0
  gap(1→3) = col_widths[1] + col_widths[2]  → 仮に合計 4.0 → 4.5以下 → 同一ティア

→ TextBlock(row=1,  col=1, text="仕様書")   : font_size=18 ≥ 11*18/11 → HEADING level=1
→ TextBlock(row=3,  col=1, text="機能一覧") : font_size=14 ≥ 11*14/11 かつ bold → HEADING level=2
→ TextBlock(row=9,  col=3, text="注意事項") : indent_level=0（col=3 が tier 0）、非bold → PARAGRAPH
```

> gap が `threshold` 以下のため col=1 と col=3 は同一ティア（どちらも `indent_level=0`）になります。
> もし列幅の合計が 4.5 を超えれば col=3 は `indent_level=1` になり、LIST_ITEM に分類されます。

### Step 3 後：markdown_renderer の出力

```markdown
# 仕様書

## 機能一覧

| 機能名 | 説明 |
| --- | --- |
| 変換 | ExcelをMarkdownに変換する |
| DL | 変換結果をダウンロードする |

注意事項

```

`row=1` と `row=3` の間は行ギャップ 2 があり `modal_row_height × 2` を超えるため BLANK が挿入されます。テーブルと "注意事項" の間（row=7→row=9）も同様です。

---

## 1. CLI レイヤー

**現在地**: `cli.py` / `__main__.py`

### エントリーポイントの構造

`python -m excel_to_markdown` で起動すると `__main__.py` が `cli.main()` を呼ぶだけです。

```python
# __main__.py（全体）
from excel_to_markdown.cli import main
if __name__ == "__main__":
    main()
```

`main()` の中では `parse_args()` と `run()` または `serve()` に分岐します：

```python
def main() -> None:
    args = parse_args()          # 引数解析
    if args.subcommand == "serve":
        sys.exit(serve(args))    # Web UI 起動（uvicorn を起動）
    else:
        sys.exit(run(args))      # 変換実行
```

`sys.exit(int)` は Java の `System.exit(int)` に相当します。戻り値（int）が OS の終了コードになります。

### serve vs convert の分岐

`parse_args()` は第1引数が `"serve"` かどうかで分岐します。Python の `argparse`（Java の picocli 相当）を2つのパーサーに分けることで、`convert` と `serve` の引数定義を完全に分離しています。

```python
def parse_args(argv):
    argv_list = list(argv) if argv else sys.argv[1:]
    if argv_list and argv_list[0] == "serve":
        return _parse_serve_args(argv_list[1:])  # --port, --no-browser
    return _parse_convert_args(argv_list)         # input, -o, -s, --debug など
```

`serve()` 内では `import uvicorn` が **その時だけ** 実行されます（モジュールトップレベルでは import しない）。

```python
def serve(args):
    try:
        import uvicorn  # ← serve 時だけ import
    except ImportError:
        print("エラー: uvicorn が必要です ...", file=sys.stderr)
        return 1
```

Java で言えば `Class.forName("...")` で動的ロードするのに近い発想です。`uvicorn` がインストールされていない環境でも `convert` モードは正常に動作します。

### _run_pipeline() — 変換の組立

単一シートの変換は `_run_pipeline()` に集約されています：

```python
def _run_pipeline(raw_cells, grid, args) -> str:
    blocks = resolve(raw_cells)                        # RawCell → TextBlock
    if args.debug:
        _dump_blocks_debug(blocks)                     # JSON を stderr に出力
    tables, remaining = find_tables(blocks, grid)      # テーブル抽出
    doc_elements = detect(remaining, grid, args.base_font_size)  # 構造分類
    all_elements = sorted(
        list(tables) + doc_elements,
        key=lambda e: e.source_row,   # ← lambda: Java の Comparator.comparing() 相当
    )
    footnotes = [e.comment_text for e in all_elements if e.comment_text]
    return render(all_elements, footnotes)
```

`list(tables) + doc_elements` は2つのリストを結合します（Java の `Stream.concat()` 相当）。

`TableElement` と `DocElement` を合流後に `source_row` でソートするのは、テーブル検出と構造検出が独立して動いているためです（どちらが先に要素を生成するか不定）。

### アトミックなファイル書き出し

変換失敗時に壊れた `.md` が残らないよう、`.tmp` ファイルに書いてから `replace`（リネーム）します：

```python
def _write_output(path: Path, content: str) -> None:
    tmp_path = path.with_suffix(path.suffix + ".tmp")  # "output.md.tmp"
    try:
        tmp_path.write_text(content, encoding="utf-8")
        tmp_path.replace(path)   # POSIX: rename(2) システムコール（アトミック）
    except OSError as e:
        tmp_path.unlink(missing_ok=True)  # 書き込み失敗時は tmp を消す
        raise PermissionError(...) from e
```

`Path.replace()` は Java の `Files.move(src, dst, StandardCopyOption.ATOMIC_MOVE)` に相当します。同一ファイルシステム内であれば中途半端な状態にならないことが保証されます。

### バッチ変換モード

`input` にディレクトリを指定すると `_run_batch()` が呼ばれます：

```python
def _run_batch(dir_path: Path, args) -> int:
    targets = sorted(
        list(dir_path.glob("**/*.xlsx")) + list(dir_path.glob("**/*.xls"))
    )
    # ...
    for file_path in targets:
        code = _convert_file(file_path, output_path, args)
        if code != 0:
            exit_code = code  # エラーがあっても残りのファイルを継続変換
    return exit_code
```

`Path.glob("**/*.xlsx")` は Java の `Files.walk()` + フィルターに相当します。`**` はサブディレクトリを再帰的に検索します。

---

## 2. 読み込みレイヤー（reader/xlsx_reader.py）

**現在地**: `excel_to_markdown/reader/xlsx_reader.py`

### 2-1. 全体の流れ

```python
def read_sheet(ws: Worksheet, print_area: str | None = None) -> list[RawCell]:
    area = get_print_area(ws)          # 印刷領域を (min_row, min_col, max_row, max_col) で取得
    hidden_rows = {r for r, rd in ws.row_dimensions.items() if rd.hidden}
    hidden_cols = {col_idx for ... if ws.column_dimensions[col].hidden}

    merge_origins = {}        # {(row, col): (row_span, col_span)}
    merge_non_origins = set() # 起点以外の結合セル (row, col)

    for row in ws.iter_rows():   # Java: worksheet の各行をイテレート
        for cell in row:
            # 印刷領域フィルタ → 非表示フィルタ → 結合非起点スキップ
            # → RawCell を生成して raw_cells に追加
    return raw_cells
```

openpyxl を **通常モード**（`read_only=False`）で開くのは `ws.merged_cells.ranges` へのアクセスのためです。read-only モードではこの API が使えません（Apache POI の `Sheet.getMergedRegions()` とは異なる制約）。

### 2-2. 印刷領域・非表示行列フィルタ

印刷領域は `ws.print_area` で取得できますが、返り値の型が不安定（文字列・リスト・`None`）なため `get_print_area()` でラップしています：

```python
def get_print_area(ws) -> tuple[int, int, int, int] | None:
    pa = ws.print_area
    if not pa:
        return None
    # pa が list の場合もあるため先頭要素を取り出す
    area_str = pa[0] if isinstance(pa, list) else str(pa)
    return _parse_area_str(area_str)  # "Sheet1!A1:D10" → (1, 1, 10, 4)
```

`isinstance(pa, list)` は Java の `pa instanceof List` に相当します。

非表示行・列は 1 回のイテレーションで `set` に収集し、セルループ内で O(1) で判定します：

```python
# セット内包表記: Java の stream().filter(...).collect(toSet()) 相当
hidden_rows: set[int] = {r for r, rd in ws.row_dimensions.items() if rd.hidden}
hidden_cols: set[int] = {
    column_index_from_string(cd)  # "A" → 1, "B" → 2 ...
    for cd in ws.column_dimensions
    if ws.column_dimensions[cd].hidden
}
```

### 2-3. 結合セルの処理

openpyxl の `ws.merged_cells.ranges` を1回走査して「起点セット」と「非起点セット」を事前構築します：

```python
for mr in ws.merged_cells.ranges:
    origin = (mr.min_row, mr.min_col)
    row_span = mr.max_row - mr.min_row + 1
    col_span = mr.max_col - mr.min_col + 1
    merge_origins[origin] = (row_span, col_span)
    for r in range(mr.min_row, mr.max_row + 1):   # Java: IntStream.rangeClosed(...)
        for c in range(mr.min_col, mr.max_col + 1):
            if (r, c) != origin:
                merge_non_origins.add((r, c))
```

`(r, c)` は Python の **タプル**（Java の `Pair<Integer, Integer>` や `record` に相当する不変の値の組）です。タプルはハッシュ可能なため `dict` のキーや `set` の要素として直接使えます。

セルループ内での判定はシンプルな `in` 演算子（O(1) の `set` 検索）：

```python
if pos in merge_non_origins:
    # 非起点セル: value=None, is_merge_origin=False で記録（後段でスキップされる）
    raw_cells.append(RawCell(row=r, col=c, value=None, is_merge_origin=False, ...))
    continue  # Java の continue と同義
```

非起点セルを「スキップ」ではなく「`value=None` で記録」する理由：下流の `merge_resolver.resolve()` で `value is None` を条件に除外するため、空セルと非起点セルを同じパスで処理できます。

### 2-4. フォントプロパティとフォントサイズ単位

`extract_font_props()` は `cell.font` オブジェクトから書式を取り出します：

```python
def extract_font_props(cell) -> tuple[bool, bool, bool, bool, float | None, str | None]:
    font = cell.font
    bold   = bool(font.bold)    # None を False に安全変換
    italic = bool(font.italic)
    strike = bool(font.strike)
    # underline は True/False ではなく "single"/"double"/"none"/None が入ることがある
    underline = bool(font.underline) and font.underline != "none"

    size = float(font.size) if font.size is not None else None  # pt 単位

    color = None
    if font.color is not None and font.color.type == "rgb":  # テーマ色は "theme" → None
        raw = font.color.rgb
        if raw and raw != "00000000":  # 透明色 (ARGB=00000000) は None 扱い
            color = str(raw)

    return bold, italic, strike, underline, size, color  # タプルで複数値を返す
```

**タプルで複数値を返す**のは Python でよく使うパターンです。Java では複数戻り値がないため専用クラスやフィールド変数を使いますが、Python では `return a, b, c` と書くだけです。

フォントサイズは openpyxl が **ポイント（pt）** 単位で返します。Apache POI の `Font.getFontHeightInPoints()` と同じ単位なので、変換は不要です。

背景色 `extract_bg_color()` ではグラデーション塗りつぶしや透明色を除外します：

```python
def extract_bg_color(cell) -> str | None:
    fill = cell.fill
    if fill.fill_type in (None, "none", "gradient"):  # Java: switch 相当
        return None
    fg = fill.fgColor
    if fg.type == "rgb":
        raw = str(fg.rgb)
        return raw if raw != "00000000" else None  # 三項演算子：Java の raw != "00000000" ? raw : null
    return None  # テーマ色 → None
```

`x if 条件 else y` は Java の三項演算子 `条件 ? x : y` に相当します。

---

## 3. 空間解析レイヤー（parser/cell_grid.py / parser/merge_resolver.py）

**現在地**: `excel_to_markdown/parser/cell_grid.py`, `parser/merge_resolver.py`

### 3-1. CellGrid: col_unit と modal_row_height

`CellGrid` はシート全体の「グリッドピッチ」の推定値を提供します。

```python
@dataclass
class CellGrid:
    cells: list[RawCell]
    col_widths: dict[int, float]   # {列番号: 列幅（Excel文字幅単位）}
    row_heights: dict[int, float]  # {行番号: 行高さ（pt）}
```

#### col_unit（列幅の中央値）

```python
@property    # ← getter メソッド。呼び出し側は grid.col_unit（括弧なし）
def col_unit(self) -> float:
    widths = [w for w in self.col_widths.values() if w and w > 0]
    return statistics.median(widths) if widths else 8.0  # デフォルト 8.0
```

`statistics.median()` は Python 標準ライブラリの中央値計算です（Apache Commons Math 等の外部ライブラリ不要）。

Excel 方眼紙では多数の列を同じ細幅で統一するため、**中央値**が「1グリッドの幅」の良い推定値になります。`compute_indent_tiers()` はこの値を閾値計算に使います。

#### modal_row_height（行高さの最頻値）

```python
@property
def modal_row_height(self) -> float:
    heights = [h for h in self.row_heights.values() if h and h > 0]
    try:
        return statistics.mode(heights)    # 最頻値
    except statistics.StatisticsError:
        return statistics.median(heights)  # 最頻値が複数ある場合は中央値にフォールバック
```

`statistics.StatisticsError` は Java の `NoSuchElementException` に近い例外です。Python の `try-except` は Java の `try-catch` に相当します。

「標準的な行の高さ」を推定するために**最頻値**を使います。方眼紙では多くの行が同じ高さで作られるため、最頻値が標準行高さの良い推定値になります。空行挿入の閾値 `modal_row_height × 2` の計算に使います。

### 3-2. resolve(): RawCell → TextBlock

`merge_resolver.resolve()` は `RawCell` リストから値のある TextBlock だけを作り、`(top_row, left_col)` でソートします：

```python
def resolve(cells: list[RawCell]) -> list[TextBlock]:
    blocks = []
    for cell in cells:
        if cell.value is None:    # Java: cell.getValue() == null
            continue
        text = cell.value.strip() # 前後の空白を除去（Java: String.strip()）
        if not text:              # 空文字は falsy（Java: text.isEmpty()）
            continue
        blocks.append(TextBlock(
            text=text,
            top_row=cell.row,
            left_col=cell.col,
            bottom_row=cell.row + cell.merge_row_span - 1,
            right_col=cell.col + cell.merge_col_span - 1,
            row_span=cell.merge_row_span,
            col_span=cell.merge_col_span,
            font_bold=cell.font_bold,
            # ... 他のフィールド
            indent_level=0,  # structure_detector が後で更新
        ))
    blocks.sort(key=lambda b: (b.top_row, b.left_col))  # 2要素タプルでソート
    return blocks
```

`blocks.sort(key=lambda b: (b.top_row, b.left_col))` は Java の `blocks.sort(Comparator.comparingInt(TextBlock::getTopRow).thenComparingInt(TextBlock::getLeftCol))` に相当します。

`bottom_row = row + merge_row_span - 1` で結合セルの終端行を計算します。この値は後の空行挿入判定（`curr.top_row - prev.bottom_row`）で使われます。

---

## 4. テーブル検出（parser/table_detector.py）

**現在地**: `excel_to_markdown/parser/table_detector.py`

### 4-1. 検出アルゴリズム

```python
def find_tables(blocks, grid) -> tuple[list[TableElement], list[TextBlock]]:
    # 1. row → col → TextBlock の2次元マップを構築
    row_col_map: dict[int, dict[int, TextBlock]] = {}
    for block in blocks:
        # setdefault: キーがなければ空 dict を作ってから追加（Java の computeIfAbsent に相当）
        row_col_map.setdefault(block.top_row, {})[block.left_col] = block

    used: set[int] = set()   # 使用済み TextBlock の id()（Java: Set<Integer>）
    tables = []

    for start_row in sorted(row_col_map.keys()):
        # 起点行の全ブロックが使用済みならスキップ
        if all(id(b) in used for b in row_col_map[start_row].values()):
            continue

        cols_in_start_row = sorted(row_col_map[start_row].keys())
        if len(cols_in_start_row) < 2:
            continue  # 1列では表にならない

        # 起点行から下に向かって同じ列構成の隣接行を探す
        table_rows = [start_row]
        for r in sorted_rows:
            if r <= start_row: continue
            prev_bottom = max(b.bottom_row for b in row_col_map[prev_row].values())
            if r > prev_bottom + 1: break  # 行ギャップあり → 表終了
            row_cols = set(row_col_map[r].keys())
            if row_cols.issubset(cols_in_start_row_set) and len(row_cols) >= 2:
                table_rows.append(r)
            else:
                break

        if len(table_rows) >= 2:
            tables.append(_build_table(table_rows, row_col_map, cols_in_start_row))
            for r in table_rows:
                for b in row_col_map[r].values():
                    used.add(id(b))   # id(b): Java の System.identityHashCode(b) 相当

    remaining = [b for b in blocks if id(b) not in used]
    return tables, remaining
```

**`dict.setdefault(key, default)`**：キーが存在しなければ `default` をセットしてから返します。Java の `Map.computeIfAbsent(key, k -> new HashMap<>())` に相当します。

**`all(条件 for b in ...)` / `any(条件 for b in ...)`**：Java の `stream().allMatch(...)` / `stream().anyMatch(...)` に相当します。

**`id(b)` を使う理由**：同じ位置・同じテキストを持つ TextBlock が複数存在しうるため、オブジェクトの**同一性**（Java の `==`）で判定します。Python の `==` はオブジェクトの「値の等価性」（Java の `equals()`）なので、`id()` を使います。

**`break`**：Python の `break` は Java と同義です。`for` ループを途中で抜けます。

### 4-2. 部分一致許容とヘッダー判定

後続行の列が起点行の**サブセット**であれば表に含めます：

```python
if row_cols.issubset(cols_in_start_row_set) and len(row_cols) >= 2:
```

`set.issubset(other)` は Java の `Set.containsAll(other)` と逆（`other.containsAll(row_cols)` に相当）。

具体例：

```
起点行: col={1, 3, 5}  "名前" "年齢" "部署"
次の行: col={1, 3, 5}  "山田" "30"   "営業"  → サブセット OK
その次: col={1, 3}     "鈴木" "25"          → サブセット OK（部署が空セル）
その次: col={1}        "田中"               → len < 2 → 停止
```

空セルがあっても表を途切れさせない、欠損値を許容した設計です。

ヘッダー判定は「1行目が全て bold かつ 2行目以降に非 bold 行がある」場合のみ `is_header=True` にします：

```python
first_row_all_bold = all(b.font_bold for b in first_row_blocks)
has_non_bold_row = any(
    not all(
        (b := row_col_map[r].get(c)) is not None and b.font_bold  # セイウチ演算子 :=
        for c in cols
    )
    for r in table_rows[1:]
)
use_header = first_row_all_bold and has_non_bold_row
```

**セイウチ演算子 `b := expr`**（Python 3.8+）：式の評価結果を変数に代入しながら条件式に使います。Java の `var b = expr; if (b != null && b.isFontBold())` を1行で書けます。

GFM テーブルはヘッダー行が必須のため、`use_header` の値に関わらず **常に1行目をヘッダーとして出力** します（`_render_table()` で `start_row=1` に固定）。

---

## 5. 構造検出（parser/structure_detector.py）

**現在地**: `excel_to_markdown/parser/structure_detector.py`

`detect()` の処理順序：

```
1. compute_indent_tiers() → TextBlock.indent_level を更新（直接書き込み）
2. _group_same_row_blocks() → 同一 top_row のブロックをグループ化
3. 各グループを _process_row_group() で DocElement に変換
   3a. グループ間で _should_insert_blank() → BLANK 要素を挿入
4. source_row でソートして返す
```

### 5-1. compute_indent_tiers(): 列幅ギャップによるティア計算

```python
def compute_indent_tiers(blocks, grid) -> dict[int, int]:
    threshold = grid.col_unit * 1.5
    sorted_cols = sorted(set(b.left_col for b in blocks))  # 重複排除してソート
    tiers = {sorted_cols[0]: 0}  # 最左列は常に tier 0
    tier = 0
    for i in range(1, len(sorted_cols)):
        # 列番号の差ではなく実際の列幅の累積でギャップを計算する
        gap = sum(
            grid.col_widths.get(c, grid.col_unit)  # なければ col_unit をデフォルト値に
            for c in range(sorted_cols[i - 1], sorted_cols[i])
        )
        if gap > threshold:
            tier += 1
        tiers[sorted_cols[i]] = tier
    return tiers  # {left_col → indent_level}
```

`set(b.left_col for b in blocks)` は **ジェネレーター式**（`()` でなく `set()` に渡す）です。`list`、`set`、`dict` などに渡せます。Java の `stream().map(...).collect(toSet())` に相当します。

`dict.get(key, default)` は Java の `map.getOrDefault(key, default)` に相当します。

**列番号の差ではなく実際の列幅の合計でギャップを計算する**のが重要です：

```
例: col_unit=3.0, threshold=4.5
sorted_cols = [1, 2, 5, 6, 9]

gap(1→2) = col_widths[1] = 2.0 → 4.5以下 → 同一ティア (tier=0)
gap(2→5) = col_widths[2] + col_widths[3] + col_widths[4] = 6.0 → 4.5超え → tier=1
gap(5→6) = col_widths[5] = 2.0 → 4.5以下 → 同一ティア (tier=1)
gap(6→9) = col_widths[6] + col_widths[7] + col_widths[8] = 6.0 → 4.5超え → tier=2

結果: {1:0, 2:0, 5:1, 6:1, 9:2}
```

計算した `tiers` は `blocks` の `indent_level` に直接書き込まれます：

```python
for block in blocks:
    block.indent_level = tiers.get(block.left_col, 0)
```

`TextBlock` は `frozen=False`（ミュータブル）なため、この代入が可能です。`frozen=True` の `RawCell` には同じ操作はできません。

### 5-2. classify_heading(): 6段階見出し判定

```python
def classify_heading(block: TextBlock, base_font_size: float) -> int | None:
    fs   = block.font_size
    bold = block.font_bold
    ind  = block.indent_level

    if fs is not None:   # Java: if (fs != null)
        if fs >= base_font_size * (18/11): return 1   # H1
        if fs >= base_font_size * (14/11) and bold: return 2   # H2
        if fs >= base_font_size * (12/11) and bold: return 3   # H3

    if bold:
        if ind == 0: return 4   # H4
        if ind == 1: return 5   # H5
        return 6                # H6（indent >= 2）

    return None  # Java: return null
```

**優先順位**：フォントサイズ系（H1/H2/H3）が `bold+indent` 系（H4/H5/H6）より先に評価されます。

**`base_font_size` パラメーター**：`--base-font-size 10.5` を指定すると 18pt 閾値が `18 × 10.5/11 ≒ 17.2pt` に調整されます。Excelテンプレートによってデフォルトフォントが異なる場合の対策です。

**`font_size=None` の場合**：フォントサイズが Excel のスタイル継承で未設定のセルは `font_size=None` になります。`if fs is not None` ブロックは丸ごとスキップされ、`bold + indent` の判定に移ります。`is not None` は Java の `!= null` に相当しますが、Python では `is None` / `is not None` を使うのが慣例です（`== None` ではなく）。

### 5-3. _process_row_group(): 行グループの分類

同一行にある TextBlock の数によって処理が変わります：

```python
def _process_row_group(group, base_font_size):
    if len(group) == 2 and is_label_value_pair(group[0], group[1]):
        # ラベル:値パターン
        left, right = group[0], group[1]   # タプルアンパック（Java にない記法）
        text = f"**{left.text}** {right.text}"   # f文字列（Java の String.format() 相当）
        return [DocElement(PARAGRAPH, text, ...)]

    if len(group) >= 3:
        # 3列以上: スペース区切りで1段落
        text = " ".join(b.text for b in group)  # Java: String.join(" ", list)
        return [DocElement(PARAGRAPH, text, ...)]

    # 1ブロック: 見出し/リスト/段落 に分類
    return [_classify_single_block(group[0], base_font_size)]
```

**f文字列**（`f"...{変数}..."`）：Java の `String.format("..%s..", var)` や文字列連結より簡潔に書けます。

**`str.join()`**：Java の `String.join()` と同義ですが、引数の順序が逆です（区切り文字.join(リスト)）。

#### ラベル:値パターンの判定

```python
def is_label_value_pair(left, right) -> bool:
    return len(left.text) <= 20  # 左セルが 20 文字以下
```

「20 文字以下」は「氏名」「プロジェクト名」「担当者」のような短いラベルを想定した経験的な閾値です。

#### 1ブロックの分類

```python
def _classify_single_block(block, base_font_size):
    heading_level = classify_heading(block, base_font_size)
    if heading_level is not None:   # Java: if (headingLevel != null)
        return DocElement(HEADING, text, level=heading_level, ...)

    if block.indent_level >= 1:   # インデントあり → リスト
        is_numbered = bool(_NUMBERED_LIST_RE.match(block.text))
        return DocElement(LIST_ITEM, text, level=block.indent_level, ...)

    # インデントなし・非見出し → 段落
    is_numbered = bool(_NUMBERED_LIST_RE.match(block.text))
    return DocElement(PARAGRAPH, text, level=0, ...)
```

番号付きリストの判定パターン（正規表現）：

```python
_NUMBERED_LIST_RE = re.compile(
    r"^(?:\d+[.)）]|（\d+）|[①-⑨]|[㊀-㊉])\s"
)
```

`re.compile()` は Java の `Pattern.compile()` に相当します。`\d+[.)）]` で `1.`/`1)`/`1）`、`[①-⑨]` で Unicode の丸数字に対応しています。

### 5-4. _should_insert_blank(): 空行挿入判定

```python
def _should_insert_blank(prev, curr, grid) -> bool:
    gap = curr.top_row - prev.bottom_row
    if gap > grid.modal_row_height * 2:   # 条件1: 行ギャップ
        return True

    if prev.bg_color != curr.bg_color:    # 条件2: 背景色の変化
        return True

    return False
```

**条件1（行ギャップ）**：`modal_row_height` は行高さの最頻値（pt）ですが、`top_row` と `bottom_row` は**行番号**（整数）です。`gap` は行番号の差なので、標準行高さ何行分の空白があるかを行番号で近似しています。行番号の差が2以上（= 間に1行以上の空行がある）で `modal_row_height × 2 = 30pt 相当` を超えると空行を挿入します。

**条件2（背景色）**：Excel 方眼紙では背景色でセクションを分けることがあります。白→有色、有色→他色、有色→白のすべてで空行を挿入します（`!=` で全パターンを包含）。Python の `!=` は `None` 同士の比較も安全です（`None != None` は `False`）。

---

## 6. Markdown 生成（renderer/markdown_renderer.py）

**現在地**: `excel_to_markdown/renderer/markdown_renderer.py`

### 6-1. render() の全体構造

```python
def render(elements: list[DocElement], footnotes: list[str]) -> str:
    parts: list[str] = []
    footnote_counter = 1   # Python に final はない。慣例として再代入しない

    for el in elements:
        md, footnote_counter = render_element(el, footnote_counter)
        parts.append(md)

    result = "".join(parts)  # Java: String.join("", parts)

    if footnotes:  # リストが空でなければ True（Java: !footnotes.isEmpty()）
        note_lines = "\n".join(f"[^{i+1}]: {fn}" for i, fn in enumerate(footnotes))
        result = result.rstrip("\n") + "\n\n" + note_lines + "\n"

    return collapse_blank_lines(result)  # 3行以上の連続空行を2行に圧縮
```

**`enumerate()`**：インデックスと要素のペアを返します。Java の `IntStream.range(0, list.size()).forEach(i -> { ... list.get(i) ... })` に相当します。

**多値の戻り値を受け取る**：`md, footnote_counter = render_element(...)` はタプルアンパックです。`render_element()` が `return md, counter` でタプルを返し、呼び出し側で2変数に展開します。

### 6-2. render_element(): 要素別変換

```python
def render_element(el: DocElement, footnote_counter: int) -> tuple[str, int]:
    text = convert_cell_newlines(el.text)   # セル内 \n → Markdown ハードブレーク

    if el.hyperlink:   # None でも空文字でもない場合 True
        text = f"[{text}]({el.hyperlink})"  # Markdown リンク形式

    if el.comment_text:   # コメントがあれば脚注番号を付記
        text = text + f"[^{footnote_counter}]"
        footnote_counter += 1

    if el.element_type == ElementType.HEADING:
        return "#" * el.level + " " + text + "\n\n", footnote_counter
    if el.element_type == ElementType.PARAGRAPH:
        return text + "\n\n", footnote_counter
    if el.element_type == ElementType.LIST_ITEM:
        indent = "  " * (el.level - 1)      # level=1 → ""、level=2 → "  "
        prefix = "1. " if el.is_numbered_list else "- "
        return indent + prefix + text + "\n", footnote_counter
    if el.element_type == ElementType.BLANK:
        return "\n", footnote_counter
    if el.element_type == ElementType.TABLE:
        return _render_table(el), footnote_counter
```

`"#" * el.level` は文字列の繰り返しです（Java にはない記法）。`"#" * 3` → `"###"`。

`"  " * (el.level - 1)` も同様で、インデントのスペースを生成します。

各 `ElementType` が生成する文字列の末尾：

| ElementType | 末尾 | 理由 |
|-------------|------|------|
| HEADING | `\n\n` | GFM で見出し後には空行が必要 |
| PARAGRAPH | `\n\n` | 段落間を1行あける |
| LIST_ITEM | `\n` | リスト項目はすぐ次の行へ |
| BLANK | `\n` | 空行1つ分 |
| TABLE | `\n\n` | テーブル後には空行 |

連続する空行は最後に `collapse_blank_lines()` で圧縮します。

### 6-3. _render_table(): GFM テーブル出力

```python
def _render_table(el: TableElement) -> str:
    col_count = el.col_count
    lines: list[str] = []

    # ヘッダー行（1行目固定）
    header_cells = [convert_cell_newlines(c.text) for c in el.rows[0]]
    # リスト内包表記: Java の stream().map(...).collect(toList()) 相当
    while len(header_cells) < col_count:
        header_cells.append("")
    lines.append("| " + " | ".join(header_cells) + " |")
    lines.append("| " + " | ".join("---" for _ in range(col_count)) + " |")

    # データ行（2行目以降）
    for row in el.rows[1:]:   # スライス: el.rows[1:] = 2番目以降
        cells = [convert_cell_newlines(c.text) for c in row]
        while len(cells) < col_count:
            cells.append("")  # 欠損セルを空文字で補完
        lines.append("| " + " | ".join(cells) + " |")

    return "\n".join(lines) + "\n\n"
```

**スライス `list[start:end]`**：`el.rows[1:]` は「インデックス1から末尾まで」、つまり「2番目以降の全要素」です。Java の `list.subList(1, list.size())` に相当します。

**`"---" for _ in range(col_count)`**：`_` は「使わない変数」の慣例名です（Java に相当する慣例はない）。`range(n)` は `0` から `n-1` までの整数シーケンスを返します（Java の `IntStream.range(0, n)` に相当）。

**「GFM はヘッダー行必須」問題**：GFM のテーブル仕様はヘッダー行が必須のため、`is_header` フラグに関わらず常に1行目をヘッダーとして出力します。

### 6-4. インライン書式・セル内改行・脚注

#### セル内改行のハードブレーク変換

```python
def convert_cell_newlines(text: str) -> str:
    return text.replace("\n", "  \n")
```

Excel の Alt+Enter（`\n`）を GFM のハードブレーク（行末スペース2つ + 改行）に変換します。Java の `String.replace()` と同義です。

#### インライン書式

`apply_inline_format()` は `InlineRun` 1つを変換します：

```python
def apply_inline_format(run: InlineRun) -> str:
    text = run.text
    if run.strikethrough:       text = f"~~{text}~~"
    if run.underline:           text = f"<u>{text}</u>"
    if run.bold and run.italic: text = f"***{text}***"
    elif run.bold:              text = f"**{text}**"
    elif run.italic:            text = f"*{text}*"
    return text
```

適用順序：`~~` と `<u>` を先に適用し、`**` / `*` はその外側に追加します。`elif` は Java の `else if` に相当します。

#### 脚注の流れ

```
xlsx_reader: cell.comment → RawCell.comment_text
merge_resolver: RawCell → TextBlock.comment_text（フィールドコピー）
structure_detector: TextBlock → DocElement.comment_text（フィールドコピー）
cli.py: all_elements から comment_text を収集 → footnotes: list[str]
markdown_renderer.render(): footnotes を末尾に一括出力
markdown_renderer.render_element(): comment_text があれば [^N] を本文に挿入
```

脚注番号 N は `render_element()` を呼ぶたびにカウンターを引数で渡す関数型スタイルで管理しています（グローバル変数やインスタンス変数を使わない設計）。

#### 空行圧縮

```python
def collapse_blank_lines(md: str) -> str:
    return re.sub(r"\n{3,}", "\n\n", md)
```

`re.sub(pattern, replacement, string)` は Java の `string.replaceAll(pattern, replacement)` に相当します。正規表現 `\n{3,}` は「3つ以上の連続した改行」にマッチします。

---

## 7. Web UI レイヤー（web/app.py）

**現在地**: `excel_to_markdown/web/app.py`

### FastAPI アプリの構成

```python
def create_app() -> FastAPI:
    app = FastAPI(title="excel-to-markdown Web UI")
    app.mount("/static", StaticFiles(directory=str(_STATIC_DIR)), name="static")

    @app.get("/")           # ← デコレーター: このパス・メソッドへのリクエストでこの関数を呼ぶ
    async def index() -> HTMLResponse:       # async def: 非同期関数（後述）
        html = (_STATIC_DIR / "index.html").read_text(encoding="utf-8")
        return HTMLResponse(content=html)

    @app.get("/health")
    async def health() -> dict[str, str]:
        return {"status": "ok"}

    @app.post("/api/convert")
    async def convert(files: Annotated[list[UploadFile], File(...)]) -> Response:
        ...

    return app
```

**デコレーター `@app.post(...)`**：Spring Boot の `@PostMapping(...)` に相当します。関数を引数に取って「拡張した新しい関数」を返す Python の機能です。ここでは FastAPI が「このパスへの POST リクエストでこの関数を呼ぶ」という登録を行います。

**`async def` / `await`**：非同期関数です。`await upload.read()` でファイルの読み込みを非同期に行います。Java で言うと `CompletableFuture` や Reactor の `Mono` に近い概念ですが、Python では `async/await` で同期的に書けます。

**ファクトリ関数パターン**：`create_app()` がアプリを生成して返します。Spring Boot の `@Configuration` + `@Bean` のように、テストコードが `create_app()` で独立したアプリを生成でき、テスト間の状態共有を避けられます。

### /api/convert のロジック

```python
@app.post("/api/convert")
async def convert(
    files: Annotated[list[UploadFile], File(...)],
) -> Response:
    results = []
    errors = []

    for upload in files:
        suffix = Path(upload.filename).suffix.lower()
        if suffix not in {".xlsx", ".xls"}:    # set の in 演算子: O(1) 検索
            if len(files) == 1:
                raise HTTPException(status_code=400, detail="...")
            errors.append(upload.filename)
            continue   # Java の continue と同義

        data = await upload.read()   # バイト列として読み込む
        if len(data) > 50 * 1024 * 1024:  # 50MB
            ...

        # tempfile に書き出して変換パイプラインを呼び出す
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp_path = Path(tmp.name)
            tmp_path.write_bytes(data)

        try:
            md = run_file(tmp_path)   # cli.py の共通ヘルパーを再利用
            results.append((md_filename, md))
        except Exception:
            errors.append(filename)
        finally:
            tmp_path.unlink(missing_ok=True)  # 変換後即削除（Java: Files.delete()）

    # 単一ファイル → text/markdown
    if len(results) == 1 and not errors:
        return Response(content=md.encode("utf-8"), media_type="text/markdown; charset=utf-8", ...)

    # 複数ファイル → application/zip
    zip_buf = io.BytesIO()   # Java: ByteArrayOutputStream
    with zipfile.ZipFile(zip_buf, mode="w", compression=ZIP_DEFLATED) as zf:
        for md_filename, md_content in results:   # タプルアンパック
            zf.writestr(md_filename, md_content.encode("utf-8"))
    return Response(content=zip_buf.getvalue(), media_type="application/zip", ...)
```

**`with` 文**：Java の try-with-resources に相当します。`with tempfile.NamedTemporaryFile(...) as tmp:` ブロックを抜けると自動的に `tmp.close()` が呼ばれます。

**`io.BytesIO()`**：Java の `ByteArrayOutputStream` に相当します。メモリ上のバイト列バッファです。

**`str.encode("utf-8")`**：Java の `str.getBytes(StandardCharsets.UTF_8)` に相当します。

**エラーヘッダー `X-Conversion-Errors`**：複数ファイル変換で一部が失敗した場合、失敗したファイル名を HTTP レスポンスヘッダーに含めます。これにより ZIP 本体には成功したファイルのみ含まれ、クライアントはどのファイルが失敗したか把握できます。

---

## 補足 A: frozen dataclass はなぜ使うのか

`RawCell` は `@dataclass(frozen=True)` で定義されています。

```python
@dataclass(frozen=True)   # ← Java の record に相当
class RawCell:
    row: int
    value: str | None
    font_bold: bool
    ...
```

`frozen=True` にすると：

1. **属性への代入が `FrozenInstanceError` になる**：`cell.value = "foo"` が実行時エラーになります。Java の `record` で宣言したフィールドは再代入できないのと同じです
2. **`__hash__` が自動生成される**：`frozenset` や `dict` のキーに使えます（Java では `equals()` と `hashCode()` を override する必要がある処理が自動化）
3. **意図の表明**：「このオブジェクトは変換パイプラインの途中で変更しない」というコードの意図が明示されます

一方、`TextBlock` は `@dataclass`（`frozen=False`）です：

```python
@dataclass   # ← Java の POJO（setter あり）に近い
class TextBlock:
    indent_level: int = 0  # structure_detector が後から書き込む
```

`structure_detector.compute_indent_tiers()` の結果を `block.indent_level = ...` で書き込むため、ミュータブルである必要があります。

Java で言えば：

| Python | Java 相当 |
|--------|-----------|
| `@dataclass(frozen=True)` | `record`（Java 16+）または `final` フィールドのみの POJO |
| `@dataclass` | 通常の POJO（setter あり） |
| `frozen=True` のフィールド代入 → `FrozenInstanceError` | `record` のコンポーネントへの代入 → コンパイルエラー |

---

## 補足 B: テスト戦略とゴールデンファイル

### ゴールデンファイルによる回帰テスト

`tests/e2e/test_houganshi_e2e.py` はゴールデンファイル（期待出力のスナップショット）との比較で回帰を検知します：

```python
def test_output_matches_golden(self, tmp_path: Path) -> None:
    actual = _convert(FIXTURE_XLSX, tmp_path)
    expected = GOLDEN_MD.read_text(encoding="utf-8")
    assert actual == expected   # Java の assertEquals(expected, actual) に相当（引数順が逆）
```

Python の `assert` は Java の `assertEquals`/`assertTrue` と異なり、JUnit のような専用メソッドではなく言語キーワードです。pytest が `assert` 式を書き換えて詳細な差分を表示します。

ゴールデンファイルを意図的に更新する場合は：

```bash
python -m excel_to_markdown tests/e2e/fixtures/sample_houganshi.xlsx \
       -o tests/e2e/golden/sample_houganshi.md
```

### フィクスチャを Python で生成する理由

バイナリの `.xlsx` をリポジトリにコミットすると：

- Git の差分が読めない（バイナリ diff）
- openpyxl のバージョンアップで細部が変わりうる
- どんなテストデータかがコードから読み取れない

そのため、フィクスチャは `make_sample_houganshi.py` のように openpyxl でプログラム的に生成します（Java の POI で Excel を生成するのと同じ発想）：

```python
# tests/e2e/fixtures/make_sample_houganshi.py（概念コード）
from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook()
ws = wb.active
ws["A1"].value = "機能要件定義書"
ws["A1"].font = Font(size=18, bold=True)
ws["A3"].value = "1. 概要"
ws["A3"].font = Font(size=14, bold=True)
...
wb.save("sample_houganshi.xlsx")
```

生成スクリプトが Git 管理されているため、どんな Excel を作っているかが完全にテキストで把握できます。

### テストの分類

| テストクラス | 目的 | Java JUnit 相当 |
|------------|------|----------------|
| `TestGoldenFile` | ゴールデンファイルとの完全一致（回帰検知） | スナップショットテスト |
| `TestContentPreservation` | 入力の全テキストが出力に含まれること | 統合テスト（assertThat(...).contains(...)） |
| `TestStructureDetection` | 見出し・テーブル・脚注が正しい記法で出力されること | 機能テスト |

Python の `class` でテストをグループ化するのは、Java の `@Nested` クラスに相当します。pytest はクラス名が `Test` で始まり、メソッド名が `test_` で始まれば自動的にテストとして認識します（`@Test` アノテーション不要）。

ユニットテスト（`tests/test_*.py`）は openpyxl を使わず Python のデータ構造を直接渡します。これにより `structure_detector` や `table_detector` を openpyxl なしで高速にテストできます。
