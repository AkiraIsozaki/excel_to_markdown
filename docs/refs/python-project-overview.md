# プロジェクト解説：Excel → Markdown 変換ツール（Python 版）

> 対象読者：Python は知っているが FastAPI・openpyxl はよく知らない方

---

## 1. このアプリが何をするか（俯瞰）

Excel ファイル（.xlsx / .xls）を受け取り、GitHub Flavored Markdown（GFM）形式のテキストに変換するツールです。

```
[ユーザー]
  ↓ input.xlsx を渡す（CLIまたはブラウザ）

[変換パイプライン]
  ↓ Excel を読み込む（openpyxl / xlrd）
  ↓ セルを空間的に解析する
  ↓ 見出し・段落・リスト・表に分類する
  ↓ Markdown 文字列に変換する

[出力]
  → output.md ファイル（CLI）
  → .md ダウンロード or .zip（Web UI）
```

---

## 2. 2 つの入口（CLI と Web UI）

このツールは **同じ変換パイプラインを共有** しながら、2 つの入口を持ちます。

```
┌─────────────────────────────────────────────────────────────┐
│ CLI（python -m excel_to_markdown input.xlsx）                │
│   - argparse で引数を受け取る                                │
│   - ファイルパスを直接操作                                    │
│   - .md ファイルをローカルに書き出す                          │
├─────────────────────────────────────────────────────────────┤
│  ここから下は共通（cli.py の run_file() 関数）                │
├─────────────────────────────────────────────────────────────┤
│ Web UI（python -m excel_to_markdown serve）                  │
│   - FastAPI で POST /api/convert を受け取る                  │
│   - アップロードファイルを tempfile に保存                    │
│   - 単一ファイル → .md / 複数ファイル → .zip を返す           │
└─────────────────────────────────────────────────────────────┘
```

**重要な点**: Web UI は Excel の中身を読まない。バイナリのまま tempfile に保存してから `run_file()` を呼ぶだけです。

---

## 3. モジュール構成

```
excel_to_markdown/
├── models.py          → パイプライン全体で使う共有データモデル
├── __main__.py        → python -m のエントリーポイント（main() を呼ぶだけ）
├── cli.py             → CLI の引数解析・パイプライン起動・ファイル書き出し
├── reader/
│   ├── xlsx_reader.py → openpyxl → list[RawCell]
│   └── xls_reader.py  → xlrd    → list[RawCell]
├── parser/
│   ├── cell_grid.py       → シート全体の空間情報（col_unit, 行高さ）を算出
│   ├── merge_resolver.py  → RawCell → TextBlock（結合セルの解決）
│   ├── structure_detector.py → TextBlock → DocElement（見出し・段落・リスト分類）
│   └── table_detector.py  → TextBlock → TableElement（グリッド表の検出）
├── renderer/
│   └── markdown_renderer.py → DocElement → Markdown 文字列
└── web/
    ├── app.py         → FastAPI アプリ定義（POST /api/convert）
    └── static/
        └── index.html → D&D UI（Vanilla JS）
```

なぜこう分けるか：`reader/` や `parser/` は HTTP のことを何も知らない。その結果、CLI でも Web API でも **同じ変換ロジックをそのまま使える**。

---

## 4. データモデルの 3 層構造

変換パイプラインは、`models.py` に定義された 3 種類のデータオブジェクトをバトンのように受け渡します。

```
[Excel セル]
    ↓ reader/ が変換
RawCell         ← openpyxl / xlrd の生データ（行・列・値・書式）
    ↓ merge_resolver.py が変換
TextBlock       ← 結合セル解決済みのブロック（indent_level は後で付与）
    ↓ table_detector + structure_detector が変換
DocElement      ← 文書要素（HEADING / PARAGRAPH / LIST_ITEM / TABLE / BLANK）
    ↓ markdown_renderer.py が変換
[Markdown 文字列]
```

### RawCell（生セルデータ）

```python
@dataclass(frozen=True)  # frozen=True → イミュータブル
class RawCell:
    row: int           # 行番号（1-based）
    col: int           # 列番号（1-based）
    value: str | None  # セルの値（文字列に変換済み）
    font_bold: bool
    font_size: float | None
    is_merge_origin: bool   # 結合セルの起点か否か
    comment_text: str | None
    # ...他の書式情報
```

**なぜ `frozen=True` か**：読み取り専用の加工データなので、途中で変わると追跡が難しくなる。Python の `frozen=True` dataclass はイミュータブル（変更不可）なので、各ステップで新しいオブジェクトを作ることを強制し、バグが起きにくい。

### TextBlock（テキストブロック）

```python
@dataclass  # こちらは mutable（indent_level を後から書き込む）
class TextBlock:
    text: str
    top_row: int; left_col: int      # 先頭の行・列位置
    bottom_row: int; right_col: int  # 末尾の行・列位置
    font_bold: bool; font_size: float | None
    indent_level: int = 0  # structure_detector が後で更新
    inline_runs: list[InlineRun] = ...  # セル内部分書式
```

### DocElement（文書要素）

```python
@dataclass
class DocElement:
    element_type: ElementType  # HEADING / PARAGRAPH / LIST_ITEM / TABLE / BLANK
    text: str
    level: int      # HEADINGは1-6、LIST_ITEMはインデント深さ
    source_row: int # 元の Excel 行番号（並び順の保持に使う）
    comment_text: str | None  # 脚注として出力するコメント
```

---

## 5. 変換パイプライン（コアロジック）

パイプラインは `cli.py` の `_run_pipeline()` 関数で組み立てられます。

```python
def _run_pipeline(raw_cells, grid, args):
    blocks = resolve(raw_cells)               # Step 1: merge_resolver
    tables, remaining = find_tables(blocks, grid)  # Step 2a: table_detector
    doc_elements = detect(remaining, grid, ...)    # Step 2b: structure_detector
    all_elements = sorted(tables + doc_elements, key=lambda e: e.source_row)
    return render(all_elements, footnotes)    # Step 3: markdown_renderer
```

### Step 1：セルを読む（reader → merge_resolver）

openpyxl で `.xlsx` の各セルから `RawCell` を作り、`merge_resolver.resolve()` で結合セルを解決して `TextBlock` に変換します。

| 取り出す情報 | 用途 |
|------------|------|
| テキスト内容 | Markdown の文字列 |
| フォントサイズ（pt） | 見出しレベル判定 |
| 太字フラグ | 見出しレベル判定 |
| 行・列の位置 | テーブル判定・インデント計算 |
| 背景色 | セクション境界の判定 |
| コメント | 脚注（`[^1]`）として出力 |

**印刷領域の扱い**：印刷領域が設定されている場合はその範囲内のセルのみを変換対象とします。非表示の行・列も除外します。

### Step 2a：テーブルを検出する（table_detector）

```
Excel のセル位置：
  [Row 3, Col 1] "名前"  [Row 3, Col 3] "年齢"  [Row 3, Col 5] "部署"
  [Row 4, Col 1] "山田"  [Row 4, Col 3] "30"    [Row 4, Col 5] "営業"
  [Row 5, Col 1] "鈴木"  [Row 5, Col 3] "25"    [Row 5, Col 5] "開発"

→ 2行×2列以上の矩形グリッド → TABLE として検出
→ テーブルに属するセルは remaining から除外される
```

テーブル検出が先に行われる理由：同一行に複数のセルがある場合、それがテーブルなのかラベル:値パターンなのかを先に確定しておく必要があるため。

### Step 2b：構造を判定する（structure_detector）

テーブル以外の残余 `TextBlock` を `DocElement` に変換します。

#### インデントレベルの計算

```
各 TextBlock の left_col を収集し、col_unit × 1.5 以内の列を同一ティアにグループ化：

col_unit = 3, cols = [1, 2, 5, 6, 9] の場合
  → [1, 2] が tier 0
  → [5, 6] が tier 1
  → [9]    が tier 2
```

`col_unit`（列幅の基準単位）はシート全体のセルが使っている列番号の中央値から算出します。Excel 方眼紙では列が細かく刻まれるため、隣接する複数列を「1段のインデント」として吸収する仕組みです。

#### 見出しレベル判定

```
base_font_size = 11pt（デフォルト）の場合：

フォントサイズ ≥ 18pt（= 11 × 18/11）          → H1
フォントサイズ ≥ 14pt（= 11 × 14/11）かつ太字  → H2
フォントサイズ ≥ 12pt（= 11 × 12/11）かつ太字  → H3
太字 かつ インデントレベル 0               → H4
太字 かつ インデントレベル 1               → H5
太字 かつ インデントレベル 2以上           → H6
それ以外                                   → PARAGRAPH / LIST_ITEM
```

`--base-font-size` オプションで閾値を文書ごとに調整できます（10.5pt を指定すると 18pt 閾値が 17.2pt に再計算）。

#### ラベル:値パターン

```
同一行に 2 つのブロックが横並びで、左のテキストが 20 文字以下 → ラベル:値
  → "氏名" + "山田太郎"  →  **氏名:** 山田太郎

3 列以上 → スペース区切りで 1 段落に結合（ラベル:値判定は不適用）
```

#### 空行の自動挿入

- 行のギャップが `modal_row_height × 2` 以上 → 空行を挿入
- 背景色が変わる箇所 → セクション境界として空行を挿入

### Step 3：Markdown に変換する（markdown_renderer）

確定した `ElementType` ごとに Markdown 記法に変換：

```
ElementType.HEADING    level=1  → "# テキスト"
ElementType.HEADING    level=2  → "## テキスト"
ElementType.LIST_ITEM  level=1  → "- テキスト"
ElementType.LIST_ITEM  level=2  → "  - テキスト"（インデント1段）
ElementType.TABLE              → GFM テーブル形式（以下参照）
ElementType.PARAGRAPH          → "テキスト"（そのまま）
ElementType.BLANK              → ""（空行）
```

**TABLE の具体例**：

```
Excel のセル群          →    Markdown のテーブル
[名前][年齢][部署]           | 名前 | 年齢 | 部署 |
[山田][30][営業]             | --- | --- | --- |
[鈴木][25][開発]             | 山田 | 30 | 営業 |
                             | 鈴木 | 25 | 開発 |
```

**ベストエフォート変換**：構造が確定できなかったセルも段落として必ず出力し、情報を捨てません。

**セルコメント → 脚注**：コメントのあるセルは脚注番号 `[^1]` をテキストに付加し、文末に `[^1]: コメント内容` を出力します。

**インライン書式**：セル内の部分的な太字・イタリック・取り消し線は `**text**` / `*text*` / `~~text~~` に変換します。

---

## 6. FastAPI による Web UI

### FastAPI とは（超簡単に）

FastAPI は「Python の型ヒントを使って HTTP API を定義できる」フレームワークです。Spring Boot と似た点として：

- 関数にデコレーターをつけると **URL へのアクセスで自動的に呼ばれる**
- ファイルアップロード（`UploadFile`）は自動でパースされて引数に渡される
- 非同期（`async def`）と同期（`def`）の両方が使える

```python
# web/app.py
@app.post("/api/convert")          # ← POST /api/convert が来たらこの関数を呼ぶ
async def convert(
    files: Annotated[list[UploadFile], File(...)],  # ← アップロードファイルが自動で入る
) -> Response:
    ...
```

### リクエストが来てから返るまで

```
POST /api/convert（multipart/form-data）
    ↓
ファイル拡張子チェック（.xlsx/.xls のみ）
ファイルサイズチェック（50MB 上限）
    ↓
tempfile にバイナリを書き出す
    ↓
run_file(tmp_path) → 変換パイプラインを呼ぶ（cli.py の共通関数）
    ↓
tempfile を削除
    ↓
単一ファイル → Content-Type: text/markdown で .md を返す
複数ファイル → Content-Type: application/zip で .zip を返す
```

**なぜ tempfile を使うか**：openpyxl はファイルパスまたはファイルオブジェクトを受け取ります。バイト列を一度ファイルに書き出すことで、CLI と同じ `run_file()` 関数をそのまま再利用できます。変換後は即座に削除するためサーバーにファイルが残りません。

### フロントエンド（Vanilla JS）

`web/static/index.html` は単一の HTML ファイルです。React は使わず、ブラウザ標準の `fetch()` API と DOM 操作で実装しています。

```
ユーザーがファイルをドロップ or 選択
    ↓
fetch('POST /api/convert', formData)
    ↓
レスポンスの Content-Type を確認
  text/markdown    → <a download> でそのまま保存
  application/zip  → Blob URL で .zip として保存
```

---

## 7. CLI の仕組み

### エントリーポイントの流れ

```
python -m excel_to_markdown input.xlsx
    ↓
__main__.py の main() → cli.py の main()
    ↓
parse_args() で引数を解析
    ↓
第1引数が "serve" → serve() → FastAPI + uvicorn を起動
それ以外          → run()  → 変換パイプラインを実行
```

### serve vs convert の分岐

```python
# cli.py
def parse_args(argv):
    if argv and argv[0] == "serve":
        return _parse_serve_args(argv[1:])   # Web UI モード
    return _parse_convert_args(argv)          # 変換モード
```

Spring Boot の `WebApplicationType.NONE` のような仕組みとは異なり、Python は「必要なモジュールを必要なときだけ import する」ことで同じ効果を得ています：

```python
def serve(args):
    import uvicorn  # ← serve 時だけ import（未インストールでもエラーにならない）
    from excel_to_markdown.web.app import create_app
    ...
```

### アトミックなファイル書き出し

変換失敗時に部分的なファイルが残らないよう、`.tmp` ファイルに書いてから `rename` します：

```python
tmp_path = path.with_suffix(".md.tmp")
tmp_path.write_text(content, encoding="utf-8")
tmp_path.replace(path)  # OS レベルのアトミック操作
```

---

## 8. 依存関係の方向

```
cli.py / web/app.py
    ↓ 呼ぶ
reader/         → RawCell を生成（openpyxl/xlrd に依存）
    ↓ 渡す
parser/         → TextBlock / DocElement を生成（models.py のみに依存）
    ↓ 渡す
renderer/       → Markdown 文字列を生成（models.py のみに依存）

models.py       → 全レイヤーが参照（Python 標準ライブラリのみ、外部依存なし）
```

**なぜ models.py に外部依存がないか**：`models.py` は全モジュールが読み込むため、ここに openpyxl などが入ると全モジュールが間接的に依存してしまう。標準ライブラリ（`dataclasses`, `enum`）のみに限定することで、`parser/` や `renderer/` のテストを openpyxl なしで書ける。

**新しいフォーマットへの対応**：`reader/` に新しいリーダーを追加するだけで、`parser/` や `renderer/` は一切変更不要（インターフェース: `list[RawCell]` を返すだけ）。

---

## 9. まとめ：データの流れ全体像

```
[CLI] input.xlsx を指定
[Web] ブラウザから .xlsx をアップロード
    ↓
cli.py: run_file() を呼ぶ（CLI・Web 共通）
    ↓
reader/xlsx_reader.py: openpyxl → list[RawCell]
    ↓
parser/cell_grid.py: col_unit・modal_row_height を算出
    ↓
parser/merge_resolver.py: RawCell → list[TextBlock]（結合セル解決）
    ↓
parser/table_detector.py: TextBlock → list[TableElement]（テーブル抽出）
    ↓
parser/structure_detector.py: 残余 TextBlock → list[DocElement]（見出し・段落・リスト分類）
    ↓
  ※ TableElement + DocElement を source_row でソートして順序を保持
    ↓
renderer/markdown_renderer.py: DocElement → Markdown 文字列
    ↓
[CLI] .md ファイルをアトミックに書き出す
[Web] text/markdown または application/zip でレスポンス
```
