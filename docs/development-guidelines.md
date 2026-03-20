# 開発ガイドライン (Development Guidelines)

## コーディング規約

### 命名規則

#### 変数・関数

```python
# ✅ 良い例
raw_cells = extract_raw_cells(worksheet)
text_blocks = resolve_merge_regions(raw_cells)
def compute_indent_tiers(blocks: list[TextBlock], grid: CellGrid) -> dict[int, int]: ...

# ❌ 悪い例
data = extract(ws)
blocks = get(data)
def calc(b, g): ...
```

**原則**:
- 変数: snake_case、内容を表す名詞または名詞句
- 関数: snake_case、動詞で始める（例: `read_sheet`, `detect_structure`, `render_markdown`）
- 定数: UPPER_SNAKE_CASE（例: `DEFAULT_BASE_FONT_SIZE = 11.0`）
- Boolean変数・戻り値: `is_`, `has_`, `can_` で始める（例: `is_merge_origin`, `has_comment`）

#### クラス・列挙型

```python
# dataclass: PascalCase、名詞
@dataclass
class TextBlock: ...

@dataclass(frozen=True)
class RawCell: ...

# Enum: PascalCase、値は意味のある文字列
class ElementType(enum.Enum):
    HEADING   = "heading"
    PARAGRAPH = "paragraph"
    LIST_ITEM = "list_item"
    TABLE     = "table"
    BLANK     = "blank"
```

---

### コードフォーマット

**ツール**: `ruff format`（blackと互換）

**設定**（pyproject.toml）:
```toml
[tool.ruff]
target-version = "py312"
line-length = 100
```

**インデント**: 4スペース（タブ禁止）

**インポート順序**: `ruff` が自動整理（標準ライブラリ → サードパーティ → ローカル）

---

### 型ヒント

Python 3.12の型ヒントを**すべての関数・メソッドに必須**とする。mypy strict モードでエラーなしを維持すること。

```python
# ✅ 良い例
def classify_heading(block: TextBlock, base_font_size: float) -> int | None:
    """見出しレベル(1-6)を返す。見出しでない場合はNone。"""
    ...

# ❌ 悪い例
def classify_heading(block, base_font_size):
    ...
```

**型ヒントのルール**:
- 戻り値が複数型の場合は `X | Y`（Union型）
- Optional は `X | None`（`Optional[X]` は非推奨）
- コレクションは `list[T]`, `dict[K, V]`（大文字の `List`, `Dict` は非推奨）

---

### コメント規約

**docstring（公開関数・クラスに必須）**:

```python
def read_sheet(ws: Worksheet, print_area: tuple[int, int, int, int] | None) -> list[RawCell]:
    """ワークシートからRawCellリストを返す。

    印刷領域が設定されている場合はその範囲内のみ処理する。
    非表示の行・列は除外する。結合セルの非起点セルはスキップする。

    Args:
        ws: openpyxlのワークシートオブジェクト
        print_area: 印刷領域 (min_row, min_col, max_row, max_col)。未設定はNone。

    Returns:
        有効なセルデータのリスト。(row, col) 順にソート済み。
    """
    ...
```

**インラインコメント（なぜそうするかを説明する）**:

```python
# ✅ 良い例: 理由を説明
# openpyxlのread-onlyモードではws.merged_cells.rangesにアクセスできないため、
# 通常モードで開く（パフォーマンスは若干低下するが仕様上必要）
wb = openpyxl.load_workbook(path, data_only=True)

# ❌ 悪い例: 何をしているか（コードを読めば分かる）
# ワークブックを開く
wb = openpyxl.load_workbook(path, data_only=True)
```

---

### dataclass設計

```python
# イミュータブルなデータには frozen=True
@dataclass(frozen=True)
class RawCell:
    row: int        # 1-based
    col: int        # 1-based
    value: str | None

# ミュータブルで後から更新するデータ（indent_level等）はfrozen=Falseのまま
@dataclass
class TextBlock:
    text: str
    top_row: int
    indent_level: int = 0   # MergeResolver後にStructureDetectorが更新
```

---

### エラーハンドリング

**原則**:
- CLIレイヤー（`cli.py`）でのみ例外をキャッチして終了コードとメッセージに変換する
- パーサー・レンダラー層は例外を上位に伝播させる（ベストエフォートはコメント挿入で対応）
- `except Exception` でのサイレント無視禁止

```python
# ✅ cli.py でのエラーハンドリング（正しいパターン）
def run(args: argparse.Namespace) -> int:
    try:
        result = convert(args.input, args)
        Path(args.output).write_text(result, encoding="utf-8")
        return 0
    except FileNotFoundError as e:
        print(f"エラー: ファイルが見つかりません: {e.filename}", file=sys.stderr)
        return 1
    except ConversionError as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"予期しないエラーが発生しました: {e}", file=sys.stderr)
        return 2

# ❌ 悪い例: サイレント無視
try:
    font_size = cell.font.size
except:
    pass  # 無視禁止。代わりに font_size = None を明示
```

---

## Git運用ルール

### ブランチ戦略

```
main
  └─ feature/xlsx-reader         # 機能追加
  └─ feature/structure-detector  # 機能追加
  └─ fix/merge-cell-detection    # バグ修正
  └─ refactor/indent-tier        # リファクタリング
  └─ docs/update-prd             # ドキュメントのみ
```

- `main`: 常に動作する状態を維持
- `feature/`: 新機能（PRDのP0/P1の機能単位）
- `fix/`: バグ修正
- `refactor/`: 機能変更を伴わないコード改善
- `docs/`: ドキュメントのみの変更

### コミットメッセージ規約

**フォーマット**（Conventional Commits）:
```
<type>(<scope>): <subject>

<body>（任意）
```

**Type**:
- `feat`: 新機能
- `fix`: バグ修正
- `docs`: ドキュメントのみ
- `refactor`: リファクタリング（機能変更なし）
- `test`: テストの追加・修正
- `chore`: ビルド・設定の変更

**例**:
```
feat(structure-detector): 見出し判定に優先順位付きルールを実装

H3とH4の判定が重複するケースを解消するため、
フォントサイズ→太字+インデントの優先順位で判定するよう変更。
PRD機能要件2の受け入れ条件に対応。
```

### プルリクエストプロセス

**作成前のチェック**:
- [ ] `pytest tests/` が全パス
- [ ] `ruff check excel_to_markdown/` がエラーなし
- [ ] `mypy excel_to_markdown/` がエラーなし
- [ ] カバレッジ80%以上を維持

**PRテンプレート**:
```markdown
## 概要
[変更内容の1-2行の説明]

## 変更理由
[PRDのどの要件に対応するか、またはどのバグを修正するか]

## 変更内容
- [変更点1]
- [変更点2]

## テスト
- [ ] ユニットテスト追加・更新
- [ ] 統合テストで確認
- [ ] `pytest tests/ -v` 全パス確認

## 関連
- PRD機能要件: #[番号]
```

---

## テスト戦略

### テストの種類と配置

| 種類 | 配置 | 目的 |
|------|------|------|
| ユニットテスト | `tests/test_{module}.py` | 各コンポーネントの単体動作確認 |
| 統合テスト | `tests/test_integration.py` | xlsxフィクスチャ→Markdownのエンドツーエンド確認 |
| パフォーマンステスト | `tests/test_integration.py` | A4/1,000行の変換時間確認 |

**カバレッジ目標**: 80%以上

### ユニットテストの書き方

**フィクスチャはopenpyxlでプログラム生成する**（バイナリ.xlsxをコミットしない）:

```python
# tests/conftest.py
import pytest
import openpyxl

@pytest.fixture
def simple_workbook():
    """シンプルな方眼紙ワークブックを生成するフィクスチャ。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    # 列幅を方眼紙スタイルに設定
    for col in range(1, 30):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 3
    return wb
```

**テスト命名規則**: `test_{対象}_{条件}_{期待結果}`

```python
# ✅ 良い例
def test_classify_heading_bold_indent0_returns_h4(): ...
def test_classify_heading_size18pt_returns_h1(): ...
def test_resolve_empty_cell_skips(): ...
def test_indent_tier_groups_nearby_columns(): ...

# ❌ 悪い例
def test_heading(): ...
def test_works(): ...
def test1(): ...
```

**テストの構造（Arrange-Act-Assert）**:

```python
def test_classify_heading_bold_indent0_returns_h4():
    # Arrange: テストデータを準備
    block = TextBlock(
        text="章タイトル",
        top_row=1, left_col=2, bottom_row=2, right_col=10,
        row_span=2, col_span=9,
        font_bold=True, font_italic=False, font_strikethrough=False,
        font_underline=False, font_size=None,
        bg_color=None, has_comment=False, comment_text=None,
        indent_level=0,
    )
    grid = CellGrid(cells=[], col_widths={}, row_heights={})

    # Act: テスト対象を実行
    result = classify_heading(block, base_font_size=11.0)

    # Assert: 期待値と比較
    assert result == 4  # H4
```

### 統合テストの書き方

```python
def test_integration_simple_heading(simple_workbook, tmp_path):
    """シンプルな見出し+段落の変換テスト。"""
    ws = simple_workbook.active
    # H1見出し（18pt）を設定
    ws["B2"].value = "報告書タイトル"
    ws["B2"].font = Font(size=18, bold=True)
    # ...フィクスチャ構築

    xlsx_path = tmp_path / "test.xlsx"
    simple_workbook.save(xlsx_path)

    result = convert(xlsx_path)
    expected = Path("tests/fixtures/simple_heading.md").read_text(encoding="utf-8")
    assert result == expected
```

### モックの使用方針

- **外部ライブラリ（openpyxl等）はモックしない**: 実際のxlsxを生成してテストする
- **ファイルシステムへの書き込みは `tmp_path`** フィクスチャを使って一時ディレクトリに行う
- **パフォーマンステストのタイムアウト**: `@pytest.mark.timeout(30)` で30秒超をfail

---

## コードレビュー基準

### レビューポイント

**機能性**:
- [ ] PRDの受け入れ条件を満たしているか
- [ ] ベストエフォート方針（情報を捨てない）が守られているか
- [ ] エッジケース（空シート・印刷領域外・非表示行）が考慮されているか

**可読性**:
- [ ] 命名が明確か（snake_case、動詞始まり）
- [ ] docstringが公開関数に付いているか
- [ ] インラインコメントが「なぜ」を説明しているか

**型安全性**:
- [ ] すべての関数に型ヒントがあるか
- [ ] `mypy strict` がエラーなしか

**レイヤー設計**:
- [ ] readerがparserに依存していないか（逆流禁止）
- [ ] models.pyが他の内部モジュールに依存していないか

**テスト**:
- [ ] 新機能にユニットテストが追加されているか
- [ ] カバレッジが80%以上を維持しているか

### レビューコメントの書き方

```markdown
# ✅ 良い例（理由と改善案を提示）
[必須] この条件式だと `font_size=None`（デフォルトスタイル継承）のセルが
H1と誤判定される可能性があります。`font_size is not None and font_size >= ...`
のようにNoneチェックを先に行うよう修正をお願いします。

[推奨] `col_unit` の計算が `compute_indent_tiers` と `CellGrid` の両方にあります。
`CellGrid` のプロパティに一元化できるかもしれません（重複排除のため）。

# ❌ 悪い例
この書き方は良くないです。
```

**優先度の明示**:
- `[必須]`: マージ前に修正が必要
- `[推奨]`: 修正を強く勧めるが、理由を示して議論可
- `[提案]`: 将来的に検討してほしいアイデア
- `[質問]`: 設計意図の確認

---

## 開発環境セットアップ

### 必要なツール

| ツール | バージョン | 提供方法 |
|--------|-----------|---------|
| Python | 3.12以上 | devcontainer に含まれる |
| pip | 最新 | Python付属 |

### セットアップ手順

```bash
# 1. リポジトリのクローン
git clone <URL>
cd excel_to_markdown

# 2. 仮想環境の作成と依存インストール
python -m venv .venv
source .venv/bin/activate       # Windows: .venv\Scripts\activate
pip install -e ".[dev]"         # 本番依存 + 開発依存

# 3. テストの実行
pytest tests/ -v

# 4. 型チェックとLintの実行
mypy excel_to_markdown/
ruff check excel_to_markdown/
ruff format --check excel_to_markdown/

# 5. ツールの実行（動作確認）
python -m excel_to_markdown --help
```

### コード品質コマンド一覧

```bash
# テスト（カバレッジ付き）
pytest tests/ --cov=excel_to_markdown --cov-report=term-missing

# 型チェック
mypy excel_to_markdown/

# Lint
ruff check excel_to_markdown/ tests/

# フォーマット
ruff format excel_to_markdown/ tests/

# 全チェック一括実行
ruff check excel_to_markdown/ && mypy excel_to_markdown/ && pytest tests/
```

### 推奨エディタ設定（VS Code）

- `ms-python.python`: Python基本拡張
- `ms-python.vscode-pylance`: 型チェックとインテリセンス
- `charliermarsh.ruff`: Ruff lint/format（保存時に自動実行）
