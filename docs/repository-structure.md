# リポジトリ構造定義書 (Repository Structure Document)

> **注意**: 本ドキュメントは実装目標として定義した構造を記述しています。
> 現在の実装状況はステアリングファイル（`.steering/`）の `tasklist.md` を参照してください。

## プロジェクト構造

```
excel_to_markdown/
├── pyproject.toml               # プロジェクト設定・依存管理・ツール設定
├── requirements.txt             # 本番依存（openpyxl等）
├── requirements-dev.txt         # 開発依存（pytest, ruff, mypy等）
│
├── excel_to_markdown/           # メインパッケージ
│   ├── __init__.py              # バージョン定義
│   ├── __main__.py              # `python -m excel_to_markdown` エントリーポイント
│   ├── cli.py                   # CLIパース（argparse）・パイプライン起動
│   ├── models.py                # 全dataclass定義（RawCell/TextBlock/DocElement等）
│   │
│   ├── reader/                  # 読み込みレイヤー
│   │   ├── __init__.py
│   │   ├── xlsx_reader.py       # openpyxl → list[RawCell]
│   │   └── xls_reader.py        # xlrd → list[RawCell]（P1）
│   │
│   ├── parser/                  # 解析レイヤー
│   │   ├── __init__.py
│   │   ├── cell_grid.py         # CellGrid（col_unit・modal_row_height算出）
│   │   ├── merge_resolver.py    # list[RawCell] → list[TextBlock]
│   │   ├── structure_detector.py # list[TextBlock] → list[DocElement]
│   │   └── table_detector.py    # グリッド表検出 → TableElement
│   │
│   └── renderer/                # 出力レイヤー
│       ├── __init__.py
│       └── markdown_renderer.py # list[DocElement] → Markdown文字列
│
├── tests/                       # テストコード
│   ├── __init__.py
│   ├── conftest.py              # 共通フィクスチャ（xlsxビルダー等）
│   ├── fixtures/                # ゴールデンファイル（期待値.md）
│   │   ├── simple_heading.md
│   │   ├── nested_list.md
│   │   ├── label_value.md
│   │   ├── mixed_document.md
│   │   ├── with_comment.md
│   │   └── print_area.md
│   ├── test_xlsx_reader.py
│   ├── test_cell_grid.py
│   ├── test_merge_resolver.py
│   ├── test_table_detector.py
│   ├── test_structure_detector.py
│   ├── test_markdown_renderer.py
│   └── test_integration.py      # エンドツーエンド統合テスト
│
├── docs/                        # プロジェクトドキュメント
│   ├── ideas/
│   │   └── initial-requirements.md  # 初期要件（アーカイブ）
│   ├── product-requirements.md
│   ├── functional-design.md
│   ├── architecture.md
│   ├── repository-structure.md  # 本ドキュメント
│   ├── development-guidelines.md
│   └── glossary.md
│
├── .claude/                     # Claude Code設定（リポジトリに含める）
│   ├── settings.json            # Claude Code設定
│   ├── settings.local.json      # ローカル設定（gitignore対象）
│   ├── commands/                # スラッシュコマンド定義
│   │   ├── setup-project.md
│   │   ├── add-feature.md
│   │   ├── review-docs.md
│   │   └── refactor.md
│   ├── skills/                  # Claude Codeスキル定義
│   │   ├── prd-writing/
│   │   ├── functional-design/
│   │   ├── architecture-design/
│   │   ├── repository-structure/
│   │   ├── development-guidelines/
│   │   ├── glossary-creation/
│   │   └── steering/
│   └── agents/                  # サブエージェント定義
│       ├── doc-reviewer.md
│       └── implementation-validator.md
│
├── .steering/                   # 開発作業ステアリングファイル（一時）
│
├── .gitignore
└── README.md                    # セットアップ・使い方
```

---

## ディレクトリ詳細

### excel_to_markdown/（メインパッケージ）

#### models.py

**役割**: パイプライン全体で使用するデータモデルを一元定義する

**配置ファイル**:
- `RawCell`, `TextBlock`, `InlineRun`, `DocElement`, `TableElement`, `TableCell`, `ElementType` を定義

**依存関係**:
- 依存可能: Python標準ライブラリのみ（dataclasses, enum）
- 依存禁止: openpyxl, xlrd, パッケージ内の他モジュール（循環依存防止）

---

#### reader/

**役割**: Excelファイルを読み込み、`list[RawCell]` を返す

**配置ファイル**:
- `xlsx_reader.py`: openpyxlでの.xlsx読み込み
- `xls_reader.py`: xlrdでの.xls読み込み（P1）

**命名規則**:
- `xlsx_reader.py` のエントリーポイント: `read_sheet(ws, print_area) -> list[RawCell]`
- `xls_reader.py` のエントリーポイント: `read_sheet_xls(sheet, wb) -> list[RawCell]`（P1、xlrdのAPIが異なるため別名）
- `cli.py` がファイル拡張子に応じてどちらを呼ぶか切り替えるアダプターとして機能する

**依存関係**:
- 依存可能: `models.py`、openpyxl、xlrd
- 依存禁止: `parser/`、`renderer/`（下流モジュールへの依存禁止）

---

#### parser/

**役割**: RawCell → TextBlock → DocElement の変換を担う解析レイヤー

**配置ファイル**:
- `cell_grid.py`: 空間クエリ（col_unit, modal_row_height等）
- `merge_resolver.py`: RawCellからTextBlockを生成
- `table_detector.py`: グリッド表の検出
- `structure_detector.py`: 文書構造の分類（最も複雑なコンポーネント）

**依存関係**:
- 依存可能: `models.py`
- 依存禁止: `reader/`、`renderer/`、`cli.py`

**モジュール内の役割**:
- 空間解析サブレイヤー: `cell_grid` → `merge_resolver`（RawCell → TextBlock）
- 構造検出サブレイヤー: `table_detector` → `structure_detector`（TextBlock → DocElement）

**モジュール内依存順序**（循環禁止）:
```
cell_grid → merge_resolver → table_detector → structure_detector
```

---

#### renderer/

**役割**: DocElementリストをMarkdown文字列に変換する

**配置ファイル**:
- `markdown_renderer.py`: レンダリングのみ。副作用（ファイル書き込み）はcli.pyが担当

**依存関係**:
- 依存可能: `models.py`
- 依存禁止: `reader/`、`parser/`、`cli.py`

---

### tests/

**役割**: 全テストコードの配置。ユニットテストと統合テストを同ディレクトリに配置

#### conftest.py

共通フィクスチャを定義する:

```python
@pytest.fixture
def xlsx_builder():
    """openpyxlでxlsxワークブックをプログラム的に生成するヘルパー。"""
    ...
```

#### fixtures/（ゴールデンファイル）

統合テストの期待値となるMarkdownファイルを配置する。
テスト用の.xlsxファイルはバイナリのためリポジトリに含めず、`conftest.py` でプログラム生成する。

| ゴールデンファイル | 対応するフィクスチャ |
|-------------------|------------------|
| `simple_heading.md` | シンプルな見出し+段落文書 |
| `nested_list.md` | 3段階ネストリスト |
| `label_value.md` | ラベル:値パターン |
| `mixed_document.md` | 見出し・表・段落・リスト混在 |
| `with_comment.md` | セルコメント付き（脚注変換） |
| `print_area.md` | 印刷領域設定あり |

---

## ファイル配置規則

### ソースファイル

| ファイル種別 | 配置先 | 命名規則 | 例 |
|------------|--------|---------|-----|
| データモデル | `excel_to_markdown/models.py` | 単一ファイル | `models.py` |
| 読み込みモジュール | `excel_to_markdown/reader/` | `{format}_reader.py` | `xlsx_reader.py` |
| 解析モジュール | `excel_to_markdown/parser/` | 処理内容を表すsnake_case | `structure_detector.py` |
| 出力モジュール | `excel_to_markdown/renderer/` | `{format}_renderer.py` | `markdown_renderer.py` |

### テストファイル

| テスト種別 | 配置先 | 命名規則 | 例 |
|-----------|--------|---------|-----|
| ユニットテスト | `tests/` | `test_{対象モジュール}.py` | `test_structure_detector.py` |
| 統合テスト | `tests/` | `test_integration.py` | — |
| ゴールデンファイル | `tests/fixtures/` | `{fixture_name}.md` | `mixed_document.md` |

**P1機能のテスト**: `test_xls_reader.py` は `.xls` 対応（P1）の実装時に追加する。

### ドキュメントファイル

| ドキュメント種別 | 配置先 | 命名規則 | 例 |
|----------------|--------|---------|-----|
| 正式版ドキュメント | `docs/` | `{document-type}.md`（ハイフン区切り） | `functional-design.md` |
| アイデア・下書き | `docs/ideas/` | 自由形式 | `initial-requirements.md` |

### 設定ファイル

| ファイル種別 | 配置先 | 命名規則 |
|------------|--------|---------|
| プロジェクト設定 | ルート | `pyproject.toml` |
| 本番依存 | ルート | `requirements.txt` |
| 開発依存 | ルート | `requirements-dev.txt` |
| Python仮想環境 | ルート（gitignore対象） | `.venv/` |

---

## 命名規則

### ディレクトリ名

- **パッケージ**: snake_case（例: `excel_to_markdown/`, `reader/`, `parser/`）
- **テスト**: `tests/`（複数形）

### ファイル名

- **モジュール**: snake_case（例: `xlsx_reader.py`, `structure_detector.py`）
- **テスト**: `test_` プレフィックス + 対象モジュール名（例: `test_xlsx_reader.py`）

### クラス・型

- **dataclass / Enum**: PascalCase（例: `RawCell`, `DocElement`, `ElementType`）

### 関数・変数

- **関数**: snake_case（例: `read_sheet`, `compute_indent_tiers`）
- **定数**: UPPER_SNAKE_CASE（例: `DEFAULT_BASE_FONT_SIZE = 11.0`）

---

## 依存関係のルール

### レイヤー間の依存（一方向のみ許可）

```
cli.py / __main__.py
    ↓
reader/ (xlsx_reader, xls_reader)
    ↓
parser/
  ├── 空間解析: cell_grid → merge_resolver   (RawCell → TextBlock)
  └── 構造検出: table_detector → structure_detector  (TextBlock → DocElement)
    ↓
renderer/ (markdown_renderer)
    ↓
models.py ←── 全モジュールが参照可能（共有レイヤー）
```

**禁止される依存**:
- `models.py` から他の内部モジュールへの依存
- `parser/` から `reader/` への依存（逆流）
- `renderer/` から `parser/` への依存（逆流）
- モジュール間の循環依存

### 外部ライブラリの利用制限

| ライブラリ | 使用可能なモジュール |
|-----------|-------------------|
| openpyxl | `reader/xlsx_reader.py` のみ |
| xlrd | `reader/xls_reader.py` のみ |
| argparse | `cli.py` のみ |

---

## スケーリング戦略

### 新しい入力フォーマットの追加

`reader/` に新しい `{format}_reader.py` を追加し、`list[RawCell]` を返すインターフェースを実装する。`cli.py` で拡張子により切り替えるだけで、下流（parser/renderer）への変更は不要。

### 新しい出力フォーマットの追加

`renderer/` に新しい `{format}_renderer.py` を追加する。`DocElement` を入力とする純関数として実装。

### ファイルサイズの管理

| ファイル行数 | 対応方針 |
|------------|---------|
| ～300行 | そのまま維持 |
| 300～500行 | リファクタリングを検討 |
| 500行超 | サブモジュールに分割（例: `parser/structure_detector/` パッケージ化） |

---

## 特殊ディレクトリ

### .steering/（ステアリングファイル）

**役割**: 開発作業ごとの要件・設計・タスクリストを定義する一時的なファイル

```
.steering/
└── 20260320-initial-implementation/
    ├── requirements.md
    ├── design.md
    └── tasklist.md
```

**.gitignore に追加**: ステアリングファイルはリポジトリには含めない。作業完了後も履歴として `.steering/` ディレクトリに保持し、削除はしない。

### .claude/（Claude Code設定）

Claude Codeのカスタマイズ設定。このディレクトリはリポジトリに含める。

| サブディレクトリ | 役割 |
|-----------------|------|
| `commands/` | `/setup-project` 等のスラッシュコマンドを `.md` ファイルで定義 |
| `skills/` | 各スキル（PRD作成・機能設計等）のガイドとテンプレートを格納 |
| `agents/` | `doc-reviewer` 等のサブエージェントの動作定義 |
| `settings.json` | Claude Codeの権限・フック設定（リポジトリに含める） |
| `settings.local.json` | ローカル固有の設定（gitignore対象） |

---

## 除外設定（.gitignore 定義）

以下の内容で `.gitignore` を作成すること。

```gitignore
# Python
.venv/
__pycache__/
*.pyc
*.pyo
*.pyd
.Python

# テスト・カバレッジ
.pytest_cache/
.coverage
htmlcov/

# ビルド
dist/
build/
*.egg-info/

# OS
.DS_Store
Thumbs.db

# 開発作業
.steering/

# ログ
*.log
```

---

## pyproject.toml 構成

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "excel-to-markdown"
version = "0.1.0"
requires-python = ">=3.12"
dependencies = [
    "openpyxl>=3.1.0,<4.0.0",
]

[project.optional-dependencies]
xls = ["xlrd>=2.0.0,<3.0.0"]   # .xls対応（P1）
dev = [
    "pytest>=8.0.0",
    "pytest-cov>=5.0.0",
    "ruff>=0.4.0",
    "mypy>=1.10.0",
]

[project.scripts]
excel-to-markdown = "excel_to_markdown.__main__:main"

[tool.ruff]
target-version = "py312"
line-length = 100

[tool.mypy]
python_version = "3.12"
strict = true

[tool.pytest.ini_options]
testpaths = ["tests"]
addopts = "--cov=excel_to_markdown --cov-report=term-missing"
```
