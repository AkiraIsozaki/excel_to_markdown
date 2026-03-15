# 開発ガイドライン (Development Guidelines)

## コーディング規約

### 命名規則

#### 変数・関数

**Python**:
```python
# 良い例
user_profile_data = fetch_user_profile()
def calculate_total_price(items: list[dict]) -> float: ...

# 悪い例
data = fetch()
def calc(arr): ...
```

**原則**:
- 変数: snake_case、名詞または名詞句
- 関数: snake_case、動詞で始める
- 定数: UPPER_SNAKE_CASE
- Boolean: `is_`, `has_`, `should_` で始める

#### クラス・列挙型

```python
# クラス: PascalCase、名詞
class GrepAnalyzer: ...
class UsageClassifier: ...

# Enum: PascalCase、値はわかりやすい文字列
from enum import Enum

class UsageType(Enum):
    CONSTANT = "定数定義"
    VARIABLE = "変数代入"
    CONDITION = "条件判定"
    RETURN = "return文"
    ARGUMENT = "メソッド引数"
    ANNOTATION = "アノテーション"
    OTHER = "その他"
```

### コードフォーマット

**インデント**: 4スペース

**行の長さ**: 最大120文字

**例**:
```python
# Python コードフォーマット例
from dataclasses import dataclass

@dataclass(frozen=True)
class GrepRecord:
    keyword: str
    ref_type: str       # 直接 / 間接 / 間接（getter経由）
    usage_type: str
    filepath: str
    lineno: str
    code: str
    src_var: str = ""
    src_file: str = ""
    src_lineno: str = ""
```

### コメント規約

**関数・クラスのドキュメント（docstring）**:
```python
def classify_usage(code: str) -> str:
    """コード行を解析し、使用タイプを返す。

    Args:
        code: 分類対象のコード行（前後の空白はtrim済み）

    Returns:
        使用タイプ文字列（7種のいずれか）

    Raises:
        ValueError: codeがNoneの場合
    """
    ...
```

**インラインコメント**:
```python
# 良い例: なぜそうするかを説明
# キャッシュを使い、同一ファイルの再解析を避ける
if filepath not in ast_cache:
    ast_cache[filepath] = parse_java_file(filepath)

# 悪い例: 何をしているか（コードを見れば分かる）
# キャッシュを確認する
if filepath not in ast_cache:
    ...
```

### エラーハンドリング

**原則**:
- 予期されるエラー: 適切な例外クラスを定義
- 予期しないエラー: 上位に伝播
- 例外を無視しない

**例**:
```python
# 例外クラス定義
class ValidationError(ValueError):
    def __init__(self, message: str, field: str, value: object):
        super().__init__(message)
        self.field = field
        self.value = value

# エラーハンドリング
try:
    records = process_file(input_path, keyword)
except ValidationError as e:
    print(f"検証エラー [{e.field}]: {e}", file=sys.stderr)
except Exception as e:
    print(f"予期しないエラー: {e}", file=sys.stderr)
    raise  # 上位に伝播
```

## Git運用ルール

### ブランチ戦略

**ブランチ種別**:
- `main`: 本番環境にデプロイ可能な状態
- `develop`: 開発の最新状態
- `feature/[機能名]`: 新機能開発
- `fix/[修正内容]`: バグ修正
- `refactor/[対象]`: リファクタリング

**フロー**:
```
main
  └─ develop
      ├─ feature/indirect-tracking
      ├─ feature/getter-tracking
      └─ fix/parse-error-handling
```

### コミットメッセージ規約

**フォーマット**:
```
<type>(<scope>): <subject>

<body>

<footer>
```

**Type**:
- `feat`: 新機能
- `fix`: バグ修正
- `docs`: ドキュメント
- `style`: コードフォーマット
- `refactor`: リファクタリング
- `test`: テスト追加・修正
- `chore`: ビルド、補助ツール等

**例**:
```
feat(tracker): getter経由の間接参照追跡機能を追加

フィールドに代入された値がgetter経由で使われる箇所を追跡する。
- 命名規則（type → getType()）によるgetter候補特定
- return文解析による非標準命名のgetter検出
- プロジェクト全体でのgetter呼び出し箇所のAST解析

Closes #12
```

### プルリクエストプロセス

**作成前のチェック**:
- [ ] 全てのテストがパス (`python -m unittest discover`)
- [ ] 構文エラーがない (`python -m py_compile analyze.py`)
- [ ] 競合が解決されている

**PRテンプレート**:
```markdown
## 概要
[変更内容の簡潔な説明]

## 変更理由
[なぜこの変更が必要か]

## 変更内容
- [変更点1]
- [変更点2]

## テスト
- [ ] ユニットテスト追加
- [ ] 手動テスト実施

## 関連Issue
Closes #[Issue番号]
```

**レビュープロセス**:
1. セルフレビュー
2. 自動テスト実行
3. レビュアーアサイン
4. レビューフィードバック対応
5. 承認後マージ

## テスト戦略

### テストの種類

#### ユニットテスト

**対象**: 個別の関数・クラス

**カバレッジ目標**: 80%

**例**:
```python
import unittest

class TestUsageClassifier(unittest.TestCase):

    def setUp(self):
        self.classifier = UsageClassifier()

    def test_classify_constant_definition(self):
        code = 'public static final String CODE = "TARGET";'
        result = self.classifier.classify(code)
        self.assertEqual(result, "定数定義")

    def test_classify_condition(self):
        code = 'if (x.equals("TARGET")) {'
        result = self.classifier.classify(code)
        self.assertEqual(result, "条件判定")

    def test_classify_empty_raises(self):
        with self.assertRaises(ValueError):
            self.classifier.classify(None)
```

#### 統合テスト

**対象**: 複数コンポーネントの連携

**例**:
```python
import unittest
from pathlib import Path

class TestAnalyzeIntegration(unittest.TestCase):

    def setUp(self):
        self.test_input = Path("tests/fixtures/input")
        self.test_output = Path("tests/fixtures/output")

    def test_full_analysis_produces_tsv(self):
        # 準備: サンプルgrep結果を入力に置く
        # 実行: 分析を実行する
        records = analyze(self.test_input, source_dir=Path("tests/fixtures/java"))
        # 検証: 期待される件数・分類が出力されている
        self.assertGreater(len(records), 0)
        direct = [r for r in records if r.ref_type == "直接"]
        self.assertGreater(len(direct), 0)
```

### テスト命名規則

**パターン**: `test_[対象]_[条件]_[期待結果]`

**例**:
```python
# 良い例
def test_parse_valid_grep_line_returns_dict(self): ...
def test_parse_binary_notice_line_returns_none(self): ...
def test_classify_static_final_as_constant_definition(self): ...

# 悪い例
def test1(self): ...
def test_works(self): ...
def test_should_work_correctly(self): ...
```

### モック・スタブの使用

**原則**:
- 外部依存（ファイルシステム、外部API）はモック化
- ビジネスロジックは実装を使用

**例**:
```python
from unittest.mock import patch, MagicMock

class TestGrepParser(unittest.TestCase):

    @patch("builtins.open", new_callable=MagicMock)
    def test_process_file_reads_grep_lines(self, mock_open):
        mock_open.return_value.__enter__.return_value = [
            "src/Constants.java:10:    public static final String CODE = \"TARGET\";\n",
        ]
        records = process_file(Path("input/TARGET.grep"), "TARGET")
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0].usage_type, "定数定義")
```

## コードレビュー基準

### レビューポイント

**機能性**:
- [ ] 要件を満たしているか
- [ ] エッジケースが考慮されているか
- [ ] エラーハンドリングが適切か

**可読性**:
- [ ] 命名が明確か（snake_case）
- [ ] コメントが適切か
- [ ] 複雑なロジックが説明されているか

**保守性**:
- [ ] 重複コードがないか
- [ ] 責務が明確に分離されているか
- [ ] 変更の影響範囲が限定的か

**パフォーマンス**:
- [ ] 不要な計算がないか
- [ ] ASTキャッシュを活用しているか
- [ ] データ構造やアルゴリズムが適切か

**セキュリティ**:
- [ ] 入力検証が適切か
- [ ] ファイルパストラバーサル対策がされているか
- [ ] 機密情報がハードコードされていないか

### レビューコメントの書き方

**建設的なフィードバック**:
```markdown
## 良い例
この実装だと、ファイル数が多い場合にAST再解析が大量発生してパフォーマンスが劣化します。
`ast_cache: dict[str, object] = {}` でキャッシュする実装を検討してはどうでしょうか？

## 悪い例
この書き方は良くないです。
```

**優先度の明示**:
- `[必須]`: 修正必須
- `[推奨]`: 修正推奨
- `[提案]`: 検討してほしい
- `[質問]`: 理解のための質問

## 開発環境セットアップ

### 必要なツール

| ツール | バージョン | インストール方法 |
|--------|-----------|-----------------|
| Python | 3.12以上 | devcontainer に含まれる |
| venv | Python標準 | `python -m venv .venv` |
| javalang | 最新安定版 | `pip install javalang` |

### セットアップ手順

```bash
# 1. リポジトリのクローン
git clone [URL]
cd grep_analyzer

# 2. venv作成と依存関係のインストール
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 3. テストの実行
python -m unittest discover

# 4. ツールの実行
python analyze.py --source-dir /path/to/javaproject
```

### 推奨開発ツール（該当する場合）

- VS Code + Python拡張機能: Python開発の基本拡張機能
- VS Code + Pylance: 型チェックとインテリセンス
