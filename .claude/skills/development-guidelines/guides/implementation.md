# 実装ガイド (Implementation Guide)

## Python 規約

### 型定義

**型ヒントを使用**:
```python
# 良い例: 型ヒントで意図を明確に
def process_items(items: list[str]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for item in items:
        counts[item] = counts.get(item, 0) + 1
    return counts

# 悪い例: 型ヒントなし（意図が不明）
def process_items(items):
    counts = {}
    for item in items:
        counts[item] = counts.get(item, 0) + 1
    return counts
```

**型エイリアスの活用**:
```python
# 複雑な型には型エイリアスを定義
from typing import TypeAlias

GrepRecord: TypeAlias = dict[str, str]
UsageType: TypeAlias = str
```

**dataclassの活用**:
```python
from dataclasses import dataclass

# データクラス: 不変データキャリア
@dataclass(frozen=True)
class GrepRecord:
    keyword: str
    usage_type: str
    filepath: str
    lineno: str
    code: str

# Enum: 固定の選択肢
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

### 命名規則

**変数・関数**:
```python
# 変数: snake_case、名詞
user_name = "John"
task_list: list[str] = []
is_completed = True

# 関数: snake_case、動詞で始める
def fetch_user_data() -> dict: ...
def validate_email(email: str) -> bool: ...
def calculate_total_price(items: list) -> float: ...

# Boolean: is_, has_, should_, can_ で始める
is_valid = True
has_permission = False
should_retry = True
can_delete = False
```

**クラス・定数**:
```python
# クラス: PascalCase、名詞
class GrepAnalyzer: ...
class UsageClassifier: ...

# 定数: UPPER_SNAKE_CASE
MAX_RETRY_COUNT = 3
API_BASE_URL = "https://api.example.com"
DEFAULT_TIMEOUT = 5000
```

**ファイル名**:
```
# モジュール: snake_case
analyze.py
usage_classifier.py
grep_parser.py

# テストファイル: test_ プレフィックス
test_analyze.py
test_usage_classifier.py
```

### メソッド設計

**単一責務の原則**:
```python
# 良い例: 単一の責務
def calculate_total_price(items: list[dict]) -> float:
    return sum(item["price"] * item["quantity"] for item in items)

def format_price(amount: float) -> str:
    return f"¥{amount:,.0f}"

# 悪い例: 複数の責務
def calculate_and_format_price(items: list[dict]) -> str:
    total = sum(item["price"] * item["quantity"] for item in items)
    return f"¥{total:,.0f}"
```

**関数の長さ**:
- 目標: 20行以内
- 推奨: 50行以内
- 100行以上: リファクタリングを検討

**パラメータの数**:
```python
# 良い例: dataclassでまとめる
@dataclass
class CreateTaskOptions:
    title: str
    description: str = ""
    priority: str = "MEDIUM"
    due_date: str | None = None

def create_task(options: CreateTaskOptions) -> dict:
    ...

# 悪い例: パラメータが多すぎる
def create_task(title, description, priority, due_date, tags, assignee):
    ...
```

### エラーハンドリング

**カスタム例外クラス**:
```python
class ValidationError(ValueError):
    def __init__(self, message: str, field: str, value: object):
        super().__init__(message)
        self.field = field
        self.value = value

class NotFoundException(Exception):
    def __init__(self, resource: str, id: str):
        super().__init__(f"{resource} not found: {id}")
        self.resource = resource
        self.id = id
```

**エラーハンドリングパターン**:
```python
# 良い例: 適切なエラーハンドリング
def get_task(task_id: str) -> dict:
    task = repository.find_by_id(task_id)
    if task is None:
        raise NotFoundException("Task", task_id)
    return task

# 悪い例: 例外を握りつぶす
def get_task(task_id: str) -> dict | None:
    try:
        return repository.find_by_id(task_id)
    except Exception:
        return None  # エラー情報が失われる
```

**エラーメッセージ**:
```python
# 良い例: 具体的で解決策を示す
raise ValidationError(
    f"タイトルは1-200文字で入力してください。現在の文字数: {len(title)}",
    field="title",
    value=title,
)

# 悪い例: 曖昧で役に立たない
raise ValueError("Invalid input")
```

### 並行処理

**concurrent.futuresの使用**:
```python
from concurrent.futures import ThreadPoolExecutor

# 良い例: ThreadPoolExecutorで並列処理
def fetch_multiple_users(ids: list[str]) -> list[dict]:
    with ThreadPoolExecutor() as executor:
        results = list(executor.map(repository.find_by_id, ids))
    return [r for r in results if r is not None]

# 同期処理で十分な場合はシンプルに
def fetch_multiple_users(ids: list[str]) -> list[dict]:
    return [repository.find_by_id(id) for id in ids if repository.find_by_id(id)]
```

## コメント規約

### ドキュメントコメント

**docstring形式**:
```python
def create_task(data: CreateTaskOptions) -> dict:
    """タスクを作成する。

    指定されたデータに基づいて新しいタスクを作成し、
    リポジトリに永続化する。

    Args:
        data: 作成するタスクのデータ

    Returns:
        作成されたタスク辞書

    Raises:
        ValidationError: データが不正な場合
    """
    ...
```

### インラインコメント

**良いコメント**:
```python
# 理由を説明
# キャッシュを無効化して最新データを取得
cache.clear()

# 複雑なロジックを説明
# Kadaneのアルゴリズムで最大部分配列和を計算
# 時間計算量: O(n)
max_so_far = arr[0]
max_ending_here = arr[0]

# TODO・FIXMEを活用
# TODO: キャッシュ機能を実装 (Issue #123)
# FIXME: 大量データでパフォーマンス劣化 (Issue #456)
```

**悪いコメント**:
```python
# コードの内容を繰り返すだけ
# iを1増やす
i += 1

# コメントアウトされたコード（削除すべき）
# old_implementation = OldClass()
```

## セキュリティ

### 入力検証

```python
# 良い例: 厳密な検証
import re

def validate_email(email: str) -> None:
    if not email or not email.strip():
        raise ValidationError("メールアドレスは必須です", "email", email)

    pattern = r'^[^\s@]+@[^\s@]+\.[^\s@]+$'
    if not re.match(pattern, email):
        raise ValidationError("メールアドレスの形式が不正です", "email", email)

    if len(email) > 254:
        raise ValidationError("メールアドレスが長すぎます", "email", email)

# 悪い例: 検証なし
def validate_email(email: str) -> None:
    pass
```

### 機密情報の管理

```python
# 良い例: 環境変数から読み込み
import os

api_key = os.environ.get("API_KEY")
if not api_key:
    raise EnvironmentError("API_KEY環境変数が設定されていません")

# 悪い例: ハードコード
api_key = "sk-1234567890abcdef"  # 絶対にしない！
```

## パフォーマンス

### データ構造の選択

```python
# 良い例: dictでO(1)アクセス
user_map = {user["id"]: user for user in users}
user = user_map.get(user_id)  # O(1)

# 悪い例: listでO(n)検索
user = next((u for u in users if u["id"] == user_id), None)  # O(n)
```

### ループの最適化

```python
# 良い例: リスト内包表記
active_items = [item for item in items if item["is_active"]]

# 良い例: ジェネレータ（大量データ）
def process_large_file(path: Path):
    with open(path, encoding="utf-8") as f:
        for line in f:  # 1行ずつ読み込み（メモリ効率良）
            yield process_line(line)

# 悪い例: 不要なリスト生成
result = list(filter(lambda x: x["is_active"], items))  # わかりにくい
```

### メモ化

```python
# 計算結果のキャッシュ
from functools import lru_cache

@lru_cache(maxsize=None)
def expensive_calculation(input: str) -> str:
    # 重い計算
    return do_expensive_work(input)

# 辞書による手動キャッシュ（ミュータブルな引数の場合）
_cache: dict[str, object] = {}

def cached_parse(filepath: str) -> object:
    if filepath not in _cache:
        _cache[filepath] = parse_file(filepath)
    return _cache[filepath]
```

## テストコード

### テストの構造 (Arrange-Act-Assert)

```python
import unittest

class TestGrepAnalyzer(unittest.TestCase):

    def setUp(self):
        self.analyzer = GrepAnalyzer()

    def test_parse_valid_line(self):
        # Arrange: 準備
        line = "src/main/java/Foo.java:42:    String msg = \"ERROR\";"

        # Act: 実行
        result = self.analyzer.parse_grep_line(line)

        # Assert: 検証
        self.assertIsNotNone(result)
        self.assertEqual(result["filepath"], "src/main/java/Foo.java")
        self.assertEqual(result["lineno"], "42")
        self.assertIn("msg", result["code"])

    def test_parse_invalid_line_returns_none(self):
        # Arrange: 準備
        line = "Binary file matches"

        # Act: 実行
        result = self.analyzer.parse_grep_line(line)

        # Assert: 検証
        self.assertIsNone(result)
```

### モックの作成

```python
from unittest.mock import patch, MagicMock

class TestGrepAnalyzer(unittest.TestCase):

    @patch("analyze.open", new_callable=MagicMock)
    def test_process_file(self, mock_open):
        # ファイル読み込みをモック
        mock_open.return_value.__enter__.return_value = [
            "src/Foo.java:10:    public static final String CODE = \"VALUE\";",
        ]

        records = process_file(Path("input/VALUE.grep"), "VALUE")

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["usage_type"], "定数定義")
```

## リファクタリング

### マジックナンバーの排除

```python
# 良い例: 定数を定義
MAX_RETRY_COUNT = 3
RETRY_DELAY_SEC = 1.0

def fetch_with_retry() -> dict:
    for i in range(MAX_RETRY_COUNT):
        try:
            return fetch_data()
        except Exception:
            if i < MAX_RETRY_COUNT - 1:
                time.sleep(RETRY_DELAY_SEC)
    raise RuntimeError("最大リトライ回数を超過しました")

# 悪い例: マジックナンバー
def fetch_with_retry() -> dict:
    for i in range(3):
        try:
            return fetch_data()
        except Exception:
            if i < 2:
                time.sleep(1)
    raise RuntimeError("失敗")
```

### 関数の抽出

```python
# 良い例: 関数を抽出
def process_order(order: dict) -> None:
    validate_order(order)
    calculate_total(order)
    apply_discounts(order)
    save_order(order)

def validate_order(order: dict) -> None:
    if not order.get("items"):
        raise ValidationError("商品が選択されていません", "items", order.get("items"))

def calculate_total(order: dict) -> None:
    order["total"] = sum(
        item["price"] * item["quantity"] for item in order["items"]
    )

# 悪い例: 長い関数
def process_order(order: dict) -> None:
    if not order.get("items"):
        raise ValidationError("商品が選択されていません", "items", order.get("items"))
    order["total"] = sum(
        item["price"] * item["quantity"] for item in order["items"]
    )
    if order.get("coupon"):
        order["total"] -= order["total"] * order["coupon"]["discount_rate"]
    repository.save(order)
```

## チェックリスト

実装完了前に確認:

### コード品質
- [ ] 命名が明確で一貫している（snake_case）
- [ ] 関数が単一の責務を持っている
- [ ] マジックナンバーがない
- [ ] 型ヒントが適切に定義されている
- [ ] エラーハンドリングが実装されている

### セキュリティ
- [ ] 入力検証が実装されている
- [ ] 機密情報がハードコードされていない
- [ ] ファイルパスのトラバーサル対策がされている（該当する場合）

### パフォーマンス
- [ ] 適切なデータ構造を使用している
- [ ] 不要な計算を避けている
- [ ] 大量データはジェネレータで処理している

### テスト
- [ ] unittestでユニットテストが書かれている
- [ ] テストがパスする（`python -m unittest`）
- [ ] エッジケースがカバーされている

### ドキュメント
- [ ] 関数にdocstringコメントがある（複雑なもののみ）
- [ ] 複雑なロジックにコメントがある
- [ ] TODOやFIXMEが記載されている（該当する場合）

### ツール
- [ ] 構文エラーがない（`python -m py_compile`）
- [ ] テストが成功する（`python -m unittest discover`）
- [ ] コードフォーマットが統一されている
