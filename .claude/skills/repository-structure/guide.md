# リポジトリ構造定義書作成ガイド

## 基本原則

### 1. 役割の明確化

各ディレクトリ（モジュール）は単一の明確な役割を持つべきです。

**悪い例**:
```
src/
├── stuff/           # 曖昧
├── misc/            # 雑多
└── utils/           # 汎用的すぎる
```

**良い例**:
```
src/
├── parser/          # 入力データのパース処理
├── classifier/      # 使用タイプの分類処理
├── tracker/         # 間接参照の追跡処理
└── reporter/        # 出力レポート生成
```

### 2. レイヤー分離の徹底

アーキテクチャのレイヤー構造をディレクトリ構造に反映させます:

```
src/
├── io/              # 入出力レイヤー（CLIパース、TSV出力）
├── analysis/        # 分析レイヤー（AST解析、分類）
└── tracking/        # 追跡レイヤー（間接参照追跡）
```

### 3. 技術要素ベースの分割(基本)

関連する技術要素ごとにモジュールを分割します:

**基本構造**:
```
src/
├── parser.py        # 入力パース処理
├── classifier.py    # 使用タイプ分類
├── tracker.py       # 間接参照追跡
└── models.py        # データモデル・型定義
```

**レイヤー構造との対応**:
```
入出力レイヤー       → parser.py, writer.py
分析レイヤー         → classifier.py, ast_analyzer.py
追跡レイヤー         → tracker.py
```

## ディレクトリ構造の設計

### レイヤー構造の表現

```python
# 悪い例: 平坦な構造
analyze.py          # 全処理が1ファイルに混在（1000行超）

# 良い例: 責務ごとに分離
src/
├── parser.py       # 入力パース
├── classifier.py   # 使用タイプ分類
├── tracker.py      # 間接参照追跡
└── writer.py       # TSV出力
```

### テストディレクトリの配置

**推奨構造**:
```
project/
├── src/
│   ├── parser.py
│   ├── classifier.py
│   └── tracker.py
├── tests/
│   ├── unit/
│   │   ├── test_parser.py
│   │   ├── test_classifier.py
│   │   └── test_tracker.py
│   └── integration/
│       └── test_analyze.py
└── analyze.py      # エントリーポイント
```

**理由**:
- テストコードが本番コードと分離
- `python -m unittest discover` で自動検出
- テストタイプごとに整理可能
- 依存関係が明確

## 命名規則のベストプラクティス

### モジュール名の原則

**1. snake_case・小文字を使う (Python慣例)**
```
✅ parser.py
✅ usage_classifier.py
✅ grep_analyzer.py

❌ Parser.py
❌ UsageClassifier.py
❌ GrepAnalyzer.py
```

理由: PEP 8のモジュール命名規約に準拠（小文字・アンダースコア）

**2. 具体的な名前を使う**
```
✅ grep_parser.py        # grep結果のパース
✅ usage_classifier.py   # 使用タイプの分類
✅ indirect_tracker.py   # 間接参照の追跡

❌ util.py               # 汎用的すぎる
❌ helper.py             # 曖昧
❌ common.py             # 意味不明
```

### ファイル名の原則

**1. モジュールファイル: snake_case**
```python
# 処理モジュール
grep_parser.py
usage_classifier.py
indirect_tracker.py

# データモデル
models.py
```

**2. テストファイル: test_ プレフィックス**
```python
# テストファイル
test_grep_parser.py
test_usage_classifier.py
test_indirect_tracker.py
```

**3. クラス: PascalCase**
```python
class GrepAnalyzer: ...
class UsageClassifier: ...
class IndirectTracker: ...
```

**4. 定数: UPPER_SNAKE_CASE**
```python
MAX_FILE_SIZE = 500 * 1024 * 1024
DEFAULT_ENCODING = "shift_jis"
```

## 依存関係の管理

### レイヤー間の依存ルール

```python
# ✅ 良い例: 上位レイヤーから下位レイヤーへの依存
# analyze.py (エントリーポイント)
from src.parser import parse_grep_file
from src.classifier import classify_usage

# ❌ 悪い例: 下位レイヤーから上位レイヤーへの依存
# src/classifier.py
from analyze import main  # 禁止！
```

### 循環依存の回避

**問題のあるコード**:
```python
# src/classifier.py
from src.tracker import IndirectTracker  # 循環依存の起点

# src/tracker.py
from src.classifier import classify_usage  # 循環依存！
```

**解決策1: 共通モデルを抽出**
```python
# src/models.py
from dataclasses import dataclass

@dataclass(frozen=True)
class GrepRecord:
    keyword: str
    usage_type: str
    filepath: str
    lineno: str
    code: str

# src/classifier.py
from src.models import GrepRecord  # モデルのみに依存

# src/tracker.py
from src.models import GrepRecord  # モデルのみに依存
```

**解決策2: 依存関係を見直す**
```python
# 共通の機能を別モジュールに抽出
# src/ast_utils.py
def parse_java_file(filepath: str) -> object:
    """Javaファイルをパースし、ASTを返す"""
    ...

# src/classifier.py
from src.ast_utils import parse_java_file

# src/tracker.py
from src.ast_utils import parse_java_file
```

## スケーリング戦略

### 推奨構造

**標準パターン（単一スクリプト）**:
```
grep_analyzer/
├── analyze.py           # エントリーポイント（全ロジック）
├── test_analyze.py      # unittest
├── requirements.txt     # 依存ライブラリ（javalang のみ）
├── setup.sh             # venv作成（Unix/Mac）
├── setup.bat            # venv作成（Windows）
├── run.sh               # 実行ラッパー（Unix/Mac）
├── run.bat              # 実行ラッパー（Windows）
└── README.txt           # 利用者向け手順書
```

**パッケージ分割パターン（大規模化時）**:
```
grep_analyzer/
├── src/
│   ├── __init__.py
│   ├── parser.py        # grep結果パース
│   ├── classifier.py    # 使用タイプ分類
│   ├── tracker.py       # 間接参照追跡
│   ├── writer.py        # TSV出力
│   └── models.py        # データモデル
├── tests/
│   ├── test_parser.py
│   ├── test_classifier.py
│   └── test_tracker.py
├── analyze.py           # エントリーポイント
└── requirements.txt
```

**理由**:
- 責務ごとに責任が明確
- 後からのリファクタリングが容易
- チーム開発で統一しやすい

### モジュール分離のタイミング

**分離を検討する兆候**:
1. ファイルの行数が300行以上
2. 関連する機能がまとまっている
3. 独立してテスト可能
4. 他の機能への依存が少ない

**分離の手順**:
```python
# Before: 全処理がanalyze.pyに集中
analyze.py  # 1000行超

# After: 責務ごとに分割
analyze.py          # エントリーポイント (50行)
src/
├── parser.py       # パース処理 (150行)
├── classifier.py   # 分類処理 (200行)
└── tracker.py      # 追跡処理 (300行)
```

## 特殊なケースの対応

### 共有コードの配置

**shared/ または utils/ モジュール**
```
src/
├── shared/
│   ├── ast_utils.py      # AST解析ユーティリティ
│   └── file_utils.py     # ファイル操作ユーティリティ
├── classifier.py
├── tracker.py
└── writer.py
```

**ルール**:
- 本当に複数のモジュールで使われるもののみ
- 単一モジュールでしか使わないものは含めない

### 設定ファイルの管理

```
project/
├── .venv/                    # Python仮想環境（gitignore対象）
├── requirements.txt          # 依存ライブラリ
└── setup.sh                  # venv作成スクリプト
```

### スクリプトの管理

```
scripts/
├── setup.sh                  # venv作成・初期化
├── setup.bat                 # Windows用
├── run.sh                    # 実行ラッパー（Unix/Mac）
└── run.bat                   # 実行ラッパー（Windows）
```

## ドキュメント配置

### ドキュメントの種類と配置先

**プロジェクトルート**:
- `README.txt`: 利用者向け手順書（zip配布用）
- `README.md`: 開発者向けドキュメント

**docs/ ディレクトリ**:
- `product-requirements.md`: PRD
- `functional-design.md`: 機能設計書
- `architecture.md`: アーキテクチャ設計書
- `repository-structure.md`: 本ドキュメント
- `development-guidelines.md`: 開発ガイドライン
- `glossary.md`: 用語集

**ソースコード内**:
- docstringコメント: 関数・クラスの説明

## チェックリスト

- [ ] 各モジュールの役割が明確に定義されている
- [ ] レイヤー構造がディレクトリに反映されている
- [ ] 命名規則が一貫している（snake_case）
- [ ] テストコードの配置方針が決まっている
- [ ] 依存関係のルールが明確である
- [ ] 循環依存がない
- [ ] スケーリング戦略が考慮されている
- [ ] 共有コードの配置ルールが定義されている
- [ ] 設定ファイルの管理方法が決まっている
- [ ] ドキュメントの配置場所が明確である
