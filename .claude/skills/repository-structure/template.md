# リポジトリ構造定義書 (Repository Structure Document)

## プロジェクト構造

```
project-root/
├── analyze.py           # エントリーポイント（またはsrc/配下に分割）
├── src/                 # ソースコード（オプション、小規模なら不要）
│   ├── __init__.py
│   ├── [module1].py     # [説明]
│   ├── [module2].py     # [説明]
│   └── [module3].py     # [説明]
├── tests/               # テストコード
│   ├── unit/            # ユニットテスト
│   └── integration/     # 統合テスト
├── docs/                # プロジェクトドキュメント
├── input/               # 入力ファイル配置ディレクトリ（.gitkeep）
├── output/              # 出力ファイル配置ディレクトリ（.gitkeep）
├── requirements.txt     # 依存ライブラリ
├── setup.sh             # venv作成スクリプト（Unix/Mac）
├── setup.bat            # venv作成スクリプト（Windows）
├── run.sh               # 実行ラッパー（Unix/Mac）
├── run.bat              # 実行ラッパー（Windows）
└── README.txt           # 利用者向け手順書
```

## ディレクトリ詳細

### src/ (ソースコードディレクトリ)

#### [モジュール1]

**役割**: [説明]

**配置ファイル**:
- [ファイルパターン1]: [説明]
- [ファイルパターン2]: [説明]

**命名規則**:
- [規則1]
- [規則2]

**依存関係**:
- 依存可能: [モジュール名]
- 依存禁止: [モジュール名]

**例**:
```
src/
├── [module1].py
└── [module2].py
```

#### [モジュール2]

**役割**: [説明]

**配置ファイル**:
- [ファイルパターン1]: [説明]

**命名規則**:
- [規則1]

**依存関係**:
- 依存可能: [モジュール名]
- 依存禁止: [モジュール名]

### tests/ (テストディレクトリ)

#### unit/

**役割**: ユニットテストの配置

**構造**:
```
tests/unit/
└── test_[module_name].py
```

**命名規則**:
- パターン: `test_[テスト対象モジュール名].py`
- 例: `parser.py` → `test_parser.py`

#### integration/

**役割**: 統合テストの配置

**構造**:
```
tests/integration/
└── test_[feature].py
```

### docs/ (ドキュメントディレクトリ)

**配置ドキュメント**:
- `product-requirements.md`: プロダクト要求定義書
- `functional-design.md`: 機能設計書
- `architecture.md`: アーキテクチャ設計書
- `repository-structure.md`: リポジトリ構造定義書(本ドキュメント)
- `development-guidelines.md`: 開発ガイドライン
- `glossary.md`: 用語集

### input/ / output/ (データディレクトリ)

**配置ファイル**:
- `input/`: ユーザーが入力ファイルを配置する
- `output/`: ツールが出力ファイルを生成する

**例**:
```
input/
├── .gitkeep
└── TARGET.grep      # ユーザーが配置するgrep結果ファイル

output/
├── .gitkeep
└── TARGET.tsv       # ツールが生成するTSVファイル
```

### config/ (設定ファイルディレクトリ - 該当する場合)

**配置ファイル**:
- 静的解析設定ファイル（flake8、mypy等）
- その他ツール設定ファイル

**例**:
```
config/
└── .flake8          # flake8設定
```

### scripts/ (スクリプトディレクトリ - 該当する場合)

**配置ファイル**:
- セットアップスクリプト
- 開発補助スクリプト

## ファイル配置規則

### ソースファイル

| ファイル種別 | 配置先 | 命名規則 | 例 |
|------------|--------|---------|-----|
| [種別1] | [ディレクトリ] | [規則] | [例] |
| [種別2] | [ディレクトリ] | [規則] | [例] |

### テストファイル

| テスト種別 | 配置先 | 命名規則 | 例 |
|-----------|--------|---------|-----|
| ユニットテスト | tests/unit/ | test_[対象].py | test_parser.py |
| 統合テスト | tests/integration/ | test_[機能].py | test_analyze.py |

### 設定ファイル

| ファイル種別 | 配置先 | 命名規則 |
|------------|--------|---------|
| 依存ライブラリ | プロジェクトルート | requirements.txt |
| Python仮想環境 | プロジェクトルート | .venv/ |
| コード品質 | プロジェクトルート | .flake8 / setup.cfg |

## 命名規則

### ディレクトリ名

- **パッケージ/モジュールディレクトリ**: snake_case
  - 例: `grep_parser/`, `usage_classifier/`
- **テストディレクトリ**: `tests/`（複数形）

### ファイル名

- **モジュールファイル**: snake_case
  - 例: `grep_parser.py`, `usage_classifier.py`
- **データクラス/Enum**: PascalCase（クラス名）
  - 例: `class GrepRecord`, `class UsageType`

### テストファイル名

- パターン: `test_[テスト対象モジュール名].py`
- 例: `test_grep_parser.py`, `test_usage_classifier.py`

## 依存関係のルール

### レイヤー間の依存

```
エントリーポイント (analyze.py)
    ↓ (OK)
処理レイヤー (src/parser.py, src/classifier.py)
    ↓ (OK)
モデル/ユーティリティ (src/models.py, src/utils.py)
```

**禁止される依存**:
- モデル → 処理レイヤー
- 処理レイヤー → エントリーポイント

### モジュール間の依存

**循環依存の禁止**:
```python
# ❌ 悪い例: 循環依存
# src/classifier.py
from src.tracker import IndirectTracker

# src/tracker.py
from src.classifier import classify_usage  # 循環依存
```

**解決策**:
```python
# ✅ 良い例: 共通モデルの抽出
# src/models.py
from dataclasses import dataclass

@dataclass(frozen=True)
class GrepRecord: ...

# src/classifier.py
from src.models import GrepRecord

# src/tracker.py
from src.models import GrepRecord
```

## スケーリング戦略

### 機能の追加

新しい機能を追加する際の配置方針:

1. **小規模機能**: 既存モジュールに関数を追加
2. **中規模機能**: 新しいモジュールファイルを作成
3. **大規模機能**: サブパッケージとして分離

**例**:
```
src/
├── parser.py               # 既存機能
└── tracker/                # 大規模機能の分離
    ├── __init__.py
    ├── constant_tracker.py
    ├── field_tracker.py
    └── getter_tracker.py
```

### ファイルサイズの管理

**ファイル分割の目安**:
- 1ファイル: 300行以下を推奨
- 300-500行: リファクタリングを検討
- 500行以上: 分割を強く推奨

**分割方法**:
```python
# 悪い例: 1ファイルに全機能
# analyze.py (1000行)

# 良い例: 責務ごとに分割
# analyze.py (100行) - エントリーポイント・CLIパース
# src/parser.py (200行) - grep結果パース
# src/classifier.py (300行) - 使用タイプ分類
# src/tracker.py (400行) - 間接参照追跡
```

## 特殊ディレクトリ

### .steering/ (ステアリングファイル)

**役割**: 特定の開発作業における「今回何をするか」を定義

**構造**:
```
.steering/
└── [YYYYMMDD]-[task-name]/
    ├── requirements.md      # 今回の作業の要求内容
    ├── design.md            # 変更内容の設計
    └── tasklist.md          # タスクリスト
```

**命名規則**: `20250115-add-user-profile` 形式

### .claude/ (Claude Code設定)

**役割**: Claude Code設定とカスタマイズ

**構造**:
```
.claude/
├── commands/                # スラッシュコマンド
├── skills/                  # タスクモード別スキル
└── agents/                  # サブエージェント定義
```

## 除外設定

### .gitignore

プロジェクトで除外すべきファイル:
- `.venv/`
- `__pycache__/`
- `*.pyc`
- `*.pyo`
- `.env`
- `.steering/` (タスク管理用の一時ファイル)
- `*.log`
- `.DS_Store`
- `dist/`
- `*.egg-info/`

### コード品質ツール

ツールで除外すべきディレクトリ:
- `.venv/`
- `__pycache__/`
- `.steering/`
