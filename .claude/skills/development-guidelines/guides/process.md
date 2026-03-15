# プロセスガイド (Process Guide)

## 基本原則

### 1. 具体例を豊富に含める

抽象的なルールだけでなく、具体的なコード例を提示します。

**悪い例**:
```
変数名は分かりやすくすること
```

**良い例**:
```python
# ✅ 良い例: 役割が明確
grep_records: list[dict] = []
ast_cache: dict[str, object] = {}

# ❌ 悪い例: 曖昧
data: list = []
cache: dict = {}
```

### 2. 理由を説明する

「なぜそうするのか」を明確にします。

**例**:
```
## エラーを無視しない

理由: エラーを無視すると、問題の原因究明が困難になります。
予期されるエラーは適切に処理し、予期しないエラーは上位に伝播させて
ログに記録できるようにします。
```

### 3. 測定可能な基準を設定

曖昧な表現を避け、具体的な数値を示します。

**悪い例**:
```
コードカバレッジは高く保つこと
```

**良い例**:
```
コードカバレッジ目標:
- ユニットテスト: 80%以上
- 統合テスト: 60%以上
- E2Eテスト: 主要フロー100%
```

## Git運用ルール

### ブランチ戦略（Git Flow採用）

**Git Flowとは**:
Vincent Driessenが提唱した、機能開発・リリース・ホットフィックスを体系的に管理するブランチモデル。明確な役割分担により、チーム開発での並行作業と安定したリリースを実現します。

**ブランチ構成**:
```
main (本番環境)
└── develop (開発・統合環境)
    ├── feature/* (新機能開発)
    ├── fix/* (バグ修正)
    └── release/* (リリース準備)※必要に応じて
```

**運用ルール**:
- **main**: 本番リリース済みの安定版コードのみを保持。タグでバージョン管理
- **develop**: 次期リリースに向けた最新の開発コードを統合。CIでの自動テスト実施
- **feature/\*、fix/\***: developから分岐し、作業完了後にPRでdevelopへマージ
- **直接コミット禁止**: すべてのブランチでPRレビューを必須とし、コード品質を担保
- **マージ方針**: feature→develop は squash merge、develop→main は merge commit を推奨

**Git Flowのメリット**:
- ブランチの役割が明確で、複数人での並行開発がしやすい
- 本番環境(main)が常にクリーンな状態に保たれる
- 緊急対応時はhotfixブランチで迅速に対応可能（必要に応じて導入）

### コミットメッセージの規約

**Conventional Commitsを推奨**:

```
<type>(<scope>): <subject>

<body>

<footer>
```

**Type一覧**:
```
feat: 新機能 (minor version up)
fix: バグ修正 (patch version up)
docs: ドキュメント
style: フォーマット (コードの動作に影響なし)
refactor: リファクタリング
perf: パフォーマンス改善
test: テスト追加・修正
build: ビルドシステム
ci: CI/CD設定
chore: その他 (依存関係更新など)

BREAKING CHANGE: 破壊的変更 (major version up)
```

**良いコミットメッセージの例**:

```
feat(tracker): getter経由の間接参照追跡機能を追加

フィールドに代入された値がgetter経由で使われる箇所を追跡する。

実装内容:
- 命名規則（type → getType()）によるgetter候補特定
- return文解析による非標準命名のgetter検出
- プロジェクト全体でのgetter呼び出し箇所をAST解析

Closes #12
```

### プルリクエストのテンプレート

**効果的なPRテンプレート**:

```markdown
## 変更の種類
- [ ] 新機能 (feat)
- [ ] バグ修正 (fix)
- [ ] リファクタリング (refactor)
- [ ] ドキュメント (docs)
- [ ] その他 (chore)

## 変更内容
### 何を変更したか
[簡潔な説明]

### なぜ変更したか
[背景・理由]

### どのように変更したか
- [変更点1]
- [変更点2]

## テスト
### 実施したテスト
- [ ] ユニットテスト追加
- [ ] 統合テスト追加
- [ ] 手動テスト実施

### テスト結果
[テスト結果の説明]

## 関連Issue
Closes #[番号]
Refs #[番号]

## レビューポイント
[レビュアーに特に見てほしい点]
```

## テスト戦略

### テストピラミッド

```
       /\
      /E2E\       少 (遅い、高コスト)
     /------\
    / 統合   \     中
   /----------\
  / ユニット   \   多 (速い、低コスト)
 /--------------\
```

**目標比率**:
- ユニットテスト: 70%
- 統合テスト: 20%
- E2Eテスト: 10%

### テストの書き方

**Arrange-Act-Assert パターン**:

```python
import unittest

class TestGrepParser(unittest.TestCase):

    def test_parse_valid_grep_line_returns_record(self):
        # Arrange: 準備
        parser = GrepParser()
        line = "src/main/Constants.java:10:    public static final String CODE = \"VALUE\";"

        # Act: 実行
        result = parser.parse_line(line)

        # Assert: 検証
        self.assertIsNotNone(result)
        self.assertEqual(result.filepath, "src/main/Constants.java")
        self.assertEqual(result.lineno, "10")

    def test_parse_binary_notice_line_returns_none(self):
        # Arrange: 準備
        parser = GrepParser()
        line = "Binary file src/main/resources/logo.png matches"

        # Act: 実行
        result = parser.parse_line(line)

        # Assert: 検証
        self.assertIsNone(result)
```

### カバレッジ目標

**測定可能な目標**:

```bash
# coverage.py を使用してカバレッジを測定
pip install coverage
coverage run -m unittest discover
coverage report --fail-under=80
coverage html  # HTMLレポート生成
```

**設定例（setup.cfg）**:
```ini
[coverage:run]
source = src
omit =
    tests/*
    .venv/*

[coverage:report]
fail_under = 80
show_missing = True
```

**理由**:
- 重要なビジネスロジック（classifier.py, tracker.py）は高いカバレッジを要求
- エントリーポイント（analyze.py）は低めでも許容
- 100%を目指さない (コストと効果のバランス)

## コードレビュープロセス

### レビューの目的

1. **品質保証**: バグの早期発見
2. **知識共有**: チーム全体でコードベースを理解
3. **学習機会**: ベストプラクティスの共有

### 効果的なレビューのポイント

**レビュアー向け**:

1. **建設的なフィードバック**
```markdown
## ❌ 悪い例
このコードはダメです。

## ✅ 良い例
この実装だと O(n²) の時間計算量になります。
dict を使うと O(1) に改善できます:

```python
ast_cache: dict[str, object] = {}
tree = ast_cache.get(filepath)
if tree is None:
    tree = parse_java_file(filepath)
    ast_cache[filepath] = tree
```
```

2. **優先度の明示**
```markdown
[必須] セキュリティ: ファイルパスのトラバーサル対策が必要です
[推奨] パフォーマンス: 同一ファイルのAST再解析を避けましょう
[提案] 可読性: この関数名をもっと明確にできませんか？
[質問] この処理の意図を教えてください
```

3. **ポジティブなフィードバックも**
```markdown
✨ この実装は分かりやすいですね！
👍 エッジケースがしっかり考慮されています
💡 このパターンは他でも使えそうです
```

**レビュイー向け**:

1. **セルフレビューを実施**
   - PR作成前に自分でコードを見直す
   - 説明が必要な箇所にコメントを追加

2. **小さなPRを心がける**
   - 1PR = 1機能
   - 変更ファイル数: 10ファイル以内を推奨
   - 変更行数: 300行以内を推奨

3. **説明を丁寧に**
   - なぜこの実装にしたか
   - 検討した代替案
   - 特に見てほしいポイント

### レビュー時間の目安

- 小規模PR (100行以下): 15分
- 中規模PR (100-300行): 30分
- 大規模PR (300行以上): 1時間以上

**原則**: 大規模PRは避け、分割する

## 自動化の推進（該当する場合）

### 品質チェックの自動化

**自動化項目と採用ツール**:

1. **静的解析**
   - **flake8**
     - PEP 8準拠チェック・未使用インポート検出
     - 設定ファイル: `.flake8` または `setup.cfg`
   - **mypy** (任意)
     - 型アノテーションの静的型チェック
     - `mypy analyze.py` で実行

2. **コードフォーマット**
   - **black** (任意)
     - Pythonコードの自動整形
     - `black .` で適用
     - レビュー時のスタイル議論を削減

3. **構文チェック**
   - **py_compile**
     - Python標準の構文チェック
     - `python -m py_compile analyze.py`

4. **テスト実行**
   - **unittest**
     - Python標準のテスティングフレームワーク
     - `python -m unittest discover` で自動検出・実行
   - **coverage** (任意)
     - `coverage run -m unittest discover` でカバレッジ測定

5. **ビルド確認（パッケージング）**
   - **Makefile**
     - `make package` でzipファイルを生成
     - `make test` でテスト実行

**実装方法**:

**1. CI/CD (GitHub Actions)**
```yaml
# .github/workflows/ci.yml
name: CI
on: [push, pull_request]
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.11'
      - run: pip install -r requirements.txt
      - run: python -m unittest discover
      - run: python -m flake8 .
```

**2. Pre-commit フック**
```bash
# .git/hooks/pre-commit
#!/bin/bash
set -e
python -m py_compile analyze.py
python -m unittest discover
python -m flake8 .
```

```bash
chmod +x .git/hooks/pre-commit
```

**導入効果**:
- コミット前に自動チェックが走り、不具合コードの混入を防止
- PR作成時に自動でCI実行され、マージ前に品質を担保
- 早期発見により、修正コストを最大80%削減（バグ検出が本番後の場合と比較）

**この構成を選んだ理由**:
- Python標準ライブラリ(unittest)を中心とした軽量な構成
- flake8とcoverageの組み合わせでテストと品質チェックがシームレスに連携
- 外部依存を最小化（javalangのみ）

## チェックリスト

- [ ] ブランチ戦略が決まっている
- [ ] コミットメッセージ規約が明確である
- [ ] PRテンプレートが用意されている
- [ ] テストの種類とカバレッジ目標が設定されている
- [ ] コードレビュープロセスが定義されている
- [ ] CI/CDパイプラインが構築されている
