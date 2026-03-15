# アーキテクチャ設計ガイド

## 基本原則

### 1. 技術選定には理由を明記

**悪い例**:
```
- Python
- javalang
```

**良い例**:
```
- Python 3.12+ (LTS相当)
  - 標準ライブラリが充実しており、外部依存を最小化できる
  - venvによる軽量な配布形式に適している
  - Windows/Mac/Linux クロスプラットフォーム動作が容易

- javalang
  - Java 7以上のソースコードをAST解析できるPythonライブラリ
  - grep結果のJavaコードを正確に構造解析するために必要
  - パースエラー時は正規表現フォールバックで処理を継続

- venv (Python標準)
  - システムのPython環境を汚染せずに依存関係を管理
  - zip配布形式でsetup.sh一発でセットアップ可能
```

### 2. レイヤー分離の原則

各レイヤーの責務を明確にし、依存関係を一方向に保ちます:

```
入力 → 分析 → 出力 (OK)
入力 ← 分析 (NG)
入力 → 出力 (NG、分析をスキップしている)
```

### 3. 測定可能な要件

すべてのパフォーマンス要件は測定可能な形で記述します。

## レイヤードアーキテクチャの設計

### 各レイヤーの責務

**入力レイヤー**:
```python
# 責務: CLIパース、grep結果ファイルの読み込みとパース
def parse_args() -> argparse.Namespace:
    """CLIオプションを解析する。"""
    ...

def parse_grep_file(path: Path, keyword: str) -> list[GrepRecord]:
    """grep結果ファイルをパースし、GrepRecordのリストを返す。"""
    ...
```

**分析レイヤー**:
```python
# 責務: ASTまたは正規表現による使用タイプ分類
def classify_usage(code: str, filepath: str, lineno: int) -> str:
    """コード行を解析し、使用タイプを返す（7種のいずれか）。"""
    ...

def track_indirect_refs(record: GrepRecord, source_dir: Path) -> list[GrepRecord]:
    """定数/変数の間接参照をプロジェクト全体で追跡する。"""
    ...
```

**出力レイヤー**:
```python
# 責務: TSVへの書き出し、処理レポートの生成
def write_tsv(records: list[GrepRecord], output_path: Path) -> None:
    """GrepRecordのリストをUTF-8 BOM付きTSVに出力する。"""
    ...
```

## パフォーマンス要件の設定

### 具体的な数値目標

```
処理速度: 4万行・500ファイル規模を30分以内を目安とする
└─ 測定方法: time.perf_counter() で処理開始〜完了まで計測
└─ 測定環境: CPU Core i5相当、メモリ8GB

ASTキャッシュ: 同一ファイルの再解析を省略
└─ 実装: ast_cache: dict[str, object] = {} で O(1) アクセス
└─ 効果: 大規模プロジェクト(1000ファイル以上)で大幅な時間短縮

網羅性優先: 処理時間の上限は設けない（数分かかっても許容）
```

## セキュリティ設計

### データ保護の3原則

1. **最小権限の原則**
```bash
# ファイルパーミッション（出力ファイル）
chmod 644 output/TARGET.tsv  # 所有者読み書き、他は読み取りのみ
```

2. **入力検証**
```python
def validate_source_dir(path: Path) -> None:
    if not path.exists():
        raise ValueError(f"ソースディレクトリが存在しません: {path}")
    if not path.is_dir():
        raise ValueError(f"ディレクトリではありません: {path}")
    # パストラバーサル対策
    resolved = path.resolve()
    if not str(resolved).startswith(str(Path.cwd().resolve())):
        # 必要に応じてチェック（ユーザー指定パスなので通常は不要）
        pass
```

3. **機密情報の管理**
```bash
# 環境変数で管理（該当する場合）
export GREP_ANALYZER_KEY="xxxxx"  # コード内にハードコードしない
```

```python
# Pythonコード内での取得
import os
api_key = os.environ.get("GREP_ANALYZER_KEY")
if not api_key:
    raise EnvironmentError("GREP_ANALYZER_KEY環境変数が設定されていません")
```

## スケーラビリティ設計

### データ増加への対応

**想定データ量**: grep結果4万行、Javaソースファイル500件

**対策**:
- ASTキャッシュによる再解析の回避
- ジェネレータによるメモリ効率的な行処理
- ファイルごとのストリーミング処理

```python
# ジェネレータで大量データをメモリ効率良く処理
def iter_grep_lines(path: Path) -> Generator[str, None, None]:
    """grep結果ファイルを1行ずつ読み込むジェネレータ。"""
    with open(path, encoding="utf-8", errors="replace") as f:
        for line in f:
            yield line.rstrip("\n")

# ASTキャッシュで再解析を省略
_ast_cache: dict[str, object] = {}

def get_ast(filepath: str, source_dir: Path) -> object | None:
    """JavaファイルのASTをキャッシュから取得、なければ解析する。"""
    if filepath not in _ast_cache:
        full_path = source_dir / filepath
        try:
            source = full_path.read_text(encoding="shift_jis", errors="replace")
            _ast_cache[filepath] = javalang.parse.parse(source)
        except javalang.parser.JavaSyntaxError:
            _ast_cache[filepath] = None  # フォールバック用
    return _ast_cache[filepath]
```

## 依存関係管理

### バージョン管理方針

```
# requirements.txt
javalang>=0.13.0,<1.0.0  # ASTパース（バグ修正のみ自動適用）
```

**方針**:
- 本番依存はバージョン範囲で管理（マイナーまで固定、パッチは許可）
- テスト依存は標準ライブラリのみ（unittest, coverage は別途インストール）
- `pip freeze > requirements-lock.txt` で再現性を確保
- `pip list --outdated` で定期的に確認

## チェックリスト

- [ ] すべての技術選定に理由が記載されている
- [ ] レイヤードアーキテクチャが明確に定義されている
- [ ] パフォーマンス要件が測定可能である
- [ ] セキュリティ考慮事項が記載されている
- [ ] スケーラビリティが考慮されている
- [ ] 依存関係管理のポリシーが明確である
- [ ] テスト戦略が定義されている
