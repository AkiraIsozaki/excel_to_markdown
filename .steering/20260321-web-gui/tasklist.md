# タスクリスト: Web GUI機能追加（FastAPI + HTML5 D&D）

## 🚨 タスク完全完了の原則

**このファイルの全タスクが完了するまで作業を継続すること**

---

## フェーズ1: ドキュメント更新（永続ドキュメント）

- [x] `docs/product-requirements.md` を更新
  - [x] 「スコープ外」から「GUIインターフェース」を削除
  - [x] 機能要件13: Webブラウザ型UI（D&D一括変換）を追加
  - [x] 非機能要件: ポート設定・ファイルサイズ上限を追記
- [x] `docs/architecture.md` を更新
  - [x] テクノロジースタック: FastAPI / uvicorn / python-multipart を追加
  - [x] アーキテクチャパターン: Webサーバーレイヤーをパイプライン図に追加
  - [x] セキュリティ: ファイルアップロードのバリデーション方針を追記
  - [x] 依存関係管理: `[web]` optional dependency を追記
- [x] `docs/functional-design.md` を更新
  - [x] システム構成図: ブラウザ → Web UIレイヤーを追加
  - [x] コンポーネント設計: web/app.py, web/static/index.html を追加
  - [x] CLIインターフェース設計: `serve` サブコマンドを追記
  - [x] エラーハンドリング: HTTPエラー種別を追記
- [x] `docs/repository-structure.md` を更新
  - [x] プロジェクト構造: `excel_to_markdown/web/` ディレクトリを追加
  - [x] `tests/test_web_app.py` を追加
  - [x] `pyproject.toml` の `[web]` optional dependency を追記

## フェーズ2: 実装

- [x] `cli.py` に `run_file()` ヘルパーを切り出し
  - [x] `run_file(input_path: Path, base_font_size: float) -> str` を定義
  - [x] 既存 `run()` から変換ロジックをリファクタリング
  - [x] 既存テストが引き続きパスすることを確認
- [x] `excel_to_markdown/web/__init__.py` を作成
- [x] `excel_to_markdown/web/app.py` を実装
  - [x] `create_app() -> FastAPI` ファクトリ関数
  - [x] `GET /health` エンドポイント
  - [x] `POST /api/convert` エンドポイント（単一・複数ファイル対応）
  - [x] 静的ファイル配信設定（`/static/` → `web/static/`）
- [x] `excel_to_markdown/web/static/index.html` を実装
  - [x] D&Dエリア（dragover/drop イベント）
  - [x] クリックでファイル選択ダイアログ
  - [x] 変換中スピナー表示
  - [x] ダウンロード処理（createObjectURL）
  - [x] エラー表示
- [x] `cli.py` に `serve` サブコマンドを追加
  - [x] `--port`（デフォルト8000）オプション
  - [x] `--no-browser` オプション
  - [x] uvicorn.run() 呼び出し
  - [x] webbrowser.open() 呼び出し
- [x] `pyproject.toml` に `[web]` optional dependency を追加

## フェーズ3: テスト

- [x] `tests/test_web_app.py` を実装
  - [x] `GET /health` → 200 `{"status": "ok"}`
  - [x] `POST /api/convert` 単一.xlsx → text/markdown レスポンス
  - [x] `POST /api/convert` 複数.xlsx → application/zip レスポンス
  - [x] `POST /api/convert` 非対応拡張子 → 400エラー
  - [x] `run_file()` の単体テスト

## フェーズ4: 品質チェック

- [x] `pytest` 実行 → 全テストパス確認（212 passed）
- [x] `ruff check excel_to_markdown/web/` → lint通過
- [x] `mypy excel_to_markdown/web/` → 型チェック通過
- [x] カバレッジ80%以上を確認（全体89%、web/app.py 88%）
- [x] `httpx` を `dev` optional dependency に追加（TestClient に必要）

---

## 実装後の振り返り

### 実装完了日
2026-03-21

### 計画と実績の差分

**計画と異なった点**:
- `parse_args()` に argparse の subparsers を使ったところ、既存テストが `parse_args(['input.xlsx'])` のように直接呼ぶため、サブコマンドの positional チェックで衝突した。サブコマンドなしの argv を detect して別パーサーに委譲する方式（`argv[0] == "serve"` チェック）に切り替えた。
- FastAPI の TestClient は `httpx` を要求するが、`[web]` optional dependency には含めていなかった。`[dev]` に `httpx>=0.23.0` を追加した。

**新たに必要になったタスク**:
- `pyproject.toml` の `[dev]` に `httpx` を追加
- cli.py の既存 ruff 指摘（f-string without placeholder、line too long）を修正

### 学んだこと

- argparse の subparsers と後方互換性の両立: subparsers を使うと既存の positional 引数テストが破綻する。第1引数のサブコマンド名チェックで分岐し、それぞれ別のパーサーに委譲するのがシンプルで安全。
- FastAPI TestClient に httpx が必要: `pip install httpx` が別途必要。`[dev]` optional dependency に明示するべき。
- `UploadFile` の読み込みはコルーチン: `await upload.read()` で全バイト取得し、tempfile に write_bytes するパターンが安全。

### 次回への改善提案
- `serve` コマンドのテスト: uvicorn 起動は副作用が大きくテストしにくい。`create_app()` のファクトリ分離でAPIテストは可能だが、起動フローのテストは別途検討。
- 変換オプション（`--base-font-size`）をUI上で設定できるようにする（現在はデフォルト値固定）。
- 大量ファイルの並列変換: 現状は直列処理。ThreadPoolExecutor + asyncio.gather で並列化可能。
