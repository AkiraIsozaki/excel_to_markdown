# 設計書

## アーキテクチャ概要

既存の変換パイプライン（CLI → Reader → Parser → Renderer）をそのまま活用し、その上に FastAPI + 静的HTML の Web UIレイヤーを追加する。変換ロジックへの変更は最小限に留め、`web/` パッケージとして独立させる。

```
┌──────────────────────────────────────────────────────────────┐
│  ブラウザ（HTML5 + Vanilla JS）                               │
│  - D&Dエリア、ファイル選択                                     │
│  - fetch() API で POST /api/convert                          │
│  - レスポンス（.md / .zip）を自動ダウンロード                   │
└────────────────────┬─────────────────────────────────────────┘
                     │ HTTP multipart/form-data
┌────────────────────▼─────────────────────────────────────────┐
│  Web UIレイヤー（excel_to_markdown/web/）                     │
│  - FastAPI app（app.py）                                      │
│  - POST /api/convert: UploadFile → Markdown / ZIP            │
│  - GET /health: ヘルスチェック                                 │
│  - 静的ファイル配信（/static/ → web/static/）                  │
└────────────────────┬─────────────────────────────────────────┘
                     │ ファイルバイト列をtempfileに書き出し
┌────────────────────▼─────────────────────────────────────────┐
│  既存変換パイプライン（cli.py の run_file() ヘルパー）           │
│  Reader → Parser → Renderer → Markdown文字列                  │
└──────────────────────────────────────────────────────────────┘
```

## コンポーネント設計

### 1. excel_to_markdown/web/app.py

**責務**:
- FastAPIアプリケーションのファクトリ・エンドポイント定義
- アップロードされたファイルをtempfileに保存し、変換パイプラインを呼び出す
- 単一ファイルは `text/markdown`、複数ファイルは `application/zip` で返す

**実装の要点**:
- `create_app() -> FastAPI` ファクトリ関数でアプリを生成（テスト容易性のため）
- `UploadFile` を受け取り、`tempfile.NamedTemporaryFile` に書き出してから既存パイプラインを呼び出す
- 変換完了後は一時ファイルを削除（`finally` ブロック）
- ファイルサイズ上限: 50MB（FastAPIの `max_upload_size` もしくはバリデーションで実装）
- 拡張子バリデーション: `.xlsx` / `.xls` 以外は 400 エラー
- ZIPはメモリ上（`io.BytesIO`）で生成してストリームレスポンスで返す

```python
@app.post("/api/convert")
async def convert(files: list[UploadFile] = File(...)) -> Response:
    """
    単一ファイル → Content-Type: text/markdown; charset=utf-8
    複数ファイル → Content-Type: application/zip
    """
    ...

@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}
```

### 2. excel_to_markdown/web/static/index.html

**責務**:
- D&Dエリアとファイル選択UIの提供
- `fetch()` APIで `/api/convert` を呼び出し、レスポンスをダウンロード

**実装の要点**:
- Vanilla JS のみ（外部CDN不要。オフライン環境でも動作）
- D&D は HTML5 Drag and Drop API（`dragover`, `drop` イベント）
- `<input type="file" multiple accept=".xlsx,.xls">` でファイル選択も可能
- 変換中は送信ボタン・D&Dエリアを無効化してスピナー表示
- ダウンロードは `URL.createObjectURL()` + `<a>` の `.click()` で実現
- エラーは画面内にインラインで表示（アラートダイアログは使わない）

### 3. excel_to_markdown/web/__main__.py（または cli.py への統合）

**責務**:
- `python -m excel_to_markdown serve` サブコマンドの実装
- uvicorn を起動し、オプションでブラウザを自動オープン

**実装の要点**:
- `cli.py` に `serve` サブコマンドを追加（`subparsers.add_parser("serve")`）
- `--port`（デフォルト8000）、`--no-browser` オプション
- `uvicorn.run("excel_to_markdown.web.app:create_app", factory=True, ...)` で起動
- `webbrowser.open(f"http://localhost:{port}")` でブラウザを自動オープン

### 4. cli.py の変換ロジック切り出し（リファクタリング）

**責務**:
- `run()` 内の「1ファイルを変換してMarkdown文字列を返す」ロジックを `run_file()` ヘルパーとして抽出
- `web/app.py` からも `cli.run_file()` を呼び出せるようにする

**実装の要点**:
```python
def run_file(input_path: Path, base_font_size: float = 11.0) -> str:
    """1つのExcelファイルをMarkdown文字列に変換して返す。"""
    ...
```

## データフロー

### 単一ファイル変換フロー

```
1. ブラウザ: ファイルをD&Dまたは選択
2. ブラウザ: FormData に file を append して POST /api/convert
3. FastAPI: UploadFile を受け取る
4. FastAPI: tempfile に保存
5. FastAPI: cli.run_file(temp_path) を呼び出す
6. FastAPI: Markdown文字列を Response(content=md, media_type="text/markdown") で返す
7. ブラウザ: Blob を生成し <a> クリックで .md ファイルをダウンロード
8. FastAPI: tempfile を削除（finally）
```

### 複数ファイル変換フロー

```
1. ブラウザ: 複数ファイルをD&D
2. ブラウザ: FormData に複数 file を append して POST /api/convert
3. FastAPI: 各UploadFileをtempfileに保存
4. FastAPI: 各ファイルに cli.run_file() を呼び出す（エラーは記録して継続）
5. FastAPI: ZIPをio.BytesIOで構築（成功分のみ含める）
6. FastAPI: Response(content=zip_bytes, media_type="application/zip") で返す
7. ブラウザ: .zip ファイルをダウンロード
8. FastAPI: 全tempfileを削除（finally）
```

## エラーハンドリング戦略

### HTTPエラーレスポンス形式

```json
{"detail": "エラーメッセージ"}
```

| エラー種別 | HTTPステータス | メッセージ |
|-----------|--------------|-----------|
| 非対応拡張子 | 400 Bad Request | `対応していないファイル形式です: .csv（.xlsx/.xls のみ対応）` |
| ファイルサイズ超過 | 413 Request Entity Too Large | `ファイルサイズが上限（50MB）を超えています` |
| 変換失敗（全ファイル） | 422 Unprocessable Entity | `すべてのファイルの変換に失敗しました` |
| サーバー内部エラー | 500 Internal Server Error | `変換中に予期しないエラーが発生しました` |

### 部分的エラー（複数ファイル）

- 一部ファイルの変換が失敗しても、成功分はZIPに含めて返す
- レスポンスヘッダー `X-Conversion-Errors` に失敗ファイル名をカンマ区切りで通知

## テスト戦略

### ユニットテスト

- `test_web_app.py`: FastAPIのTestClientを使ったAPIテスト
  - 単一ファイルアップロード → text/markdown レスポンス
  - 複数ファイルアップロード → application/zip レスポンス
  - 非対応拡張子 → 400
  - ヘルスチェック → 200 `{"status": "ok"}`
  - run_file() の単体テスト

### 統合テスト

- 既存の統合テスト（test_integration.py）が引き続きパスすること
- cli.py の既存テストが引き続きパスすること

## 依存ライブラリ

```toml
[project.optional-dependencies]
web = [
    "fastapi>=0.110.0,<1.0.0",
    "uvicorn[standard]>=0.29.0,<1.0.0",
    "python-multipart>=0.0.9",   # FastAPI のファイルアップロードに必要
]
```

`pip install -e ".[web]"` でインストール。

## ディレクトリ構造

```
excel_to_markdown/
├── web/                          # 新規: Webレイヤー
│   ├── __init__.py
│   ├── app.py                   # FastAPIアプリ・エンドポイント
│   └── static/
│       └── index.html           # D&D UI（Vanilla JS）
│
├── cli.py                       # 変更: serve サブコマンド追加、run_file() 切り出し
│
tests/
├── test_web_app.py              # 新規: Web APIテスト
```

## 実装の順序

1. `cli.py` に `run_file()` ヘルパーを切り出し（既存 `run()` のリファクタリング）
2. `web/__init__.py`, `web/app.py` を実装（FastAPI + エンドポイント）
3. `web/static/index.html` を実装（D&D UI）
4. `cli.py` に `serve` サブコマンドを追加（uvicorn起動）
5. `tests/test_web_app.py` を実装
6. `pyproject.toml` に `[web]` optional dependency を追加
7. ドキュメント更新（永続ドキュメントはステアリングフェーズ1で実施済み）

## セキュリティ考慮事項

- アップロードファイルはtempfileに保存し、変換後即削除（サーバーに残さない）
- ファイルサイズ上限 50MB でDoS緩和
- 拡張子バリデーションで非Excelファイルを早期拒否
- ローカル起動前提のため認証は不要（bindアドレスは `127.0.0.1` 固定）
- パストラバーサル: tempfileを使うためユーザー指定パスは使わない

## パフォーマンス考慮事項

- 複数ファイルは並列変換を検討（asyncio.gather + run_in_executor）
  - ただし変換パイプラインはCPUバウンドのため、ThreadPoolExecutor を使用
  - 初期実装は直列でも可（性能要件が問題になれば並列化）
- メモリ: ZIPはio.BytesIOでメモリ上に生成。大量ファイルの場合は要注意だが、現スコープは「複数ファイル程度」の想定

## 将来の拡張性

- 変換オプション（--base-font-size等）のUIサポート: フォームに追加可能な設計
- WebSocket進捗通知: FastAPIはWebSocket対応済みのため追加容易
- Docker化: `uvicorn` をそのままコンテナで動かせる
