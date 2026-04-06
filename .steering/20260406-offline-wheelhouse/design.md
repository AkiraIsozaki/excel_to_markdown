# 設計書

## アプローチ

`pip download` コマンドで依存パッケージのwheelをダウンロードし、`wheelhouse/` に格納する。
インストール時は `pip install --no-index --find-links=./wheelhouse` を使う。

```
[配布パッケージ]
├── wheelhouse/          ← pip download で事前取得したwheelファイル群
├── install.sh           ← オフラインインストールスクリプト
├── pyproject.toml
└── README.md            ← オフライン手順を追記
```

## ダウンロード対象

| グループ | パッケージ |
|---|---|
| 本体 | openpyxl + 依存 |
| Web UIオプション | fastapi, uvicorn[standard], python-multipart |
| xlsオプション | xlrd |
| ビルドバックエンド | hatchling + 依存 (pip install . に必要) |

## pip download のオプション

```bash
pip download \
  --dest ./wheelhouse \
  --platform linux_x86_64 \
  --python-version 312 \
  --only-binary :all: \
  ".[web,xls]"
```

- `--platform linux_x86_64`: Ubuntu/Linux向けのwheelを指定
- `--python-version 312`: Python 3.12向け
- `--only-binary :all:`: wheelのみ（ソース配布物を除外）
- hatchling はビルド時に必要なので別途ダウンロード

## install.sh の動作

1. `pip install --no-index --find-links=./wheelhouse .` で本体インストール
2. オプションでWebUI (`.[web]`) や xls (`.[xls]`) も選択可能なメッセージを表示

## 実装の順序

1. wheelhouse/ へ pip download を実行
2. install.sh を作成
3. README.md にオフライン手順を追記
4. 動作確認（--no-index でインストールできるか）
