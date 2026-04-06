# 要求内容

## 概要

オフライン環境（Ubuntu系Linux）でこのプロジェクトをインストール・実行できるよう、
必要なPythonパッケージのwheelファイルを事前ダウンロードし同梱する。

## 背景

配布先がインターネット接続のないオフライン環境のため、`pip install` 時にPyPIへのアクセスが不要な状態にしたい。
配布先にはPythonが入っている前提。OSはUbuntu系Linux。

## 実装対象の機能

### 1. wheelhouse ディレクトリ

- 依存パッケージ（本体 + オプション含む）のwheelファイルを `wheelhouse/` に格納
- Linux/Ubuntu向けのwheelを取得する

### 2. インストールスクリプト

- `install.sh` を作成し、`pip install --no-index --find-links=./wheelhouse` を使ってオフラインインストールできるようにする

### 3. README / 手順書

- オフラインインストール手順をREADME.mdに追記する

## 受け入れ条件

- [ ] `wheelhouse/` に openpyxl およびその依存関係のwheelが含まれている
- [ ] `wheelhouse/` に fastapi / uvicorn / python-multipart のwheelが含まれている（Web UIオプション）
- [ ] `wheelhouse/` に xlrd のwheelが含まれている（xlsオプション）
- [ ] `wheelhouse/` に hatchling（ビルドバックエンド）のwheelが含まれている
- [ ] `install.sh` を実行するだけでオフラインインストールが完了する
- [ ] README.mdにオフラインインストール手順が記載されている

## スコープ外

- Windows / macOS 向けのwheel配布
- Python自体のオフライン配布
- dev依存（pytest, ruff, mypy）のwheel同梱
