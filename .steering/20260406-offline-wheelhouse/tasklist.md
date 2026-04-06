# タスクリスト

## 🚨 タスク完全完了の原則

全タスクを `[x]` にすること。未完了タスクを残したまま終了しない。

---

## フェーズ1: wheelhouse の作成

- [x] pip download で本体依存（openpyxl）のwheelを取得
- [x] pip download で Web UI オプション（fastapi, uvicorn, python-multipart）のwheelを取得
- [x] pip download で xls オプション（xlrd）のwheelを取得
- [x] pip download で hatchling（ビルドバックエンド）のwheelを取得
- [x] wheelhouse/ に必要なwheelが揃っているか確認

## フェーズ2: インストールスクリプト作成

- [x] install.sh を作成
  - [x] `pip install --no-index --find-links=./wheelhouse .` を実行
  - [x] Web UI / xls オプションの案内メッセージを表示
  - [x] 実行権限を付与

## フェーズ3: README.md 更新

- [x] README.md にオフラインインストールのセクションを追記
  - [x] 前提条件（Python 3.12+）
  - [x] install.sh を使ったインストール手順
  - [x] pip コマンドで直接インストールする手順も記載

## フェーズ4: 動作確認

- [x] install.sh が正常に実行できるか確認
- [x] `pip install --no-index --find-links=./wheelhouse .` が成功するか確認

---

## 実装後の振り返り

### 実装完了日
2026-04-06

### 計画と実績の差分

**計画と異なった点**:
- 特になし。pip download / install.sh / README 追記のすべてが計画通りに完了。

### 学んだこと
- `--platform manylinux2014_x86_64` と `linux_x86_64` を併用することでUbuntu系で確実に動くwheelを取得できる
- `-e .`（editable install）でも `--no-index --find-links` は機能する
