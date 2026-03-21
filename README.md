# excel-to-markdown

Excel方眼紙（日本のSIer文化で広く使われるグリッド状Excelドキュメント）をMarkdownに変換するツールです。

フォント書式・セル位置・結合情報から見出し・段落・リスト・表を自動認識し、AIが扱えるテキストに変換します。

## 特徴

- **構造の自動認識**: フォントサイズ・太字・列位置から見出し/段落/リスト/表を判定
- **情報を捨てないベストエフォート変換**: 構造が不明なセルも段落として必ず出力し、`<!-- WARNING -->` コメントで通知
- **CLIとWeb UIの両対応**: コマンドライン1発変換と、ブラウザへのD&Dによる一括変換
- **xlsx / xls 両対応**: openpyxl（.xlsx）・xlrd（.xls）で同一パイプラインに通す
- **バッチ変換**: ディレクトリを指定して配下の全ファイルを一括変換
- **ハイパーリンク変換**: セル内のリンクを `[テキスト](URL)` 形式に変換

## インストール

```bash
git clone <このリポジトリ>
cd excel_to_markdown
pip install -e .
```

`.xls`（旧形式）も扱う場合:

```bash
pip install -e ".[xls]"
```

ブラウザ型Web UIを使う場合:

```bash
pip install -e ".[web]"
```

開発環境（テスト・lint・型チェック含む）:

```bash
pip install -e ".[xls,web,dev]"
```

## 使い方

### CLI

```bash
# 基本（入力と同名の .md を出力）
python -m excel_to_markdown input.xlsx

# 出力先を指定
python -m excel_to_markdown input.xlsx -o output.md

# 特定のシートのみ変換（名前またはインデックスで指定）
python -m excel_to_markdown input.xlsx -s "Sheet1"
python -m excel_to_markdown input.xlsx -s 0

# ディレクトリを指定してバッチ変換（配下の xlsx/xls を全変換）
python -m excel_to_markdown ./設計書フォルダ/

# フォントサイズ基準値を変更（見出し判定の閾値を調整）
python -m excel_to_markdown input.xlsx --base-font-size 10.5

# デバッグ出力（TextBlockリストをJSONでstderrに出力）
python -m excel_to_markdown input.xlsx --debug
```

### Web UI（ブラウザでD&D変換）

```bash
# サーバー起動（ブラウザが自動で開きます）
python -m excel_to_markdown serve

# ポートを変更する場合
python -m excel_to_markdown serve --port 9000

# ブラウザを自動で開かない場合
python -m excel_to_markdown serve --no-browser
```

起動後、`http://localhost:8000` にアクセスして .xlsx / .xls ファイルをドラッグ＆ドロップします。

- 単一ファイル → `.md` ファイルをダウンロード
- 複数ファイル → `.zip` ファイルをダウンロード（成功分のみ含む）

## 変換ルール

### 見出し判定（優先順位順）

| 条件 | 出力 |
| --- | --- |
| フォントサイズ 18pt 以上 | `# H1` |
| フォントサイズ 14pt 以上 かつ 太字 | `## H2` |
| フォントサイズ 12pt 以上 かつ 太字 | `### H3` |
| 太字 かつ インデントレベル 0 | `#### H4` |
| 太字 かつ インデントレベル 1 | `##### H5` |
| 太字 かつ インデントレベル 2 以上 | `###### H6` |

フォントサイズ閾値は `--base-font-size` で調整できます（デフォルト: 11pt）。

### インデントとリスト

セルの列位置から階層を算出します。列幅の中央値（`col_unit`）を基準単位として、`col_unit × 1.5` 以内の列を同一階層とみなすことで、方眼紙の細かな位置ズレを吸収します。

- インデントレベル 1 以上かつ非見出し → `- リスト項目`
- `1.` や `1)` で始まるテキスト → `1. 番号付きリスト`

### 表（テーブル）

2行 × 2列以上の矩形グリッドを GFM テーブルに変換します。

```markdown
| 列1 | 列2 | 列3 |
| --- | --- |
| データ1 | データ2 | データ3 |
```

### ラベル:値パターン

同一行に2ブロックが横並びで、左のテキストが 20 文字以下の場合:

```markdown
**氏名:** 山田太郎
```

3列以上横並びの場合はスペース区切りで1段落に結合します。

### その他

- セル内改行（Alt+Enter）→ Markdown ハードブレーク（行末スペース2つ + 改行）
- 太字 → `**text**` / イタリック → `*text*` / 取り消し線 → `~~text~~` / 下線 → `<u>text</u>`
- ハイパーリンク → `[テキスト](URL)` 形式
- セルコメント → 脚注（`[^1]` 形式）
- 複数シートは `# シート名` と `---` で区切って1ファイルに統合

## サンプルファイル

`samples/` ディレクトリに動作確認用の方眼紙サンプルを収録しています。

| ファイル | 内容 |
| --- | --- |
| `samples/議事録_システム開発定例.xlsx` | セル結合なし・列ずらしで横並びを表現した議事録スタイル |
| `samples/詳細設計メモ_在庫照会.xls` | 旧形式。項目名と補足をそれぞれ別セルに書いた設計メモ |

再生成する場合:

```bash
python samples/make_samples.py
```

## 開発

### 環境セットアップ

```bash
pip install -e ".[xls,web,dev]"
```

### テスト

```bash
pytest
```

### リント・型チェック

```bash
ruff check .
mypy excel_to_markdown
```

## 要件

- Python 3.12+
- openpyxl 3.1+（必須）
- xlrd 2.x（`.xls` 変換時に必要）
- FastAPI / uvicorn / python-multipart（Web UI使用時に必要）

## スコープ外

- 画像・図形・グラフの変換
- マクロ・VBAの処理
- Google Sheets 対応
- `pip install` によるパッケージ公開（リポジトリをクローンして使用）
