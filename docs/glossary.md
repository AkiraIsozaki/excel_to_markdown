# プロジェクト用語集 (Glossary)

## 概要

このドキュメントは、excel-to-markdownプロジェクトで使用される用語の定義を管理します。

**更新日**: 2026-03-20

---

## ドメイン用語

プロジェクト固有のビジネス概念・変換対象に関する用語。

### エクセル方眼紙

**定義**: セルを極小サイズ（方眼紙状）に設定し、ワープロ的に文書作成に使うExcelの利用スタイル。

**説明**:
日本のビジネス・行政・SIer現場で広く使われる文書作成手法。通常のスプレッドシートとは異なり、データ管理ではなく文書レイアウトに特化している。セルの結合・位置・フォント書式を組み合わせて見出しや段落を表現する。AIへの入力に使えないことが課題。

**特徴**:
- セルを結合してテキストブロックを作る
- 列位置でインデント・階層を表現する
- フォントサイズ・太字で見出しレベルを表現する
- 行の空きでセクション区切りを表現する

**関連用語**: TextBlock、インデントレベル、印刷領域

**英語表記**: Excel Graph Paper Style / Excel-as-Word-Processor

---

### 印刷領域（Print Area）

**定義**: Excelシートで「印刷範囲」として設定されたセル範囲。

**説明**:
エクセル方眼紙は印刷を前提に作成されることが多いため、印刷領域の外側には下書き・作業メモが含まれる場合がある。本ツールは印刷領域が設定されている場合、その内側のみを変換対象とする。

**本プロジェクトでの扱い**:
- 設定あり → 印刷領域内のセルのみ変換
- 設定なし → シート全体を変換

**関連用語**: エクセル方眼紙、RawCell

---

### ラベル:値パターン

**定義**: 同一行に横並びの2ブロックが「ラベル（短いテキスト）」と「値」の形式で配置されているパターン。

**説明**:
方眼紙でよく見られる「氏名: 山田太郎」「所属: 開発部」のような横並び構造。左のブロックが20文字以下の場合にラベル候補として判定する。

**変換例**:
```
Excel: [氏名:]  [山田太郎]
→ Markdown: **氏名:** 山田太郎
```

**関連用語**: TextBlock、インデントレベル、DocElement

---

### ベストエフォート変換

**定義**: 構造が認識できない場合でも、セルの内容を必ず段落として出力する変換方針。

**説明**:
フォントサイズ未設定・曖昧なレイアウト等、構造を正確に判定できないケースでも情報を捨てない。認識が不確かな要素にはMarkdownコメント（`<!-- WARNING: ... -->`）を挿入して変換担当者に通知する。

**関連用語**: DocElement（PARAGRAPH）、WARNINGコメント

---

## データモデル用語

変換パイプラインの各ステージで使用するデータ構造。

### RawCell

**定義**: openpyxl（またはxlrd）からセルの生データを取り出した中間データ構造。

**主要フィールド**:
- `row`, `col`: 1-based行・列番号
- `value`: セルの文字列値（数値は変換済み）
- `font_bold`, `font_italic`, `font_size`: フォント書式
- `bg_color`: 背景色（ARGB hex）
- `is_merge_origin`: 結合セルの起点か否か
- `merge_row_span`, `merge_col_span`: 結合スパン
- `has_comment`, `comment_text`: セルコメント

**制約**: 結合セルの非起点セルは `value=None` で処理をスキップする

**実装箇所**: `excel_to_markdown/models.py`

**関連エンティティ**: TextBlock（変換後）

---

### TextBlock

**定義**: 結合セル1単位を1つのテキストブロックとして表現した中間データ構造。RawCellをMergeResolverで変換したもの。

**主要フィールド**:
- `text`: テキスト内容
- `top_row`, `left_col`, `bottom_row`, `right_col`: 領域の座標
- `indent_level`: インデントレベル（StructureDetectorで後から付与）
- `inline_runs`: 部分的な書式（リッチテキスト）の分割リスト

**制約**: 空白のみのセルはTextBlockを生成しない

**実装箇所**: `excel_to_markdown/models.py`

**関連エンティティ**: RawCell（変換元）、DocElement（変換後）

---

### DocElement

**定義**: 文書要素を表すデータ構造。StructureDetectorがTextBlockを分類した結果。

**主要フィールド**:
- `element_type`: 要素の種別（ElementType列挙型）
- `text`: テキスト内容
- `level`: 見出しレベル（1-6）またはリストのインデント深さ
- `source_row`: 元のExcel行番号（ソート用）

**関連エンティティ**: TextBlock（変換元）、TableElement（表の特殊型）

**実装箇所**: `excel_to_markdown/models.py`

---

### TableElement

**定義**: DocElementのサブタイプ。グリッド状のセル配置から検出した表を表す。

**主要フィールド**:
- `rows`: TableCellの2次元リスト
- `col_count`: 列数

**制約**: 2行×2列以上のグリッドのみTableElementとして検出する

**実装箇所**: `excel_to_markdown/models.py`

---

### InlineRun

**定義**: セル内の部分的な書式（リッチテキスト）を表す最小単位。

**説明**: Excelではセル内の一部だけに太字・イタリックを適用できる。この単位ごとに対応するMarkdown記法に変換する。

**主要フィールド**:
- `text`: このRun内のテキスト
- `bold`, `italic`, `strikethrough`, `underline`: 書式フラグ

**実装箇所**: `excel_to_markdown/models.py`

---

## アーキテクチャ用語

### 変換パイプライン

**定義**: Excel → RawCell → TextBlock → DocElement → Markdown文字列 という5段階の一方向データフロー。

**本プロジェクトでの適用**: 各ステージは前段の出力のみに依存し、後段を知らない。これによりreader層の変更（xlsx→xls追加等）が後段に影響しない。

```
[.xlsx/.xls] → [RawCell] → [TextBlock] → [DocElement] → [Markdown]
               Reader        MergeResolver  StructureDetector  Renderer
```

**関連コンポーネント**: xlsx_reader、merge_resolver、structure_detector、markdown_renderer

---

### col_unit

**定義**: シート全体の列幅の中央値。方眼紙のグリッドピッチの推定値として使用する。

**説明**:
方眼紙では全セルをほぼ同じ幅に設定するため、列幅の中央値が方眼紙の基本グリッド幅になる。インデントレベルの計算や、空行挿入の行ギャップ判定に使用する。

**計算式**: `col_unit = median(column_widths.values())`

**実装箇所**: `excel_to_markdown/parser/cell_grid.py` の `CellGrid.col_unit` プロパティ

**関連用語**: インデントレベル、インデントティア

---

### インデントレベル

**定義**: TextBlockの左端列位置から算出した階層の深さ。0が最も左（基準）。

**説明**: 方眼紙では列位置でインデントを表現する。同一インデントを `col_unit × 1.5` の許容幅でグループ化する（位置ズレ吸収）。

**実装箇所**: `excel_to_markdown/parser/structure_detector.py` の `compute_indent_tiers()`

**関連用語**: col_unit、インデントティア

---

### インデントティア

**定義**: `col_unit × 1.5` 以内の列をまとめた「同一インデントレベルのグループ」。

**説明**:
方眼紙の著者が列位置を完全に揃えないケースを吸収するためのグループ化。列番号のソート列に対して貪欲法でグループを形成し、グループ番号＝インデントレベルとする。

**計算例**:
```
col_unit=3, 許容幅=4.5
列番号: [2, 3, 8, 9, 14]
→ tier0: [2, 3]  tier1: [8, 9]  tier2: [14]
→ {2:0, 3:0, 8:1, 9:1, 14:2}
```

**実装箇所**: `excel_to_markdown/parser/structure_detector.py`

---

### modal_row_height

**定義**: 行高さの最頻値。方眼紙の行グリッドピッチの推定値として使用する。

**説明**: 行ギャップが `modal_row_height × 2` を超える場合に空行（BLANK要素）を挿入する。

**実装箇所**: `excel_to_markdown/parser/cell_grid.py` の `CellGrid.modal_row_height` プロパティ

---

## ElementType（要素種別）

DocElementの`element_type`フィールドで使用する列挙型。

| 値 | 意味 | Markdown出力例 |
|----|------|--------------|
| `HEADING` | 見出し（H1〜H6） | `# タイトル` |
| `PARAGRAPH` | 段落 | `テキスト` |
| `LIST_ITEM` | リスト項目 | `- 項目` |
| `TABLE` | 表（GFMテーブル） | `\| col \| col \|` |
| `BLANK` | 空行区切り | （空行） |

---

## 技術用語

### openpyxl

**定義**: Pythonで.xlsxファイルを読み書きするためのライブラリ。

**本プロジェクトでの用途**: .xlsx形式のExcelファイルを読み込み、セル値・フォント書式・結合情報・コメントを取得する。

**バージョン**: `>=3.1.0,<4.0.0`

**注意事項**: `read_only=True` では `ws.merged_cells.ranges` にアクセスできないため、通常モードで使用する。

---

### xlrd

**定義**: Pythonで.xls（旧形式Excel）ファイルを読み込むためのライブラリ。

**本プロジェクトでの用途**: .xls形式のExcelファイルを読み込み、openpyxlと同一のRawCellデータモデルに変換する（P1対応）。

**バージョン**: `>=2.0.0,<3.0.0`（.xlsxには非対応）

---

### GFM（GitHub Flavored Markdown）

**定義**: GitHubが拡張したMarkdown方言。テーブル・取り消し線・タスクリスト等が追加されている。

**本プロジェクトでの用途**: 出力するMarkdownの形式として採用。特にテーブル（`| col |`形式）と取り消し線（`~~text~~`）を使用する。

---

## 略語・頭字語

### PRD

**正式名称**: Product Requirements Document（プロダクト要求定義書）

**本プロジェクトでの使用**: `docs/product-requirements.md`

---

### CLI

**正式名称**: Command Line Interface（コマンドラインインターフェース）

**本プロジェクトでの使用**: `python -m excel_to_markdown` コマンドで提供するインターフェース。

---

### WARNINGコメント

**定義**: 変換結果のMarkdownに挿入する構造認識失敗の通知コメント。

**形式**: `<!-- WARNING: 構造を認識できませんでした -->`

**本プロジェクトでの使用**: ベストエフォート変換でセルを段落として出力した際に、認識失敗を明示するために挿入する。

---

## アルゴリズム用語

### 見出し判定アルゴリズム

**定義**: TextBlockのフォント書式とインデントレベルから見出しレベル（H1〜H6）を判定する優先順位付きルール集。

**判定式**（優先度順・`base = base_font_size`）:

```
1. font_size >= base * (18/11) → H1
2. font_size >= base * (14/11) かつ bold → H2
3. font_size >= base * (12/11) かつ bold → H3
4. bold かつ indent_level == 0 → H4
5. bold かつ indent_level == 1 → H5
6. bold かつ indent_level >= 2 → H6
7. 該当なし → 見出しでない
```

**実装箇所**: `excel_to_markdown/parser/structure_detector.py` の `classify_heading()`

---

### グリッド表検出

**定義**: TextBlockの空間配置からGFMテーブルに変換すべき矩形グリッドを検出するアルゴリズム。

**検出条件**:
- 2行以上 × 2列以上の矩形
- 各行の列境界（left_col）が一致
- 曖昧なケース（不完全グリッド）は非検出（保守的判定）

**実装箇所**: `excel_to_markdown/parser/table_detector.py` の `find_tables()`
