# 要求内容

## 概要

Excelのオブジェクト（図形）とコネクタで描かれた業務フロー図・システムフロー図・シーケンス図を、Mermaid形式に変換する機能を追加する。

## 背景

日本の現場ではExcel方眼紙上に図形（四角形・ひし形など）をコネクタ（矢印線）で接続し、業務フローやシステムフローを表現することが多い。
これらをテキストベースのMermaid形式に変換することで、Git管理・ドキュメントへの埋め込み・再利用が可能になる。

## 実装対象の機能

### 1. サンプルExcelファイルの生成

- DrawingML形式（xlsx内部のXML）でフローチャートを持つサンプルExcelを生成するスクリプト
- 以下の図形タイプを含む:
  - 開始/終了: `flowChartTerminator`（スタジアム形）
  - 処理: `flowChartProcess`（四角形）
  - 判断: `flowChartDecision`（ひし形）
- 図形同士を正式なDrawingMLコネクタ（`<xdr:cxnSp>`）で接続
- コネクタにラベル（「はい」「いいえ」など）を持つシーケンス

### 2. DrawingML図形抽出モジュール（`drawing/extractor.py`）

- xlsxファイル内の `xl/drawings/drawing*.xml` をパース
- 図形（`<xdr:sp>`）を `DiagramShape` データクラスとして抽出:
  - ID、名前、テキスト、位置（EMU単位）、形状タイプ（prst）
- コネクタ（`<xdr:cxnSp>`）を `DiagramConnector` データクラスとして抽出:
  - 始点Shape ID、終点Shape ID、ラベル
- 位置情報（行・列・オフセット）も保持

### 3. Mermaid変換モジュール（`renderer/mermaid_renderer.py`）

- `DiagramShape` + `DiagramConnector` → Mermaid flowchart 文字列に変換
- 形状タイプのMermaidノード種別マッピング:
  - `flowChartTerminator` → `([テキスト])`
  - `flowChartProcess` / `rect` → `[テキスト]`
  - `flowChartDecision` / `diamond` → `{テキスト}`
  - `ellipse` / `flowChartConnector` → `((テキスト))`
  - その他 → `[テキスト]`（デフォルト）
- グラフ方向の自動推定（TD / LR）: 図形配置の縦横比から判定

### 4. CLIへの統合

- `--diagram` オプションを追加: 図形抽出モードでMermaid出力
- 出力はMarkdownコードブロック（` ```mermaid ` ）として埋め込む

## 受け入れ条件

### サンプルファイル
- [ ] `tests/e2e/fixtures/make_sample_flowchart.py` を実行するとxlsxが生成される
- [ ] 生成されたxlsxをExcelで開いたとき、フローチャートとして見える（ユーザーが確認）

### 抽出モジュール
- [ ] サンプルxlsxから DiagramShape / DiagramConnector が正しく抽出できる
- [ ] 形状テキスト、形状タイプ、接続関係が正確に取得できる

### Mermaid変換
- [ ] サンプルxlsxのフローチャートが正しいMermaid文字列に変換される
- [ ] 変換後のMermaidをMermaid Live Editorで表示したとき、元図と同等の構造が再現される

### テスト
- [ ] 各モジュールのユニットテストが存在する
- [ ] 全テストが pass する

## スコープ外

- VML形式（古い .xls ファイルの図形）
- 図形画像（PNG/JPEG等の埋め込み画像）
- SmartArt
- コネクタなし（位置のみで接続判定する）ケース
- シーケンス図の泳ぎレーン自動認識

## 参照ドキュメント

- `docs/architecture.md` - アーキテクチャ設計書
- `docs/repository-structure.md` - リポジトリ構造定義書
