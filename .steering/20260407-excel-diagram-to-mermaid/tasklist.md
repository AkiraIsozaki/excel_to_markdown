# タスクリスト

## 🚨 タスク完全完了の原則

**このファイルの全タスクが完了するまで作業を継続すること**

---

## フェーズ1: データモデル追加

- [x] `models.py` に `DiagramShape` dataclass を追加
- [x] `models.py` に `DiagramConnector` dataclass を追加

## フェーズ2: サンプルxlsx生成スクリプト

- [x] `tests/e2e/fixtures/make_sample_flowchart.py` を作成
  - [x] 受注処理フロー（開始→受注受付→在庫確認?→出荷処理→完了 / 発注処理→入荷待ち→出荷処理）
  - [x] 6種の図形（flowChartTerminator×2, flowChartProcess×4, flowChartDecision×1）
  - [x] DrawingMLコネクタ（stCxn/endCxnで接続、ラベル付き2本）
  - [x] スクリプト実行で `tests/e2e/fixtures/sample_flowchart.xlsx` が生成される
- [x] スクリプトを実行してxlsxが生成されることを確認

## フェーズ3: DrawingML抽出モジュール

- [x] `excel_to_markdown/drawing/__init__.py` を作成
- [x] `excel_to_markdown/drawing/extractor.py` を実装
  - [x] `extract_diagrams(xlsx_path)` 関数: ZipFileを開いてdrawingを検索
  - [x] `<xdr:sp>` → DiagramShape の変換ロジック
  - [x] `<xdr:cxnSp>` → DiagramConnector の変換ロジック
  - [x] TwoCellAnchor / OneCellAnchor 両方に対応

## フェーズ4: 抽出モジュールのテスト

- [x] `tests/test_drawing_extractor.py` を作成
  - [x] サンプルxlsxから正しい数のShapeが抽出できることを検証
  - [x] 各Shapeのshape_type・textが正確に取得できることを検証
  - [x] コネクタのstart_shape_id / end_shape_idが正確に取得できることを検証
  - [x] drawingなしのxlsxに対して空リストが返ることを検証
- [x] テストが全て pass することを確認

## フェーズ5: Mermaidレンダラー

- [x] `excel_to_markdown/renderer/mermaid_renderer.py` を実装
  - [x] 形状タイプ → Mermaidノード記法のマッピング
  - [x] グラフ方向自動判定（TD / LR）
  - [x] ノードID生成（`N{shape_id}` 形式）
  - [x] コネクタのエッジ生成（ラベルあり/なし）
  - [x] Markdownコードブロック（` ```mermaid ` ）として出力する関数

## フェーズ6: Mermaidレンダラーのテスト

- [x] `tests/test_mermaid_renderer.py` を作成
  - [x] 形状タイプマッピングのテスト（各種prst → Mermaidノード記法）
  - [x] コネクタ→エッジのテスト（ラベルあり/なし）
  - [x] グラフ方向判定のテスト
  - [x] サンプルxlsxから生成したMermaid文字列の内容検証
- [x] テストが全て pass することを確認

## フェーズ7: CLIへの統合

- [x] `cli.py` に `--diagram` フラグを追加
  - [x] `--diagram` 指定時に図形変換パイプラインを実行
  - [x] drawingが存在しない場合の適切なメッセージ出力
- [x] CLI統合テスト（`tests/test_cli_diagram.py` に作成: test_cli.py は xlwt skipのため分離）

## フェーズ8: 品質チェック

- [x] `python -m pytest tests/test_drawing_extractor.py tests/test_mermaid_renderer.py -v` が全て pass
- [x] `python -m pytest` で全テストが pass（既存テストの回帰なし: 175 passed）
- [x] `ruff check excel_to_markdown/drawing/ excel_to_markdown/renderer/mermaid_renderer.py` が pass

## フェーズ9: ドキュメント更新と振り返り

- [x] 振り返りをこのファイルに記録

---

## 実装後の振り返り

### 実装完了日
2026-04-07

### 計画と実績の差分

**計画と異なった点**:
- Mermaidのひし形記法テンプレートで中括弧の扱いが問題に。テンプレートプレースホルダーを `{text}` から `__TEXT__` に変更してクリーンに解決
- `test_cli.py` が xlwt の `pytest.importorskip` でファイル全体がスキップされるため、CLIテストを `tests/test_cli_diagram.py` に分離

**新たに必要になったタスク**:
- `tests/test_cli_diagram.py` の新規作成（test_cli.pyへの追記から変更）

### 学んだこと

**技術的な学び**:
- xlsx は ZIP ファイルであり、DrawingML の図形情報は `xl/drawings/drawing1.xml` に格納される
- openpyxl のモデルクラス（`ConnectorShape`, `Shape`）はあるが、シートからの `_drawing` アクセスは限定的のため、ZipFile + ElementTree で直接パースする方が確実
- Mermaid のひし形記法は `{text}` (シングル中括弧)、六角形は `{{text}}`
- `pytest.importorskip` をクラスレベルで使うとファイル全体がスキップされる

### 次回への改善提案
- コネクタなし（位置ベース接続推定）のケースへの対応
- シーケンス図 → `sequenceDiagram` 形式への変換
- 複数シートの drawing を統合する際の名前空間管理
