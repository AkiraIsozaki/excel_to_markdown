# タスクリスト: セル内容+Drawing 統合出力

## 🚨 タスク完全完了の原則
全タスクが`[x]`になるまで作業を継続する。

---

## フェーズ1: Extractor拡張（シート→Drawing マッピング）

- [x] `drawing/extractor.py` に `extract_sheet_drawing_map(xlsx_path)` を追加
  - [x] xlsx の workbook.xml + _rels からシート名→drawingパスのマッピングを構築
  - [x] シートごとに `(shapes, connectors)` を返す（row spanはshapesから計算）

## フェーズ2: Mermaidレンダラー改善

- [x] 空テキストシェイプの処理: テキスト空の場合はシェイプ名をフォールバックに
- [x] 未知シェイプタイプの追加マッピング（flowChartMagneticDisk, wedgeRoundRectCallout 等）
- [x] 孤立ノード（接続なし）をMermaidに含めるか否かのロジック整備（コネクタあり時は孤立除外、なし時は全表示）

## フェーズ3: CLI統合 - 自動結合モード

- [x] `cli.py` に `_build_drawing_map` は不要（`extract_sheet_drawing_map` を直接利用）
  - [x] `_process_xlsx()` 内で `extract_sheet_drawing_map` を呼び出してシート名→図形情報マップを取得
- [x] `_process_xlsx()` を拡張
  - [x] 各シートで drawingの有無を確認
  - [x] drawingありの場合: `_convert_sheet_combined()` を呼ぶ
  - [x] drawingなしの場合: 既存の `_run_pipeline()` を呼ぶ（変更なし）
- [x] `_convert_sheet_combined(ws, raw_cells, shapes, connectors, drawing_top_row, drawing_bottom_row, args)` を実装
  - [x] セル要素を drawing_top_row より前 / 後 に分割
  - [x] 前部分 → markdown
  - [x] drawing行範囲 → Mermaid block
  - [x] 後部分 → markdown
  - [x] 3つを結合して返す

## フェーズ4: サンプルファイルで動作確認

- [x] `tests/e2e/fixtures/1.xlsx` で変換実行し出力を確認
  - [x] '共通要件' シート: セルのみ → 通常markdown ✓
  - [x] '画面遷移' シート: セル(ヘッダー) + Mermaid(ログインフロー) → 結合 ✓
  - [x] 'トップ画面' シート: セル(ヘッダー) + Mermaid(UIモックアップ全ノード) → 結合 ✓
- [x] `tests/e2e/fixtures/gyoumuflow_answer.xlsx` で変換実行し出力を確認
  - [x] 'はじめに' シート: セルのみ → 通常markdown ✓
  - [x] '単一業務' シート: Drawing(凡例5種) → Mermaid ✓（特記事項セルは描画範囲内のため除外）
  - [x] '複数部署' シート: Drawing(業務フロー全ノード+エッジ) → Mermaid ✓

## フェーズ5: テスト追加

- [x] `tests/test_drawing_extractor.py` に `extract_sheet_drawing_map` のテストを追加
- [x] `tests/test_cli_diagram.py` に統合変換テストを追加（`TestAutoMermaidIntegration` クラス）
  - [x] drawing付きシートが Mermaid を含む出力になることを検証
  - [x] drawing なしシートは通常 markdown のみになることを検証
- [x] 全テスト pass 確認

## フェーズ6: 品質チェック

- [x] `python -m pytest tests/test_drawing_extractor.py tests/test_mermaid_renderer.py tests/test_cli_diagram.py -v` が全て pass（56 passed）
- [x] 既存テストへの回帰なし（187 passed 合計）

---

## 実装後の振り返り

### 実装完了日
2026-04-07

### 計画と実績の差分

**計画と異なった点**:
- `_build_drawing_map` は不要で、`extract_sheet_drawing_map` を `_process_xlsx` に直接組み込んだ
- `sample_flowchart.xlsx` の sheet rels に `2006` が欠落していた（既存バグ修正）
- openpyxl が workbook.xml.rels の Target に絶対パスを使うケースに対応（`split("/")[-1]` でファイル名のみ取得）

**新たに必要になったタスク**:
- `make_sample_flowchart.py` の関係タイプURL修正（`/officeDocument/relationships/` → `/officeDocument/2006/relationships/`）
- `extract_sheet_drawing_map` のパス正規化修正

### 学んだこと

**技術的な学び**:
- xlsx の OPC (Open Packaging Conventions) 関係タイプURLには `2006` が含まれるが、openpyxl が生成するRels XMLでは省略されることがある
- workbook.xml.rels の Target は相対パス(`worksheets/sheetN.xml`)と絶対パス(`/xl/worksheets/sheetN.xml`)両方ありうる
- Mermaid の孤立ノード判定（接続あり時は除外、接続なし時は全表示）でUIモックアップと業務フローを同一レンダラーで処理できる

### 次回への改善提案
- 空テキストシェイプの隣接シェイプからのテキスト補完（Excelでよくある「図形+ラベルテキストボックス」パターン）
- 泳ぎレーン（swim lane）のMermaid subgraph変換
- 複数のDrawingを持つシートへの対応
