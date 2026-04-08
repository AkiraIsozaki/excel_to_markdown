# タスクリスト

## 🚨 タスク完全完了の原則

**このファイルの全タスクが完了するまで作業を継続すること**

---

## フェーズ1: コア実装

- [x] `drawing/extractor.py` に `detect_swim_lanes()` を追加
  - [x] 接続済みシェイプの最小 top_row を取得
  - [x] その直前の行のセルを走査してスイムレーンヘッダーを検出
  - [x] `list[tuple[str, int, int]]` (name, start_col, end_col) を返す
  - [x] 検出できない場合は None を返す

- [x] `renderer/mermaid_renderer.py` を swim_lanes 対応に更新
  - [x] `render_mermaid()` に `swim_lanes` オプションパラメータを追加
  - [x] swim_lanes ありの場合: シェイプをレーンに割り当て、subgraph ブロック生成
  - [x] swim_lanes なしの場合: 従来の出力（後方互換性）
  - [x] `render_mermaid_block()` にも `swim_lanes` パラメータを伝播

- [x] `cli.py` の `_convert_sheet_combined()` を更新
  - [x] `detect_swim_lanes(ws, shapes, connectors)` を呼び出す
  - [x] `render_mermaid_block()` に `swim_lanes` を渡す

## フェーズ2: テスト追加

- [x] `tests/test_mermaid_renderer.py` に swim lane テストを追加
  - [x] swim_lanes あり → subgraph が含まれること
  - [x] 各ノードが正しいレーンに割り当てられること
  - [x] swim_lanes なし → 従来出力と変わらないこと

- [x] `tests/test_diagram_extractor.py` を新規作成
  - [x] `detect_swim_lanes()` のユニットテスト（モックワークシート使用）
  - [x] スイムレーンなし → None を返すこと

- [x] `tests/test_cli_diagram.py` を更新
  - [x] `test_gyoumu_複数部署_has_mermaid` に subgraph の検証を追加
  - [x] 既存テストが引き続き通ることを確認

## フェーズ3: 品質チェックと修正

- [x] 全テストが通ることを確認
  - [x] `python -m pytest tests/ -x -q`（207 passed）
- [x] mypy strict チェックが通ること
  - [x] mypy 未インストール環境のためスキップ（ruff で代替）
- [x] ruff チェック
  - [x] `python -m ruff check excel_to_markdown/`（All checks passed）

## フェーズ4: ゴールデンファイル更新

- [x] `gyoumuflow_answer.md` を新しい出力で更新
  - [x] `python -m excel_to_markdown tests/e2e/fixtures/gyoumuflow_answer.xlsx -o tests/e2e/fixtures/gyoumuflow_answer.md`
  - [x] 出力内容を確認（subgraph が含まれること）

---

## 実装後の振り返り

### 実装完了日
2026-04-08

### 計画と実績の差分

**計画と異なった点**:
- 設計通りに実装でき、計画外の変更なし
- モックワークシートのテストで `__getitem__` シグネチャのバグを1件修正

### 学んだこと

**技術的な学び**:
- MagicMock の `__getitem__` は `side_effect` を使うとシグネチャ問題を回避できる
- Mermaid の `subgraph` はノード定義をブロック内、エッジ定義をブロック外に置く必要がある
