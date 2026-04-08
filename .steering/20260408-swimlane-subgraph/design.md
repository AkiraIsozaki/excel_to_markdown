# 設計書

## アーキテクチャ概要

既存の drawing/extractor → renderer/mermaid_renderer → cli パイプラインに、
スイムレーン検出ステップを追加する。

```
cli._convert_sheet_combined(ws, shapes, connectors, ...)
    │
    ├─ detect_swim_lanes(ws, shapes)  ← 新規
    │     returns: list[SwimLane] | None
    │
    └─ render_mermaid_block(shapes, connectors, swim_lanes=...)
          ├─ (swim_lanes=None) → 従来のフラット出力
          └─ (swim_lanes=[...]) → subgraph ブロック出力
```

## コンポーネント設計

### 1. `drawing/extractor.py` - `detect_swim_lanes()`

**責務**: ワークシートのセルからスイムレーンヘッダーを検出し、列範囲を特定する

**アルゴリズム**:
1. 接続されたシェイプ（コネクタの端点にあるシェイプ）の最小 top_row を求める
2. その行（0-based）の直前にある最も近い非空セル行を探す（1-based row で検索）
3. その行に 2+ の非空セルがあれば、各セルの列位置をスイムレーン開始位置とする
4. スイムレーン終了位置: 次のスイムレーン開始 - 1（最後は無限大）
5. 戻り値: `list[tuple[str, int, int]]` = [(name, start_col_0based, end_col_0based)]
   - `end_col_0based` は最後のレーンは大きな値（999）を使う

**判定ロジック**:
```python
def detect_swim_lanes(ws, shapes) -> list[tuple[str, int, int]] | None:
    connected_ids = {c.start_shape_id, c.end_shape_id for c in connectors} ...
    # shapes を connectors とともに受け取る必要あり
```

→ シグネチャ: `detect_swim_lanes(ws, shapes, connectors) -> list[tuple[str, int, int]] | None`

**スイムレーン割り当て**:
- 各シェイプの center_col = (left_col + right_col) / 2
- center_col が [start_col, end_col] の範囲に入るレーンに割り当て
- どのレーンにも入らない場合: 最も近いレーンに割り当て（孤立ノード防止）

### 2. `renderer/mermaid_renderer.py` - `render_mermaid()` 拡張

**責務**: swim_lanes パラメータが渡された場合に subgraph ブロックを生成する

**出力例**:
```
flowchart TD
    subgraph lane_0 [お客様]
        N7(窓ガラスが割れた...)
        N30[現場訪問に立ち合う]
    end
    subgraph lane_1 [〇〇ガラス店]
        N12[受付用紙に記入する]
        N16[現場確認のための...]
    end
    subgraph lane_2 [問屋]
        N58[見積依頼を受領し...]
    end
    N7 --> N12
    N12 --> N16
    ...
```

**実装の要点**:
- ノード定義は各 subgraph 内に出力
- エッジはすべての subgraph の後に出力（Mermaidの仕様上）
- 孤立ノード（コネクタなし）もスイムレーンに含める
- swim_lanes=None の場合は従来通りの出力（後方互換性）

### 3. `cli.py` - `_convert_sheet_combined()` 更新

**変更箇所**:
```python
# 変更前
parts.append(render_mermaid_block(shapes, connectors))

# 変更後
swim_lanes = detect_swim_lanes(ws, shapes, connectors)
parts.append(render_mermaid_block(shapes, connectors, swim_lanes=swim_lanes))
```

`detect_swim_lanes` を `drawing.extractor` からインポート。

## データフロー

### スイムレーン付き変換
```
1. extract_sheet_drawing_map() → shapes, connectors
2. detect_swim_lanes(ws, shapes, connectors) → swim_lanes
3. render_mermaid_block(shapes, connectors, swim_lanes=swim_lanes)
4. → subgraph を含む Mermaid 文字列
```

## テスト戦略

### ユニットテスト（`tests/test_mermaid_renderer.py` に追加）
- swim_lanes あり → subgraph が含まれることを検証
- swim_lanes なし → 従来通りのフラット出力
- ノードのレーン割り当てが正しいこと

### ユニットテスト（`tests/test_diagram_extractor.py` 新規作成）
- detect_swim_lanes: ヘッダー行からレーンを検出
- スイムレーンなし → None を返す

### 統合テスト（`tests/test_cli_diagram.py` 更新）
- gyoumuflow 複数部署: subgraph が出力に含まれる
- スイムレーンなしのケース: 従来出力を維持

## ディレクトリ構造

```
excel_to_markdown/
  drawing/
    extractor.py          ← detect_swim_lanes() 追加
  renderer/
    mermaid_renderer.py   ← swim_lanes パラメータ追加
  cli.py                  ← swim lane 検出の呼び出し追加
tests/
  test_mermaid_renderer.py  ← swim lane テスト追加
  test_diagram_extractor.py ← 新規（detect_swim_lanes テスト）
  test_cli_diagram.py       ← gyoumu subgraph テスト追加
```

## 実装の順序

1. `drawing/extractor.py` に `detect_swim_lanes()` を追加
2. `renderer/mermaid_renderer.py` を swim_lanes 対応に更新
3. `cli.py` を更新して swim lane 検出・渡し込みを追加
4. テスト追加・更新
5. `gyoumuflow_answer.md` ゴールデンファイル更新
