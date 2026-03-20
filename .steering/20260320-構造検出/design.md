# 設計: 構造検出

## 実装アプローチ

### ファイル作成順序

```
parser/table_detector.py
    ↓
parser/structure_detector.py
    ↓
tests/test_table_detector.py
tests/test_structure_detector.py
```

### table_detector.py の設計

```
ステップ1: 全TextBlockをrow→colの2次元マップに変換
ステップ2: 各ブロックを起点として矩形拡張を試みる
ステップ3: 候補矩形に対して列境界一致性を検証（全行で同一の left_col セット）
ステップ4: 最大の有効矩形を選択（貪欲法）
ステップ5: 検出した矩形に属するTextBlockを除外リストに追加
```

ヘッダー判定: 1行目全セルが bold かつ 2行目以降が非bold の場合のみ is_header=True

### structure_detector.py の設計

**インデントティア計算**:
```python
sorted_cols = sorted(set(b.left_col for b in blocks))
tier = 0
for i in range(1, len(sorted_cols)):
    if sorted_cols[i] - sorted_cols[i-1] > grid.col_unit * 1.5:
        tier += 1
    tiers[sorted_cols[i]] = tier
```

**見出し優先順位**:
1. font_size >= base * (18/11) → H1
2. font_size >= base * (14/11) かつ bold → H2
3. font_size >= base * (12/11) かつ bold → H3
4. bold かつ indent_level == 0 → H4
5. bold かつ indent_level == 1 → H5
6. bold かつ indent_level >= 2 → H6

**空行挿入ルール**:
- 行ギャップ > modal_row_height × 2 → BLANK
- 背景色が異なる → BLANK
- 前ブロック背景が FFFFFFFF 以外 → 後ブロック白/None なら BLANK

**ラベル:値パターン**: 同一行2ブロック、左ブロック20文字以下 → `**ラベル:** 値` で PARAGRAPH

**番号付きリスト判定**: text が `1.` / `1)` / `（1）` / `①` で始まる場合 is_numbered_list=True
