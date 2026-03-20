"""TextBlock リストを DocElement リストに変換する構造検出エンジン。

依存可能: models.py, parser/cell_grid.py
依存禁止: reader/, renderer/, cli.py
"""

from __future__ import annotations

import re

from excel_to_markdown.models import DocElement, ElementType, TextBlock
from excel_to_markdown.parser.cell_grid import CellGrid

# 見出し判定のフォントサイズ倍率
_H1_RATIO = 18 / 11
_H2_RATIO = 14 / 11
_H3_RATIO = 12 / 11

# 白色の ARGB hex 表現
_WHITE = "FFFFFFFF"

# 番号付きリストの先頭パターン
_NUMBERED_LIST_RE = re.compile(
    r"^(?:\d+[.)）]|（\d+）|[①-⑨]|[㊀-㊉])\s"
)


def detect(
    blocks: list[TextBlock],
    grid: CellGrid,
    base_font_size: float = 11.0,
) -> list[DocElement]:
    """table_detector.find_tables() 後の残余 TextBlock を DocElement に変換する。

    処理順序:
    1. TextBlock にインデントレベルを付与
    2. 行グループ（同一 top_row）ごとに分類
    3. BLANK 要素を挿入
    4. source_row でソートして返す
    """
    if not blocks:
        return []

    # 1. インデントレベルを付与
    tiers = compute_indent_tiers(blocks, grid)
    for block in blocks:
        block.indent_level = tiers.get(block.left_col, 0)

    # 2. 行グループごとに分類
    elements: list[DocElement] = []
    row_groups = _group_same_row_blocks(blocks)

    prev_block: TextBlock | None = None
    for group in row_groups:
        # 空行挿入判定
        if prev_block is not None:
            group_top = group[0].top_row
            if _should_insert_blank(prev_block, group[0], grid):
                elements.append(
                    DocElement(
                        element_type=ElementType.BLANK,
                        text="",
                        level=0,
                        source_row=group_top - 1,
                    )
                )

        group_elements = _process_row_group(group, base_font_size)
        elements.extend(group_elements)
        prev_block = group[-1]

    return sorted(elements, key=lambda e: e.source_row)


def compute_indent_tiers(blocks: list[TextBlock], grid: CellGrid) -> dict[int, int]:
    """全 TextBlock の left_col を収集し、col_unit * 1.5 以内の列を同一ティアにグループ化する。

    戻り値: {left_col → indent_level} のマッピング

    例: col_unit=3, cols=[2, 3, 6, 7, 10] の場合
        tier0: [2, 3] (差=1 ≤ 4.5)
        tier1: [6, 7] (差=1 ≤ 4.5)
        tier2: [10]
    戻り値: {2: 0, 3: 0, 6: 1, 7: 1, 10: 2}
    """
    if not blocks:
        return {}

    threshold = grid.col_unit * 1.5
    sorted_cols = sorted(set(b.left_col for b in blocks))
    tiers: dict[int, int] = {sorted_cols[0]: 0}
    tier = 0
    for i in range(1, len(sorted_cols)):
        if sorted_cols[i] - sorted_cols[i - 1] > threshold:
            tier += 1
        tiers[sorted_cols[i]] = tier
    return tiers


def classify_heading(block: TextBlock, base_font_size: float) -> int | None:
    """見出しレベル (1-6) を返す。見出しでない場合は None を返す。

    判定は優先順位順 (functional-design.md 参照):
    1. font_size >= base * (18/11) → H1
    2. font_size >= base * (14/11) かつ bold → H2
    3. font_size >= base * (12/11) かつ bold → H3
    4. bold かつ indent_level == 0 → H4
    5. bold かつ indent_level == 1 → H5
    6. bold かつ indent_level >= 2 → H6

    前提: font_size=None の場合、優先度1〜3のフォントサイズ条件はすべて不成立。
    """
    fs = block.font_size
    bold = block.font_bold
    indent = block.indent_level

    if fs is not None:
        if fs >= base_font_size * _H1_RATIO:
            return 1
        if fs >= base_font_size * _H2_RATIO and bold:
            return 2
        if fs >= base_font_size * _H3_RATIO and bold:
            return 3

    if bold:
        if indent == 0:
            return 4
        if indent == 1:
            return 5
        return 6

    return None


def is_label_value_pair(left: TextBlock, right: TextBlock) -> bool:
    """同一行の2ブロックがラベル:値パターンか判定する。

    判定条件: left.text が 20 文字以下
    前提: top_row が一致する2ブロックの組に対して呼び出すこと。
    """
    return len(left.text) <= 20


# ---------------------------------------------------------------------------
# 内部ヘルパー
# ---------------------------------------------------------------------------


def _group_same_row_blocks(blocks: list[TextBlock]) -> list[list[TextBlock]]:
    """同一 top_row のブロックをグループ化して返す（top_row 昇順）。"""
    groups: dict[int, list[TextBlock]] = {}
    for b in blocks:
        groups.setdefault(b.top_row, []).append(b)
    for g in groups.values():
        g.sort(key=lambda b: b.left_col)
    return [groups[r] for r in sorted(groups.keys())]


def _process_row_group(group: list[TextBlock], base_font_size: float) -> list[DocElement]:
    """1行グループを DocElement リストに変換する。

    - 2ブロック かつ is_label_value_pair → PARAGRAPH (ラベル:値)
    - 3ブロック以上 → スペース区切りで1段落に結合
    - 1ブロック → 見出し/段落/リスト に分類
    """
    if len(group) == 2 and is_label_value_pair(group[0], group[1]):
        left, right = group[0], group[1]
        text = f"**{left.text}** {right.text}"
        return [
            DocElement(
                element_type=ElementType.PARAGRAPH,
                text=text,
                level=0,
                source_row=left.top_row,
                comment_text=left.comment_text or right.comment_text,
            )
        ]

    if len(group) >= 3:
        text = " ".join(b.text for b in group)
        return [
            DocElement(
                element_type=ElementType.PARAGRAPH,
                text=text,
                level=0,
                source_row=group[0].top_row,
                comment_text=next((b.comment_text for b in group if b.comment_text), None),
            )
        ]

    # 1ブロック
    block = group[0]
    return [_classify_single_block(block, base_font_size)]


def _classify_single_block(block: TextBlock, base_font_size: float) -> DocElement:
    """単一 TextBlock を DocElement に変換する。"""
    heading_level = classify_heading(block, base_font_size)
    if heading_level is not None:
        return DocElement(
            element_type=ElementType.HEADING,
            text=block.text,
            level=heading_level,
            source_row=block.top_row,
            comment_text=block.comment_text,
        )

    # リスト判定: インデントレベルが1以上の非見出し
    if block.indent_level >= 1:
        is_numbered = bool(_NUMBERED_LIST_RE.match(block.text))
        return DocElement(
            element_type=ElementType.LIST_ITEM,
            text=block.text,
            level=block.indent_level,
            source_row=block.top_row,
            is_numbered_list=is_numbered,
            comment_text=block.comment_text,
        )

    # 段落
    is_numbered = bool(_NUMBERED_LIST_RE.match(block.text))
    return DocElement(
        element_type=ElementType.PARAGRAPH,
        text=block.text,
        level=0,
        source_row=block.top_row,
        is_numbered_list=is_numbered,
        comment_text=block.comment_text,
    )


def _should_insert_blank(prev: TextBlock, curr: TextBlock, grid: CellGrid) -> bool:
    """前後の TextBlock 間に BLANK 要素を挿入すべきか判定する。

    条件:
    1. ブロック間の行ギャップ > modal_row_height × 2
    2. 前後の背景色が異なる
    3. 前ブロックの背景色が白以外で後ブロックが白またはNone
    """
    gap = curr.top_row - prev.bottom_row
    if gap > grid.modal_row_height * 2:
        return True

    prev_bg = prev.bg_color
    curr_bg = curr.bg_color

    # 背景色が異なる場合（有色→白・有色→他色・白→有色など全ケースを包含）
    if prev_bg != curr_bg:
        return True

    return False
