"""TextBlock リストからグリッド表を検出して TableElement に変換する。

依存可能: models.py, parser/cell_grid.py
依存禁止: reader/, renderer/, cli.py
"""

from __future__ import annotations

from excel_to_markdown.models import TableCell, TableElement, TextBlock
from excel_to_markdown.parser.cell_grid import CellGrid


def find_tables(
    blocks: list[TextBlock],
    grid: CellGrid,
) -> tuple[list[TableElement], list[TextBlock]]:
    """グリッド状の配置を検出して TableElement に変換する。

    戻り値: (検出した TableElement リスト, 表に含まれなかった TextBlock リスト)

    検出条件:
    - 2行以上 × 2列以上の矩形
    - 各行の列境界 (left_col) が一致している
    - 曖昧なケース（不完全グリッド）は検出しない（保守的）
    """
    # row → col → TextBlock の2次元マップを構築
    row_col_map: dict[int, dict[int, TextBlock]] = {}
    for block in blocks:
        row_col_map.setdefault(block.top_row, {})[block.left_col] = block

    used: set[int] = set()  # 使用済み TextBlock の id
    tables: list[TableElement] = []

    # 行・列でソートした起点から貪欲に矩形を探す
    sorted_rows = sorted(row_col_map.keys())
    for start_row in sorted_rows:
        row_blocks = row_col_map[start_row]
        # 既に表に取り込まれた起点はスキップ
        if all(id(b) in used for b in row_blocks.values()):
            continue

        cols_in_start_row = sorted(row_blocks.keys())
        if len(cols_in_start_row) < 2:
            continue  # 1列のみでは表にならない

        # 起点行から下に向かって同じ列境界を持つ隣接行を探す
        # 「隣接」: 直前の行の bottom_row の直後の行（行番号が連続している）
        table_rows: list[int] = [start_row]
        for r in sorted_rows:
            if r <= start_row:
                continue
            if r not in row_col_map:
                break
            # 前の行の直後の行番号のみを許可（行ギャップがあれば停止）
            prev_row = table_rows[-1]
            prev_bottom = max(
                b.bottom_row for b in row_col_map[prev_row].values()
            )
            if r > prev_bottom + 1:
                break  # 行ギャップあり → 表終了
            if sorted(row_col_map[r].keys()) == cols_in_start_row:
                # 全ブロックが未使用であることを確認
                if any(id(b) in used for b in row_col_map[r].values()):
                    break
                table_rows.append(r)
            else:
                break  # 列境界が崩れたら終了

        if len(table_rows) < 2:
            continue  # 1行のみでは表にならない

        # 表として確定。TableElement を構築
        table_element = _build_table(table_rows, row_col_map, cols_in_start_row)
        tables.append(table_element)
        for r in table_rows:
            for b in row_col_map[r].values():
                used.add(id(b))

    remaining = [b for b in blocks if id(b) not in used]
    return tables, remaining


def _build_table(
    table_rows: list[int],
    row_col_map: dict[int, dict[int, TextBlock]],
    cols: list[int],
) -> TableElement:
    """検出済み矩形から TableElement を構築する。"""
    rows: list[list[TableCell]] = []
    col_count = len(cols)

    # ヘッダー判定: 1行目全セルが bold かつ 2行目以降の少なくとも1行が非bold
    first_row_blocks = [row_col_map[table_rows[0]][c] for c in cols]
    first_row_all_bold = all(b.font_bold for b in first_row_blocks)
    has_non_bold_row = any(
        not all(row_col_map[r][c].font_bold for c in cols) for r in table_rows[1:]
    )
    use_header = first_row_all_bold and has_non_bold_row

    for rel_row, r in enumerate(table_rows):
        row_cells: list[TableCell] = []
        is_header = use_header and rel_row == 0
        for rel_col, c in enumerate(cols):
            block = row_col_map[r][c]
            tc = TableCell(text=block.text, row=rel_row, col=rel_col, is_header=is_header)
            row_cells.append(tc)
        rows.append(row_cells)

    # source_row は表の先頭行
    return TableElement(
        text="",
        level=0,
        source_row=table_rows[0],
        rows=rows,
        col_count=col_count,
    )
