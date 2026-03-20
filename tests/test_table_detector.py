"""table_detector.py のユニットテスト。"""

from __future__ import annotations

from excel_to_markdown.models import TextBlock
from excel_to_markdown.parser.cell_grid import CellGrid
from excel_to_markdown.parser.table_detector import find_tables


def _make_block(
    row: int,
    col: int,
    text: str = "X",
    bold: bool = False,
) -> TextBlock:
    return TextBlock(
        text=text,
        top_row=row,
        left_col=col,
        bottom_row=row,
        right_col=col,
        row_span=1,
        col_span=1,
        font_bold=bold,
        font_italic=False,
        font_strikethrough=False,
        font_underline=False,
        font_size=None,
        bg_color=None,
        has_comment=False,
        comment_text=None,
    )


def _make_grid() -> CellGrid:
    return CellGrid(cells=[], col_widths={}, row_heights={})


class TestFindTablesBasic:
    def test_2x2_grid_detected(self) -> None:
        blocks = [
            _make_block(1, 1, "A"), _make_block(1, 3, "B"),
            _make_block(2, 1, "C"), _make_block(2, 3, "D"),
        ]
        tables, remaining = find_tables(blocks, _make_grid())
        assert len(tables) == 1
        assert len(remaining) == 0
        t = tables[0]
        assert t.col_count == 2
        assert len(t.rows) == 2

    def test_3x3_grid_detected(self) -> None:
        blocks = [
            _make_block(1, 1), _make_block(1, 3), _make_block(1, 5),
            _make_block(2, 1), _make_block(2, 3), _make_block(2, 5),
            _make_block(3, 1), _make_block(3, 3), _make_block(3, 5),
        ]
        tables, remaining = find_tables(blocks, _make_grid())
        assert len(tables) == 1
        assert tables[0].col_count == 3
        assert len(tables[0].rows) == 3

    def test_remaining_blocks_excluded_from_table(self) -> None:
        """表外のブロックは remaining に含まれる。"""
        blocks = [
            _make_block(1, 1), _make_block(1, 3),
            _make_block(2, 1), _make_block(2, 3),
            _make_block(5, 2, "Standalone"),  # 表外
        ]
        tables, remaining = find_tables(blocks, _make_grid())
        assert len(tables) == 1
        assert len(remaining) == 1
        assert remaining[0].text == "Standalone"

    def test_table_source_row_is_first_row(self) -> None:
        blocks = [
            _make_block(3, 1), _make_block(3, 5),
            _make_block(4, 1), _make_block(4, 5),
        ]
        tables, _ = find_tables(blocks, _make_grid())
        assert tables[0].source_row == 3


class TestFindTablesIncomplete:
    def test_1_row_not_detected(self) -> None:
        blocks = [_make_block(1, 1), _make_block(1, 3)]
        tables, remaining = find_tables(blocks, _make_grid())
        assert len(tables) == 0
        assert len(remaining) == 2

    def test_1_col_not_detected(self) -> None:
        blocks = [_make_block(1, 1), _make_block(2, 1)]
        tables, remaining = find_tables(blocks, _make_grid())
        assert len(tables) == 0
        assert len(remaining) == 2

    def test_mismatched_columns_not_detected(self) -> None:
        """列境界が一致しない行は表として検出しない（保守的）。"""
        blocks = [
            _make_block(1, 1), _make_block(1, 3),
            _make_block(2, 1), _make_block(2, 5),  # 2列目が異なる
        ]
        tables, remaining = find_tables(blocks, _make_grid())
        assert len(tables) == 0
        assert len(remaining) == 4

    def test_empty_blocks(self) -> None:
        tables, remaining = find_tables([], _make_grid())
        assert tables == []
        assert remaining == []


class TestFindTablesHeader:
    def test_header_detected_when_first_row_bold(self) -> None:
        """1行目全 bold, 2行目非 bold → is_header=True。"""
        blocks = [
            _make_block(1, 1, bold=True), _make_block(1, 3, bold=True),
            _make_block(2, 1, bold=False), _make_block(2, 3, bold=False),
        ]
        tables, _ = find_tables(blocks, _make_grid())
        assert tables[0].rows[0][0].is_header is True
        assert tables[0].rows[0][1].is_header is True
        assert tables[0].rows[1][0].is_header is False

    def test_no_header_when_all_bold(self) -> None:
        """全行 bold → ヘッダー判定できず is_header=False。"""
        blocks = [
            _make_block(1, 1, bold=True), _make_block(1, 3, bold=True),
            _make_block(2, 1, bold=True), _make_block(2, 3, bold=True),
        ]
        tables, _ = find_tables(blocks, _make_grid())
        assert all(c.is_header is False for c in tables[0].rows[0])

    def test_no_header_when_no_bold(self) -> None:
        """全行非 bold → is_header=False。"""
        blocks = [
            _make_block(1, 1, bold=False), _make_block(1, 3, bold=False),
            _make_block(2, 1, bold=False), _make_block(2, 3, bold=False),
        ]
        tables, _ = find_tables(blocks, _make_grid())
        assert all(c.is_header is False for row in tables[0].rows for c in row)


class TestFindTablesContent:
    def test_cell_text_preserved(self) -> None:
        blocks = [
            _make_block(1, 1, "名前"), _make_block(1, 3, "年齢"),
            _make_block(2, 1, "田中"), _make_block(2, 3, "30"),
        ]
        tables, _ = find_tables(blocks, _make_grid())
        assert tables[0].rows[0][0].text == "名前"
        assert tables[0].rows[0][1].text == "年齢"
        assert tables[0].rows[1][0].text == "田中"
        assert tables[0].rows[1][1].text == "30"

    def test_multiple_separate_tables(self) -> None:
        """行が離れた2つの表がそれぞれ別のTableElementとして検出される。"""
        blocks = [
            _make_block(1, 1), _make_block(1, 3),
            _make_block(2, 1), _make_block(2, 3),
            _make_block(10, 1), _make_block(10, 3),
            _make_block(11, 1), _make_block(11, 3),
        ]
        tables, remaining = find_tables(blocks, _make_grid())
        assert len(tables) == 2
        assert len(remaining) == 0
