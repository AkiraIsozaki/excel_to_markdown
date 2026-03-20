"""cell_grid.py のユニットテスト。"""

from __future__ import annotations

from excel_to_markdown.models import RawCell
from excel_to_markdown.parser.cell_grid import CellGrid


def _make_cell(row: int, col: int, value: str | None = "X") -> RawCell:
    return RawCell(
        row=row,
        col=col,
        value=value,
        font_bold=False,
        font_italic=False,
        font_strikethrough=False,
        font_underline=False,
        font_size=None,
        font_color=None,
        bg_color=None,
        is_merge_origin=False,
        merge_row_span=1,
        merge_col_span=1,
        has_comment=False,
        comment_text=None,
    )


class TestBaselineCol:
    def test_returns_leftmost_col_with_value(self) -> None:
        cells = [_make_cell(1, 3), _make_cell(1, 5), _make_cell(2, 4)]
        grid = CellGrid(cells=cells)
        assert grid.baseline_col == 3

    def test_ignores_none_value_cells(self) -> None:
        cells = [_make_cell(1, 1, None), _make_cell(1, 4, "X")]
        grid = CellGrid(cells=cells)
        assert grid.baseline_col == 4

    def test_empty_cells_returns_1(self) -> None:
        grid = CellGrid(cells=[])
        assert grid.baseline_col == 1


class TestColUnit:
    def test_median_of_col_widths(self) -> None:
        grid = CellGrid(cells=[], col_widths={1: 2.0, 2: 4.0, 3: 6.0})
        assert grid.col_unit == 4.0

    def test_empty_col_widths_returns_default(self) -> None:
        grid = CellGrid(cells=[], col_widths={})
        assert grid.col_unit == 8.0

    def test_zero_widths_ignored(self) -> None:
        grid = CellGrid(cells=[], col_widths={1: 0.0, 2: 4.0, 3: 8.0})
        assert grid.col_unit == 6.0


class TestModalRowHeight:
    def test_mode_of_row_heights(self) -> None:
        grid = CellGrid(cells=[], row_heights={1: 15.0, 2: 15.0, 3: 20.0})
        assert grid.modal_row_height == 15.0

    def test_empty_row_heights_returns_default(self) -> None:
        grid = CellGrid(cells=[], row_heights={})
        assert grid.modal_row_height == 15.0

    def test_all_same_returns_that_value(self) -> None:
        grid = CellGrid(cells=[], row_heights={1: 18.0, 2: 18.0, 3: 18.0})
        assert grid.modal_row_height == 18.0


class TestIsEmptyRow:
    def test_row_with_value_is_not_empty(self) -> None:
        cells = [_make_cell(2, 1, "Hello")]
        grid = CellGrid(cells=cells)
        assert grid.is_empty_row(2) is False

    def test_row_without_value_is_empty(self) -> None:
        cells = [_make_cell(1, 1, "Hello")]
        grid = CellGrid(cells=cells)
        assert grid.is_empty_row(2) is True

    def test_row_with_none_value_is_empty(self) -> None:
        cells = [_make_cell(1, 1, None)]
        grid = CellGrid(cells=cells)
        assert grid.is_empty_row(1) is True
