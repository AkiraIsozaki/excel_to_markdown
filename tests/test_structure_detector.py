"""structure_detector.py のユニットテスト。"""

from __future__ import annotations

from excel_to_markdown.models import ElementType, TextBlock
from excel_to_markdown.parser.cell_grid import CellGrid
from excel_to_markdown.parser.structure_detector import (
    classify_heading,
    compute_indent_tiers,
    detect,
    is_label_value_pair,
)


def _make_block(
    row: int,
    col: int = 1,
    text: str = "Hello",
    bold: bool = False,
    font_size: float | None = None,
    bg_color: str | None = None,
    has_comment: bool = False,
    comment_text: str | None = None,
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
        font_size=font_size,
        bg_color=bg_color,
        has_comment=has_comment,
        comment_text=comment_text,
    )


def _make_grid(col_unit: float = 3.0, modal_row_height: float = 15.0) -> CellGrid:
    """指定した col_unit / modal_row_height を持つ CellGrid を返す。"""
    # col_widths の中央値が col_unit になるように設定
    return CellGrid(
        cells=[],
        col_widths={1: col_unit},
        row_heights={1: modal_row_height},
    )


# ---------------------------------------------------------------------------
# compute_indent_tiers
# ---------------------------------------------------------------------------


class TestComputeIndentTiers:
    def test_single_col_is_tier_0(self) -> None:
        blocks = [_make_block(1, col=2)]
        grid = _make_grid(col_unit=3.0)
        tiers = compute_indent_tiers(blocks, grid)
        assert tiers[2] == 0

    def test_close_cols_same_tier(self) -> None:
        """差が col_unit * 1.5 以内 → 同一ティア。"""
        blocks = [_make_block(1, col=2), _make_block(2, col=3)]
        grid = _make_grid(col_unit=3.0)  # threshold = 4.5
        tiers = compute_indent_tiers(blocks, grid)
        assert tiers[2] == tiers[3] == 0

    def test_distant_cols_different_tiers(self) -> None:
        """差が col_unit * 1.5 超 → 異なるティア。"""
        blocks = [_make_block(1, col=2), _make_block(2, col=8)]  # 差=6 > 4.5
        grid = _make_grid(col_unit=3.0)
        tiers = compute_indent_tiers(blocks, grid)
        assert tiers[2] == 0
        assert tiers[8] == 1

    def test_three_tiers(self) -> None:
        blocks = [
            _make_block(1, col=2),
            _make_block(2, col=8),
            _make_block(3, col=14),
        ]
        grid = _make_grid(col_unit=3.0)
        tiers = compute_indent_tiers(blocks, grid)
        assert tiers[2] == 0
        assert tiers[8] == 1
        assert tiers[14] == 2

    def test_empty_blocks(self) -> None:
        tiers = compute_indent_tiers([], _make_grid())
        assert tiers == {}


# ---------------------------------------------------------------------------
# classify_heading
# ---------------------------------------------------------------------------


class TestClassifyHeading:
    BASE = 11.0

    def test_h1_by_large_font(self) -> None:
        block = _make_block(1, font_size=18.0)  # 18 >= 11 * 18/11
        assert classify_heading(block, self.BASE) == 1

    def test_h2_by_font_and_bold(self) -> None:
        block = _make_block(1, bold=True, font_size=14.0)
        assert classify_heading(block, self.BASE) == 2

    def test_h3_by_font_and_bold(self) -> None:
        block = _make_block(1, bold=True, font_size=12.0)
        assert classify_heading(block, self.BASE) == 3

    def test_h4_bold_indent0(self) -> None:
        block = _make_block(1, bold=True)
        block.indent_level = 0
        assert classify_heading(block, self.BASE) == 4

    def test_h5_bold_indent1(self) -> None:
        block = _make_block(1, bold=True)
        block.indent_level = 1
        assert classify_heading(block, self.BASE) == 5

    def test_h6_bold_indent2(self) -> None:
        block = _make_block(1, bold=True)
        block.indent_level = 2
        assert classify_heading(block, self.BASE) == 6

    def test_not_heading(self) -> None:
        block = _make_block(1, bold=False, font_size=11.0)
        assert classify_heading(block, self.BASE) is None

    def test_none_fontsize_with_bold_is_h4(self) -> None:
        """font_size=None かつ bold → H4（H2/H3のサイズ条件は不成立）。"""
        block = _make_block(1, bold=True, font_size=None)
        block.indent_level = 0
        assert classify_heading(block, self.BASE) == 4

    def test_large_font_without_bold_is_h1(self) -> None:
        """H1 は bold 不要。"""
        block = _make_block(1, bold=False, font_size=18.0)
        assert classify_heading(block, self.BASE) == 1

    def test_h2_requires_bold(self) -> None:
        """H2 の font_size 条件を満たしても bold なしでは H2 にならない。"""
        block = _make_block(1, bold=False, font_size=14.0)
        # H2 条件不成立, H1 条件も不成立(18未満), bold なし → None
        assert classify_heading(block, self.BASE) is None


# ---------------------------------------------------------------------------
# is_label_value_pair
# ---------------------------------------------------------------------------


class TestIsLabelValuePair:
    def test_short_left_is_label(self) -> None:
        left = _make_block(1, col=1, text="氏名")  # 2文字
        right = _make_block(1, col=5, text="山田太郎")
        assert is_label_value_pair(left, right) is True

    def test_exactly_20_chars_is_label(self) -> None:
        left = _make_block(1, col=1, text="a" * 20)
        right = _make_block(1, col=5, text="value")
        assert is_label_value_pair(left, right) is True

    def test_21_chars_is_not_label(self) -> None:
        left = _make_block(1, col=1, text="a" * 21)
        right = _make_block(1, col=5, text="value")
        assert is_label_value_pair(left, right) is False


# ---------------------------------------------------------------------------
# detect: 統合テスト
# ---------------------------------------------------------------------------


class TestDetect:
    def test_empty_returns_empty(self) -> None:
        assert detect([], _make_grid()) == []

    def test_single_bold_block_is_heading(self) -> None:
        blocks = [_make_block(1, bold=True, text="タイトル")]
        elements = detect(blocks, _make_grid())
        assert len(elements) == 1
        assert elements[0].element_type == ElementType.HEADING
        assert elements[0].text == "タイトル"

    def test_non_bold_block_is_paragraph(self) -> None:
        blocks = [_make_block(1, text="本文テキスト")]
        elements = detect(blocks, _make_grid())
        assert len(elements) == 1
        assert elements[0].element_type == ElementType.PARAGRAPH

    def test_indented_block_is_list_item(self) -> None:
        blocks = [
            _make_block(1, col=1, text="見出し", bold=True),
            _make_block(2, col=7, text="リスト項目"),  # 差=6 > 4.5 → tier1
        ]
        grid = _make_grid(col_unit=3.0)
        elements = detect(blocks, grid)
        list_items = [e for e in elements if e.element_type == ElementType.LIST_ITEM]
        assert len(list_items) == 1
        assert list_items[0].text == "リスト項目"
        assert list_items[0].level == 1

    def test_label_value_pair_becomes_paragraph(self) -> None:
        blocks = [
            _make_block(1, col=1, text="氏名"),
            _make_block(1, col=5, text="山田太郎"),
        ]
        elements = detect(blocks, _make_grid())
        assert len(elements) == 1
        el = elements[0]
        assert el.element_type == ElementType.PARAGRAPH
        assert "**氏名**" in el.text
        assert "山田太郎" in el.text

    def test_3_blocks_same_row_merged_to_paragraph(self) -> None:
        blocks = [
            _make_block(1, col=1, text="A"),
            _make_block(1, col=3, text="B"),
            _make_block(1, col=5, text="C"),
        ]
        elements = detect(blocks, _make_grid())
        assert len(elements) == 1
        assert elements[0].element_type == ElementType.PARAGRAPH
        assert elements[0].text == "A B C"

    def test_blank_inserted_on_large_row_gap(self) -> None:
        """行ギャップが modal_row_height * 2 を超える場合 BLANK が挿入される。"""
        blocks = [
            _make_block(1, text="First"),
            _make_block(50, text="Second"),  # 大きな行ギャップ
        ]
        grid = _make_grid(modal_row_height=15.0)
        elements = detect(blocks, grid)
        types = [e.element_type for e in elements]
        assert ElementType.BLANK in types

    def test_blank_inserted_on_bg_color_change(self) -> None:
        """背景色が変わる場合 BLANK が挿入される。"""
        blocks = [
            _make_block(1, bg_color="FFFF0000", text="Red section"),
            _make_block(2, bg_color="FF00FF00", text="Green section"),
        ]
        grid = _make_grid(modal_row_height=15.0)
        elements = detect(blocks, grid)
        types = [e.element_type for e in elements]
        assert ElementType.BLANK in types

    def test_sorted_by_source_row(self) -> None:
        blocks = [
            _make_block(5, text="Last"),
            _make_block(1, text="First"),
            _make_block(3, text="Middle"),
        ]
        elements = detect(blocks, _make_grid())
        rows = [e.source_row for e in elements if e.element_type != ElementType.BLANK]
        assert rows == sorted(rows)

    def test_comment_transferred_to_element(self) -> None:
        blocks = [_make_block(1, text="With comment", has_comment=True, comment_text="注釈")]
        elements = detect(blocks, _make_grid())
        assert elements[0].comment_text == "注釈"

    def test_numbered_list_detected(self) -> None:
        blocks = [
            _make_block(1, col=1, text="見出し", bold=True),
            _make_block(2, col=7, text="1. 最初のステップ"),
        ]
        grid = _make_grid(col_unit=3.0)
        elements = detect(blocks, grid)
        list_items = [e for e in elements if e.element_type == ElementType.LIST_ITEM]
        assert list_items[0].is_numbered_list is True
