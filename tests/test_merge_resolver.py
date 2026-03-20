"""merge_resolver.py のユニットテスト。"""

from __future__ import annotations

from excel_to_markdown.models import RawCell
from excel_to_markdown.parser.merge_resolver import resolve, to_inline_runs


def _make_cell(
    row: int,
    col: int,
    value: str | None = "X",
    merge_row_span: int = 1,
    merge_col_span: int = 1,
    bold: bool = False,
    font_size: float | None = None,
    bg_color: str | None = None,
    has_comment: bool = False,
    comment_text: str | None = None,
) -> RawCell:
    return RawCell(
        row=row,
        col=col,
        value=value,
        font_bold=bold,
        font_italic=False,
        font_strikethrough=False,
        font_underline=False,
        font_size=font_size,
        font_color=None,
        bg_color=bg_color,
        is_merge_origin=merge_row_span > 1 or merge_col_span > 1,
        merge_row_span=merge_row_span,
        merge_col_span=merge_col_span,
        has_comment=has_comment,
        comment_text=comment_text,
    )


class TestResolveBasic:
    def test_returns_text_blocks_sorted(self) -> None:
        cells = [
            _make_cell(3, 1, "C"),
            _make_cell(1, 2, "A"),
            _make_cell(2, 1, "B"),
        ]
        blocks = resolve(cells)
        assert [b.text for b in blocks] == ["A", "B", "C"]

    def test_none_value_skipped(self) -> None:
        cells = [_make_cell(1, 1, None), _make_cell(1, 2, "Keep")]
        blocks = resolve(cells)
        assert len(blocks) == 1
        assert blocks[0].text == "Keep"

    def test_whitespace_only_skipped(self) -> None:
        cells = [_make_cell(1, 1, "   "), _make_cell(1, 2, "Keep")]
        blocks = resolve(cells)
        assert len(blocks) == 1
        assert blocks[0].text == "Keep"

    def test_text_stripped(self) -> None:
        cells = [_make_cell(1, 1, "  Hello  ")]
        blocks = resolve(cells)
        assert blocks[0].text == "Hello"


class TestResolveMergedCell:
    def test_merge_span_preserved(self) -> None:
        cells = [_make_cell(1, 1, "Merged", merge_row_span=3, merge_col_span=2)]
        blocks = resolve(cells)
        assert blocks[0].row_span == 3
        assert blocks[0].col_span == 2
        assert blocks[0].bottom_row == 3
        assert blocks[0].right_col == 2

    def test_single_cell_span_is_1(self) -> None:
        cells = [_make_cell(2, 3, "Normal")]
        blocks = resolve(cells)
        assert blocks[0].row_span == 1
        assert blocks[0].col_span == 1
        assert blocks[0].bottom_row == 2
        assert blocks[0].right_col == 3


class TestResolveAttributes:
    def test_font_attrs_transferred(self) -> None:
        cells = [_make_cell(1, 1, "Bold", bold=True, font_size=14.0)]
        blocks = resolve(cells)
        assert blocks[0].font_bold is True
        assert blocks[0].font_size == 14.0

    def test_bg_color_transferred(self) -> None:
        cells = [_make_cell(1, 1, "Colored", bg_color="FFFF0000")]
        blocks = resolve(cells)
        assert blocks[0].bg_color == "FFFF0000"

    def test_comment_transferred(self) -> None:
        cells = [_make_cell(1, 1, "WithComment", has_comment=True, comment_text="注釈")]
        blocks = resolve(cells)
        assert blocks[0].has_comment is True
        assert blocks[0].comment_text == "注釈"

    def test_indent_level_initialized_to_zero(self) -> None:
        cells = [_make_cell(1, 1, "X")]
        blocks = resolve(cells)
        assert blocks[0].indent_level == 0

    def test_inline_runs_empty_by_default(self) -> None:
        cells = [_make_cell(1, 1, "X")]
        blocks = resolve(cells)
        assert blocks[0].inline_runs == []


class TestResolveCellNewline:
    def test_newline_preserved_in_text(self) -> None:
        cells = [_make_cell(1, 1, "Line1\nLine2")]
        blocks = resolve(cells)
        assert blocks[0].text == "Line1\nLine2"


# ---------------------------------------------------------------------------
# to_inline_runs
# ---------------------------------------------------------------------------


class TestToInlineRuns:
    def test_plain_string_returns_empty(self) -> None:
        """通常の str 値はリッチテキストでないので空リストを返す。"""
        result = to_inline_runs("plain text")
        assert result == []

    def test_none_returns_empty(self) -> None:
        result = to_inline_runs(None)
        assert result == []

    def test_int_returns_empty(self) -> None:
        result = to_inline_runs(42)
        assert result == []

    def test_rich_text_plain_string_parts(self) -> None:
        """CellRichText の文字列パーツが InlineRun に変換される。"""
        try:
            from openpyxl.cell.rich_text import CellRichText
        except ImportError:
            pytest.skip("openpyxl rich text not available")
        rt = CellRichText("Hello", " World")
        result = to_inline_runs(rt)
        assert len(result) == 2
        assert result[0].text == "Hello"
        assert result[1].text == " World"
        assert result[0].bold is False
        assert result[0].italic is False

    def test_rich_text_empty_string_parts_skipped(self) -> None:
        """空文字列パーツはスキップされる。"""
        try:
            from openpyxl.cell.rich_text import CellRichText
        except ImportError:
            pytest.skip("openpyxl rich text not available")
        rt = CellRichText("", "Keep")
        result = to_inline_runs(rt)
        assert len(result) == 1
        assert result[0].text == "Keep"

    def test_rich_text_textblock_with_bold(self) -> None:
        """CellRichText の TextBlock パーツで bold が正しく変換される。"""
        try:
            from openpyxl.cell.rich_text import CellRichText
            from openpyxl.cell.rich_text import TextBlock as OxlTextBlock
            from openpyxl.cell.text import InlineFont
        except ImportError:
            pytest.skip("openpyxl rich text not available")
        rt = CellRichText(OxlTextBlock(InlineFont(b=True), "BoldText"))
        result = to_inline_runs(rt)
        assert len(result) == 1
        assert result[0].text == "BoldText"
        assert result[0].bold is True
        assert result[0].italic is False

    def test_rich_text_textblock_with_italic(self) -> None:
        try:
            from openpyxl.cell.rich_text import CellRichText
            from openpyxl.cell.rich_text import TextBlock as OxlTextBlock
            from openpyxl.cell.text import InlineFont
        except ImportError:
            pytest.skip("openpyxl rich text not available")
        rt = CellRichText(OxlTextBlock(InlineFont(i=True), "ItalicText"))
        result = to_inline_runs(rt)
        assert result[0].italic is True
        assert result[0].bold is False

    def test_rich_text_textblock_with_strikethrough(self) -> None:
        try:
            from openpyxl.cell.rich_text import CellRichText
            from openpyxl.cell.rich_text import TextBlock as OxlTextBlock
            from openpyxl.cell.text import InlineFont
        except ImportError:
            pytest.skip("openpyxl rich text not available")
        rt = CellRichText(OxlTextBlock(InlineFont(strike=True), "StrikeText"))
        result = to_inline_runs(rt)
        assert result[0].strikethrough is True

    def test_rich_text_textblock_with_underline(self) -> None:
        try:
            from openpyxl.cell.rich_text import CellRichText
            from openpyxl.cell.rich_text import TextBlock as OxlTextBlock
            from openpyxl.cell.text import InlineFont
        except ImportError:
            pytest.skip("openpyxl rich text not available")
        rt = CellRichText(OxlTextBlock(InlineFont(u="single"), "UnderText"))
        result = to_inline_runs(rt)
        assert result[0].underline is True

    def test_rich_text_textblock_empty_text_skipped(self) -> None:
        """TextBlock で text が空のものはスキップされる。"""
        try:
            from openpyxl.cell.rich_text import CellRichText
            from openpyxl.cell.rich_text import TextBlock as OxlTextBlock
            from openpyxl.cell.text import InlineFont
        except ImportError:
            pytest.skip("openpyxl rich text not available")
        rt = CellRichText(OxlTextBlock(InlineFont(b=True), ""), "Keep")
        result = to_inline_runs(rt)
        # 空 TextBlock はスキップ、文字列 "Keep" は変換
        texts = [r.text for r in result]
        assert "Keep" in texts
        assert "" not in texts
