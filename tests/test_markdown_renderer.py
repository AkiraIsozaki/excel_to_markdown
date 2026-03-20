"""markdown_renderer.py のユニットテスト。"""

from __future__ import annotations

from excel_to_markdown.models import (
    DocElement,
    ElementType,
    InlineRun,
    TableCell,
    TableElement,
)
from excel_to_markdown.renderer.markdown_renderer import (
    apply_inline_format,
    collapse_blank_lines,
    convert_cell_newlines,
    render,
    render_element,
    render_inline,
)


def _el(
    etype: ElementType,
    text: str = "",
    level: int = 0,
    source_row: int = 1,
    is_numbered: bool = False,
    comment: str | None = None,
) -> DocElement:
    return DocElement(
        element_type=etype,
        text=text,
        level=level,
        source_row=source_row,
        is_numbered_list=is_numbered,
        comment_text=comment,
    )


def _table(rows_texts: list[list[str]], first_row_header: bool = False) -> TableElement:
    rows = []
    for ri, row in enumerate(rows_texts):
        cells = [
            TableCell(text=t, row=ri, col=ci, is_header=(first_row_header and ri == 0))
            for ci, t in enumerate(row)
        ]
        rows.append(cells)
    return TableElement(text="", level=0, source_row=1, rows=rows, col_count=len(rows[0]))


# ---------------------------------------------------------------------------
# convert_cell_newlines
# ---------------------------------------------------------------------------


class TestConvertCellNewlines:
    def test_newline_becomes_hard_break(self) -> None:
        assert convert_cell_newlines("A\nB") == "A  \nB"

    def test_no_newline_unchanged(self) -> None:
        assert convert_cell_newlines("Hello") == "Hello"

    def test_multiple_newlines(self) -> None:
        assert convert_cell_newlines("A\nB\nC") == "A  \nB  \nC"


# ---------------------------------------------------------------------------
# collapse_blank_lines
# ---------------------------------------------------------------------------


class TestCollapseBlankLines:
    def test_3_blanks_collapsed_to_2(self) -> None:
        assert collapse_blank_lines("A\n\n\nB") == "A\n\nB"

    def test_4_blanks_collapsed_to_2(self) -> None:
        assert collapse_blank_lines("A\n\n\n\nB") == "A\n\nB"

    def test_2_blanks_unchanged(self) -> None:
        assert collapse_blank_lines("A\n\nB") == "A\n\nB"


# ---------------------------------------------------------------------------
# apply_inline_format
# ---------------------------------------------------------------------------


class TestApplyInlineFormat:
    def test_bold(self) -> None:
        assert apply_inline_format(InlineRun(text="X", bold=True)) == "**X**"

    def test_italic(self) -> None:
        assert apply_inline_format(InlineRun(text="X", italic=True)) == "*X*"

    def test_bold_italic(self) -> None:
        assert apply_inline_format(InlineRun(text="X", bold=True, italic=True)) == "***X***"

    def test_strikethrough(self) -> None:
        assert apply_inline_format(InlineRun(text="X", strikethrough=True)) == "~~X~~"

    def test_underline(self) -> None:
        assert apply_inline_format(InlineRun(text="X", underline=True)) == "<u>X</u>"

    def test_no_format(self) -> None:
        assert apply_inline_format(InlineRun(text="X")) == "X"

    def test_empty_text(self) -> None:
        assert apply_inline_format(InlineRun(text="", bold=True)) == ""


# ---------------------------------------------------------------------------
# render_inline
# ---------------------------------------------------------------------------


class TestRenderInline:
    def test_no_runs_returns_text(self) -> None:
        assert render_inline("Hello", []) == "Hello"

    def test_runs_applied(self) -> None:
        runs = [InlineRun(text="bold", bold=True), InlineRun(text=" plain")]
        result = render_inline("ignored", runs)
        assert result == "**bold** plain"


# ---------------------------------------------------------------------------
# render_element: 各 ElementType
# ---------------------------------------------------------------------------


class TestRenderElement:
    def test_heading_1(self) -> None:
        el = _el(ElementType.HEADING, text="見出し", level=1)
        md, ctr = render_element(el, 1)
        assert md == "# 見出し\n\n"
        assert ctr == 1

    def test_heading_3(self) -> None:
        el = _el(ElementType.HEADING, text="小見出し", level=3)
        md, _ = render_element(el, 1)
        assert md == "### 小見出し\n\n"

    def test_paragraph(self) -> None:
        el = _el(ElementType.PARAGRAPH, text="本文")
        md, _ = render_element(el, 1)
        assert md == "本文\n\n"

    def test_list_item_level1(self) -> None:
        el = _el(ElementType.LIST_ITEM, text="項目", level=1)
        md, _ = render_element(el, 1)
        assert md == "- 項目\n"

    def test_list_item_level2(self) -> None:
        el = _el(ElementType.LIST_ITEM, text="サブ項目", level=2)
        md, _ = render_element(el, 1)
        assert md == "  - サブ項目\n"

    def test_list_item_numbered(self) -> None:
        el = _el(ElementType.LIST_ITEM, text="ステップ", level=1, is_numbered=True)
        md, _ = render_element(el, 1)
        assert md == "1. ステップ\n"

    def test_blank(self) -> None:
        el = _el(ElementType.BLANK)
        md, _ = render_element(el, 1)
        assert md == "\n"

    def test_footnote_appended(self) -> None:
        el = _el(ElementType.PARAGRAPH, text="本文", comment="注釈")
        md, ctr = render_element(el, 1)
        assert "[^1]" in md
        assert ctr == 2

    def test_footnote_counter_increments(self) -> None:
        el = _el(ElementType.PARAGRAPH, text="X", comment="注")
        _, ctr = render_element(el, 3)
        assert ctr == 4


# ---------------------------------------------------------------------------
# render_element: TABLE
# ---------------------------------------------------------------------------


class TestRenderTable:
    def test_basic_gfm_table(self) -> None:
        table = _table([["名前", "年齢"], ["田中", "30"]], first_row_header=True)
        md, _ = render_element(table, 1)
        assert "| 名前 | 年齢 |" in md
        assert "| --- | --- |" in md
        assert "| 田中 | 30 |" in md

    def test_table_without_explicit_header(self) -> None:
        """ヘッダー行なし: 1行目をヘッダーとして出力し、セパレータを挿入。"""
        table = _table([["A", "B"], ["C", "D"]], first_row_header=False)
        md, _ = render_element(table, 1)
        lines = md.strip().split("\n")
        assert "---" in lines[1]

    def test_empty_table_returns_empty(self) -> None:
        table = TableElement(text="", level=0, source_row=1, rows=[], col_count=0)
        md, _ = render_element(table, 1)
        assert md == ""


# ---------------------------------------------------------------------------
# render: 統合
# ---------------------------------------------------------------------------


class TestRender:
    def test_render_heading_paragraph(self) -> None:
        elements = [
            _el(ElementType.HEADING, text="タイトル", level=1),
            _el(ElementType.PARAGRAPH, text="本文"),
        ]
        md = render(elements, [])
        assert "# タイトル" in md
        assert "本文" in md

    def test_footnotes_appended(self) -> None:
        elements = [_el(ElementType.PARAGRAPH, text="X", comment="注釈1")]
        md = render(elements, ["注釈1"])
        assert "[^1]: 注釈1" in md

    def test_blank_lines_collapsed(self) -> None:
        elements = [
            _el(ElementType.BLANK),
            _el(ElementType.BLANK),
            _el(ElementType.BLANK),
            _el(ElementType.PARAGRAPH, text="X"),
        ]
        md = render(elements, [])
        assert "\n\n\n" not in md

    def test_no_elements(self) -> None:
        assert render([], []) == ""
