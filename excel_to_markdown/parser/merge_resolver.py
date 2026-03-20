"""list[RawCell] を list[TextBlock] に変換する。

依存可能: models.py
依存禁止: reader/, renderer/, cli.py
"""

from __future__ import annotations

from excel_to_markdown.models import InlineRun, RawCell, TextBlock

# openpyxl のリッチテキスト型（実行時に resolve するため遅延インポート）
try:
    from openpyxl.cell.rich_text import CellRichText as _CellRichText
    from openpyxl.cell.rich_text import TextBlock as _OxlTextBlock
except ImportError:  # pragma: no cover
    _CellRichText = None  # type: ignore[assignment,misc]
    _OxlTextBlock = None  # type: ignore[assignment,misc]


def resolve(cells: list[RawCell]) -> list[TextBlock]:
    """非空の RawCell を TextBlock に変換し、(top_row, left_col) でソートして返す。

    - 値が None または空白のみのセルはスキップ
    - リッチテキスト（部分書式）は InlineRun リストとして保持
    - indent_level は 0 で初期化（後で structure_detector が更新）
    - セル内改行（\\n）はそのまま text に保持する（Markdown変換は renderer が担当）
    """
    blocks: list[TextBlock] = []

    for cell in cells:
        if cell.value is None:
            continue
        text = cell.value.strip()
        if not text:
            continue

        # セル全体書式を使った InlineRun を生成（リッチテキスト判定はreaderが担当済み）
        # ここでは value が既に str として渡されているため、inline_runs は空リストのみ
        # （リッチテキストは xlsx_reader で文字列化済み）
        inline_runs: list[InlineRun] = []

        blocks.append(
            TextBlock(
                text=text,
                top_row=cell.row,
                left_col=cell.col,
                bottom_row=cell.row + cell.merge_row_span - 1,
                right_col=cell.col + cell.merge_col_span - 1,
                row_span=cell.merge_row_span,
                col_span=cell.merge_col_span,
                font_bold=cell.font_bold,
                font_italic=cell.font_italic,
                font_strikethrough=cell.font_strikethrough,
                font_underline=cell.font_underline,
                font_size=cell.font_size,
                bg_color=cell.bg_color,
                has_comment=cell.has_comment,
                comment_text=cell.comment_text,
                indent_level=0,
                inline_runs=inline_runs,
            )
        )

    blocks.sort(key=lambda b: (b.top_row, b.left_col))
    return blocks


def to_inline_runs(value: object) -> list[InlineRun]:
    """openpyxl のリッチテキスト (_CellRichText) を InlineRun リストに変換する。

    リッチテキストでない場合は空リストを返す。
    """
    if _CellRichText is None or not isinstance(value, _CellRichText):
        return []

    runs: list[InlineRun] = []
    for part in value:
        if isinstance(part, str):
            if part:
                runs.append(InlineRun(text=part))
        elif _OxlTextBlock is not None and isinstance(part, _OxlTextBlock):
            text = str(part.text) if part.text else ""
            if not text:
                continue
            font = part.font
            bold = bool(font.bold) if font and font.bold is not None else False
            italic = bool(font.italic) if font and font.italic is not None else False
            strike = bool(font.strike) if font and font.strike is not None else False
            underline = (
                bool(font.underline) and font.underline != "none"
                if font and font.underline is not None
                else False
            )
            runs.append(
                InlineRun(
                    text=text,
                    bold=bold,
                    italic=italic,
                    strikethrough=strike,
                    underline=underline,
                )
            )

    return runs
