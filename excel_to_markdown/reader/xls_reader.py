"""xlrd 2.x を使って .xls シートから list[RawCell] を抽出する。

依存可能: models.py, xlrd
依存禁止: parser/, renderer/, cli.py

xlrd 2.x は .xls 形式専用。行・列インデックスは 0-based のため 1-based に変換して返す。
コメント・ハイパーリンクは xlrd 2.x でサポートされないため None を返す。
"""

from __future__ import annotations

import xlrd
import xlrd.sheet

from excel_to_markdown.models import RawCell

# xlrd セル型定数
_XL_CELL_EMPTY = 0
_XL_CELL_TEXT = 1
_XL_CELL_NUMBER = 2
_XL_CELL_DATE = 3
_XL_CELL_BOOLEAN = 4
_XL_CELL_ERROR = 5
_XL_CELL_BLANK = 6

# ARGB 白色
_WHITE_ARGB = "FFFFFFFF"


def read_sheet_xls(sheet: xlrd.sheet.Sheet, book: xlrd.Book) -> list[RawCell]:
    """xlrd シートから RawCell リストを返す。xlsx_reader と同一の出力型。

    - 非表示行・列の除外: xlrd 2.x では行の表示/非表示情報にアクセスできないためスキップしない
    - 結合セルの非起点セルは value=None, is_merge_origin=False で記録
    - コメント・ハイパーリンクは未サポートのため None
    """
    # 結合セル情報を事前収集（xlrd の merged_cells は 0-based）
    merge_origins: dict[tuple[int, int], tuple[int, int]] = {}
    merge_non_origins: set[tuple[int, int]] = set()
    for rlo, rhi, clo, chi in sheet.merged_cells:
        # 1-based に変換
        origin = (rlo + 1, clo + 1)
        row_span = rhi - rlo
        col_span = chi - clo
        merge_origins[origin] = (row_span, col_span)
        for r in range(rlo, rhi):
            for c in range(clo, chi):
                pos = (r + 1, c + 1)
                if pos != origin:
                    merge_non_origins.add(pos)

    raw_cells: list[RawCell] = []

    for r0 in range(sheet.nrows):
        r = r0 + 1  # 1-based
        for c0 in range(sheet.ncols):
            c = c0 + 1  # 1-based
            pos = (r, c)

            if pos in merge_non_origins:
                raw_cells.append(
                    RawCell(
                        row=r,
                        col=c,
                        value=None,
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
                )
                continue

            value = _cell_value_to_str(sheet, book, r0, c0)
            font_bold, font_italic, font_strike, font_underline, font_size, font_color = (
                _extract_font_props(sheet, book, r0, c0)
            )
            bg_color = _extract_bg_color(sheet, book, r0, c0)
            row_span, col_span = merge_origins.get(pos, (1, 1))

            raw_cells.append(
                RawCell(
                    row=r,
                    col=c,
                    value=value,
                    font_bold=font_bold,
                    font_italic=font_italic,
                    font_strikethrough=font_strike,
                    font_underline=font_underline,
                    font_size=font_size,
                    font_color=font_color,
                    bg_color=bg_color,
                    is_merge_origin=pos in merge_origins,
                    merge_row_span=row_span,
                    merge_col_span=col_span,
                    has_comment=False,
                    comment_text=None,
                )
            )

    return raw_cells


def _cell_value_to_str(
    sheet: xlrd.sheet.Sheet, book: xlrd.Book, r: int, c: int
) -> str | None:
    """セルの値を文字列に変換する。空・エラーは None を返す。"""
    cell_type = sheet.cell_type(r, c)
    if cell_type in (_XL_CELL_EMPTY, _XL_CELL_BLANK):
        return None
    if cell_type == _XL_CELL_ERROR:
        return None

    val = sheet.cell_value(r, c)

    if cell_type == _XL_CELL_BOOLEAN:
        return "TRUE" if val else "FALSE"

    if cell_type == _XL_CELL_NUMBER:
        # 整数値は整数表記、小数はそのまま文字列化
        if isinstance(val, float) and val == int(val):
            text = str(int(val))
        else:
            text = str(val)
        return text if text.strip() else None

    if cell_type == _XL_CELL_DATE:
        # 日付は xlrd の datemode を使って変換
        try:
            dt = xlrd.xldate_as_datetime(val, book.datemode)
            text = dt.strftime("%Y-%m-%d %H:%M:%S").rstrip(" 00:00:00") or dt.strftime("%Y-%m-%d")
            return text
        except Exception:
            return str(val)

    text = str(val)
    return text if text.strip() else None


def _extract_font_props(
    sheet: xlrd.sheet.Sheet, book: xlrd.Book, r: int, c: int
) -> tuple[bool, bool, bool, bool, float | None, str | None]:
    """セルのフォントプロパティ (bold, italic, strike, underline, size, color) を返す。"""
    try:
        xf_index = sheet.cell_xf_index(r, c)
        xf = book.xf_list[xf_index]
        font = book.font_list[xf.font_index]
    except (IndexError, AttributeError):
        return False, False, False, False, None, None

    bold = bool(font.bold)
    italic = bool(font.italic)
    strike = bool(font.struck_out)
    underline = font.underline_type != 0  # 0=none, 1=single, etc.

    # xlrd のフォントサイズは 1/20 pt 単位
    size: float | None = None
    if font.height:
        size = font.height / 20.0

    color: str | None = None
    colour_index = font.colour_index
    if colour_index is not None and colour_index in book.colour_map:
        rgb = book.colour_map[colour_index]
        if rgb is not None:
            r_val, g_val, b_val = rgb
            color = f"FF{r_val:02X}{g_val:02X}{b_val:02X}"

    return bold, italic, strike, underline, size, color


def _extract_bg_color(
    sheet: xlrd.sheet.Sheet, book: xlrd.Book, r: int, c: int
) -> str | None:
    """セルの背景色を ARGB hex で返す。なければ None。"""
    try:
        xf_index = sheet.cell_xf_index(r, c)
        xf = book.xf_list[xf_index]
        bg = xf.background
    except (IndexError, AttributeError):
        return None

    colour_index = bg.pattern_colour_index
    if colour_index is None:
        return None
    if colour_index not in book.colour_map:
        return None

    rgb = book.colour_map[colour_index]
    if rgb is None:
        return None

    r_val, g_val, b_val = rgb
    argb = f"FF{r_val:02X}{g_val:02X}{b_val:02X}"
    # 白色は None 扱い
    return argb if argb != _WHITE_ARGB else None
