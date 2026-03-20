"""xls_reader.py のユニットテスト。

xlwt で .xls ファイルを生成して xlrd で読み込み、read_sheet_xls の出力を検証する。
"""

from __future__ import annotations

from pathlib import Path

import pytest
import xlrd

xlwt = pytest.importorskip("xlwt")

from excel_to_markdown.reader.xls_reader import read_sheet_xls


def _open_sheet(path: Path) -> tuple[xlrd.sheet.Sheet, xlrd.Book]:
    book = xlrd.open_workbook(str(path), formatting_info=True)
    sheet = book.sheet_by_index(0)
    return sheet, book


class TestReadSheetXlsBasic:
    def test_text_cell(self, tmp_path: Path) -> None:
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        ws.write(0, 0, "テキスト値")
        path = tmp_path / "test.xls"
        wb.save(str(path))

        sheet, book = _open_sheet(path)
        cells = read_sheet_xls(sheet, book)
        values = [c.value for c in cells if c.value]
        assert "テキスト値" in values

    def test_number_cell(self, tmp_path: Path) -> None:
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        ws.write(0, 0, 42)
        path = tmp_path / "num.xls"
        wb.save(str(path))

        sheet, book = _open_sheet(path)
        cells = read_sheet_xls(sheet, book)
        values = [c.value for c in cells if c.value]
        assert "42" in values

    def test_float_number(self, tmp_path: Path) -> None:
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        ws.write(0, 0, 3.14)
        path = tmp_path / "float.xls"
        wb.save(str(path))

        sheet, book = _open_sheet(path)
        cells = read_sheet_xls(sheet, book)
        values = [c.value for c in cells if c.value]
        assert any("3.14" in v for v in values)

    def test_empty_cell_skipped(self, tmp_path: Path) -> None:
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        ws.write(0, 0, "")
        ws.write(0, 1, "Keep")
        path = tmp_path / "empty.xls"
        wb.save(str(path))

        sheet, book = _open_sheet(path)
        cells = read_sheet_xls(sheet, book)
        values = [c.value for c in cells if c.value]
        assert "Keep" in values

    def test_1based_row_col(self, tmp_path: Path) -> None:
        """row/col が 1-based で返ること。"""
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        ws.write(2, 3, "ポジション")  # 0-based (row=2, col=3) → 1-based (row=3, col=4)
        path = tmp_path / "pos.xls"
        wb.save(str(path))

        sheet, book = _open_sheet(path)
        cells = read_sheet_xls(sheet, book)
        target = next((c for c in cells if c.value == "ポジション"), None)
        assert target is not None
        assert target.row == 3
        assert target.col == 4

    def test_no_comment_or_hyperlink(self, tmp_path: Path) -> None:
        """has_comment=False, comment_text=None, hyperlink=None が返ること。"""
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        ws.write(0, 0, "テキスト")
        path = tmp_path / "noc.xls"
        wb.save(str(path))

        sheet, book = _open_sheet(path)
        cells = read_sheet_xls(sheet, book)
        cell = next(c for c in cells if c.value)
        assert cell.has_comment is False
        assert cell.comment_text is None
        assert cell.hyperlink is None


class TestReadSheetXlsMerge:
    def test_merged_cells(self, tmp_path: Path) -> None:
        """結合セルの起点と非起点が正しく分類されること。"""
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        ws.write_merge(0, 1, 0, 1, "結合")  # rows 0-1, cols 0-1 を結合
        path = tmp_path / "merge.xls"
        wb.save(str(path))

        sheet, book = _open_sheet(path)
        cells = read_sheet_xls(sheet, book)

        # 起点 (1, 1) には値あり
        origin = next((c for c in cells if c.row == 1 and c.col == 1), None)
        assert origin is not None
        assert origin.value == "結合"
        assert origin.is_merge_origin is True
        assert origin.merge_row_span == 2
        assert origin.merge_col_span == 2

        # 非起点セルは value=None
        non_origins = [c for c in cells if c.row == 1 and c.col == 2]
        assert non_origins
        assert non_origins[0].value is None
        assert non_origins[0].is_merge_origin is False


class TestReadSheetXlsFont:
    def test_bold_font(self, tmp_path: Path) -> None:
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        bold_style = xlwt.easyxf("font: bold true")
        ws.write(0, 0, "太字", bold_style)
        ws.write(1, 0, "通常")
        path = tmp_path / "bold.xls"
        wb.save(str(path))

        sheet, book = _open_sheet(path)
        cells = read_sheet_xls(sheet, book)
        bold_cell = next(c for c in cells if c.value == "太字")
        normal_cell = next(c for c in cells if c.value == "通常")
        assert bold_cell.font_bold is True
        assert normal_cell.font_bold is False

    def test_italic_font(self, tmp_path: Path) -> None:
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        italic_style = xlwt.easyxf("font: italic true")
        ws.write(0, 0, "斜体", italic_style)
        path = tmp_path / "italic.xls"
        wb.save(str(path))

        sheet, book = _open_sheet(path)
        cells = read_sheet_xls(sheet, book)
        cell = next(c for c in cells if c.value == "斜体")
        assert cell.font_italic is True

    def test_font_size(self, tmp_path: Path) -> None:
        """フォントサイズが pt 単位で返ること（xlrd の height は 1/20 pt）。"""
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        big_style = xlwt.easyxf("font: height 360")  # 360 / 20 = 18pt
        ws.write(0, 0, "大きいフォント", big_style)
        path = tmp_path / "size.xls"
        wb.save(str(path))

        sheet, book = _open_sheet(path)
        cells = read_sheet_xls(sheet, book)
        cell = next(c for c in cells if c.value == "大きいフォント")
        assert cell.font_size == pytest.approx(18.0)
