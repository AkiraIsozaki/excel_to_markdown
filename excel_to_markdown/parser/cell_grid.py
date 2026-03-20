"""RawCell リストから空間的な統計・クエリを提供する。

依存可能: models.py
依存禁止: reader/, renderer/, cli.py
"""

from __future__ import annotations

import statistics
from dataclasses import dataclass, field

from excel_to_markdown.models import RawCell

# 列幅・行高さが未設定の場合のデフォルト値
_DEFAULT_COL_WIDTH: float = 8.0
_DEFAULT_ROW_HEIGHT: float = 15.0


@dataclass
class CellGrid:
    """RawCell リストと列幅・行高さ情報を保持し、空間クエリを提供する。

    col_widths: 列番号 → 列幅 (Excel文字幅単位)
    row_heights: 行番号 → 行高さ (Excelポイント単位)
    """

    cells: list[RawCell]
    col_widths: dict[int, float] = field(default_factory=dict)
    row_heights: dict[int, float] = field(default_factory=dict)

    @property
    def baseline_col(self) -> int:
        """コンテンツが存在する最左の列番号。セルがなければ1を返す。"""
        cols = [cell.col for cell in self.cells if cell.value is not None]
        return min(cols) if cols else 1

    @property
    def col_unit(self) -> float:
        """列幅の中央値（方眼紙のグリッドピッチの推定値）。

        col_widths が空または全て0の場合はデフォルト値を返す。
        """
        widths = [w for w in self.col_widths.values() if w and w > 0]
        if not widths:
            return _DEFAULT_COL_WIDTH
        return statistics.median(widths)

    @property
    def modal_row_height(self) -> float:
        """行高さの最頻値（方眼紙の行ピッチの推定値）。

        row_heights が空の場合はデフォルト値を返す。
        """
        heights = [h for h in self.row_heights.values() if h and h > 0]
        if not heights:
            return _DEFAULT_ROW_HEIGHT
        try:
            return statistics.mode(heights)
        except statistics.StatisticsError:
            # 最頻値が複数ある場合は中央値にフォールバック
            return statistics.median(heights)

    def is_empty_row(self, row: int) -> bool:
        """指定行にテキストを含むセルが存在しない場合 True を返す。"""
        return not any(cell.row == row and cell.value is not None for cell in self.cells)
