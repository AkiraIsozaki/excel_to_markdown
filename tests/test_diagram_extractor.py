"""detect_swim_lanes() のユニットテスト。"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import openpyxl
import pytest

from excel_to_markdown.drawing.extractor import detect_swim_lanes
from excel_to_markdown.models import DiagramConnector, DiagramShape

SAMPLE_GYOUMU = Path(__file__).parent / "e2e/fixtures/gyoumuflow_answer.xlsx"


def _make_shape(
    shape_id: int,
    left_col: int,
    right_col: int,
    top_row: int = 5,
    bottom_row: int = 7,
) -> DiagramShape:
    return DiagramShape(
        shape_id=shape_id,
        name=f"shape_{shape_id}",
        text=f"text_{shape_id}",
        shape_type="rect",
        left_col=left_col,
        top_row=top_row,
        right_col=right_col,
        bottom_row=bottom_row,
    )


def _make_connector(connector_id: int, start: int, end: int) -> DiagramConnector:
    return DiagramConnector(
        connector_id=connector_id,
        name=f"conn_{connector_id}",
        start_shape_id=start,
        end_shape_id=end,
        label="",
    )


def _make_ws_with_header_row(row_1based: int, cells: dict[int, str]) -> MagicMock:
    """指定行に複数セルを持つモックワークシートを生成する。

    cells: {col_1based: value}
    """
    ws = MagicMock()

    def fake_getitem(key: int) -> list[MagicMock]:
        if key == row_1based:
            result = []
            for col, val in sorted(cells.items()):
                cell = MagicMock()
                cell.value = val
                cell.column = col
                result.append(cell)
            return result
        return []

    ws.__getitem__ = MagicMock(side_effect=fake_getitem)
    return ws


class TestDetectSwimLanes:
    def test_detects_three_lanes_from_real_file(self) -> None:
        """実際の gyoumuflow_answer.xlsx から3レーンが検出されること。"""
        if not SAMPLE_GYOUMU.exists():
            pytest.skip("gyoumuflow_answer.xlsx が見つかりません")

        wb = openpyxl.load_workbook(str(SAMPLE_GYOUMU))
        ws = wb["複数部署"]

        from excel_to_markdown.drawing.extractor import extract_sheet_drawing_map
        drawing_map = extract_sheet_drawing_map(SAMPLE_GYOUMU)
        shapes, connectors = drawing_map["複数部署"]

        result = detect_swim_lanes(ws, shapes, connectors)
        assert result is not None
        assert len(result) == 3
        names = [r[0] for r in result]
        assert "お客様" in names
        assert "〇〇ガラス店" in names
        assert "問屋" in names

    def test_no_connectors_returns_none(self) -> None:
        """コネクタがない場合は None を返すこと。"""
        ws = MagicMock()
        shapes = [_make_shape(1, 0, 3)]
        connectors: list[DiagramConnector] = []
        result = detect_swim_lanes(ws, shapes, connectors)
        assert result is None

    def test_single_header_cell_returns_none(self) -> None:
        """ヘッダー行に1つしかセルがない場合は None を返すこと（スイムレーンとみなさない）。"""
        # shapes start at top_row=5 (0-based), header row = 5 (1-based)
        shapes = [_make_shape(1, 0, 3, top_row=5), _make_shape(2, 6, 9, top_row=5)]
        connectors = [_make_connector(10, 1, 2)]
        ws = _make_ws_with_header_row(5, {1: "単一レーン"})
        result = detect_swim_lanes(ws, shapes, connectors)
        assert result is None

    def test_two_header_cells_returns_two_lanes(self) -> None:
        """ヘッダー行に2つのセルがある場合は2レーンが返されること。"""
        # connected shapes at top_row=5 (0-based)
        shapes = [_make_shape(1, 0, 3, top_row=5), _make_shape(2, 8, 11, top_row=5)]
        connectors = [_make_connector(10, 1, 2)]
        ws = _make_ws_with_header_row(5, {1: "レーンA", 8: "レーンB"})
        result = detect_swim_lanes(ws, shapes, connectors)
        assert result is not None
        assert len(result) == 2
        assert result[0][0] == "レーンA"
        assert result[1][0] == "レーンB"
        # レーンAの開始: col 1 (1-based) → 0 (0-based)
        assert result[0][1] == 0
        # レーンAの終了: レーンB開始 col 8 (1-based) - 2 = 6 (0-based)
        assert result[0][2] == 6
        # レーンBの終了: 999 (最後のレーン)
        assert result[1][2] == 999

    def test_no_shapes_returns_none(self) -> None:
        """シェイプがない場合は None を返すこと。"""
        ws = MagicMock()
        shapes: list[DiagramShape] = []
        connectors = [_make_connector(10, 1, 2)]
        result = detect_swim_lanes(ws, shapes, connectors)
        assert result is None
