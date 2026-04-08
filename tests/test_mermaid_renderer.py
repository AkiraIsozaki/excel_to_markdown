"""Mermaidレンダラーのユニットテスト。"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_to_markdown.drawing.extractor import extract_diagrams
from excel_to_markdown.models import DiagramConnector, DiagramShape
from excel_to_markdown.renderer.mermaid_renderer import (
    _detect_direction,
    _node_notation,
    render_mermaid,
    render_mermaid_block,
)

SAMPLE_FLOWCHART = Path(__file__).parent / "e2e/fixtures/sample_flowchart.xlsx"


# ---------------------------------------------------------------------------
# ノード記法マッピングのテスト
# ---------------------------------------------------------------------------

class TestNodeNotation:
    def test_flowchart_terminator(self) -> None:
        assert _node_notation("flowChartTerminator", "開始") == "([開始])"

    def test_flowchart_process(self) -> None:
        assert _node_notation("flowChartProcess", "受注") == "[受注]"

    def test_flowchart_decision(self) -> None:
        result = _node_notation("flowChartDecision", "確認?")
        assert result == "{確認?}"  # Mermaid ひし形記法: {text}

    def test_rect(self) -> None:
        assert _node_notation("rect", "処理") == "[処理]"

    def test_ellipse(self) -> None:
        assert _node_notation("ellipse", "A") == "((A))"

    def test_diamond(self) -> None:
        result = _node_notation("diamond", "判断")
        assert result == "{判断}"

    def test_unknown_type_defaults_to_rect(self) -> None:
        assert _node_notation("unknownShapeType", "テキスト") == "[テキスト]"

    def test_double_quote_escaped(self) -> None:
        result = _node_notation("flowChartProcess", 'テスト"引用')
        assert '"' not in result or "#quot;" in result


# ---------------------------------------------------------------------------
# グラフ方向判定のテスト
# ---------------------------------------------------------------------------

class TestDetectDirection:
    def _shape(self, top_row: int, bottom_row: int, left_col: int, right_col: int) -> DiagramShape:
        return DiagramShape(
            shape_id=1, name="", text="", shape_type="rect",
            left_col=left_col, top_row=top_row, right_col=right_col, bottom_row=bottom_row,
        )

    def test_empty_shapes_returns_td(self) -> None:
        assert _detect_direction([]) == "TD"

    def test_tall_layout_returns_td(self) -> None:
        # 縦長レイアウト: height=20, width=4
        shapes = [
            self._shape(0, 10, 0, 2),
            self._shape(10, 20, 0, 2),
        ]
        assert _detect_direction(shapes) == "TD"

    def test_wide_layout_returns_lr(self) -> None:
        # 横長レイアウト: height=2, width=20
        shapes = [
            self._shape(0, 1, 0, 10),
            self._shape(0, 1, 10, 20),
        ]
        assert _detect_direction(shapes) == "LR"


# ---------------------------------------------------------------------------
# render_mermaid のテスト
# ---------------------------------------------------------------------------

class TestRenderMermaid:
    def _make_shape(self, shape_id: int, text: str, shape_type: str = "flowChartProcess",
                    top_row: int = 0, bottom_row: int = 2) -> DiagramShape:
        return DiagramShape(
            shape_id=shape_id, name=text, text=text, shape_type=shape_type,
            left_col=0, top_row=top_row, right_col=2, bottom_row=bottom_row,
        )

    def _make_connector(self, cid: int, src: int, dst: int, label: str = "") -> DiagramConnector:
        return DiagramConnector(
            connector_id=cid, name="", start_shape_id=src, end_shape_id=dst, label=label,
        )

    def test_empty_shapes_returns_header_only(self) -> None:
        result = render_mermaid([], [])
        assert result.strip() == "flowchart TD"

    def test_single_shape(self) -> None:
        shapes = [self._make_shape(1, "処理A")]
        result = render_mermaid(shapes, [])
        assert "flowchart" in result
        assert "N1[処理A]" in result

    def test_connector_edge(self) -> None:
        shapes = [self._make_shape(1, "A"), self._make_shape(2, "B", top_row=3, bottom_row=5)]
        connectors = [self._make_connector(10, 1, 2)]
        result = render_mermaid(shapes, connectors)
        assert "N1 --> N2" in result

    def test_connector_edge_with_label(self) -> None:
        shapes = [self._make_shape(1, "判断", "flowChartDecision"),
                  self._make_shape(2, "処理")]
        connectors = [self._make_connector(10, 1, 2, label="はい")]
        result = render_mermaid(shapes, connectors)
        assert "N1 -->|はい| N2" in result

    def test_connector_with_unknown_shape_id_skipped(self) -> None:
        shapes = [self._make_shape(1, "A")]
        connectors = [self._make_connector(10, 1, 999)]  # 999は存在しない
        result = render_mermaid(shapes, connectors)
        assert "N999" not in result

    def test_direction_override(self) -> None:
        shapes = [self._make_shape(1, "A")]
        result = render_mermaid(shapes, [], direction="LR")
        assert result.startswith("flowchart LR")

    def test_terminator_node_notation(self) -> None:
        shapes = [self._make_shape(2, "開始", "flowChartTerminator")]
        result = render_mermaid(shapes, [])
        assert "N2([開始])" in result

    def test_decision_node_notation(self) -> None:
        shapes = [self._make_shape(4, "条件?", "flowChartDecision")]
        result = render_mermaid(shapes, [])
        assert "N4{条件?}" in result

    def test_connector_none_endpoints_skipped(self) -> None:
        shapes = [self._make_shape(1, "A")]
        connectors = [DiagramConnector(
            connector_id=1, name="", start_shape_id=None, end_shape_id=None, label="",
        )]
        result = render_mermaid(shapes, connectors)
        assert "-->" not in result


# ---------------------------------------------------------------------------
# render_mermaid_block のテスト
# ---------------------------------------------------------------------------

class TestRenderMermaidBlock:
    def test_wraps_in_code_block(self) -> None:
        shapes = [DiagramShape(
            shape_id=1, name="A", text="処理",
            shape_type="flowChartProcess",
            left_col=0, top_row=0, right_col=2, bottom_row=2,
        )]
        result = render_mermaid_block(shapes, [])
        assert result.startswith("```mermaid\n")
        assert result.endswith("```\n")


# ---------------------------------------------------------------------------
# サンプルxlsxからの統合テスト
# ---------------------------------------------------------------------------

class TestSampleFlowchartMermaid:
    @pytest.fixture(autouse=True)
    def setup(self) -> None:
        if not SAMPLE_FLOWCHART.exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_FLOWCHART}")
        results = extract_diagrams(SAMPLE_FLOWCHART)
        self.shapes, self.connectors = results[0]

    def test_mermaid_contains_all_nodes(self) -> None:
        result = render_mermaid(self.shapes, self.connectors)
        assert "開始" in result
        assert "完了" in result
        assert "受注受付" in result
        assert "在庫確認?" in result
        assert "出荷処理" in result
        assert "発注処理" in result
        assert "入荷待ち" in result

    def test_mermaid_contains_labels(self) -> None:
        result = render_mermaid(self.shapes, self.connectors)
        assert "はい" in result
        assert "いいえ" in result

    def test_mermaid_direction_is_td(self) -> None:
        result = render_mermaid(self.shapes, self.connectors)
        assert "flowchart TD" in result

    def test_mermaid_edge_count(self) -> None:
        result = render_mermaid(self.shapes, self.connectors)
        edge_count = result.count("-->")
        assert edge_count == 7

    def test_mermaid_block_format(self) -> None:
        result = render_mermaid_block(self.shapes, self.connectors)
        assert result.startswith("```mermaid\n")
        assert result.endswith("```\n")


# ---------------------------------------------------------------------------
# スイムレーン (subgraph) のテスト
# ---------------------------------------------------------------------------

class TestSwimLanes:
    """swim_lanes パラメータによる subgraph 出力のテスト。"""

    def _make_shape(
        self,
        shape_id: int,
        text: str,
        left_col: int,
        right_col: int,
    ) -> DiagramShape:
        return DiagramShape(
            shape_id=shape_id,
            name=f"shape_{shape_id}",
            text=text,
            shape_type="rect",
            left_col=left_col,
            top_row=0,
            right_col=right_col,
            bottom_row=2,
        )

    def _make_connector(
        self, connector_id: int, start: int, end: int
    ) -> DiagramConnector:
        return DiagramConnector(
            connector_id=connector_id,
            name=f"conn_{connector_id}",
            start_shape_id=start,
            end_shape_id=end,
            label="",
        )

    def test_swim_lanes_produce_subgraph(self) -> None:
        """swim_lanes あり → subgraph ブロックが含まれること。"""
        shapes = [
            self._make_shape(1, "受付", 0, 3),
            self._make_shape(2, "処理", 6, 9),
        ]
        connectors = [self._make_connector(10, 1, 2)]
        swim_lanes = [("部署A", 0, 5), ("部署B", 6, 11)]

        result = render_mermaid(shapes, connectors, swim_lanes=swim_lanes)
        assert "subgraph lane_0" in result
        assert "subgraph lane_1" in result
        assert '"部署A"' in result
        assert '"部署B"' in result
        assert "end" in result

    def test_swim_lanes_nodes_assigned_correctly(self) -> None:
        """各ノードが正しいスイムレーンに配置されること。"""
        shapes = [
            self._make_shape(1, "受付", 0, 3),   # center=1.5 → lane_0 (0-5)
            self._make_shape(2, "処理", 6, 9),   # center=7.5 → lane_1 (6-11)
        ]
        connectors = [self._make_connector(10, 1, 2)]
        swim_lanes = [("部署A", 0, 5), ("部署B", 6, 11)]

        result = render_mermaid(shapes, connectors, swim_lanes=swim_lanes)
        lines = result.split("\n")

        # subgraph lane_0 内に N1 があること
        in_lane_0 = False
        lane_0_has_n1 = False
        lane_1_has_n2 = False
        in_lane_1 = False
        for line in lines:
            stripped = line.strip()
            if "subgraph lane_0" in stripped:
                in_lane_0 = True
                in_lane_1 = False
            elif "subgraph lane_1" in stripped:
                in_lane_1 = True
                in_lane_0 = False
            elif stripped == "end":
                in_lane_0 = False
                in_lane_1 = False
            elif in_lane_0 and stripped.startswith("N1"):
                lane_0_has_n1 = True
            elif in_lane_1 and stripped.startswith("N2"):
                lane_1_has_n2 = True

        assert lane_0_has_n1, "N1 は lane_0 に配置されるべき"
        assert lane_1_has_n2, "N2 は lane_1 に配置されるべき"

    def test_swim_lanes_edges_outside_subgraph(self) -> None:
        """エッジはすべての subgraph の後に出力されること。"""
        shapes = [
            self._make_shape(1, "受付", 0, 3),
            self._make_shape(2, "処理", 6, 9),
        ]
        connectors = [self._make_connector(10, 1, 2)]
        swim_lanes = [("部署A", 0, 5), ("部署B", 6, 11)]

        result = render_mermaid(shapes, connectors, swim_lanes=swim_lanes)
        last_end_pos = result.rfind("    end")
        edge_pos = result.find("N1 --> N2")
        assert last_end_pos < edge_pos, "エッジは最後の subgraph end の後に出力されるべき"

    def test_no_swim_lanes_flat_output(self) -> None:
        """swim_lanes=None の場合は従来のフラット出力になること。"""
        shapes = [
            self._make_shape(1, "受付", 0, 3),
            self._make_shape(2, "処理", 6, 9),
        ]
        connectors = [self._make_connector(10, 1, 2)]

        result = render_mermaid(shapes, connectors, swim_lanes=None)
        assert "subgraph" not in result
        assert "N1[受付]" in result
        assert "N2[処理]" in result
