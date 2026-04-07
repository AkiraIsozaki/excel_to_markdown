"""DrawingML図形・コネクタ抽出モジュール。

xlsxファイル内の xl/drawings/drawing*.xml をパースし、
DiagramShape / DiagramConnector のリストを返す。

依存: 標準ライブラリのみ（zipfile, xml.etree.ElementTree）
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from excel_to_markdown.models import DiagramConnector, DiagramShape

# ---------------------------------------------------------------------------
# DrawingML 名前空間
# ---------------------------------------------------------------------------
_NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# drawing XML のパスパターン
_DRAWING_PATTERN = re.compile(r"^xl/drawings/drawing\d+\.xml$")

# OPC Relationships の名前空間
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_WB_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_DRAWING_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
_WORKSHEET_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)


def extract_sheet_drawing_map(
    xlsx_path: Path,
) -> dict[str, tuple[list[DiagramShape], list[DiagramConnector]]]:
    """シート名 → (shapes, connectors) のマッピングを返す。

    drawingを持たないシートはマップに含まれない。
    """
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        all_names = set(zf.namelist())

        # --- workbook.xml.rels: rId → worksheet ファイル名 ---
        rid_to_sheet_file: dict[str, str] = {}
        wb_rels_path = "xl/_rels/workbook.xml.rels"
        if wb_rels_path in all_names:
            rels_root = ET.fromstring(zf.read(wb_rels_path))
            for rel in rels_root:
                if rel.get("Type") == _WORKSHEET_REL_TYPE:
                    rid = rel.get("Id", "")
                    target = rel.get("Target", "")
                    # Target の形式:
                    #   "worksheets/sheetN.xml"         (相対)
                    #   "/xl/worksheets/sheetN.xml"      (絶対, openpyxl edge case)
                    # → ファイル名だけ取り出す
                    sheet_file = target.split("/")[-1]
                    rid_to_sheet_file[rid] = sheet_file

        # --- workbook.xml: シート名 → rId ---
        sheet_name_to_file: dict[str, str] = {}
        wb_path = "xl/workbook.xml"
        if wb_path in all_names:
            wb_root = ET.fromstring(zf.read(wb_path))
            for sh in wb_root.iter(f"{{{_WB_NS}}}sheet"):
                rid = sh.get(f"{{{_R_NS}}}id", "")
                name = sh.get("name", "")
                if rid in rid_to_sheet_file:
                    sheet_name_to_file[name] = rid_to_sheet_file[rid]

        # --- 各シートの _rels: sheet file → drawing path ---
        sheet_file_to_drawing: dict[str, str] = {}
        for sheet_file in rid_to_sheet_file.values():
            rels_path = f"xl/worksheets/_rels/{sheet_file}.rels"
            if rels_path not in all_names:
                continue
            rels_root = ET.fromstring(zf.read(rels_path))
            for rel in rels_root:
                if rel.get("Type") == _DRAWING_REL_TYPE:
                    # Target: "../drawings/drawingN.xml" → "xl/drawings/drawingN.xml"
                    target = rel.get("Target", "")
                    drawing_path = "xl/" + target.lstrip("../")
                    sheet_file_to_drawing[sheet_file] = drawing_path
                    break  # 1シートに1drawingと仮定

        # --- drawing XML をパースしてシート名にマップ ---
        result: dict[str, tuple[list[DiagramShape], list[DiagramConnector]]] = {}
        drawing_cache: dict[str, tuple[list[DiagramShape], list[DiagramConnector]]] = {}

        for sheet_name, sheet_file in sheet_name_to_file.items():
            drawing_path = sheet_file_to_drawing.get(sheet_file)
            if drawing_path is None or drawing_path not in all_names:
                continue
            if drawing_path not in drawing_cache:
                xml_bytes = zf.read(drawing_path)
                drawing_cache[drawing_path] = _parse_drawing_xml(xml_bytes)
            result[sheet_name] = drawing_cache[drawing_path]

    return result


def extract_diagrams(
    xlsx_path: Path,
) -> list[tuple[list[DiagramShape], list[DiagramConnector]]]:
    """xlsxの全drawingファイルから図形・コネクタを抽出する。

    Returns:
        drawingファイルごとのタプル (shapes, connectors) のリスト。
        drawingが存在しない場合は空リストを返す。
    """
    results: list[tuple[list[DiagramShape], list[DiagramConnector]]] = []

    with zipfile.ZipFile(xlsx_path, "r") as zf:
        drawing_paths = [
            name for name in zf.namelist() if _DRAWING_PATTERN.match(name)
        ]
        for path in sorted(drawing_paths):
            xml_bytes = zf.read(path)
            shapes, connectors = _parse_drawing_xml(xml_bytes)
            results.append((shapes, connectors))

    return results


def _parse_drawing_xml(
    xml_bytes: bytes,
) -> tuple[list[DiagramShape], list[DiagramConnector]]:
    """drawing XML をパースして DiagramShape / DiagramConnector を返す。"""
    root = ET.fromstring(xml_bytes)  # noqa: S314 - ローカルファイルのみ処理

    shapes: list[DiagramShape] = []
    connectors: list[DiagramConnector] = []

    # TwoCellAnchor / OneCellAnchor / AbsoluteAnchor を統一処理
    anchor_tags = {
        f"{{{_NS['xdr']}}}twoCellAnchor",
        f"{{{_NS['xdr']}}}oneCellAnchor",
        f"{{{_NS['xdr']}}}absoluteAnchor",
    }
    for anchor in root:
        if anchor.tag not in anchor_tags:
            continue

        from_col, from_row = _extract_from(anchor)
        to_col, to_row = _extract_to(anchor, from_col, from_row)

        # 図形
        sp = anchor.find("xdr:sp", _NS)
        if sp is not None:
            shape = _parse_shape(sp, from_col, from_row, to_col, to_row)
            if shape is not None:
                shapes.append(shape)
            continue

        # コネクタ
        cxn_sp = anchor.find("xdr:cxnSp", _NS)
        if cxn_sp is not None:
            connector = _parse_connector(cxn_sp)
            if connector is not None:
                connectors.append(connector)

    return shapes, connectors


def _extract_from(anchor: ET.Element) -> tuple[int, int]:
    """アンカー要素から from の col/row を取得する（0-based）。"""
    from_elem = anchor.find("xdr:from", _NS)
    if from_elem is None:
        return 0, 0
    col = int(from_elem.findtext("xdr:col", "0", _NS))
    row = int(from_elem.findtext("xdr:row", "0", _NS))
    return col, row


def _extract_to(
    anchor: ET.Element, from_col: int, from_row: int
) -> tuple[int, int]:
    """アンカー要素から to の col/row を取得する。

    oneCellAnchor / absoluteAnchor は to がないため from + ext で近似。
    """
    to_elem = anchor.find("xdr:to", _NS)
    if to_elem is not None:
        col = int(to_elem.findtext("xdr:col", str(from_col + 2), _NS))
        row = int(to_elem.findtext("xdr:row", str(from_row + 2), _NS))
        return col, row
    return from_col + 2, from_row + 2


def _extract_text(sp: ET.Element) -> str:
    """<xdr:txBody> から全テキストを結合して返す。"""
    tx_body = sp.find("xdr:txBody", _NS)
    if tx_body is None:
        return ""
    parts: list[str] = []
    for t in tx_body.iter(f"{{{_NS['a']}}}t"):
        if t.text:
            parts.append(t.text)
    return "".join(parts).strip()


def _parse_shape(
    sp: ET.Element,
    from_col: int,
    from_row: int,
    to_col: int,
    to_row: int,
) -> DiagramShape | None:
    """<xdr:sp> 要素を DiagramShape に変換する。"""
    nv_sp_pr = sp.find("xdr:nvSpPr", _NS)
    if nv_sp_pr is None:
        return None

    cnv_pr = nv_sp_pr.find("xdr:cNvPr", _NS)
    if cnv_pr is None:
        return None

    shape_id = int(cnv_pr.get("id", "0"))
    name = cnv_pr.get("name", "")

    # 形状タイプ: spPr/prstGeom[@prst]
    sp_pr = sp.find("xdr:spPr", _NS)
    shape_type = "rect"
    if sp_pr is not None:
        prst_geom = sp_pr.find("a:prstGeom", _NS)
        if prst_geom is not None:
            shape_type = prst_geom.get("prst", "rect")

    text = _extract_text(sp)

    return DiagramShape(
        shape_id=shape_id,
        name=name,
        text=text,
        shape_type=shape_type,
        left_col=from_col,
        top_row=from_row,
        right_col=to_col,
        bottom_row=to_row,
    )


def _parse_connector(cxn_sp: ET.Element) -> DiagramConnector | None:
    """<xdr:cxnSp> 要素を DiagramConnector に変換する。"""
    nv_cxn_sp_pr = cxn_sp.find("xdr:nvCxnSpPr", _NS)
    if nv_cxn_sp_pr is None:
        return None

    cnv_pr = nv_cxn_sp_pr.find("xdr:cNvPr", _NS)
    if cnv_pr is None:
        return None

    connector_id = int(cnv_pr.get("id", "0"))
    name = cnv_pr.get("name", "")

    # stCxn / endCxn
    cnv_cxn_sp_pr = nv_cxn_sp_pr.find("xdr:cNvCxnSpPr", _NS)
    start_shape_id: int | None = None
    end_shape_id: int | None = None
    if cnv_cxn_sp_pr is not None:
        st = cnv_cxn_sp_pr.find("a:stCxn", _NS)
        if st is not None:
            start_shape_id = int(st.get("id", "0"))
        end = cnv_cxn_sp_pr.find("a:endCxn", _NS)
        if end is not None:
            end_shape_id = int(end.get("id", "0"))

    label = _extract_text(cxn_sp)

    return DiagramConnector(
        connector_id=connector_id,
        name=name,
        start_shape_id=start_shape_id,
        end_shape_id=end_shape_id,
        label=label,
    )
