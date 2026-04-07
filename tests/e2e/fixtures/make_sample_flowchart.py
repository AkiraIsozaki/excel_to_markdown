"""受注処理フロー図を含むサンプルxlsxを生成するスクリプト。

DrawingML (xl/drawings/drawing1.xml) に直接図形とコネクタを書き込む。
openpyxl はシェイプ追加APIを持たないため、ZipFileを直接操作する。

生成されるフロー:
    [開始] → [受注受付] → {在庫確認?} → はい → [出荷処理] → [完了]
                                        ↓ いいえ
                                     [発注処理] → [入荷待ち] → [出荷処理]
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import openpyxl

OUTPUT_PATH = Path(__file__).parent / "sample_flowchart.xlsx"

# ---------------------------------------------------------------------------
# DrawingML 名前空間
# ---------------------------------------------------------------------------
_NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# ---------------------------------------------------------------------------
# EMU ユーティリティ（1 cm = 360000 EMU）
# ---------------------------------------------------------------------------
CM = 360_000  # 1 cm in EMU


def _col_emu(col: int, width_cm: float) -> int:
    """列インデックス(0-based)をEMUで返す（簡易近似）。"""
    return int(col * width_cm * CM)


# ---------------------------------------------------------------------------
# セルアンカー XML ヘルパー
# ---------------------------------------------------------------------------

def _cell_anchor(col: int, row: int, col_off: int = 0, row_off: int = 0) -> str:
    return (
        f"<xdr:col>{col}</xdr:col>"
        f"<xdr:colOff>{col_off}</xdr:colOff>"
        f"<xdr:row>{row}</xdr:row>"
        f"<xdr:rowOff>{row_off}</xdr:rowOff>"
    )


def _sp(
    shape_id: int,
    name: str,
    text: str,
    prst: str,
    from_col: int,
    from_row: int,
    to_col: int,
    to_row: int,
) -> str:
    """<xdr:sp>（図形）のXML文字列を返す。"""
    return f"""  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>{_cell_anchor(from_col, from_row)}</xdr:from>
    <xdr:to>{_cell_anchor(to_col, to_row)}</xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="{shape_id}" name="{name}"/>
        <xdr:cNvSpPr><a:spLocks noGrp="1"/></xdr:cNvSpPr>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
        </a:xfrm>
        <a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        <a:ln><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>"""


def _cxn(
    connector_id: int,
    name: str,
    start_id: int,
    start_idx: int,
    end_id: int,
    end_idx: int,
    from_col: int,
    from_row: int,
    to_col: int,
    to_row: int,
    label: str = "",
) -> str:
    """<xdr:cxnSp>（コネクタ）のXML文字列を返す。"""
    label_xml = ""
    if label:
        label_xml = f"""
      <xdr:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:t>{label}</a:t></a:r></a:p>
      </xdr:txBody>"""
    return f"""  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>{_cell_anchor(from_col, from_row)}</xdr:from>
    <xdr:to>{_cell_anchor(to_col, to_row)}</xdr:to>
    <xdr:cxnSp macro="">
      <xdr:nvCxnSpPr>
        <xdr:cNvPr id="{connector_id}" name="{name}"/>
        <xdr:cNvCxnSpPr>
          <a:stCxn id="{start_id}" idx="{start_idx}"/>
          <a:endCxn id="{end_id}" idx="{end_idx}"/>
        </xdr:cNvCxnSpPr>
      </xdr:nvCxnSpPr>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
        </a:xfrm>
        <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
        <a:ln>
          <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
          <a:tailEnd type="arrow"/>
        </a:ln>
      </xdr:spPr>{label_xml}
    </xdr:cxnSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>"""


# ---------------------------------------------------------------------------
# connection point index の慣例:
#   0 = 左, 1 = 上, 2 = 右, 3 = 下
# ---------------------------------------------------------------------------

def build_drawing_xml() -> str:
    """受注処理フローのDrawingML XMLを返す。

    図形レイアウト（セル座標、0-based）:
      列: 0  1  2  3  4  5  6  7  8  9  10
      行: 0
          2  [開始 id=2]
          4  [受注受付 id=3]
          6  {在庫確認? id=4}
          8  [出荷処理 id=5]    [発注処理 id=6] (col=6)
          10 [完了 id=7]        [入荷待ち id=8] (col=6)

    コネクタ（id=10〜16）:
      10: 開始(2) → 受注受付(3)
      11: 受注受付(3) → 在庫確認(4)
      12: 在庫確認(4) → 出荷処理(5)  label="はい"
      13: 在庫確認(4) → 発注処理(6)  label="いいえ"
      14: 出荷処理(5) → 完了(7)
      15: 発注処理(6) → 入荷待ち(8)
      16: 入荷待ち(8) → 出荷処理(5)
    """
    shapes = [
        # (shape_id, name, text, prst, from_col, from_row, to_col, to_row)
        _sp(2,  "開始",      "開始",      "flowChartTerminator", 1, 0,  4, 1),
        _sp(3,  "受注受付",  "受注受付",  "flowChartProcess",    1, 2,  4, 3),
        _sp(4,  "在庫確認",  "在庫確認?", "flowChartDecision",   1, 4,  4, 6),
        _sp(5,  "出荷処理",  "出荷処理",  "flowChartProcess",    1, 8,  4, 9),
        _sp(7,  "完了",      "完了",      "flowChartTerminator", 1, 11, 4, 12),
        _sp(6,  "発注処理",  "発注処理",  "flowChartProcess",    6, 4,  9, 5),
        _sp(8,  "入荷待ち",  "入荷待ち",  "flowChartProcess",    6, 7,  9, 8),
    ]

    connectors = [
        # (id, name, start_id, start_idx, end_id, end_idx, from_col, from_row, to_col, to_row, label)
        _cxn(10, "コネクタ_開始→受注",     2, 3, 3, 1,  2, 1,  2, 2),
        _cxn(11, "コネクタ_受注→在庫",     3, 3, 4, 1,  2, 3,  2, 4),
        _cxn(12, "コネクタ_在庫→出荷",     4, 3, 5, 1,  2, 6,  2, 8,  "はい"),
        _cxn(13, "コネクタ_在庫→発注",     4, 2, 6, 1,  4, 5,  6, 5,  "いいえ"),
        _cxn(14, "コネクタ_出荷→完了",     5, 3, 7, 1,  2, 9,  2, 11),
        _cxn(15, "コネクタ_発注→入荷待ち", 6, 3, 8, 1,  7, 5,  7, 7),
        _cxn(16, "コネクタ_入荷待ち→出荷", 8, 0, 5, 2,  6, 8,  1, 8),
    ]

    inner = "\n".join(shapes) + "\n" + "\n".join(connectors)

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<xdr:wsDr'
        ' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        '>\n'
        + inner
        + "\n</xdr:wsDr>\n"
    )


# ---------------------------------------------------------------------------
# xlsx 組み立て
# ---------------------------------------------------------------------------

def build_xlsx() -> bytes:
    """openpyxlでベースxlsxを生成し、DrawingMLを注入してbytesで返す。"""
    # ベースxlsx生成
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "受注処理フロー"  # type: ignore[union-attr]
    ws["A1"] = "受注処理フロー図（サンプル）"  # type: ignore[union-attr]

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    base_bytes = buf.read()

    # ZipFile を操作して Drawing を注入
    in_buf = io.BytesIO(base_bytes)
    out_buf = io.BytesIO()

    drawing_xml = build_drawing_xml()
    drawing_path = "xl/drawings/drawing1.xml"
    drawing_rels_path = "xl/drawings/_rels/drawing1.xml.rels"
    sheet_rels_path = "xl/worksheets/_rels/sheet1.xml.rels"

    with zipfile.ZipFile(in_buf, "r") as zin:
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            existing_names = set(zin.namelist())

            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == "[Content_Types].xml":
                    # Drawing の ContentType を追加
                    data = data.replace(
                        b"</Types>",
                        b'<Override PartName="/xl/drawings/drawing1.xml"'
                        b' ContentType="application/vnd.openxmlformats-officedocument'
                        b".drawing+xml"
                        b'"/>'
                        b"</Types>",
                    )

                elif item.filename == sheet_rels_path:
                    # sheet1 → drawing1 のリレーションシップを追加
                    data = data.replace(
                        b"</Relationships>",
                        b'<Relationship Id="rId10"'
                        b' Type="http://schemas.openxmlformats.org/officeDocument'
                        b"/relationships/drawing"
                        b'" Target="../drawings/drawing1.xml"/>'
                        b"</Relationships>",
                    )

                elif item.filename == "xl/worksheets/sheet1.xml":
                    # <drawing r:id="rId10"/> を sheetData の後に追加
                    # r 名前空間を worksheet 要素に付与してから drawing 要素を挿入
                    if b"<drawing" not in data:
                        r_ns = (
                            b' xmlns:r="http://schemas.openxmlformats.org'
                            b'/officeDocument/2006/relationships"'
                        )
                        # <worksheet ... > に xmlns:r を追加（既になければ）
                        if b"xmlns:r" not in data:
                            data = data.replace(b"<worksheet ", b"<worksheet" + r_ns + b" ", 1)
                        data = data.replace(
                            b"</worksheet>",
                            b'<drawing r:id="rId10"/></worksheet>',
                        )

                zout.writestr(item, data)

            # sheet1 の _rels が存在しない場合は新規作成（openpyxlは通常生成しない）
            if sheet_rels_path not in existing_names:
                rels_xml = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org'
                    '/package/2006/relationships">'
                    '<Relationship Id="rId10"'
                    ' Type="http://schemas.openxmlformats.org/officeDocument'
                    '/2006/relationships/drawing"'
                    ' Target="../drawings/drawing1.xml"/>'
                    "</Relationships>"
                )
                zout.writestr(sheet_rels_path, rels_xml)

            # drawing XML 本体
            zout.writestr(drawing_path, drawing_xml)

            # drawing の _rels（空でも必要な場合あり）
            drawing_rels_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org'
                '/package/2006/relationships"/>'
            )
            zout.writestr(drawing_rels_path, drawing_rels_xml)

    out_buf.seek(0)
    return out_buf.read()


def main() -> None:
    xlsx_bytes = build_xlsx()
    OUTPUT_PATH.write_bytes(xlsx_bytes)
    print(f"生成完了: {OUTPUT_PATH}")

    # 生成したxlsxの内部構造を確認
    with zipfile.ZipFile(OUTPUT_PATH) as z:
        print("内包ファイル:")
        for name in sorted(z.namelist()):
            print(f"  {name}")
        if "xl/drawings/drawing1.xml" in z.namelist():
            print("\n--- xl/drawings/drawing1.xml の先頭 ---")
            xml = z.read("xl/drawings/drawing1.xml").decode()
            print(xml[:500])


if __name__ == "__main__":
    main()
