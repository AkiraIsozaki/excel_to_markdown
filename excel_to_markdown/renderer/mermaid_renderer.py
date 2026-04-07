"""Mermaidフローチャートレンダラー。

DiagramShape / DiagramConnector のリストを Mermaid flowchart 文字列に変換する。
"""

from __future__ import annotations

from excel_to_markdown.models import DiagramConnector, DiagramShape

# ---------------------------------------------------------------------------
# 形状タイプ → Mermaid ノード記法マッピング
#
# Mermaid の記法:
#   [text]       四角形（プロセス）
#   ([text])     スタジアム形（開始/終了）
#   {text}       ひし形（判断）
#   ((text))     二重円（コネクタ・楕円）
#   >text]       非対称（手動入力）
#   [(text)]     シリンダ（データベース）
# ---------------------------------------------------------------------------
# プレースホルダーに __TEXT__ を使用して中括弧との混乱を避ける
_SHAPE_TO_MERMAID_TEMPLATE: dict[str, str] = {
    # フローチャート専用形状
    "flowChartTerminator": "([__TEXT__])",        # スタジアム形（開始/終了）
    "flowChartProcess": "[__TEXT__]",              # 四角形（処理）
    "flowChartDecision": "{__TEXT__}",             # ひし形（判断）
    "flowChartConnector": "((__TEXT__))",          # 二重円（接続子）
    "flowChartManualInput": ">__TEXT__]",          # 非対称（手動入力）
    "flowChartDatabase": "[(__TEXT__)]",           # シリンダ（データベース）
    "flowChartPredefinedProcess": "[[__TEXT__]]",  # サブルーティン
    "flowChartDelay": "[/__TEXT__/]",              # 遅延
    "flowChartManualOperation": "[\\__TEXT__\\]",  # 手動操作
    "flowChartDocument": "[/__TEXT__\\]",          # 文書
    "flowChartMagneticDisk": "[(__TEXT__)]",       # 磁気ディスク（システム）→ シリンダで代替
    "flowChartOfflineStorage": "[(__TEXT__)]",     # オフラインストレージ
    "flowChartOnlineStorage": "[(__TEXT__)]",      # オンラインストレージ
    # 汎用形状
    "rect": "[__TEXT__]",
    "roundRect": "(__TEXT__)",
    "ellipse": "((__TEXT__))",
    "diamond": "{__TEXT__}",
    "triangle": "[__TEXT__]",
    "parallelogram": "[/__TEXT__/]",
    "hexagon": "{{__TEXT__}}",
    # 吹き出し系（ノートスタイルで代替）
    "wedgeRoundRectCallout": "[__TEXT__]",
    "wedgeRectCallout": "[__TEXT__]",
    "cloudCallout": "[__TEXT__]",
    "callout1": "[__TEXT__]",
    "callout2": "[__TEXT__]",
}

_DEFAULT_TEMPLATE = "[__TEXT__]"

# テキストが空のシェイプのフォールバック用プレフィックス
_EMPTY_TEXT_PREFIX = "node_"


def _safe_text(shape: DiagramShape) -> str:
    """シェイプのテキストを返す。空の場合はシェイプ名またはIDをフォールバックに使う。"""
    if shape.text:
        return shape.text
    if shape.name:
        return shape.name
    return f"{_EMPTY_TEXT_PREFIX}{shape.shape_id}"


def _node_notation(shape_type: str, text: str) -> str:
    """形状タイプとテキストからMermaidノード記法文字列を返す。"""
    template = _SHAPE_TO_MERMAID_TEMPLATE.get(shape_type, _DEFAULT_TEMPLATE)
    # テキスト内の特殊文字をエスケープ（改行は空白に変換）
    safe = text.replace("\n", " ").replace('"', "#quot;")
    return template.replace("__TEXT__", safe)


def _node_id(shape_id: int) -> str:
    """シェイプIDからMermaidノードIDを生成する。"""
    return f"N{shape_id}"


def _detect_direction(shapes: list[DiagramShape]) -> str:
    """図形の配置から最適なグラフ方向を返す（'TD' または 'LR'）。

    全図形のバウンディングボックスを計算し、
    縦幅 > 横幅 × 1.2 なら TD、それ以外は LR。
    """
    if not shapes:
        return "TD"

    min_col = min(s.left_col for s in shapes)
    max_col = max(s.right_col for s in shapes)
    min_row = min(s.top_row for s in shapes)
    max_row = max(s.bottom_row for s in shapes)

    width = max_col - min_col
    height = max_row - min_row

    return "TD" if height > width * 1.2 else "LR"


def render_mermaid(
    shapes: list[DiagramShape],
    connectors: list[DiagramConnector],
    *,
    direction: str | None = None,
) -> str:
    """図形・コネクタリストをMermaid flowchart文字列に変換する。

    Args:
        shapes: DiagramShape のリスト
        connectors: DiagramConnector のリスト
        direction: グラフ方向（'TD'/'LR'）。None で自動判定

    Returns:
        Mermaid flowchart 文字列（先頭行は 'flowchart TD' など）
    """
    if not shapes:
        return "flowchart TD\n"

    dir_ = direction or _detect_direction(shapes)
    lines: list[str] = [f"flowchart {dir_}"]

    shape_ids = {s.shape_id for s in shapes}

    # 有効なエッジ（両端が既知のシェイプ）を先に収集
    valid_edges: list[DiagramConnector] = [
        c for c in connectors
        if c.start_shape_id is not None
        and c.end_shape_id is not None
        and c.start_shape_id in shape_ids
        and c.end_shape_id in shape_ids
    ]

    # 接続されているシェイプIDセット（孤立ノード判定に使用）
    connected_ids: set[int] = set()
    for c in valid_edges:
        if c.start_shape_id is not None:
            connected_ids.add(c.start_shape_id)
        if c.end_shape_id is not None:
            connected_ids.add(c.end_shape_id)

    # コネクタが存在する場合は孤立ノードを除外（ラベルのみの吹き出し等を除く）
    # コネクタが存在しない場合（UIモックアップ等）はすべてのノードを表示
    if valid_edges:
        render_shapes = [s for s in shapes if s.shape_id in connected_ids]
    else:
        render_shapes = shapes

    # ノード定義
    for shape in render_shapes:
        node_id = _node_id(shape.shape_id)
        text = _safe_text(shape)
        notation = _node_notation(shape.shape_type, text)
        lines.append(f"    {node_id}{notation}")

    # エッジ定義
    for conn in valid_edges:
        src = _node_id(conn.start_shape_id)  # type: ignore[arg-type]
        dst = _node_id(conn.end_shape_id)    # type: ignore[arg-type]
        if conn.label:
            lines.append(f'    {src} -->|{conn.label}| {dst}')
        else:
            lines.append(f"    {src} --> {dst}")

    return "\n".join(lines) + "\n"


def render_mermaid_block(
    shapes: list[DiagramShape],
    connectors: list[DiagramConnector],
    *,
    direction: str | None = None,
) -> str:
    """Mermaid文字列をMarkdownコードブロックとして返す。

    Returns:
        ```mermaid\\n...\\n``` 形式の文字列
    """
    mermaid_content = render_mermaid(shapes, connectors, direction=direction)
    return f"```mermaid\n{mermaid_content}```\n"
