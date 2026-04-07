"""パイプライン全体で使用する共有データモデル。

依存可能: Python標準ライブラリのみ（dataclasses, enum）
依存禁止: openpyxl, xlrd, パッケージ内の他モジュール
"""

from __future__ import annotations

import enum
from dataclasses import dataclass, field

# ---------------------------------------------------------------------------
# レイヤー1: RawCell（生セルデータ）
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class RawCell:
    """openpyxl / xlrd から抽出した1セルの生データ。不変値オブジェクト。

    - is_merge_origin=False かつ結合領域内のセルは value=None として扱い、変換対象から除外する
    - font_size=None はExcelのデフォルトスタイル継承を意味する
    """

    row: int  # 1-based 行番号 (openpyxl準拠)
    col: int  # 1-based 列番号 (openpyxl準拠)
    value: str | None  # セルの文字列値（数値は文字列に変換済み）
    font_bold: bool  # 太字フラグ
    font_italic: bool  # イタリックフラグ
    font_strikethrough: bool  # 取り消し線フラグ
    font_underline: bool  # 下線フラグ
    font_size: float | None  # フォントサイズ（pt）。未設定時はNone
    font_color: str | None  # 文字色 ARGB hex（例: "FF000000"）。テーマ色はNone
    bg_color: str | None  # 背景色 ARGB hex（例: "FFFFFFFF"）
    is_merge_origin: bool  # 結合セルの起点か否か
    merge_row_span: int  # 結合の行スパン（非結合・非起点は1）
    merge_col_span: int  # 結合の列スパン（非結合・非起点は1）
    has_comment: bool  # セルコメントの有無
    comment_text: str | None  # セルコメントのテキスト
    hyperlink: str | None = None  # ハイパーリンクURL（なければNone）


# ---------------------------------------------------------------------------
# レイヤー2: TextBlock（テキストブロック）
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class InlineRun:
    """セル内の部分書式適用テキスト。セル全体に書式がある場合はTextBlockに統合。"""

    text: str
    bold: bool = False
    italic: bool = False
    strikethrough: bool = False
    underline: bool = False


@dataclass
class TextBlock:
    """結合セルを解決した後の1テキストブロック。

    indent_level は 0 で初期化され、structure_detector が後で更新する。
    """

    text: str  # セルのテキスト内容（前後の空白を除去済み）
    top_row: int  # 先頭行 (1-based)
    left_col: int  # 先頭列 (1-based)
    bottom_row: int  # 末尾行 (1-based)
    right_col: int  # 末尾列 (1-based)
    row_span: int  # 行スパン
    col_span: int  # 列スパン
    font_bold: bool
    font_italic: bool
    font_strikethrough: bool
    font_underline: bool
    font_size: float | None
    bg_color: str | None  # 背景色 ARGB hex（セクション境界判定に使用）
    has_comment: bool
    comment_text: str | None
    indent_level: int = 0  # 後処理で計算。左端列位置から算出したインデント階層
    inline_runs: list[InlineRun] = field(default_factory=list)
    hyperlink: str | None = None  # ハイパーリンクURL（なければNone）


# ---------------------------------------------------------------------------
# レイヤー3: DocElement（文書要素）
# ---------------------------------------------------------------------------


class ElementType(enum.Enum):
    """文書要素の種別。"""

    HEADING = "heading"
    PARAGRAPH = "paragraph"
    LIST_ITEM = "list_item"
    TABLE = "table"
    BLANK = "blank"


@dataclass
class DocElement:
    """1つの文書要素。HEADING / PARAGRAPH / LIST_ITEM / BLANK に使用する。"""

    element_type: ElementType
    text: str  # HEADING/PARAGRAPH/LIST_ITEMのテキスト。TABLE/BLANKは空文字
    level: int  # HEADINGは1-6、LIST_ITEMはインデント深さ(1+)、その他は0
    source_row: int  # 元のExcel行番号（ソート・デバッグ用）
    is_numbered_list: bool = False  # 番号付きリスト（`1.`/`1)` 始まり）か否か
    comment_text: str | None = None  # 脚注として出力するセルコメント
    hyperlink: str | None = None  # ハイパーリンクURL（なければNone）


@dataclass
class TableCell:
    """テーブル内の1セル。"""

    text: str
    row: int  # テーブル相対行インデックス (0-based)
    col: int  # テーブル相対列インデックス (0-based)
    is_header: bool  # 最初の行のセルはTrue
    # 注意: 表内の結合セルはセルスパンのMarkdown表現が未サポート。
    #       非起点セルは空文字("")として出力される（情報の欠落を抑制するためベストエフォート）。


@dataclass
class TableElement(DocElement):
    """テーブル要素。DocElement の element_type は常に TABLE。"""

    element_type: ElementType = field(default=ElementType.TABLE, init=False)
    rows: list[list[TableCell]] = field(default_factory=list)
    col_count: int = 0


# ---------------------------------------------------------------------------
# レイヤー4: 図形データモデル（DrawingML図形変換用）
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class DiagramShape:
    """DrawingML の <xdr:sp> から抽出した図形1つ。

    位置はセル座標（0-based）で保持する。
    shape_type は prstGeom の prst 値（例: "flowChartProcess"）。
    """

    shape_id: int  # cNvPr id
    name: str  # cNvPr name
    text: str  # txBody のテキスト（空白除去済み）
    shape_type: str  # prstGeom の prst 値。未設定時は "rect"
    left_col: int  # TwoCellAnchor.from_.col（0-based）
    top_row: int  # TwoCellAnchor.from_.row（0-based）
    right_col: int  # TwoCellAnchor.to_.col（0-based）
    bottom_row: int  # TwoCellAnchor.to_.row（0-based）


@dataclass(frozen=True)
class DiagramConnector:
    """DrawingML の <xdr:cxnSp> から抽出したコネクタ1本。

    start_shape_id / end_shape_id は stCxn / endCxn の id 属性。
    未接続（属性なし）の場合は None。
    """

    connector_id: int  # cNvPr id
    name: str  # cNvPr name
    start_shape_id: int | None  # stCxn id（Noneは未接続）
    end_shape_id: int | None  # endCxn id（Noneは未接続）
    label: str  # コネクタ上のテキストラベル（なければ空文字）
