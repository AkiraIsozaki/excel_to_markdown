"""CLI 引数のパース・バリデーション、変換パイプラインの起動。

依存可能: 全モジュール (reader/, parser/, renderer/, models.py)
"""

from __future__ import annotations

import argparse
import dataclasses
import json
import sys
from pathlib import Path
from typing import Sequence

import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

from excel_to_markdown import __version__
from excel_to_markdown.models import DocElement, RawCell, TextBlock
from excel_to_markdown.parser.cell_grid import CellGrid
from excel_to_markdown.parser.merge_resolver import resolve
from excel_to_markdown.parser.structure_detector import detect
from excel_to_markdown.parser.table_detector import find_tables
from excel_to_markdown.reader.xlsx_reader import read_sheet
from excel_to_markdown.renderer.markdown_renderer import render

# デフォルト値
DEFAULT_BASE_FONT_SIZE: float = 11.0
DEFAULT_COL_WIDTH: float = 8.0
DEFAULT_ROW_HEIGHT: float = 15.0


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """CLI 引数を解析する。バリデーションエラーは argparse が処理。"""
    parser = argparse.ArgumentParser(
        prog="python -m excel_to_markdown",
        description="Excel方眼紙 (.xlsx/.xls) を Markdown に変換する",
    )
    parser.add_argument(
        "input",
        help="変換する .xlsx または .xls ファイルのパス",
    )
    parser.add_argument(
        "--output",
        "-o",
        default=None,
        metavar="OUTPUT",
        help="出力 .md ファイルパス（省略時: 入力と同名 .md）",
    )
    parser.add_argument(
        "--sheet",
        "-s",
        default=None,
        metavar="SHEET",
        help="シート名または 0-based インデックス（省略時: 全シートを統合）",
    )
    parser.add_argument(
        "--base-font-size",
        type=float,
        default=DEFAULT_BASE_FONT_SIZE,
        metavar="SIZE",
        help=f"見出し判定の基準フォントサイズ（デフォルト: {DEFAULT_BASE_FONT_SIZE}）",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        default=False,
        help="TextBlock リストを JSON 形式で stderr に出力",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    return parser.parse_args(argv)


def run(args: argparse.Namespace) -> int:
    """変換パイプライン全体を実行し、exit code を返す。

    複数シート統合ロジック:
    1. 全シートを順番に処理する（--sheet 指定時は1シートのみ）
    2. 各シートの Markdown 文字列を生成する
    3. シート数が2枚以上の場合: シート名を `# シート名\\n\\n---\\n\\n` で区切り1つに連結する
    4. 最終 Markdown 文字列を出力ファイルに書き出す
    """
    output_path: Path | None = None
    try:
        input_path = Path(args.input).resolve()
        _validate_input(input_path)

        output_path = _resolve_output_path(input_path, args.output)

        wb = _open_workbook(input_path)

        sheets = _select_sheets(wb, args.sheet)

        sheet_markdowns: list[tuple[str, str]] = []  # (sheet_name, markdown)

        for ws in sheets:
            sheet_name: str = ws.title
            raw_cells = read_sheet(ws)
            if not any(c.value for c in raw_cells):
                print(
                    f'警告: シート "{sheet_name}" にコンテンツがありません。スキップします',
                    file=sys.stderr,
                )
                continue

            grid = _build_grid(ws, raw_cells)
            blocks = resolve(raw_cells)

            if args.debug:
                _dump_blocks_debug(blocks)

            tables, remaining = find_tables(blocks, grid)
            doc_elements = detect(remaining, grid, args.base_font_size)

            all_elements: list[DocElement] = sorted(
                list(tables) + doc_elements,
                key=lambda e: e.source_row,
            )

            footnotes = [e.comment_text for e in all_elements if e.comment_text]
            md = render(all_elements, footnotes)
            sheet_markdowns.append((sheet_name, md))

        if not sheet_markdowns:
            return 0

        if len(sheet_markdowns) == 1:
            final_md = sheet_markdowns[0][1]
        else:
            parts: list[str] = []
            for sheet_name, md in sheet_markdowns:
                parts.append(f"# {sheet_name}\n\n---\n\n{md}")
            final_md = "\n".join(parts)

        _write_output(output_path, final_md)
        return 0

    except FileNotFoundError as e:
        print(f"エラー: ファイルが見つかりません: {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1
    except PermissionError:
        print(f"エラー: 出力ファイルに書き込めません: {output_path}", file=sys.stderr)
        return 1
    except Exception as e:  # noqa: BLE001
        print(f"予期しないエラーが発生しました: {e}", file=sys.stderr)
        return 2


def main() -> None:
    """CLI エントリーポイント。"""
    args = parse_args()
    sys.exit(run(args))


# ---------------------------------------------------------------------------
# 内部ヘルパー
# ---------------------------------------------------------------------------


def _validate_input(path: Path) -> None:
    """入力ファイルのバリデーション。"""
    if not path.exists():
        raise FileNotFoundError(path)
    if path.suffix.lower() not in {".xlsx", ".xls"}:
        raise ValueError(
            f"対応していないファイル形式です: {path.suffix}（.xlsx/.xls のみ対応）"
        )


def _resolve_output_path(input_path: Path, output_arg: str | None) -> Path:
    """出力ファイルパスを決定する。"""
    if output_arg is not None:
        return Path(output_arg).resolve()
    return input_path.with_suffix(".md")


def _open_workbook(path: Path) -> openpyxl.Workbook:
    """ワークブックを開く。パスワード保護されている場合は ValueError を送出。"""
    try:
        return openpyxl.load_workbook(str(path), data_only=True)
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypted" in msg or "protect" in msg:
            raise ValueError("パスワード保護されたファイルは変換できません") from e
        raise


def _select_sheets(
    wb: openpyxl.Workbook, sheet_arg: str | None
) -> list[Worksheet]:
    """対象シートのリストを返す。"""
    if sheet_arg is None:
        return [wb[name] for name in wb.sheetnames]

    # 数値インデックス (0-based)
    if sheet_arg.isdigit():
        idx = int(sheet_arg)
        if idx >= len(wb.sheetnames):
            raise ValueError(
                f"シートが見つかりません: {sheet_arg}"
                f"（存在するシート: {wb.sheetnames}）"
            )
        return [wb[wb.sheetnames[idx]]]

    # シート名で検索
    if sheet_arg not in wb.sheetnames:
        raise ValueError(
            f"シートが見つかりません: {sheet_arg}"
            f"（存在するシート: {wb.sheetnames}）"
        )
    return [wb[sheet_arg]]


def _build_grid(ws: Worksheet, raw_cells: list[RawCell]) -> CellGrid:
    """ワークシートから CellGrid を構築する。"""
    col_widths: dict[int, float] = {
        column_index_from_string(col): ws.column_dimensions[col].width or DEFAULT_COL_WIDTH
        for col in ws.column_dimensions
    }
    row_heights: dict[int, float] = {
        r: ws.row_dimensions[r].height or DEFAULT_ROW_HEIGHT
        for r in ws.row_dimensions
    }
    return CellGrid(cells=raw_cells, col_widths=col_widths, row_heights=row_heights)


def _dump_blocks_debug(blocks: list[TextBlock]) -> None:
    """TextBlock リストを JSON 形式で stderr に出力する。"""
    data = [dataclasses.asdict(b) for b in blocks]
    print(json.dumps(data, ensure_ascii=False, indent=2), file=sys.stderr)


def _write_output(path: Path, content: str) -> None:
    """Markdown を UTF-8 で書き出す。"""
    try:
        path.write_text(content, encoding="utf-8")
    except OSError as e:
        raise PermissionError(f"出力ファイルに書き込めません: {path}") from e
