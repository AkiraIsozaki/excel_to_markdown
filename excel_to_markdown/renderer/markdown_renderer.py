"""DocElement リストを Markdown 文字列に変換する。

依存可能: models.py
依存禁止: reader/, parser/, cli.py
"""

from __future__ import annotations

import re

from excel_to_markdown.models import DocElement, ElementType, InlineRun, TableElement


def render(elements: list[DocElement], footnotes: list[str]) -> str:
    """DocElement リストを Markdown 文字列に変換する。

    footnotes: 脚注テキストリスト（セルコメント）。
               呼び出し元 (cli.py) が DocElement の comment_text を走査して収集し渡す。
               脚注番号は render_element() が [^1] から連番で付与する。
               末尾に [^1]: 内容 の形式で一括出力する。
    """
    parts: list[str] = []
    footnote_counter = 1

    for el in elements:
        md, footnote_counter = render_element(el, footnote_counter)
        parts.append(md)

    result = "".join(parts)

    if footnotes:
        note_lines = "\n".join(f"[^{i + 1}]: {fn}" for i, fn in enumerate(footnotes))
        result = result.rstrip("\n") + "\n\n" + note_lines + "\n"

    return collapse_blank_lines(result)


def render_element(el: DocElement, footnote_counter: int) -> tuple[str, int]:
    """1要素を Markdown 文字列に変換する。脚注付き要素は連番を更新して返す。"""
    text = convert_cell_newlines(el.text)

    # ハイパーリンク変換
    if el.hyperlink:
        text = f"[{text}]({el.hyperlink})"

    # 脚注マーカーを付記
    if el.comment_text:
        text = text + f"[^{footnote_counter}]"
        footnote_counter += 1

    if el.element_type == ElementType.HEADING:
        return "#" * el.level + " " + text + "\n\n", footnote_counter

    if el.element_type == ElementType.PARAGRAPH:
        return text + "\n\n", footnote_counter

    if el.element_type == ElementType.LIST_ITEM:
        indent = "  " * (el.level - 1)
        prefix = "1. " if el.is_numbered_list else "- "
        return indent + prefix + text + "\n", footnote_counter

    if el.element_type == ElementType.BLANK:
        return "\n", footnote_counter

    if el.element_type == ElementType.TABLE:
        if not isinstance(el, TableElement):
            return "", footnote_counter
        return _render_table(el), footnote_counter

    return text + "\n\n", footnote_counter


def _render_table(el: TableElement) -> str:
    """TableElement を GFM テーブル形式の Markdown に変換する。"""
    if not el.rows:
        return ""

    col_count = el.col_count or (len(el.rows[0]) if el.rows else 0)

    lines: list[str] = []

    # GFM はヘッダー行必須のため、is_header フラグに関わらず1行目をヘッダーとして出力
    header_cells = [convert_cell_newlines(c.text) for c in el.rows[0]]
    while len(header_cells) < col_count:
        header_cells.append("")
    lines.append("| " + " | ".join(header_cells) + " |")
    lines.append("| " + " | ".join("---" for _ in range(col_count)) + " |")

    start_row = 1  # 1行目はヘッダーとして出力済み
    for row in el.rows[start_row:]:
        cells = [convert_cell_newlines(c.text) for c in row]
        while len(cells) < col_count:
            cells.append("")
        lines.append("| " + " | ".join(cells) + " |")

    return "\n".join(lines) + "\n\n"


def render_inline(text: str, runs: list[InlineRun]) -> str:
    """テキストと InlineRun リストを結合してインライン書式付き Markdown 文字列を生成する。

    runs が空の場合は text をそのまま返す。
    """
    if not runs:
        return text
    return "".join(apply_inline_format(run) for run in runs)


def apply_inline_format(run: InlineRun) -> str:
    """InlineRun を対応する Markdown 記法に変換する。

    bold → **text**, italic → *text*, strikethrough → ~~text~~, underline → <u>text</u>
    複数書式の組み合わせ: bold+italic → ***text***
    """
    text = run.text
    if not text:
        return ""

    if run.strikethrough:
        text = f"~~{text}~~"
    if run.underline:
        text = f"<u>{text}</u>"
    if run.bold and run.italic:
        text = f"***{text}***"
    elif run.bold:
        text = f"**{text}**"
    elif run.italic:
        text = f"*{text}*"

    return text


def convert_cell_newlines(text: str) -> str:
    """セル内改行（\\n）を Markdown ハードブレーク（行末スペース2つ + \\n）に変換する。

    例: "1行目\\n2行目" → "1行目  \\n2行目"
    """
    return text.replace("\n", "  \n")


def collapse_blank_lines(md: str) -> str:
    """3行以上の連続空行を2行に圧縮する。"""
    return re.sub(r"\n{3,}", "\n\n", md)
