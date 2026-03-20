"""実際の Excel 方眼紙ドキュメントを使ったエンドツーエンドテスト。

フィクスチャ: tests/e2e/fixtures/sample_houganshi.xlsx
  - SIer の要件定義書スタイルを再現した方眼紙 xlsx（2シート）
  - 生成スクリプト: tests/e2e/fixtures/make_sample_houganshi.py

ゴールデンファイル: tests/e2e/golden/sample_houganshi.md
  - 初回実行時に生成された変換結果の基準出力
  - ゴールデンファイルを意図的に更新する場合は:
      python -m excel_to_markdown tests/e2e/fixtures/sample_houganshi.xlsx \\
             -o tests/e2e/golden/sample_houganshi.md

テスト方針:
  1. ゴールデン比較 — 出力全体がゴールデンファイルと一致すること（回帰検知）
  2. コンテンツ保全 — 入力のすべてのテキストが出力に含まれること
  3. 構造検出 — 見出し・テーブル・脚注が期待する Markdown 記法で出力されること
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_to_markdown.cli import parse_args, run

FIXTURES_DIR = Path(__file__).parent / "fixtures"
GOLDEN_DIR = Path(__file__).parent / "golden"

FIXTURE_XLSX = FIXTURES_DIR / "sample_houganshi.xlsx"
GOLDEN_MD = GOLDEN_DIR / "sample_houganshi.md"


def _convert(xlsx_path: Path, tmp_path: Path) -> str:
    """xlsx を変換して Markdown 文字列を返す。"""
    out = tmp_path / "output.md"
    args = parse_args([str(xlsx_path), "-o", str(out)])
    code = run(args)
    assert code == 0, f"変換が失敗しました (exit code={code})"
    return out.read_text(encoding="utf-8")


# ---------------------------------------------------------------------------
# ゴールデンファイル比較（回帰検知）
# ---------------------------------------------------------------------------


class TestGoldenFile:
    def test_output_matches_golden(self, tmp_path: Path) -> None:
        """変換結果がゴールデンファイルと完全一致すること。

        失敗した場合はゴールデンファイルを意図的に更新する必要があるか、
        または変換ロジックにリグレッションが発生している。
        """
        assert GOLDEN_MD.exists(), (
            f"ゴールデンファイルが存在しません: {GOLDEN_MD}\n"
            "以下のコマンドで生成してください:\n"
            f"  python -m excel_to_markdown {FIXTURE_XLSX} -o {GOLDEN_MD}"
        )
        actual = _convert(FIXTURE_XLSX, tmp_path)
        expected = GOLDEN_MD.read_text(encoding="utf-8")
        assert actual == expected, (
            "変換結果がゴールデンファイルと一致しません。\n"
            "意図した変更の場合はゴールデンファイルを更新してください:\n"
            f"  python -m excel_to_markdown {FIXTURE_XLSX} -o {GOLDEN_MD}"
        )


# ---------------------------------------------------------------------------
# コンテンツ保全（全テキストが出力に含まれること）
# ---------------------------------------------------------------------------


class TestContentPreservation:
    """入力 Excel のすべてのテキストが変換後 Markdown に含まれること。"""

    ALL_TEXTS = [
        # タイトル・プロジェクト情報
        "機能要件定義書",
        "在庫管理システム刷新プロジェクト",
        "田中 誠一",
        "2025-01-15",
        # セクション見出し
        "1. 概要",
        "2. 機能一覧",
        "3. 詳細要件",
        "4. 非機能要件",
        # 概要本文
        "本書は在庫管理システム刷新プロジェクトにおける機能要件を定義するものである",
        "リアルタイムな在庫状況の把握を可能にする",
        # 機能一覧テーブルデータ
        "在庫照会",
        "入庫登録",
        "出庫登録",
        "在庫調整",
        "アラート通知",
        "バーコード対応",
        "承認フロー必要",
        "要件#12参照",
        # 詳細要件
        "3.1 在庫照会",
        "JANコード",
        "ページング機能",
        "CSV形式でエクスポート",
        "3.2 入庫登録",
        "入庫伝票番号",
        "バーコードリーダー",
        "在庫マスタに自動反映",
        "入庫履歴は最低5年間保持",
        # 非機能要件テーブル
        "画面の初期表示時間",
        "3秒以内",
        "在庫検索のレスポンス",
        "99.5%以上",
        "2要素認証",
        "操作ログ: 1年間",
        # セルコメント（脚注）
        "SLAレポートのテンプレートは別途資料参照",
        "佐藤PM",
        # 画面一覧シート
        "画面一覧",
        "SCR-001",
        "ログイン画面",
        "SCR-010",
        "メニュー画面",
        "SCR-101",
        "在庫照会画面",
        "SCR-201",
        "入庫登録画面",
        "SCR-301",
        "出庫登録画面",
    ]

    def test_all_texts_preserved(self, tmp_path: Path) -> None:
        md = _convert(FIXTURE_XLSX, tmp_path)
        missing = [t for t in self.ALL_TEXTS if t not in md]
        assert not missing, (
            f"以下のテキストが変換結果に含まれていません:\n"
            + "\n".join(f"  - {t}" for t in missing)
        )


# ---------------------------------------------------------------------------
# 構造検出
# ---------------------------------------------------------------------------


class TestStructureDetection:
    def test_title_is_h1(self, tmp_path: Path) -> None:
        """18pt フォントのタイトルが H1 見出しに変換されること。"""
        md = _convert(FIXTURE_XLSX, tmp_path)
        assert "# 機能要件定義書" in md

    def test_section_headers_are_h2(self, tmp_path: Path) -> None:
        """14pt+Bold のセクション見出しが H2 に変換されること。"""
        md = _convert(FIXTURE_XLSX, tmp_path)
        assert "## 1. 概要" in md
        assert "## 2. 機能一覧" in md
        assert "## 3. 詳細要件" in md
        assert "## 4. 非機能要件" in md

    def test_subsection_headers_are_h3(self, tmp_path: Path) -> None:
        """12pt+Bold の小見出しが H3 に変換されること。"""
        md = _convert(FIXTURE_XLSX, tmp_path)
        assert "### 3.1 在庫照会" in md
        assert "### 3.2 入庫登録" in md

    def test_nonfunctional_table_detected(self, tmp_path: Path) -> None:
        """非機能要件テーブルが GFM テーブルとして出力されること。"""
        md = _convert(FIXTURE_XLSX, tmp_path)
        assert "| 分類 |" in md
        assert "| 要件 |" in md
        assert "| 目標値 |" in md
        assert "| 性能 |" in md
        assert "| 3秒以内 |" in md

    def test_screen_list_table_detected(self, tmp_path: Path) -> None:
        """画面一覧テーブルが GFM テーブルとして出力されること。"""
        md = _convert(FIXTURE_XLSX, tmp_path)
        assert "| 画面ID |" in md
        assert "| SCR-001 |" in md
        assert "| SCR-101 |" in md

    def test_cell_comment_becomes_footnote(self, tmp_path: Path) -> None:
        """セルコメントが脚注として出力されること。"""
        md = _convert(FIXTURE_XLSX, tmp_path)
        assert "[^1]" in md
        assert "SLAレポートのテンプレートは別途資料参照" in md

    def test_multiple_sheets_merged(self, tmp_path: Path) -> None:
        """複数シートが H1 区切りで統合されること。"""
        md = _convert(FIXTURE_XLSX, tmp_path)
        assert "# 機能要件\n" in md
        assert "# 画面一覧\n" in md

    def test_project_info_preserved(self, tmp_path: Path) -> None:
        """プロジェクト情報（ラベル:値）が出力に含まれること。"""
        md = _convert(FIXTURE_XLSX, tmp_path)
        assert "プロジェクト名" in md
        assert "在庫管理システム刷新プロジェクト" in md
