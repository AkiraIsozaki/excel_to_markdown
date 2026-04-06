#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WHEELHOUSE="${SCRIPT_DIR}/wheelhouse"
VENV_DIR="${SCRIPT_DIR}/.venv"

echo "=== excel-to-markdown オフラインインストーラー ==="
echo ""

if [ ! -d "${WHEELHOUSE}" ]; then
  echo "エラー: wheelhouse/ ディレクトリが見つかりません。"
  echo "このスクリプトはプロジェクトのルートディレクトリで実行してください。"
  exit 1
fi

# インストールモードを引数で切り替え
MODE="${1:-base}"

case "${MODE}" in
  base|web|xls|all) ;;
  *)
    echo "使用方法: $0 [base|web|xls|all]"
    echo ""
    echo "  base  本体のみ（デフォルト）"
    echo "  web   本体 + Web UI（FastAPI）"
    echo "  xls   本体 + .xlsファイルサポート"
    echo "  all   すべてのオプション込み"
    exit 1
    ;;
esac

# venv を作成（既存の場合はスキップ）
if [ ! -d "${VENV_DIR}" ]; then
  echo "仮想環境を作成しています: ${VENV_DIR}"
  python3 -m venv "${VENV_DIR}"
else
  echo "既存の仮想環境を使用します: ${VENV_DIR}"
fi

PIP="${VENV_DIR}/bin/pip"

case "${MODE}" in
  base)
    echo "インストール対象: 本体のみ (openpyxl)"
    "${PIP}" install --no-index --find-links="${WHEELHOUSE}" "${SCRIPT_DIR}"
    ;;
  web)
    echo "インストール対象: 本体 + Web UI (fastapi / uvicorn)"
    "${PIP}" install --no-index --find-links="${WHEELHOUSE}" "${SCRIPT_DIR}[web]"
    ;;
  xls)
    echo "インストール対象: 本体 + xlsサポート (xlrd)"
    "${PIP}" install --no-index --find-links="${WHEELHOUSE}" "${SCRIPT_DIR}[xls]"
    ;;
  all)
    echo "インストール対象: 本体 + Web UI + xlsサポート"
    "${PIP}" install --no-index --find-links="${WHEELHOUSE}" "${SCRIPT_DIR}[web,xls]"
    ;;
esac

echo ""
echo "インストール完了！"
echo ""
echo "使い方:"
echo "  仮想環境を有効化: source ${VENV_DIR}/bin/activate"
echo "  CLIで変換: excel-to-markdown input.xlsx"
if [ "${MODE}" = "web" ] || [ "${MODE}" = "all" ]; then
  echo "  Web UIを起動: uvicorn excel_to_markdown.web:app --reload"
fi
