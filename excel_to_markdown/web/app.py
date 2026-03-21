"""FastAPI アプリケーション定義。

POST /api/convert : Excel ファイルを Markdown に変換して返す
GET  /health      : ヘルスチェック
GET  /            : D&D UI (index.html)
"""

from __future__ import annotations

import io
import tempfile
import zipfile
from pathlib import Path
from typing import Annotated

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles

from excel_to_markdown.cli import DEFAULT_BASE_FONT_SIZE, run_file

_STATIC_DIR = Path(__file__).parent / "static"
_MAX_UPLOAD_BYTES = 50 * 1024 * 1024  # 50 MB
_ALLOWED_SUFFIXES = {".xlsx", ".xls"}


def create_app() -> FastAPI:
    """FastAPI アプリのファクトリ。テスト時はこれを直接使用する。"""
    app = FastAPI(title="excel-to-markdown Web UI")

    # 静的ファイル配信
    app.mount("/static", StaticFiles(directory=str(_STATIC_DIR)), name="static")

    @app.get("/", response_class=HTMLResponse)
    async def index() -> HTMLResponse:
        html = (_STATIC_DIR / "index.html").read_text(encoding="utf-8")
        return HTMLResponse(content=html)

    @app.get("/health")
    async def health() -> dict[str, str]:
        return {"status": "ok"}

    @app.post("/api/convert")
    async def convert(
        files: Annotated[list[UploadFile], File(...)],
    ) -> Response:
        """Excel ファイルを Markdown に変換する。

        単一ファイル → Content-Type: text/markdown
        複数ファイル → Content-Type: application/zip
        """
        results: list[tuple[str, str]] = []   # (元ファイル名, markdown)
        errors: list[str] = []

        for upload in files:
            filename = upload.filename or "unknown"
            suffix = Path(filename).suffix.lower()

            if suffix not in _ALLOWED_SUFFIXES:
                if len(files) == 1:
                    raise HTTPException(
                        status_code=400,
                        detail=f"対応していないファイル形式です: {suffix}（.xlsx/.xls のみ対応）",
                    )
                errors.append(filename)
                continue

            # ファイルサイズチェック
            data = await upload.read()
            if len(data) > _MAX_UPLOAD_BYTES:
                if len(files) == 1:
                    raise HTTPException(
                        status_code=413,
                        detail="ファイルサイズが上限（50MB）を超えています",
                    )
                errors.append(filename)
                continue

            # tempfile に書き出して変換パイプラインを呼び出す
            try:
                with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                    tmp_path = Path(tmp.name)
                    tmp_path.write_bytes(data)

                md = run_file(tmp_path, base_font_size=DEFAULT_BASE_FONT_SIZE)
                md_filename = Path(filename).with_suffix(".md").name
                results.append((md_filename, md))
            except Exception:
                errors.append(filename)
            finally:
                try:
                    tmp_path.unlink(missing_ok=True)
                except Exception:
                    pass

        if not results:
            raise HTTPException(
                status_code=422,
                detail="すべてのファイルの変換に失敗しました",
            )

        # 単一ファイルは .md を直接返す
        if len(results) == 1 and not errors:
            md_filename, md_content = results[0]
            return Response(
                content=md_content.encode("utf-8"),
                media_type="text/markdown; charset=utf-8",
                headers={
                    "Content-Disposition": f'attachment; filename="{md_filename}"',
                },
            )

        # 複数ファイル（または部分エラー）は ZIP で返す
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for md_filename, md_content in results:
                zf.writestr(md_filename, md_content.encode("utf-8"))
        zip_bytes = zip_buf.getvalue()

        headers: dict[str, str] = {
            "Content-Disposition": 'attachment; filename="converted.zip"',
        }
        if errors:
            headers["X-Conversion-Errors"] = ",".join(errors)

        return Response(
            content=zip_bytes,
            media_type="application/zip",
            headers=headers,
        )

    return app
