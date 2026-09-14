"""FastAPI Application for csv2xlsx converter."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

from app.converter import (
    MAX_FILE_SIZE_BYTES,
    ConversionError,
    convert_csv_to_xlsx,
    read_csv_preview,
)

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"

app = FastAPI(
    title="csv2xlsx Converter",
    description="Convert CSV/TSV files into beautifully styled Excel spreadsheets",
    version="1.0.0",
)

# CORS middleware for open development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def remove_temp_file(path: str) -> None:
    """Safe background cleanup of temporary output files."""
    try:
        if os.path.exists(path):
            os.unlink(path)
    except OSError:
        pass


@app.get("/api/health")
async def health_check():
    """Health check endpoint for Docker & CI/CD monitoring."""
    return {"status": "ok"}


@app.post("/api/preview")
async def preview_csv(
    file: UploadFile = File(...),
    delimiter: Optional[str] = Form(None),
):
    """Generate preview and metadata for uploaded CSV."""
    if not file.filename:
        raise HTTPException(status_code=400, detail="Nama file tidak valid.")

    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE_BYTES:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"Ukuran file melebihi batas maksimum 50 MB.",
        )

    try:
        # Pass empty or sanitized delimiter
        clean_delim = delimiter.strip() if delimiter and delimiter.strip() else None
        preview_data = read_csv_preview(
            file_bytes=contents,
            manual_delimiter=clean_delim,
            max_preview_rows=20,
            filename=file.filename,
        )
        return JSONResponse(content=preview_data)
    except ConversionError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal memproses file: {str(e)}")


@app.post("/api/convert")
async def convert_csv(
    file: UploadFile = File(...),
    sheet_name: Optional[str] = Form(None),
    delimiter: Optional[str] = Form(None),
    apply_autofilter: bool = Form(True),
    apply_autowidth: bool = Form(True),
):
    """Convert CSV to formatted Excel .xlsx and stream to client."""
    if not file.filename:
        raise HTTPException(status_code=400, detail="Nama file tidak valid.")

    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE_BYTES:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"Ukuran file melebihi batas maksimum 50 MB.",
        )

    # Prepare base filename without extension
    stem_name = Path(file.filename).stem or "export"
    resolved_sheet = sheet_name.strip() if sheet_name and sheet_name.strip() else stem_name
    clean_delim = delimiter.strip() if delimiter and delimiter.strip() else None

    # Write output to temporary file for memory efficiency on large sheets
    temp_fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="csv2xlsx_")
    try:
        with os.fdopen(temp_fd, "wb") as f_out:
            convert_csv_to_xlsx(
                file_bytes=contents,
                output_stream=f_out,
                sheet_name=resolved_sheet,
                manual_delimiter=clean_delim,
                apply_autofilter=apply_autofilter,
                apply_autowidth=apply_autowidth,
            )

        xlsx_filename = f"{stem_name}.xlsx"
        return FileResponse(
            path=temp_path,
            filename=xlsx_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            background=BackgroundTask(remove_temp_file, temp_path),
        )
    except ConversionError as e:
        remove_temp_file(temp_path)
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        remove_temp_file(temp_path)
        raise HTTPException(status_code=500, detail=f"Terjadi kesalahan saat konversi: {str(e)}")


# Serve static files if directory exists
if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")
