"""Integration tests for FastAPI endpoints."""

import io
import openpyxl
import pytest
from fastapi.testclient import TestClient

from app.main import app

client = TestClient(app)


def test_health_check():
    response = client.get("/api/health")
    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_preview_endpoint_success():
    csv_content = b"colA,colB,colC\nval1,val2,100\nval3,val4,200\n"
    response = client.post(
        "/api/preview",
        files={"file": ("sample.csv", csv_content, "text/csv")},
    )
    assert response.status_code == 200
    data = response.json()
    assert data["success"] is True
    assert data["filename"] == "sample.csv"
    assert data["detected_delimiter"] == ","
    assert data["total_rows"] == 3
    assert data["total_columns"] == 3
    assert data["headers"] == ["colA", "colB", "colC"]
    assert len(data["preview_rows"]) == 2


def test_preview_endpoint_empty_file():
    response = client.post(
        "/api/preview",
        files={"file": ("empty.csv", b"", "text/csv")},
    )
    assert response.status_code == 400
    assert "File CSV kosong" in response.json()["detail"]


def test_preview_endpoint_binary_rejection():
    bad_data = b"\x00\x01\x02\x03\x04\x05\x00"
    response = client.post(
        "/api/preview",
        files={"file": ("bad.bin", bad_data, "application/octet-stream")},
    )
    assert response.status_code == 400
    assert "biner" in response.json()["detail"]


def test_convert_endpoint_success():
    csv_content = b"nama,gaji,status\nBudi,5000000,true\nAni,6500000,false\n"
    response = client.post(
        "/api/convert",
        files={"file": ("karyawan.csv", csv_content, "text/csv")},
        data={"sheet_name": "Daftar Karyawan", "apply_autofilter": "true", "apply_autowidth": "true"},
    )
    assert response.status_code == 200
    assert (
        response.headers["content-type"]
        == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert 'filename="karyawan.xlsx"' in response.headers["content-disposition"]

    # Verify that the generated binary is indeed a readable XLSX file
    wb = openpyxl.load_workbook(io.BytesIO(response.content))
    assert "Daftar Karyawan" in wb.sheetnames
    ws = wb["Daftar Karyawan"]
    assert ws.cell(row=1, column=1).value == "nama"
    assert ws.cell(row=2, column=2).value == 5000000


def test_index_page_served():
    response = client.get("/")
    assert response.status_code == 200
    assert "csv2xlsx" in response.text
