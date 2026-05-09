from pathlib import Path
import pytest

from scripts.verify_workbook import scan_errors, ERROR_MARKERS
from scripts.init_workbook import create_workbook


def test_scan_errors_clean_workbook(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)
    # New workbook has no cached values yet (Excel hasn't computed)
    # so scan_errors should return empty list — cells contain formulas (str starting with =)
    # or None or numbers, not error markers.
    errors = scan_errors(out)
    assert errors == [], f"unexpected errors in fresh workbook: {errors}"


def test_scan_errors_detects_injected_marker(tmp_path: Path):
    import openpyxl
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)
    wb = openpyxl.load_workbook(out)
    wb["Dashboard"]["Z1"] = "#REF!"
    wb.save(out)

    errors = scan_errors(out)
    assert any("Dashboard!Z1" in e for e in errors), errors
