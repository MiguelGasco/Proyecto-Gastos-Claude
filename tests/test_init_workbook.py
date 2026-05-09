from pathlib import Path
import openpyxl
from scripts.init_workbook import create_workbook


def test_create_workbook_has_four_named_sheets(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["Dashboard", "Movimientos", "Categorías", "_aux"]


def test_create_workbook_fails_if_file_exists(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    import pytest
    with pytest.raises(FileExistsError):
        create_workbook(out)
