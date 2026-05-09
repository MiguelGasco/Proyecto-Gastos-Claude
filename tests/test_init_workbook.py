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


def test_movimientos_has_table_with_expected_columns(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    ws = wb["Movimientos"]
    assert "tblMov" in ws.tables
    headers = [c.value for c in ws[1]]
    assert headers == ["Fecha", "Tipo", "Categoría", "Importe", "Descripción"]


def test_categorias_sheet_lists_all_categories(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    ws = wb["Categorías"]
    pairs = {(row[0].value, row[1].value) for row in ws.iter_rows(min_row=2) if row[0].value}
    # Spot-check: at least these pairs must be present
    assert ("Gasto", "Alimentación") in pairs
    assert ("Gasto", "Restauración") in pairs
    assert ("Ingreso", "Nómina") in pairs
    # And total count matches our source of truth
    from scripts.categories import CATEGORIES
    expected = sum(len(v) for v in CATEGORIES.values())
    assert len(pairs) == expected


def test_movimientos_has_data_validations(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    ws = wb["Movimientos"]
    # At least two validations attached: one for Tipo (col B), one for Categoría (col C)
    dv_ranges = [str(dv.sqref) for dv in ws.data_validations.dataValidation]
    assert any("B" in r for r in dv_ranges), f"Missing Tipo validation, got {dv_ranges}"
    assert any("C" in r for r in dv_ranges), f"Missing Categoría validation, got {dv_ranges}"
