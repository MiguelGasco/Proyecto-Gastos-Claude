from datetime import date
from pathlib import Path
import pytest
import openpyxl

from scripts.init_workbook import create_workbook
from scripts.add_movement import (
    add_movement,
    edit_last,
    delete_last,
    InvalidMovementError,
)


@pytest.fixture()
def fresh_book(tmp_path: Path) -> Path:
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)
    return out


def _read_data_rows(path: Path) -> list[tuple]:
    wb = openpyxl.load_workbook(path)
    ws = wb["Movimientos"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        rows.append(row)
    return rows


def test_add_appends_a_row(fresh_book: Path):
    add_movement(
        fresh_book,
        fecha=date(2026, 5, 9),
        tipo="Gasto",
        categoria="Restauración",
        importe=30.0,
        descripcion="Comida con amigos",
    )
    rows = _read_data_rows(fresh_book)
    assert len(rows) == 1
    fecha, tipo, cat, imp, desc = rows[0]
    assert fecha == date(2026, 5, 9) or str(fecha).startswith("2026-05-09")
    assert tipo == "Gasto"
    assert cat == "Restauración"
    assert imp == 30.0
    assert desc == "Comida con amigos"


def test_add_two_rows_appends_in_order(fresh_book: Path):
    add_movement(fresh_book, date(2026, 5, 1), "Ingreso", "Nómina", 2100.0, "Mayo")
    add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Restauración", 30.0, "Comida")
    rows = _read_data_rows(fresh_book)
    assert len(rows) == 2
    assert rows[0][1] == "Ingreso"
    assert rows[1][1] == "Gasto"


def test_add_rejects_invalid_category(fresh_book: Path):
    with pytest.raises(InvalidMovementError, match="categoría"):
        add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Inexistente", 10.0, "")


def test_add_rejects_wrong_pair(fresh_book: Path):
    # Nómina is only valid for Ingreso
    with pytest.raises(InvalidMovementError, match="categoría"):
        add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Nómina", 10.0, "")


def test_add_rejects_invalid_tipo(fresh_book: Path):
    with pytest.raises(InvalidMovementError, match="tipo"):
        add_movement(fresh_book, date(2026, 5, 9), "Otro", "Nómina", 10.0, "")


def test_add_rejects_negative_importe(fresh_book: Path):
    with pytest.raises(InvalidMovementError, match="importe"):
        add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Otros", -5.0, "")


def test_edit_last_changes_specified_fields(fresh_book: Path):
    add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Restauración", 30.0, "Comida")
    edit_last(fresh_book, importe=35.0, descripcion="Comida (corregido)")
    rows = _read_data_rows(fresh_book)
    assert rows[0][3] == 35.0
    assert rows[0][4] == "Comida (corregido)"
    # Untouched fields remain
    assert rows[0][1] == "Gasto"
    assert rows[0][2] == "Restauración"


def test_edit_last_on_empty_book_raises(fresh_book: Path):
    with pytest.raises(InvalidMovementError, match="vacía|empty"):
        edit_last(fresh_book, importe=10.0)


def test_delete_last_removes_only_last_row(fresh_book: Path):
    add_movement(fresh_book, date(2026, 5, 1), "Ingreso", "Nómina", 2100.0, "Mayo")
    add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Restauración", 30.0, "Comida")
    delete_last(fresh_book)
    rows = _read_data_rows(fresh_book)
    assert len(rows) == 1
    assert rows[0][1] == "Ingreso"


def test_delete_last_on_empty_book_raises(fresh_book: Path):
    with pytest.raises(InvalidMovementError, match="vacía|empty"):
        delete_last(fresh_book)
