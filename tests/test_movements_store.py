from datetime import date
from pathlib import Path
import json
import pytest

from scripts.movements_store import (
    load,
    save,
    append,
    edit_last,
    delete_last,
    InvalidMovementError,
)


@pytest.fixture()
def fresh_store(tmp_path: Path) -> Path:
    out = tmp_path / "movements.json"
    save(out, {"version": 1, "movements": []})
    return out


def test_load_returns_versioned_dict(fresh_store):
    s = load(fresh_store)
    assert s["version"] == 1
    assert s["movements"] == []


def test_load_missing_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        load(tmp_path / "nope.json")


def test_append_adds_movement_with_id(fresh_store):
    new_id = append(fresh_store, date(2026, 5, 9), "Gasto", "Restauración", 30.0, "Comida")
    assert new_id == 1
    s = load(fresh_store)
    assert len(s["movements"]) == 1
    m = s["movements"][0]
    assert m["id"] == 1
    assert m["fecha"] == "2026-05-09"
    assert m["tipo"] == "Gasto"
    assert m["categoria"] == "Restauración"
    assert m["importe"] == 30.0
    assert m["descripcion"] == "Comida"


def test_append_increments_id(fresh_store):
    append(fresh_store, date(2026, 5, 1), "Ingreso", "Nómina", 2100.0, "Mayo")
    new_id = append(fresh_store, date(2026, 5, 9), "Gasto", "Restauración", 30.0, "Comida")
    assert new_id == 2


def test_append_rejects_invalid_category(fresh_store):
    with pytest.raises(InvalidMovementError, match="categoría"):
        append(fresh_store, date(2026, 5, 9), "Gasto", "Inexistente", 10.0, "")


def test_append_rejects_wrong_pair(fresh_store):
    with pytest.raises(InvalidMovementError, match="categoría"):
        append(fresh_store, date(2026, 5, 9), "Gasto", "Nómina", 10.0, "")


def test_append_rejects_invalid_tipo(fresh_store):
    with pytest.raises(InvalidMovementError, match="tipo"):
        append(fresh_store, date(2026, 5, 9), "Otro", "Nómina", 10.0, "")


def test_append_rejects_zero_importe(fresh_store):
    with pytest.raises(InvalidMovementError, match="importe"):
        append(fresh_store, date(2026, 5, 9), "Gasto", "Otros", 0.0, "")


def test_append_rejects_negative_importe(fresh_store):
    with pytest.raises(InvalidMovementError, match="importe"):
        append(fresh_store, date(2026, 5, 9), "Gasto", "Otros", -5.0, "")


def test_edit_last_changes_specified_fields(fresh_store):
    append(fresh_store, date(2026, 5, 9), "Gasto", "Restauración", 30.0, "Comida")
    edit_last(fresh_store, importe=35.0, descripcion="Comida (corregido)")
    s = load(fresh_store)
    m = s["movements"][0]
    assert m["importe"] == 35.0
    assert m["descripcion"] == "Comida (corregido)"
    assert m["tipo"] == "Gasto"
    assert m["categoria"] == "Restauración"


def test_edit_last_on_empty_raises(fresh_store):
    with pytest.raises(InvalidMovementError, match="vacía|empty"):
        edit_last(fresh_store, importe=10.0)


def test_delete_last_pops_only_last(fresh_store):
    append(fresh_store, date(2026, 5, 1), "Ingreso", "Nómina", 2100.0, "Mayo")
    append(fresh_store, date(2026, 5, 9), "Gasto", "Restauración", 30.0, "Comida")
    delete_last(fresh_store)
    s = load(fresh_store)
    assert len(s["movements"]) == 1
    assert s["movements"][0]["tipo"] == "Ingreso"


def test_delete_last_on_empty_raises(fresh_store):
    with pytest.raises(InvalidMovementError, match="vacía|empty"):
        delete_last(fresh_store)


def test_save_creates_parent_dir(tmp_path):
    out = tmp_path / "subdir" / "deeper" / "movements.json"
    save(out, {"version": 1, "movements": []})
    assert out.exists()
    data = json.loads(out.read_text(encoding="utf-8"))
    assert data == {"version": 1, "movements": []}
