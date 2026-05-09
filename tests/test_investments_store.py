"""Tests for scripts.investments_store."""
import json
from pathlib import Path

import pytest

from scripts.investments_store import (
    load,
    save,
    append,
    edit_last,
    delete_last,
    InvalidOperationError,
)


@pytest.fixture()
def fresh_store(tmp_path: Path) -> Path:
    out = tmp_path / "investments.json"
    save(out, {"version": 1, "operations": []})
    return out


def test_load_returns_versioned_dict(fresh_store):
    s = load(fresh_store)
    assert s["version"] == 1
    assert s["operations"] == []


def test_load_missing_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        load(tmp_path / "nope.json")


def test_append_adds_operation_with_id(fresh_store):
    new_id = append(
        fresh_store,
        fecha="2026-01-15",
        tipo="compra",
        ticker="AAPL",
        nombre="Apple Inc.",
        unidades=5.0,
        precio=165.00,
        comision=1.00,
        moneda="EUR",
    )
    assert new_id == 1
    s = load(fresh_store)
    assert len(s["operations"]) == 1
    op = s["operations"][0]
    assert op["id"] == 1
    assert op["fecha"] == "2026-01-15"
    assert op["tipo"] == "compra"
    assert op["ticker"] == "AAPL"
    assert op["nombre"] == "Apple Inc."
    assert op["unidades"] == 5.0
    assert op["precio"] == 165.0
    assert op["comision"] == 1.0
    assert op["moneda"] == "EUR"


def test_append_increments_id(fresh_store):
    append(fresh_store, "2026-01-15", "compra", "AAPL", "Apple Inc.", 5.0, 165.0)
    new_id = append(fresh_store, "2026-02-10", "compra", "NVDA", "NVIDIA Corp.", 10.0, 875.0)
    assert new_id == 2


def test_append_rejects_invalid_tipo(fresh_store):
    with pytest.raises(InvalidOperationError, match="tipo"):
        append(fresh_store, "2026-01-15", "donación", "AAPL", "Apple Inc.", 5.0, 165.0)


def test_append_rejects_zero_or_negative_unidades(fresh_store):
    with pytest.raises(InvalidOperationError, match="unidades"):
        append(fresh_store, "2026-01-15", "compra", "AAPL", "Apple Inc.", 0.0, 165.0)
    with pytest.raises(InvalidOperationError, match="unidades"):
        append(fresh_store, "2026-01-15", "compra", "AAPL", "Apple Inc.", -1.0, 165.0)


def test_append_rejects_negative_precio(fresh_store):
    with pytest.raises(InvalidOperationError, match="precio"):
        append(fresh_store, "2026-01-15", "compra", "AAPL", "Apple Inc.", 5.0, 0.0)
    with pytest.raises(InvalidOperationError, match="precio"):
        append(fresh_store, "2026-01-15", "compra", "AAPL", "Apple Inc.", 5.0, -10.0)


def test_edit_last_changes_specified_fields(fresh_store):
    append(fresh_store, "2026-01-15", "compra", "AAPL", "Apple Inc.", 5.0, 165.0)
    edit_last(fresh_store, precio=170.0, nombre="Apple Inc. (corrected)")
    s = load(fresh_store)
    op = s["operations"][0]
    assert op["precio"] == 170.0
    assert op["nombre"] == "Apple Inc. (corrected)"
    assert op["tipo"] == "compra"
    assert op["ticker"] == "AAPL"
    assert op["unidades"] == 5.0


def test_edit_last_on_empty_raises(fresh_store):
    with pytest.raises(InvalidOperationError, match="vacía|empty"):
        edit_last(fresh_store, precio=100.0)


def test_delete_last_pops_only_last(fresh_store):
    append(fresh_store, "2026-01-15", "compra", "AAPL", "Apple Inc.", 5.0, 165.0)
    append(fresh_store, "2026-02-10", "compra", "NVDA", "NVIDIA Corp.", 10.0, 875.0)
    delete_last(fresh_store)
    s = load(fresh_store)
    assert len(s["operations"]) == 1
    assert s["operations"][0]["ticker"] == "AAPL"


def test_delete_last_on_empty_raises(fresh_store):
    with pytest.raises(InvalidOperationError, match="vacía|empty"):
        delete_last(fresh_store)


def test_save_creates_parent_dir(tmp_path):
    out = tmp_path / "subdir" / "deeper" / "investments.json"
    save(out, {"version": 1, "operations": []})
    assert out.exists()
    data = json.loads(out.read_text(encoding="utf-8"))
    assert data == {"version": 1, "operations": []}
