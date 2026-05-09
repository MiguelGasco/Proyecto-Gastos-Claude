"""Tests for scripts.add_movement (JSON-backed)."""
from datetime import date
from pathlib import Path
import json
import subprocess
import sys
import pytest

from scripts.movements_store import save


@pytest.fixture()
def fresh_setup(tmp_path: Path, monkeypatch) -> tuple:
    """Return (data_path, project_dir) with empty movements.json and project structure."""
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    data_path = data_dir / "movements.json"
    save(data_path, {"version": 1, "movements": []})
    output_dir = tmp_path
    return data_path, output_dir


def _run_add(data_path: Path, output_path, *args: str) -> subprocess.CompletedProcess:
    """Run add_movement.py as a subprocess against the test data path."""
    project_root = Path(__file__).resolve().parent.parent
    cmd = [sys.executable, "-m", "scripts.add_movement",
           "--path", str(data_path), *args]
    return subprocess.run(cmd, capture_output=True, text=True, cwd=str(project_root), env={
        **__import__("os").environ,
    })


def _read_movements(data_path: Path) -> list:
    return json.loads(data_path.read_text(encoding="utf-8"))["movements"]


def test_add_appends_a_row(fresh_setup):
    data_path, _ = fresh_setup
    r = _run_add(data_path, _,
                 "--fecha", "2026-05-09",
                 "--tipo", "Gasto",
                 "--categoria", "Restauración",
                 "--importe", "30",
                 "--descripcion", "Comida con amigos")
    assert r.returncode == 0, r.stderr
    movs = _read_movements(data_path)
    assert len(movs) == 1
    assert movs[0]["fecha"] == "2026-05-09"
    assert movs[0]["tipo"] == "Gasto"
    assert movs[0]["categoria"] == "Restauración"
    assert movs[0]["importe"] == 30.0
    assert movs[0]["descripcion"] == "Comida con amigos"


def test_add_rejects_invalid_category(fresh_setup):
    data_path, _ = fresh_setup
    r = _run_add(data_path, _,
                 "--fecha", "2026-05-09",
                 "--tipo", "Gasto",
                 "--categoria", "NoExiste",
                 "--importe", "10")
    assert r.returncode != 0
    assert "categoría" in r.stderr.lower() or "categoría" in r.stdout.lower()


def test_add_missing_required_flags(fresh_setup):
    data_path, _ = fresh_setup
    r = _run_add(data_path, _,
                 "--tipo", "Gasto",
                 "--categoria", "Otros",
                 "--importe", "10")  # missing --fecha
    assert r.returncode != 0


def test_edit_last_updates_fields(fresh_setup):
    data_path, _ = fresh_setup
    _run_add(data_path, _, "--fecha", "2026-05-09", "--tipo", "Gasto",
             "--categoria", "Restauración", "--importe", "30")
    r = _run_add(data_path, _, "--edit-last", "--importe", "35")
    assert r.returncode == 0, r.stderr
    movs = _read_movements(data_path)
    assert movs[0]["importe"] == 35.0


def test_delete_last_removes_row(fresh_setup):
    data_path, _ = fresh_setup
    _run_add(data_path, _, "--fecha", "2026-05-09", "--tipo", "Gasto",
             "--categoria", "Restauración", "--importe", "30")
    r = _run_add(data_path, _, "--delete-last")
    assert r.returncode == 0, r.stderr
    movs = _read_movements(data_path)
    assert movs == []
