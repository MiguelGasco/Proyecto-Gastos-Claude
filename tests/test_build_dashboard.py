"""Tests for scripts.build_dashboard."""
from datetime import date
from pathlib import Path
import json
import re

from scripts.movements_store import save, append
from scripts.build_dashboard import render


def test_render_creates_html_file(tmp_path: Path):
    data = tmp_path / "movements.json"
    save(data, {"version": 1, "movements": []})
    out = tmp_path / "dashboard.html"
    template = Path(__file__).resolve().parent.parent / "templates" / "dashboard.template.html"
    result = render(data_path=data, output_path=out, template_path=template)
    assert result == out
    assert out.exists()
    assert out.read_text(encoding="utf-8").startswith("<!DOCTYPE html>")


def test_render_embeds_movements_json(tmp_path: Path):
    data = tmp_path / "movements.json"
    save(data, {"version": 1, "movements": []})
    append(data, date(2026, 5, 9), "Gasto", "Compras", 110.0, "Monitor")
    out = tmp_path / "dashboard.html"
    template = Path(__file__).resolve().parent.parent / "templates" / "dashboard.template.html"
    render(data_path=data, output_path=out, template_path=template)
    html = out.read_text(encoding="utf-8")

    # Find the MOVEMENTS embedded JSON
    match = re.search(r"const MOVEMENTS = (\[.*?\]);", html, re.DOTALL)
    assert match, "MOVEMENTS array not found in HTML"
    embedded = json.loads(match.group(1))
    assert len(embedded) == 1
    assert embedded[0]["descripcion"] == "Monitor"
    assert embedded[0]["importe"] == 110.0


def test_render_embeds_categories_json(tmp_path: Path):
    data = tmp_path / "movements.json"
    save(data, {"version": 1, "movements": []})
    out = tmp_path / "dashboard.html"
    template = Path(__file__).resolve().parent.parent / "templates" / "dashboard.template.html"
    render(data_path=data, output_path=out, template_path=template)
    html = out.read_text(encoding="utf-8")
    match = re.search(r"const CATEGORIES = (\{.*?\});", html, re.DOTALL)
    assert match
    embedded = json.loads(match.group(1))
    assert "Gasto" in embedded
    assert "Restauración" in embedded["Gasto"]


def test_render_loads_chartjs_cdn(tmp_path: Path):
    data = tmp_path / "movements.json"
    save(data, {"version": 1, "movements": []})
    out = tmp_path / "dashboard.html"
    template = Path(__file__).resolve().parent.parent / "templates" / "dashboard.template.html"
    render(data_path=data, output_path=out, template_path=template)
    html = out.read_text(encoding="utf-8")
    assert "chart.js" in html.lower()


def test_render_handles_missing_data_file(tmp_path: Path):
    # If movements.json doesn't exist, render should still succeed with empty list
    out = tmp_path / "dashboard.html"
    template = Path(__file__).resolve().parent.parent / "templates" / "dashboard.template.html"
    render(data_path=tmp_path / "missing.json", output_path=out, template_path=template)
    html = out.read_text(encoding="utf-8")
    assert "const MOVEMENTS = [];" in html or "const MOVEMENTS = []" in html
