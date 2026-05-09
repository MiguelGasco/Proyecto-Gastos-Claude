"""Tests for scripts.build_investments."""
import json
from pathlib import Path

from scripts.investments_store import save as save_investments, append
from scripts.build_investments import render

TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "templates" / "investments.template.html"


def _make_investments(path: Path) -> Path:
    """Create a minimal investments.json at path."""
    save_investments(path, {"version": 1, "operations": []})
    return path


def _make_quotes(path: Path, quotes: dict = None, updated_at: str = "") -> Path:
    """Create a minimal quotes.json at path."""
    store = {
        "version": 1,
        "updated_at": updated_at,
        "quotes": quotes or {},
    }
    path.write_text(json.dumps(store, ensure_ascii=False), encoding="utf-8")
    return path


def test_render_creates_html_file(tmp_path: Path):
    inv = _make_investments(tmp_path / "investments.json")
    quotes = _make_quotes(tmp_path / "quotes.json")
    out = tmp_path / "investments.html"
    result = render(
        investments_path=inv,
        quotes_path=quotes,
        output_path=out,
        template_path=TEMPLATE_PATH,
    )
    assert result == out
    assert out.exists()
    assert out.read_text(encoding="utf-8").startswith("<!DOCTYPE html>")


def test_render_embeds_operations_json(tmp_path: Path):
    inv = tmp_path / "investments.json"
    _make_investments(inv)
    append(inv, "2026-01-15", "compra", "AAPL", "Apple Inc.", 5.0, 165.0, 1.0)
    quotes = _make_quotes(tmp_path / "quotes.json")
    out = tmp_path / "investments.html"
    render(
        investments_path=inv,
        quotes_path=quotes,
        output_path=out,
        template_path=TEMPLATE_PATH,
    )
    html = out.read_text(encoding="utf-8")
    # Template embeds: const OPERATIONS = __OPERATIONS_JSON__;
    assert "AAPL" in html
    assert "Apple Inc." in html
    # The placeholder must be replaced
    assert "__OPERATIONS_JSON__" not in html


def test_render_embeds_quotes_json(tmp_path: Path):
    inv = _make_investments(tmp_path / "investments.json")
    quotes = _make_quotes(
        tmp_path / "quotes.json",
        quotes={"AAPL": {"precio": 179.5, "moneda": "EUR", "cambio_pct_1d": 0.012}},
    )
    out = tmp_path / "investments.html"
    render(
        investments_path=inv,
        quotes_path=quotes,
        output_path=out,
        template_path=TEMPLATE_PATH,
    )
    html = out.read_text(encoding="utf-8")
    assert "179.5" in html
    assert "__QUOTES_JSON__" not in html


def test_render_embeds_quotes_updated_at(tmp_path: Path):
    inv = _make_investments(tmp_path / "investments.json")
    quotes = _make_quotes(
        tmp_path / "quotes.json",
        updated_at="10/05/2026 14:30",
    )
    out = tmp_path / "investments.html"
    render(
        investments_path=inv,
        quotes_path=quotes,
        output_path=out,
        template_path=TEMPLATE_PATH,
    )
    html = out.read_text(encoding="utf-8")
    assert "10/05/2026 14:30" in html
    assert "__QUOTES_UPDATED__" not in html


def test_render_loads_chartjs_cdn(tmp_path: Path):
    inv = _make_investments(tmp_path / "investments.json")
    quotes = _make_quotes(tmp_path / "quotes.json")
    out = tmp_path / "investments.html"
    render(
        investments_path=inv,
        quotes_path=quotes,
        output_path=out,
        template_path=TEMPLATE_PATH,
    )
    html = out.read_text(encoding="utf-8")
    assert "chart.js" in html.lower()


def test_render_handles_missing_files(tmp_path: Path):
    """render() should succeed with empty arrays when JSON files don't exist."""
    out = tmp_path / "investments.html"
    render(
        investments_path=tmp_path / "missing_investments.json",
        quotes_path=tmp_path / "missing_quotes.json",
        output_path=out,
        template_path=TEMPLATE_PATH,
    )
    html = out.read_text(encoding="utf-8")
    assert out.exists()
    assert "__OPERATIONS_JSON__" not in html
    assert "__QUOTES_JSON__" not in html
    # Operations should be an empty array
    assert "[]" in html
