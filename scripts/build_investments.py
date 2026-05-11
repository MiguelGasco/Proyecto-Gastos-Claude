"""Render investments.html from operations + quotes."""
from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path

from scripts.investments_store import load as load_investments


PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "investments.template.html"
DEFAULT_INVESTMENTS_PATH = PROJECT_ROOT / "data" / "investments.json"
DEFAULT_QUOTES_PATH = PROJECT_ROOT / "data" / "quotes.json"
DEFAULT_OUTPUT_PATH = PROJECT_ROOT / "investments.html"


def render(
    investments_path: Path = DEFAULT_INVESTMENTS_PATH,
    quotes_path: Path = DEFAULT_QUOTES_PATH,
    output_path: Path = DEFAULT_OUTPUT_PATH,
    template_path: Path = TEMPLATE_PATH,
) -> Path:
    template = template_path.read_text(encoding="utf-8")

    operations = []
    deposits = []
    if investments_path.exists():
        store = load_investments(investments_path)
        operations = store.get("operations", [])
        deposits = store.get("deposits", [])

    quotes = {}
    quotes_updated = ""
    if quotes_path.exists():
        with open(quotes_path, encoding="utf-8") as f:
            qstore = json.load(f)
        quotes = qstore.get("quotes", {})
        quotes_updated = qstore.get("updated_at", "")

    operations_json = json.dumps(operations, ensure_ascii=False)
    deposits_json = json.dumps(deposits, ensure_ascii=False)
    quotes_json = json.dumps(quotes, ensure_ascii=False)
    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = (
        template
        .replace("__OPERATIONS_JSON__", operations_json)
        .replace("__DEPOSITS_JSON__", deposits_json)
        .replace("__QUOTES_JSON__", quotes_json)
        .replace("__QUOTES_UPDATED__", quotes_updated)
        .replace("__GENERATED_AT__", generated_at)
    )
    output_path.write_text(html, encoding="utf-8")
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Genera investments.html desde data/investments.json + data/quotes.json y el template."
    )
    parser.add_argument("--investments", default=str(DEFAULT_INVESTMENTS_PATH))
    parser.add_argument("--quotes", default=str(DEFAULT_QUOTES_PATH))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_PATH))
    parser.add_argument("--template", default=str(TEMPLATE_PATH))
    args = parser.parse_args()
    out = render(
        Path(args.investments),
        Path(args.quotes),
        Path(args.output),
        Path(args.template),
    )
    print(f"OK: {out}")


if __name__ == "__main__":
    main()
