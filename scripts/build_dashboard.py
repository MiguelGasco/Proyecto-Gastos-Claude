"""Render dashboard.html from movements.json + template."""
from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path

from scripts.movements_store import load
from scripts.categories import CATEGORIES


PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "dashboard.template.html"
DEFAULT_DATA_PATH = PROJECT_ROOT / "data" / "movements.json"
DEFAULT_OUTPUT_PATH = PROJECT_ROOT / "dashboard.html"


def render(
    data_path: Path = DEFAULT_DATA_PATH,
    output_path: Path = DEFAULT_OUTPUT_PATH,
    template_path: Path = TEMPLATE_PATH,
) -> Path:
    template = template_path.read_text(encoding="utf-8")
    if data_path.exists():
        store = load(data_path)
        movements = store.get("movements", [])
    else:
        movements = []
    movements_json = json.dumps(movements, ensure_ascii=False)
    categories_json = json.dumps(CATEGORIES, ensure_ascii=False)
    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")
    html = (
        template
        .replace("__MOVEMENTS_JSON__", movements_json)
        .replace("__CATEGORIES_JSON__", categories_json)
        .replace("__GENERATED_AT__", generated_at)
    )
    output_path.write_text(html, encoding="utf-8")
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Genera dashboard.html desde data/movements.json y el template."
    )
    parser.add_argument("--data", default=str(DEFAULT_DATA_PATH))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_PATH))
    parser.add_argument("--template", default=str(TEMPLATE_PATH))
    args = parser.parse_args()
    out = render(Path(args.data), Path(args.output), Path(args.template))
    print(f"OK: {out}")


if __name__ == "__main__":
    main()
