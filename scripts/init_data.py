"""Initialise the movements and investments JSON stores plus rendered HTML pages."""
from __future__ import annotations

import argparse
from pathlib import Path

from scripts.movements_store import save as save_movements
from scripts.investments_store import save as save_investments
from scripts.build_dashboard import render as render_dashboard, DEFAULT_DATA_PATH, DEFAULT_OUTPUT_PATH
from scripts.build_investments import (
    render as render_investments,
    DEFAULT_INVESTMENTS_PATH,
    DEFAULT_QUOTES_PATH,
)


def init_data(path: Path = DEFAULT_DATA_PATH, force: bool = False) -> int:
    """Create movements.json (empty or from legacy xlsx). Returns count of imported movements."""
    if path.exists() and not force:
        raise FileExistsError(
            f"{path} ya existe. Borralo a mano si querés re-crearlo desde cero."
        )

    movements = []
    legacy = path.parent.parent / "gastos.xlsx"
    if legacy.exists():
        try:
            import openpyxl

            wb = openpyxl.load_workbook(legacy, data_only=False)
            ws = wb["Movimientos"]
            for i, row in enumerate(
                ws.iter_rows(min_row=2, max_col=5, values_only=True), start=1
            ):
                fecha, tipo, cat, imp, desc = row
                if fecha is None:
                    break
                fecha_iso = (
                    fecha.date().isoformat()
                    if hasattr(fecha, "date")
                    else str(fecha)
                )
                movements.append(
                    {
                        "id": i,
                        "fecha": fecha_iso,
                        "tipo": tipo,
                        "categoria": cat,
                        "importe": float(imp) if imp is not None else 0.0,
                        "descripcion": desc or "",
                    }
                )
        except (ImportError, KeyError):
            pass

    save_movements(path, {"version": 1, "movements": movements})
    render_dashboard()

    # Initialise investments.json if it doesn't exist yet
    if not DEFAULT_INVESTMENTS_PATH.exists():
        save_investments(DEFAULT_INVESTMENTS_PATH, {"version": 1, "operations": []})

    # Initialise quotes.json if it doesn't exist yet
    if not DEFAULT_QUOTES_PATH.exists():
        import json
        DEFAULT_QUOTES_PATH.parent.mkdir(parents=True, exist_ok=True)
        DEFAULT_QUOTES_PATH.write_text(
            json.dumps({"version": 1, "updated_at": "", "quotes": {}}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    render_investments()

    return len(movements)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Inicializa data/movements.json, data/investments.json, dashboard.html e investments.html."
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Sobrescribir movements.json existente.",
    )
    args = parser.parse_args()
    n = init_data(force=args.force)
    if n > 0:
        print(
            f"Creado: data/movements.json con {n} movimientos importados desde gastos.xlsx."
        )
    else:
        print("Creado: data/movements.json vacío.")
    print("Creado: dashboard.html. Abrilo con doble clic.")
    print("Creado: investments.html. Abrilo con doble clic.")


if __name__ == "__main__":
    main()
