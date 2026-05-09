"""Append / edit-last / delete-last movements in data/movements.json + regenerate dashboard.html."""
from __future__ import annotations

import argparse
from datetime import date, datetime
from pathlib import Path

from scripts.categories import VALID_TYPES
from scripts.movements_store import (
    append as store_append,
    edit_last as store_edit_last,
    delete_last as store_delete_last,
    InvalidMovementError,
)
from scripts.build_dashboard import render, DEFAULT_DATA_PATH


def _parse_fecha(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Añade, edita o borra el último movimiento."
    )
    parser.add_argument("--path", default=str(DEFAULT_DATA_PATH))
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--edit-last", action="store_true", dest="edit_last_flag")
    mode.add_argument("--delete-last", action="store_true", dest="delete_last_flag")
    parser.add_argument("--fecha", type=_parse_fecha)
    parser.add_argument("--tipo", choices=VALID_TYPES)
    parser.add_argument("--categoria")
    parser.add_argument("--importe", type=float)
    parser.add_argument("--descripcion", default=None)

    args = parser.parse_args()
    path = Path(args.path)

    if args.delete_last_flag:
        store_delete_last(path)
        if path == DEFAULT_DATA_PATH:
            render()
        print("OK: última fila borrada. Refresca dashboard.html (F5).")
        return

    if args.edit_last_flag:
        store_edit_last(
            path,
            fecha=args.fecha,
            tipo=args.tipo,
            categoria=args.categoria,
            importe=args.importe,
            descripcion=args.descripcion,
        )
        if path == DEFAULT_DATA_PATH:
            render()
        print("OK: última fila actualizada. Refresca dashboard.html (F5).")
        return

    missing = [
        f
        for f, v in [
            ("--fecha", args.fecha),
            ("--tipo", args.tipo),
            ("--categoria", args.categoria),
            ("--importe", args.importe),
        ]
        if v is None
    ]
    if missing:
        parser.error(f"faltan flags requeridos para añadir: {', '.join(missing)}")

    new_id = store_append(
        path,
        fecha=args.fecha,
        tipo=args.tipo,
        categoria=args.categoria,
        importe=args.importe,
        descripcion=args.descripcion or "",
    )
    if path == DEFAULT_DATA_PATH:
        render()
    print(f"OK: movimiento #{new_id} añadido. Refresca dashboard.html (F5).")


if __name__ == "__main__":
    main()
