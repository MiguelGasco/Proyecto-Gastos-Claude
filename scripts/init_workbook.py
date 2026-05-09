"""Build gastos.xlsx from scratch."""
from pathlib import Path
import argparse
import openpyxl


SHEET_ORDER = ["Dashboard", "Movimientos", "Categorías", "_aux"]


def create_workbook(path: Path) -> None:
    path = Path(path)
    if path.exists():
        raise FileExistsError(
            f"{path} ya existe. Borralo a mano si querés re-crearlo desde cero."
        )

    wb = openpyxl.Workbook()
    # Workbook starts with one default sheet — rename it and add the rest.
    default = wb.active
    default.title = SHEET_ORDER[0]
    for name in SHEET_ORDER[1:]:
        wb.create_sheet(name)

    wb.save(path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Crea gastos.xlsx desde cero.")
    parser.add_argument(
        "--path",
        default="gastos.xlsx",
        help="Ruta del archivo de salida (por defecto: gastos.xlsx en el cwd).",
    )
    args = parser.parse_args()
    create_workbook(Path(args.path))
    print(f"Creado: {args.path}")


if __name__ == "__main__":
    main()
