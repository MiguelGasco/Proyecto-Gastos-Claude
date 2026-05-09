"""Build gastos.xlsx from scratch."""
from pathlib import Path
import argparse
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from scripts.categories import CATEGORIES, VALID_TYPES


SHEET_ORDER = ["Dashboard", "Movimientos", "Categorías", "_aux"]

MOV_HEADERS = ["Fecha", "Tipo", "Categoría", "Importe", "Descripción"]
MOV_LAST_DATA_ROW = 1000  # pre-size the table for data validations and table ref


def _build_movimientos(ws) -> None:
    for col_idx, header in enumerate(MOV_HEADERS, start=1):
        ws.cell(row=1, column=col_idx, value=header)

    last_col = get_column_letter(len(MOV_HEADERS))
    table = Table(
        displayName="tblMov",
        ref=f"A1:{last_col}{MOV_LAST_DATA_ROW}",
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showRowStripes=True,
    )
    ws.add_table(table)

    # Number/date formats on data rows
    for row in range(2, MOV_LAST_DATA_ROW + 1):
        ws.cell(row=row, column=1).number_format = "dd/mm/yyyy"
        ws.cell(row=row, column=4).number_format = '#,##0.00 "€"'

    # Sensible column widths
    widths = {"A": 12, "B": 10, "C": 16, "D": 14, "E": 40}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def _build_categorias(ws) -> None:
    ws.cell(row=1, column=1, value="Tipo")
    ws.cell(row=1, column=2, value="Categoría")
    r = 2
    for tipo, cats in CATEGORIES.items():
        for cat in cats:
            ws.cell(row=r, column=1, value=tipo)
            ws.cell(row=r, column=2, value=cat)
            r += 1
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 18


def _add_validations(ws_mov) -> None:
    # Tipo: list of two values
    dv_tipo = DataValidation(
        type="list",
        formula1=f'"{",".join(VALID_TYPES)}"',
        allow_blank=True,
    )
    dv_tipo.add(f"B2:B{MOV_LAST_DATA_ROW}")
    ws_mov.add_data_validation(dv_tipo)

    # Categoría: union of all categories (Excel doesn't easily do dependent dropdowns
    # via openpyxl; we accept any valid category and rely on add_movement.py to
    # enforce the (tipo, categoria) pairing).
    all_cats = sorted({c for cats in CATEGORIES.values() for c in cats})
    dv_cat = DataValidation(
        type="list",
        formula1=f'"{",".join(all_cats)}"',
        allow_blank=True,
    )
    dv_cat.add(f"C2:C{MOV_LAST_DATA_ROW}")
    ws_mov.add_data_validation(dv_cat)


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

    _build_movimientos(wb["Movimientos"])
    _build_categorias(wb["Categorías"])
    _add_validations(wb["Movimientos"])

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
