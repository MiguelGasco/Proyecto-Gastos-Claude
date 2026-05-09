"""Build gastos.xlsx from scratch."""
from pathlib import Path
import argparse
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.chart import PieChart, LineChart, BarChart, Reference
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


def _build_aux(ws) -> None:
    # --- B1: selected month (driven by Dashboard!B1, which user picks) ---
    ws["A1"] = "Mes seleccionado"
    ws["B1"] = "=Dashboard!B1"
    ws["B1"].number_format = "dd/mm/yyyy"
    ws["A2"] = "Etiqueta"
    ws["B2"] = '=TEXT(B1,"mm/yyyy")'

    # --- KPIs ---
    # Date range filter: >= B1 (first of month) and < EOMonth(B1)+1
    month_start = "$B$1"
    month_end_excl = 'DATE(YEAR($B$1),MONTH($B$1)+1,1)'

    ws["C1"] = "Ingresos mes"
    ws["D1"] = (
        f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Ingreso",'
        f'tblMov[Fecha],">="&{month_start},'
        f'tblMov[Fecha],"<"&{month_end_excl})'
    )
    ws["C2"] = "Gastos mes"
    ws["D2"] = (
        f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
        f'tblMov[Fecha],">="&{month_start},'
        f'tblMov[Fecha],"<"&{month_end_excl})'
    )
    ws["C3"] = "Ahorro neto"
    ws["D3"] = "=D1-D2"
    ws["C4"] = "% ahorro"
    ws["D4"] = "=IFERROR(D3/D1,0)"

    for r in (1, 2, 3):
        ws.cell(row=r, column=4).number_format = '#,##0.00 "€"'
    ws["D4"].number_format = "0.0%"

    # --- Category-of-month block (F1:G11) ---
    ws["F1"] = "Categoría"
    ws["G1"] = "Total mes"
    for i, cat in enumerate(CATEGORIES["Gasto"], start=2):
        ws.cell(row=i, column=6, value=cat)
        ws.cell(
            row=i,
            column=7,
            value=(
                f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
                f'tblMov[Categoría],F{i},'
                f'tblMov[Fecha],">="&{month_start},'
                f'tblMov[Fecha],"<"&{month_end_excl})'
            ),
        )
        ws.cell(row=i, column=7).number_format = '#,##0.00 "€"'

    # --- Monthly evolution block (I1:M13), rolling 12 months oldest→newest ---
    ws["I1"] = "Mes"
    ws["J1"] = "Ingresos"
    ws["K1"] = "Gastos"
    ws["L1"] = "Ahorro"
    ws["M1"] = "% Ahorro"
    for offset in range(12):
        r = 2 + offset
        # offset 0 = oldest (11 months back), offset 11 = current (selected) month
        months_back = 11 - offset
        ws.cell(
            row=r,
            column=9,
            value=f'=TEXT(DATE(YEAR($B$1),MONTH($B$1)-{months_back},1),"mm/yyyy")',
        )
        start_ref = f'DATE(YEAR($B$1),MONTH($B$1)-{months_back},1)'
        end_ref = f'DATE(YEAR($B$1),MONTH($B$1)-{months_back}+1,1)'
        ws.cell(
            row=r,
            column=10,
            value=(
                f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Ingreso",'
                f'tblMov[Fecha],">="&{start_ref},'
                f'tblMov[Fecha],"<"&{end_ref})'
            ),
        )
        ws.cell(
            row=r,
            column=11,
            value=(
                f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
                f'tblMov[Fecha],">="&{start_ref},'
                f'tblMov[Fecha],"<"&{end_ref})'
            ),
        )
        ws.cell(row=r, column=12, value=f"=J{r}-K{r}")
        ws.cell(row=r, column=13, value=f"=IFERROR(L{r}/J{r},0)")
        for c in (10, 11, 12):
            ws.cell(row=r, column=c).number_format = '#,##0.00 "€"'
        ws.cell(row=r, column=13).number_format = "0.0%"

    # --- Top-5 block (N1:P6). Uses FILTER+SORT+INDEX for correct ranking ---
    ws["N1"] = "Fecha"
    ws["O1"] = "Descripción"
    ws["P1"] = "Importe"
    filter_base = (
        'FILTER(Movimientos!$A$2:$E$1000,'
        '(Movimientos!$B$2:$B$1000="Gasto")'
        '*(Movimientos!$A$2:$A$1000>=$B$1)'
        f'*(Movimientos!$A$2:$A$1000<{month_end_excl}))'
    )
    sorted_expr = f'SORT({filter_base},4,-1)'
    for i in range(5):
        r = 2 + i
        rank = i + 1
        # Fecha (N): column 1 of sorted result
        ws.cell(
            row=r,
            column=14,
            value=f'=IFERROR(INDEX({sorted_expr},{rank},1),"")',
        )
        # Descripción (O): column 5 of sorted result
        ws.cell(
            row=r,
            column=15,
            value=f'=IFERROR(INDEX({sorted_expr},{rank},5),"")',
        )
        # Importe (P): column 4 of sorted result
        ws.cell(
            row=r,
            column=16,
            value=f'=IFERROR(INDEX({sorted_expr},{rank},4),"")',
        )
        ws.cell(row=r, column=14).number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=16).number_format = '#,##0.00 "€"'

    # --- Cumulative daily expenses block (R1:S32) ---
    ws["R1"] = "Día"
    ws["S1"] = "Acumulado"
    for i in range(31):
        r = 2 + i
        ws.cell(row=r, column=18, value=f"=$B$1+{i}")
        ws.cell(row=r, column=18).number_format = "dd/mm"
        ws.cell(
            row=r,
            column=19,
            value=(
                f'=IF(R{r}>=DATE(YEAR($B$1),MONTH($B$1)+1,1),NA(),'
                f'SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
                f'tblMov[Fecha],">="&$B$1,'
                f'tblMov[Fecha],"<="&R{r}))'
            ),
        )
        ws.cell(row=r, column=19).number_format = '#,##0.00 "€"'

    # --- Category comparison: this month vs previous month (U1:W11) ---
    ws["U1"] = "Categoría"
    ws["V1"] = "Este mes"
    ws["W1"] = "Mes anterior"
    prev_start = 'DATE(YEAR($B$1),MONTH($B$1)-1,1)'
    for i, cat in enumerate(CATEGORIES["Gasto"], start=2):
        ws.cell(row=i, column=21, value=cat)
        ws.cell(
            row=i,
            column=22,
            value=(
                f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
                f'tblMov[Categoría],U{i},'
                f'tblMov[Fecha],">="&{month_start},'
                f'tblMov[Fecha],"<"&{month_end_excl})'
            ),
        )
        ws.cell(
            row=i,
            column=23,
            value=(
                f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
                f'tblMov[Categoría],U{i},'
                f'tblMov[Fecha],">="&{prev_start},'
                f'tblMov[Fecha],"<"&{month_start})'
            ),
        )
        ws.cell(row=i, column=22).number_format = '#,##0.00 "€"'
        ws.cell(row=i, column=23).number_format = '#,##0.00 "€"'

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["O"].width = 30


def _build_dashboard(ws) -> None:
    # --- Month selector ---
    ws["A1"] = "Mes seleccionado:"
    ws["A1"].font = Font(bold=True)
    ws["B1"] = "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)"
    ws["B1"].number_format = "dd/mm/yyyy"
    ws["B1"].fill = PatternFill("solid", fgColor="FFF2CC")  # soft yellow = editable

    # --- KPI labels and values ---
    kpi_rows = [
        ("Ingresos:", "=_aux!D1", '#,##0.00 "€"'),
        ("Gastos:", "=_aux!D2", '#,##0.00 "€"'),
        ("Ahorro neto:", "=_aux!D3", '#,##0.00 "€"'),
        ("% Ahorro:", "=_aux!D4", "0.0%"),
    ]
    for i, (label, formula, fmt) in enumerate(kpi_rows, start=1):
        ws.cell(row=i, column=3, value=label).font = Font(bold=True)
        c = ws.cell(row=i, column=4, value=formula)
        c.number_format = fmt
        c.font = Font(size=14, bold=True)

    # --- Top-5 section ---
    ws["A6"] = "TOP 5 GASTOS DEL MES"
    ws["A6"].font = Font(bold=True, size=12)
    ws["A7"] = "Fecha"
    ws["B7"] = "Descripción"
    ws["C7"] = "Importe"
    for cell in (ws["A7"], ws["B7"], ws["C7"]):
        cell.font = Font(bold=True)
    for i in range(5):
        r = 8 + i
        ws.cell(row=r, column=1, value=f"=_aux!N{2 + i}").number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=2, value=f"=_aux!O{2 + i}")
        ws.cell(row=r, column=3, value=f"=_aux!P{2 + i}").number_format = '#,##0.00 "€"'

    # Column widths
    widths = {"A": 18, "B": 30, "C": 16, "D": 18}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def _add_charts(wb) -> None:
    aux = wb["_aux"]
    dash = wb["Dashboard"]

    # --- (1) Pie chart: gastos por categoría del mes seleccionado ---
    pie = PieChart()
    pie.title = "Gastos por categoría (mes)"
    labels = Reference(aux, min_col=6, min_row=2, max_row=11)        # F2:F11
    data = Reference(aux, min_col=7, min_row=1, max_row=11)          # G1:G11 (incl header)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.height = 9   # cm
    pie.width = 14
    dash.add_chart(pie, "F1")

    # --- (2) Line chart: evolución mensual últimos 12 meses (oldest→newest) ---
    # _aux!I2:I13 = Mes (categories), J2:J13 / K2:K13 / L2:L13 = Ingresos / Gastos / Ahorro
    line = LineChart()
    line.title = "Evolución mensual (últimos 12 meses)"
    cats = Reference(aux, min_col=9, min_row=2, max_row=13)          # I2:I13
    series_data = Reference(aux, min_col=10, max_col=12, min_row=1, max_row=13)
    line.add_data(series_data, titles_from_data=True)
    line.set_categories(cats)
    line.height = 9
    line.width = 18
    dash.add_chart(line, "F18")

    # --- (3) Line chart: gasto acumulado del mes ---
    acum = LineChart()
    acum.title = "Gasto acumulado del mes"
    acum_cats = Reference(aux, min_col=18, min_row=2, max_row=32)    # R2:R32
    acum_data = Reference(aux, min_col=19, min_row=1, max_row=32)    # S1:S32 (header included)
    acum.add_data(acum_data, titles_from_data=True)
    acum.set_categories(acum_cats)
    acum.height = 9
    acum.width = 18
    dash.add_chart(acum, "F35")

    # --- (4) Bar chart: comparativa categorías este mes vs mes anterior ---
    comp = BarChart()
    comp.type = "bar"
    comp.style = 11
    comp.title = "Categorías: este mes vs mes anterior"
    comp_cats = Reference(aux, min_col=21, min_row=2, max_row=11)    # U2:U11
    comp_data = Reference(aux, min_col=22, max_col=23, min_row=1, max_row=11)  # V1:W11
    comp.add_data(comp_data, titles_from_data=True)
    comp.set_categories(comp_cats)
    comp.height = 11
    comp.width = 18
    dash.add_chart(comp, "F52")

    # --- (5) Line chart: % ahorro mensual 12 meses ---
    pct = LineChart()
    pct.title = "% Ahorro mensual (12 meses)"
    pct_cats = Reference(aux, min_col=9, min_row=2, max_row=13)      # I2:I13 (months)
    pct_data = Reference(aux, min_col=13, min_row=1, max_row=13)     # M1:M13 (% ahorro incl header)
    pct.add_data(pct_data, titles_from_data=True)
    pct.set_categories(pct_cats)
    pct.height = 9
    pct.width = 18
    dash.add_chart(pct, "F70")


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
    _build_aux(wb["_aux"])
    _build_dashboard(wb["Dashboard"])
    _add_charts(wb)

    wb["_aux"].sheet_state = "hidden"
    wb.active = wb.sheetnames.index("Dashboard")

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
