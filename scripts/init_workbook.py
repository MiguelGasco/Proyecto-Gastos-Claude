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

# --- Skill-convention number formats ---
_FMT_CURRENCY = '#,##0.00 "€";(#,##0.00 "€");"-"'
_FMT_PERCENT  = '0.0%;(0.0%);"-"'
_FMT_DATE     = "dd/mm/yyyy"
_FMT_DAY      = "dd/mm"

# --- Skill-convention colors ---
_COLOR_BLUE  = "FF0000FF"   # hardcoded inputs
_COLOR_BLACK = "FF000000"   # formulas / calculations
_COLOR_GREEN = "FF008000"   # cross-sheet links


def _apply_font(cell, *, color: str = _COLOR_BLACK, size: int = 11, bold: bool = False) -> None:
    """Apply Arial font with skill color conventions."""
    cell.font = Font(name="Arial", size=size, bold=bold, color=color)


def _blue_input(cell) -> None:
    """Blue = hardcoded input the user will change."""
    _apply_font(cell, color=_COLOR_BLUE)


def _green_link(cell, *, size: int = 11, bold: bool = False) -> None:
    """Green = cross-sheet link within same workbook."""
    _apply_font(cell, color=_COLOR_GREEN, size=size, bold=bold)


def _black_formula(cell, *, size: int = 11, bold: bool = False) -> None:
    """Black = formula / calculation."""
    _apply_font(cell, color=_COLOR_BLACK, size=size, bold=bold)


def _currency(cell) -> None:
    cell.number_format = _FMT_CURRENCY


def _percent(cell) -> None:
    cell.number_format = _FMT_PERCENT


def _build_movimientos(ws) -> None:
    for col_idx, header in enumerate(MOV_HEADERS, start=1):
        c = ws.cell(row=1, column=col_idx, value=header)
        _black_formula(c, bold=True)

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

    # Number/date formats on data rows (font/color applied per row in add_movement.py)
    for row in range(2, MOV_LAST_DATA_ROW + 1):
        ws.cell(row=row, column=1).number_format = _FMT_DATE
        ws.cell(row=row, column=4).number_format = _FMT_CURRENCY

    # Sensible column widths
    widths = {"A": 12, "B": 10, "C": 16, "D": 14, "E": 40}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def _build_categorias(ws) -> None:
    for col_idx, header in enumerate(["Tipo", "Categoría"], start=1):
        c = ws.cell(row=1, column=col_idx, value=header)
        _black_formula(c, bold=True)
    r = 2
    for tipo, cats in CATEGORIES.items():
        for cat in cats:
            c1 = ws.cell(row=r, column=1, value=tipo)
            c2 = ws.cell(row=r, column=2, value=cat)
            _black_formula(c1)
            _black_formula(c2)
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
    _black_formula(ws["A1"])
    ws["A1"] = "Mes seleccionado"
    c = ws["B1"]
    c.value = "=Dashboard!B1"
    c.number_format = _FMT_DATE
    _black_formula(c)

    _black_formula(ws["A2"])
    ws["A2"] = "Etiqueta"
    c2 = ws["B2"]
    c2.value = '=TEXT(B1,"mm/yyyy")'
    _black_formula(c2)

    # --- KPIs ---
    # Date range filter: >= B1 (first of month) and < EOMonth(B1)+1
    month_start = "$B$1"
    month_end_excl = 'DATE(YEAR($B$1),MONTH($B$1)+1,1)'

    ws["C1"] = "Ingresos mes"
    _black_formula(ws["C1"])
    ws["D1"] = (
        f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Ingreso",'
        f'tblMov[Fecha],">="&{month_start},'
        f'tblMov[Fecha],"<"&{month_end_excl})'
    )
    _black_formula(ws["D1"])
    ws["C2"] = "Gastos mes"
    _black_formula(ws["C2"])
    ws["D2"] = (
        f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
        f'tblMov[Fecha],">="&{month_start},'
        f'tblMov[Fecha],"<"&{month_end_excl})'
    )
    _black_formula(ws["D2"])
    ws["C3"] = "Ahorro neto"
    _black_formula(ws["C3"])
    ws["D3"] = "=D1-D2"
    _black_formula(ws["D3"])
    ws["C4"] = "% ahorro"
    _black_formula(ws["C4"])
    ws["D4"] = "=IFERROR(D3/D1,0)"
    _black_formula(ws["D4"])

    for r in (1, 2, 3):
        ws.cell(row=r, column=4).number_format = _FMT_CURRENCY
    ws["D4"].number_format = _FMT_PERCENT

    # --- Category-of-month block (F1:G11) ---
    ws["F1"] = "Categoría"
    _black_formula(ws["F1"], bold=True)
    ws["G1"] = "Total mes"
    _black_formula(ws["G1"], bold=True)
    for i, cat in enumerate(CATEGORIES["Gasto"], start=2):
        c_label = ws.cell(row=i, column=6, value=cat)
        _black_formula(c_label)
        c_sum = ws.cell(
            row=i,
            column=7,
            value=(
                f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
                f'tblMov[Categoría],F{i},'
                f'tblMov[Fecha],">="&{month_start},'
                f'tblMov[Fecha],"<"&{month_end_excl})'
            ),
        )
        _black_formula(c_sum)
        c_sum.number_format = _FMT_CURRENCY

    # --- Monthly evolution block (I1:M13), fixed Jan-Dec of YEAR($B$1) ---
    for hdr, col in [("Mes", 9), ("Ingresos", 10), ("Gastos", 11), ("Ahorro", 12), ("% Ahorro", 13)]:
        c = ws.cell(row=1, column=col, value=hdr)
        _black_formula(c, bold=True)
    for month_num in range(1, 13):
        r = 1 + month_num  # row 2 = January, row 13 = December
        start_ref = f'DATE(YEAR($B$1),{month_num},1)'
        end_ref = f'DATE(YEAR($B$1),{month_num + 1},1)'
        c9 = ws.cell(row=r, column=9, value=f'=TEXT(DATE(YEAR($B$1),{month_num},1),"mm/yyyy")')
        _black_formula(c9)
        c10 = ws.cell(
            row=r,
            column=10,
            value=(
                f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Ingreso",'
                f'tblMov[Fecha],">="&{start_ref},'
                f'tblMov[Fecha],"<"&{end_ref})'
            ),
        )
        _black_formula(c10)
        c11 = ws.cell(
            row=r,
            column=11,
            value=(
                f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
                f'tblMov[Fecha],">="&{start_ref},'
                f'tblMov[Fecha],"<"&{end_ref})'
            ),
        )
        _black_formula(c11)
        c12 = ws.cell(row=r, column=12, value=f"=J{r}-K{r}")
        _black_formula(c12)
        c13 = ws.cell(row=r, column=13, value=f"=IFERROR(L{r}/J{r},0)")
        _black_formula(c13)
        for c in (c10, c11, c12):
            c.number_format = _FMT_CURRENCY
        c13.number_format = _FMT_PERCENT

    # --- Top-5 block (N1:P6). Uses AGGREGATE(14,6,...) for Excel 2010+ compatibility ---
    for hdr, col in [("Fecha", 14), ("Descripción", 15), ("Importe", 16)]:
        c = ws.cell(row=1, column=col, value=hdr)
        _black_formula(c, bold=True)
    mask_expr = (
        '(Movimientos!$B$2:$B$1000="Gasto")'
        '*(Movimientos!$A$2:$A$1000>=$B$1)'
        f'*(Movimientos!$A$2:$A$1000<{month_end_excl})'
    )
    for i in range(5):
        r = 2 + i
        rank = i + 1
        cp = ws.cell(
            row=r,
            column=16,
            value=(
                f'=IFERROR(AGGREGATE(14,6,'
                f'Movimientos!$D$2:$D$1000/({mask_expr}),{rank}),"")'
            ),
        )
        _black_formula(cp)
        cn = ws.cell(
            row=r,
            column=14,
            value=(
                f'=IFERROR(INDEX(Movimientos!$A$2:$A$1000,'
                f'SUMPRODUCT((Movimientos!$D$2:$D$1000=P{r})'
                f'*({mask_expr})'
                f'*ROW(Movimientos!$A$2:$A$1000))-1),"")'
            ),
        )
        _black_formula(cn)
        co = ws.cell(
            row=r,
            column=15,
            value=(
                f'=IFERROR(INDEX(Movimientos!$E$2:$E$1000,'
                f'SUMPRODUCT((Movimientos!$D$2:$D$1000=P{r})'
                f'*({mask_expr})'
                f'*ROW(Movimientos!$A$2:$A$1000))-1),"")'
            ),
        )
        _black_formula(co)
        cn.number_format = _FMT_DATE
        cp.number_format = _FMT_CURRENCY

    # --- Cumulative daily expenses block (R1:S32) ---
    c = ws["R1"]
    c.value = "Día"
    _black_formula(c, bold=True)
    c = ws["S1"]
    c.value = "Acumulado"
    _black_formula(c, bold=True)
    for i in range(31):
        r = 2 + i
        cr = ws.cell(row=r, column=18, value=f"=$B$1+{i}")
        _black_formula(cr)
        cr.number_format = _FMT_DAY
        cs = ws.cell(
            row=r,
            column=19,
            value=(
                f'=IF(R{r}>=DATE(YEAR($B$1),MONTH($B$1)+1,1),NA(),'
                f'SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
                f'tblMov[Fecha],">="&$B$1,'
                f'tblMov[Fecha],"<="&R{r}))'
            ),
        )
        _black_formula(cs)
        cs.number_format = _FMT_CURRENCY

    # --- Category comparison: this month vs previous month (U1:W11) ---
    for hdr, col in [("Categoría", 21), ("Este mes", 22), ("Mes anterior", 23)]:
        c = ws.cell(row=1, column=col, value=hdr)
        _black_formula(c, bold=True)
    prev_start = 'DATE(YEAR($B$1),MONTH($B$1)-1,1)'
    for i, cat in enumerate(CATEGORIES["Gasto"], start=2):
        cu = ws.cell(row=i, column=21, value=cat)
        _black_formula(cu)
        cv = ws.cell(
            row=i,
            column=22,
            value=(
                f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
                f'tblMov[Categoría],U{i},'
                f'tblMov[Fecha],">="&{month_start},'
                f'tblMov[Fecha],"<"&{month_end_excl})'
            ),
        )
        _black_formula(cv)
        cw = ws.cell(
            row=i,
            column=23,
            value=(
                f'=SUMIFS(tblMov[Importe],tblMov[Tipo],"Gasto",'
                f'tblMov[Categoría],U{i},'
                f'tblMov[Fecha],">="&{prev_start},'
                f'tblMov[Fecha],"<"&{month_start})'
            ),
        )
        _black_formula(cw)
        cv.number_format = _FMT_CURRENCY
        cw.number_format = _FMT_CURRENCY

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["O"].width = 30


def _build_dashboard(ws) -> None:
    # --- Month selector ---
    ws["A1"] = "Mes seleccionado:"
    _black_formula(ws["A1"], bold=True)
    b1 = ws["B1"]
    b1.value = "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)"
    b1.number_format = _FMT_DATE
    b1.fill = PatternFill("solid", fgColor="FFF2CC")  # soft yellow = key assumption
    _blue_input(b1)  # blue = hardcoded input the user will change

    # --- KPI labels and values ---
    kpi_rows = [
        ("Ingresos:", "=_aux!D1", _FMT_CURRENCY),
        ("Gastos:", "=_aux!D2", _FMT_CURRENCY),
        ("Ahorro neto:", "=_aux!D3", _FMT_CURRENCY),
        ("% Ahorro:", "=_aux!D4", _FMT_PERCENT),
    ]
    for i, (label, formula, fmt) in enumerate(kpi_rows, start=1):
        lbl = ws.cell(row=i, column=3, value=label)
        _black_formula(lbl, bold=True)
        c = ws.cell(row=i, column=4, value=formula)
        c.number_format = fmt
        _green_link(c, size=14, bold=True)

    # --- Top-5 section ---
    ws["A6"] = "TOP 5 GASTOS DEL MES"
    _black_formula(ws["A6"], size=12, bold=True)
    for hdr, col in [("Fecha", 1), ("Descripción", 2), ("Importe", 3)]:
        c = ws.cell(row=7, column=col, value=hdr)
        _black_formula(c, bold=True)
    for i in range(5):
        r = 8 + i
        ca = ws.cell(row=r, column=1, value=f"=_aux!N{2 + i}")
        ca.number_format = _FMT_DATE
        _green_link(ca)
        cb = ws.cell(row=r, column=2, value=f"=_aux!O{2 + i}")
        _green_link(cb)
        cc = ws.cell(row=r, column=3, value=f"=_aux!P{2 + i}")
        cc.number_format = _FMT_CURRENCY
        _green_link(cc)

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
