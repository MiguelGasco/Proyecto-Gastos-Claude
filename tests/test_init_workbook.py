from pathlib import Path
import openpyxl
from scripts.init_workbook import create_workbook


def test_create_workbook_has_four_named_sheets(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["Dashboard", "Movimientos", "Categorías", "_aux"]


def test_create_workbook_fails_if_file_exists(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    import pytest
    with pytest.raises(FileExistsError):
        create_workbook(out)


def test_movimientos_has_table_with_expected_columns(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    ws = wb["Movimientos"]
    assert "tblMov" in ws.tables
    headers = [c.value for c in ws[1]]
    assert headers == ["Fecha", "Tipo", "Categoría", "Importe", "Descripción"]


def test_categorias_sheet_lists_all_categories(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    ws = wb["Categorías"]
    pairs = {(row[0].value, row[1].value) for row in ws.iter_rows(min_row=2) if row[0].value}
    # Spot-check: at least these pairs must be present
    assert ("Gasto", "Alimentación") in pairs
    assert ("Gasto", "Restauración") in pairs
    assert ("Ingreso", "Nómina") in pairs
    # And total count matches our source of truth
    from scripts.categories import CATEGORIES
    expected = sum(len(v) for v in CATEGORIES.values())
    assert len(pairs) == expected


def test_movimientos_has_data_validations(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    ws = wb["Movimientos"]
    # At least two validations attached: one for Tipo (col B), one for Categoría (col C)
    dv_ranges = [str(dv.sqref) for dv in ws.data_validations.dataValidation]
    assert any("B" in r for r in dv_ranges), f"Missing Tipo validation, got {dv_ranges}"
    assert any("C" in r for r in dv_ranges), f"Missing Categoría validation, got {dv_ranges}"


def test_aux_has_selected_month_and_kpi_formulas(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    aux = wb["_aux"]

    # Selected month is driven from Dashboard
    assert aux["B1"].value == "=Dashboard!B1"

    # KPIs use SUMIFS over tblMov filtered by month
    ingresos = aux["D1"].value or ""
    gastos = aux["D2"].value or ""
    assert "SUMIFS" in ingresos.upper()
    assert "tblMov" in ingresos
    assert "SUMIFS" in gastos.upper()

    # Ahorro is Ingresos - Gastos (or arithmetic of D1/D2)
    ahorro = aux["D3"].value or ""
    assert "D1" in ahorro and "D2" in ahorro

    # % ahorro guards against /0
    pct = aux["D4"].value or ""
    assert "SI.ERROR" in pct.upper() or "IFERROR" in pct.upper()


def test_aux_category_block_has_ten_rows(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    aux = wb["_aux"]

    cats = [aux.cell(row=r, column=6).value for r in range(2, 12)]
    sums = [aux.cell(row=r, column=7).value for r in range(2, 12)]
    assert all(cats), f"Missing category labels: {cats}"
    assert all(s and "SUMIFS" in s.upper() for s in sums), f"Missing SUMIFS: {sums}"


def test_aux_monthly_evolution_block_has_twelve_rows(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    aux = wb["_aux"]

    # Row 2 is the most recent month, row 13 is 11 months earlier (rolling window)
    months = [aux.cell(row=r, column=9).value for r in range(2, 14)]
    ingresos = [aux.cell(row=r, column=10).value for r in range(2, 14)]
    gastos = [aux.cell(row=r, column=11).value for r in range(2, 14)]
    assert all(m for m in months), f"Missing month labels: {months}"
    assert all("SUMIFS" in (s or "").upper() for s in ingresos)
    assert all("SUMIFS" in (s or "").upper() for s in gastos)


def test_dashboard_has_month_selector_default_to_today(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    dash = wb["Dashboard"]

    assert dash["A1"].value == "Mes seleccionado:"
    formula = (dash["B1"].value or "").upper()
    assert "FECHA" in formula or "DATE" in formula
    assert "HOY" in formula or "TODAY" in formula


def test_dashboard_kpis_link_to_aux(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    dash = wb["Dashboard"]

    assert dash["D1"].value == "=_aux!D1"
    assert dash["D2"].value == "=_aux!D2"
    assert dash["D3"].value == "=_aux!D3"
    assert dash["D4"].value == "=_aux!D4"


def test_dashboard_top5_links_to_aux(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    dash = wb["Dashboard"]

    # 5 rows starting at row 8, three columns A/B/C linked to _aux!N/O/P 2..6
    for i in range(5):
        r = 8 + i
        assert dash.cell(row=r, column=1).value == f"=_aux!N{2 + i}"
        assert dash.cell(row=r, column=2).value == f"=_aux!O{2 + i}"
        assert dash.cell(row=r, column=3).value == f"=_aux!P{2 + i}"


def test_dashboard_has_pie_and_line_charts(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    dash = wb["Dashboard"]

    # openpyxl exposes charts via ws._charts
    chart_types = {type(c).__name__ for c in dash._charts}
    assert "PieChart" in chart_types, f"got {chart_types}"
    assert "LineChart" in chart_types, f"got {chart_types}"


def test_aux_sheet_is_hidden(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    assert wb["_aux"].sheet_state == "hidden"
    # Dashboard is the active/visible default sheet
    assert wb.active.title == "Dashboard"


def test_aux_evolution_is_calendar_year(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)
    wb = openpyxl.load_workbook(out)
    aux = wb["_aux"]
    # Row 2 = January of YEAR($B$1); row 13 = December of YEAR($B$1)
    row2 = aux.cell(row=2, column=9).value or ""
    row13 = aux.cell(row=13, column=9).value or ""
    assert "DATE(YEAR($B$1),1,1)" in row2, f"row 2 should be January: {row2}"
    assert "DATE(YEAR($B$1),12,1)" in row13, f"row 13 should be December: {row13}"


def test_aux_evolution_has_pct_ahorro_column(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)
    wb = openpyxl.load_workbook(out)
    aux = wb["_aux"]
    assert aux["M1"].value == "% Ahorro"
    for r in range(2, 14):
        f = aux.cell(row=r, column=13).value or ""
        assert "IFERROR" in f.upper() and f"L{r}" in f and f"J{r}" in f


def test_aux_top5_uses_aggregate(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)
    wb = openpyxl.load_workbook(out)
    aux = wb["_aux"]
    p2 = (aux["P2"].value or "").upper()
    n2 = (aux["N2"].value or "").upper()
    assert "AGGREGATE" in p2, f"P2 should use AGGREGATE: {p2}"
    assert "SUMPRODUCT" in n2, f"N2 should use SUMPRODUCT: {n2}"


def test_aux_cumulative_daily_block(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)
    wb = openpyxl.load_workbook(out)
    aux = wb["_aux"]
    assert aux["R1"].value == "Día"
    assert aux["S1"].value == "Acumulado"
    # 31 rows of dates (R2..R32) and 31 rows of cumulative SUMIFS (S2..S32)
    for r in range(2, 33):
        date_formula = aux.cell(row=r, column=18).value or ""
        sum_formula = (aux.cell(row=r, column=19).value or "").upper()
        assert "$B$1" in date_formula
        assert "SUMIFS" in sum_formula and "NA()" in sum_formula


def test_aux_category_comparison_block(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)
    wb = openpyxl.load_workbook(out)
    aux = wb["_aux"]
    assert aux["U1"].value == "Categoría"
    assert aux["V1"].value == "Este mes"
    assert aux["W1"].value == "Mes anterior"
    from scripts.categories import CATEGORIES
    cats = [aux.cell(row=r, column=21).value for r in range(2, 12)]
    assert cats == CATEGORIES["Gasto"]
    for r in range(2, 12):
        v = (aux.cell(row=r, column=22).value or "").upper()
        w = (aux.cell(row=r, column=23).value or "").upper()
        assert "SUMIFS" in v and "SUMIFS" in w


def test_dashboard_has_five_charts(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)
    wb = openpyxl.load_workbook(out)
    dash = wb["Dashboard"]
    types = [type(c).__name__ for c in dash._charts]
    # Original: Pie + Line. New: Line (acumulado), Bar (comparativa), Line (% ahorro)
    assert types.count("PieChart") == 1
    assert types.count("LineChart") == 3
    assert types.count("BarChart") == 1
    assert len(types) == 5
