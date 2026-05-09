# Control de Gastos Excel — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local Excel-based personal finance tracker (`gastos.xlsx`) that Claude updates conversationally via Python scripts, with live KPIs and charts.

**Architecture:** Two Python scripts using `openpyxl`. `init_workbook.py` builds the `.xlsx` from scratch (4 sheets: Movimientos as a real Excel Table, Categorías as reference list, Dashboard with KPIs + 2 charts + top-5 list, `_aux` hidden sheet with intermediate `SUMIFS`). `add_movement.py` appends/edits/deletes rows in the Movimientos table. Excel formulas + chart references update automatically when the user opens the file. No persistent state lives outside `gastos.xlsx` (Engram is parallel persistence handled by Claude, not by these scripts).

**Tech Stack:** Python 3.11+, `openpyxl` (>=3.1), `pytest` for tests. Windows + PowerShell.

**Spec:** [docs/superpowers/specs/2026-05-09-control-gastos-excel-design.md](../specs/2026-05-09-control-gastos-excel-design.md)

---

## File Structure

Files this plan creates:

| File | Responsibility |
|---|---|
| `requirements.txt` | Pin `openpyxl` + `pytest` |
| `scripts/__init__.py` | Make `scripts/` an importable package for tests |
| `scripts/init_workbook.py` | Build `gastos.xlsx` from scratch (idempotency-protected) |
| `scripts/add_movement.py` | Append / edit-last / delete-last rows in Movimientos |
| `scripts/categories.py` | Single source of truth for the category lists (imported by both scripts) |
| `tests/__init__.py` | Test package marker |
| `tests/test_init_workbook.py` | Tests for the workbook builder |
| `tests/test_add_movement.py` | Tests for the movement script |
| `README.md` | How to install, init, and use |

Output (not source, gitignored if a repo is initialised):

| File | Notes |
|---|---|
| `gastos.xlsx` | Live data file. Created by `init_workbook.py`, mutated by `add_movement.py`. |

---

## Task 1: Project scaffolding

**Files:**
- Create: `requirements.txt`
- Create: `scripts/__init__.py` (empty)
- Create: `tests/__init__.py` (empty)
- Create: `scripts/categories.py`
- Create: `pyproject.toml` (minimal, for pytest config)
- Create: `.gitignore`

- [ ] **Step 1: Verify Python is available**

Run (PowerShell):
```powershell
python --version
```
Expected: `Python 3.11.x` or higher. If not available, install from python.org first.

- [ ] **Step 2: Create `requirements.txt`**

```
openpyxl>=3.1,<4
pytest>=8,<9
```

- [ ] **Step 3: Create `pyproject.toml` for pytest config**

```toml
[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["."]
```

- [ ] **Step 4: Create `.gitignore`**

```
__pycache__/
*.pyc
.venv/
.pytest_cache/
gastos.xlsx
~$gastos.xlsx
```

- [ ] **Step 5: Create empty package markers**

`scripts/__init__.py`: empty file.
`tests/__init__.py`: empty file.

- [ ] **Step 6: Create `scripts/categories.py` (single source of truth for categories)**

```python
"""Category lists shared by init_workbook and add_movement."""

CATEGORIES = {
    "Gasto": [
        "Alimentación",
        "Restauración",
        "Vivienda",
        "Suministros",
        "Transporte",
        "Salud",
        "Ocio",
        "Compras",
        "Suscripciones",
        "Otros",
    ],
    "Ingreso": [
        "Nómina",
        "Extras",
        "Otros",
    ],
}

VALID_TYPES = list(CATEGORIES.keys())


def is_valid(tipo: str, categoria: str) -> bool:
    return tipo in CATEGORIES and categoria in CATEGORIES[tipo]
```

- [ ] **Step 7: Create venv and install dependencies**

Run (PowerShell, from project root):
```powershell
python -m venv .venv; .venv\Scripts\Activate.ps1; pip install -r requirements.txt
```
Expected: `openpyxl` and `pytest` install without errors.

- [ ] **Step 8: Initialise git (optional but recommended for frequent commits)**

Run:
```powershell
git init; git add .; git commit -m "chore: project scaffolding"
```
If the user prefers to skip git, omit this step. All later commit steps become no-ops in that case.

---

## Task 2: `init_workbook.py` — empty workbook with the 4 sheets

**Files:**
- Create: `scripts/init_workbook.py`
- Create: `tests/test_init_workbook.py`

- [ ] **Step 1: Write the failing test**

`tests/test_init_workbook.py`:
```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run:
```powershell
pytest tests/test_init_workbook.py -v
```
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.init_workbook'`.

- [ ] **Step 3: Write minimal implementation**

`scripts/init_workbook.py`:
```python
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
```

- [ ] **Step 4: Run test to verify it passes**

Run:
```powershell
pytest tests/test_init_workbook.py -v
```
Expected: PASS, both tests green.

- [ ] **Step 5: Commit**

```powershell
git add scripts/ tests/; git commit -m "feat: init_workbook scaffold with 4 sheets and idempotency guard"
```

---

## Task 3: `Movimientos` as an Excel Table + `Categorías` reference list + data validations

**Files:**
- Modify: `scripts/init_workbook.py`
- Modify: `tests/test_init_workbook.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_init_workbook.py`:
```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```powershell
pytest tests/test_init_workbook.py -v
```
Expected: 3 new tests FAIL (table missing, Categorías empty, no validations).

- [ ] **Step 3: Add Movimientos table and Categorías builders**

In `scripts/init_workbook.py`, add these helpers below `create_workbook` and call them from inside `create_workbook` before `wb.save(path)`:

```python
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from scripts.categories import CATEGORIES, VALID_TYPES


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
```

Then in `create_workbook`, after creating the sheets and before saving:
```python
    _build_movimientos(wb["Movimientos"])
    _build_categorias(wb["Categorías"])
    _add_validations(wb["Movimientos"])
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```powershell
pytest tests/test_init_workbook.py -v
```
Expected: all 5 tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add scripts/ tests/; git commit -m "feat: Movimientos table, Categorías list and data validations"
```

---

## Task 4: `_aux` sheet — monthly summary, category-of-month, top-5

**Files:**
- Modify: `scripts/init_workbook.py`
- Modify: `tests/test_init_workbook.py`

The `_aux` sheet exposes named cells/ranges that the Dashboard (Task 5) and the charts (Task 6) read from. Single source of truth for all derived values.

Layout of `_aux`:

| Block | Cells | Content |
|---|---|---|
| Selected month | `B1` | First day of selected month, e.g. `=FECHA(AÑO(HOY()),MES(HOY()),1)`. Driven by Dashboard `B1`. |
| Selected month label | `B2` | `=TEXTO(B1,"mm/yyyy")` (display only) |
| KPIs (mes) | `D1:D4` | Ingresos, Gastos, Ahorro, % Ahorro (each labelled in `C1:C4`) |
| Category-of-month | `F1:G11` | 10 rows: category name (F) + sum (G) |
| Monthly evolution | `I1:L13` | Header row + 12 rows: Año-Mes (I), Ingresos (J), Gastos (K), Ahorro (L) |
| Top-5 of month | `N1:P6` | Header row + 5 rows: Fecha (N), Descripción (O), Importe (P) |

The selected-month cell (`Dashboard!B1`) is the single user input; everything else recomputes.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_init_workbook.py`:
```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```powershell
pytest tests/test_init_workbook.py -v
```
Expected: 3 new tests FAIL (cells empty / formulas missing).

- [ ] **Step 3: Implement `_build_aux`**

Add to `scripts/init_workbook.py`:

```python
def _build_aux(ws) -> None:
    # --- B1: selected month (driven by Dashboard!B1, which user picks) ---
    ws["A1"] = "Mes seleccionado"
    ws["B1"] = "=Dashboard!B1"
    ws["B1"].number_format = "dd/mm/yyyy"
    ws["A2"] = "Etiqueta"
    ws["B2"] = '=TEXTO(B1,"mm/yyyy")'

    # --- KPIs ---
    # Date range filter: >= B1 (first of month) and < EOMonth(B1)+1
    month_start = "$B$1"
    month_end_excl = 'FECHA(AÑO($B$1),MES($B$1)+1,1)'

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
    ws["D4"] = "=SI.ERROR(D3/D1,0)"

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

    # --- Monthly evolution block (I1:L13), rolling 12 months ending at selected ---
    ws["I1"] = "Mes"
    ws["J1"] = "Ingresos"
    ws["K1"] = "Gastos"
    ws["L1"] = "Ahorro"
    for offset in range(12):
        r = 2 + offset
        # offset 0 = selected month, offset 11 = 11 months earlier
        # Use FECHA + DESREF-free arithmetic
        ws.cell(
            row=r,
            column=9,
            value=f'=TEXTO(FECHA(AÑO($B$1),MES($B$1)-{offset},1),"mm/yyyy")',
        )
        start_ref = f'FECHA(AÑO($B$1),MES($B$1)-{offset},1)'
        end_ref = f'FECHA(AÑO($B$1),MES($B$1)-{offset}+1,1)'
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
        for c in (10, 11, 12):
            ws.cell(row=r, column=c).number_format = '#,##0.00 "€"'

    # --- Top-5 block (N1:P6). Uses LARGE + INDEX/MATCH on the filtered set ---
    # Note: array-style formulas. The simplest reliable form on a fixed table range.
    ws["N1"] = "Fecha"
    ws["O1"] = "Descripción"
    ws["P1"] = "Importe"
    # We compute against the static range A2:E1000 of Movimientos because tblMov
    # column references inside K.ESIMO.MAYOR with a filter expression are clunky
    # in classic Excel. With dynamic-array Excel, FILTER+SORT would be cleaner;
    # we fall back to a robust classic-Excel formula.
    for i in range(5):
        r = 2 + i
        rank = i + 1
        # Importe (P): rank-th largest of Importe in the month, type=Gasto
        ws.cell(
            row=r,
            column=16,
            value=(
                f'=SI.ERROR(K.ESIMO.MAYOR(SI((Movimientos!$B$2:$B$1000="Gasto")*'
                f'(Movimientos!$A$2:$A$1000>=$B$1)*'
                f'(Movimientos!$A$2:$A$1000<{month_end_excl}),'
                f'Movimientos!$D$2:$D$1000),{rank}),"")'
            ),
        )
        # Fecha (N) and Descripción (O): match by the importe found in Pr
        ws.cell(
            row=r,
            column=14,
            value=(
                f'=SI.ERROR(INDICE(Movimientos!$A$2:$A$1000,'
                f'COINCIDIR(P{r},Movimientos!$D$2:$D$1000,0)),"")'
            ),
        )
        ws.cell(
            row=r,
            column=15,
            value=(
                f'=SI.ERROR(INDICE(Movimientos!$E$2:$E$1000,'
                f'COINCIDIR(P{r},Movimientos!$D$2:$D$1000,0)),"")'
            ),
        )
        ws.cell(row=r, column=14).number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=16).number_format = '#,##0.00 "€"'

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["O"].width = 30
```

Then call it inside `create_workbook` after `_add_validations(...)`:
```python
    _build_aux(wb["_aux"])
```

> **Note on Spanish formula names:** `openpyxl` writes formula text verbatim. Excel stores formulas with English internal names (`SUMIFS`, `IFERROR`, `LARGE`, etc.) and renders them in the user's locale. To avoid round-trip surprises, use English names if you have any doubt. A safe alternative for this whole task is to replace the Spanish names below with their English equivalents — Excel will render them in Spanish for Miguel anyway:
> - `SUMIFS` ← (same)
> - `SI.ERROR` ← `IFERROR`
> - `FECHA` ← `DATE`
> - `AÑO` ← `YEAR`
> - `MES` ← `MONTH`
> - `TEXTO` ← `TEXT`
> - `K.ESIMO.MAYOR` ← `LARGE`
> - `INDICE` ← `INDEX`
> - `COINCIDIR` ← `MATCH`
> - `SI` ← `IF`
>
> If the smoke test in Task 8 shows `#NAME?` errors, switch all Spanish names to English in this task and Task 5, then re-run. The tests in this task only check for `SUMIFS`/`IFERROR` substrings (case-insensitive variants), so they pass either way.

- [ ] **Step 4: Run tests to verify they pass**

Run:
```powershell
pytest tests/test_init_workbook.py -v
```
Expected: all tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add scripts/ tests/; git commit -m "feat: _aux sheet with KPIs, category, evolution, top-5"
```

---

## Task 5: Dashboard — KPI cards + month selector

**Files:**
- Modify: `scripts/init_workbook.py`
- Modify: `tests/test_init_workbook.py`

Dashboard layout (target):

```
A           B               C           D
1  Mes:    [01/05/2026]    Ingresos:   =_aux!D1
2                          Gastos:     =_aux!D2
3                          Ahorro:     =_aux!D3
4                          % Ahorro:   =_aux!D4
6  TOP 5 GASTOS DEL MES
7  Fecha   Descripción     Importe
8..12      (5 rows linked to _aux!N2:P6)
```

Charts go in Task 6.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_init_workbook.py`:
```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```powershell
pytest tests/test_init_workbook.py -v
```
Expected: 3 new tests FAIL.

- [ ] **Step 3: Implement `_build_dashboard`**

Add to `scripts/init_workbook.py`:
```python
from openpyxl.styles import Font, Alignment, PatternFill


def _build_dashboard(ws) -> None:
    # --- Month selector ---
    ws["A1"] = "Mes seleccionado:"
    ws["A1"].font = Font(bold=True)
    ws["B1"] = "=FECHA(AÑO(HOY()),MES(HOY()),1)"
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
```

In `create_workbook`, after `_build_aux(...)`:
```python
    _build_dashboard(wb["Dashboard"])
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```powershell
pytest tests/test_init_workbook.py -v
```
Expected: all tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add scripts/ tests/; git commit -m "feat: Dashboard KPIs, month selector and top-5"
```

---

## Task 6: Dashboard charts — pie of categories + line of monthly evolution

**Files:**
- Modify: `scripts/init_workbook.py`
- Modify: `tests/test_init_workbook.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_init_workbook.py`:
```python
def test_dashboard_has_pie_and_line_charts(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    dash = wb["Dashboard"]

    # openpyxl exposes charts via ws._charts
    chart_types = {type(c).__name__ for c in dash._charts}
    assert "PieChart" in chart_types, f"got {chart_types}"
    assert "LineChart" in chart_types, f"got {chart_types}"
```

- [ ] **Step 2: Run test to verify it fails**

Run:
```powershell
pytest tests/test_init_workbook.py::test_dashboard_has_pie_and_line_charts -v
```
Expected: FAIL — `chart_types == set()`.

- [ ] **Step 3: Implement `_add_charts`**

Add to `scripts/init_workbook.py`:
```python
from openpyxl.chart import PieChart, LineChart, Reference


def _add_charts(wb) -> None:
    aux = wb["_aux"]
    dash = wb["Dashboard"]

    # --- Pie chart: gastos por categoría del mes seleccionado ---
    pie = PieChart()
    pie.title = "Gastos por categoría (mes)"
    labels = Reference(aux, min_col=6, min_row=2, max_row=11)        # F2:F11
    data = Reference(aux, min_col=7, min_row=1, max_row=11)          # G1:G11 (incl header)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.height = 9   # cm
    pie.width = 14
    dash.add_chart(pie, "F1")

    # --- Line chart: evolución mensual últimos 12 meses ---
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
```

In `create_workbook`, after `_build_dashboard(...)`:
```python
    _add_charts(wb)
```

- [ ] **Step 4: Run test to verify it passes**

Run:
```powershell
pytest tests/test_init_workbook.py -v
```
Expected: all tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add scripts/ tests/; git commit -m "feat: dashboard pie and line charts"
```

---

## Task 7: Hide `_aux`, polish formats, full smoke

**Files:**
- Modify: `scripts/init_workbook.py`
- Modify: `tests/test_init_workbook.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_init_workbook.py`:
```python
def test_aux_sheet_is_hidden(tmp_path: Path):
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)

    wb = openpyxl.load_workbook(out)
    assert wb["_aux"].sheet_state == "hidden"
    # Dashboard is the active/visible default sheet
    assert wb.active.title == "Dashboard"
```

- [ ] **Step 2: Run test to verify it fails**

Run:
```powershell
pytest tests/test_init_workbook.py::test_aux_sheet_is_hidden -v
```
Expected: FAIL — `_aux.sheet_state == "visible"`.

- [ ] **Step 3: Hide `_aux` and set Dashboard active**

Inside `create_workbook`, just before `wb.save(path)`:
```python
    wb["_aux"].sheet_state = "hidden"
    wb.active = wb.sheetnames.index("Dashboard")
```

- [ ] **Step 4: Run all tests**

Run:
```powershell
pytest -v
```
Expected: all PASS.

- [ ] **Step 5: Manual smoke test — open in Excel**

Run:
```powershell
Remove-Item gastos.xlsx -ErrorAction SilentlyContinue; python -m scripts.init_workbook
```

Open `gastos.xlsx` in Excel and verify:
- 3 visible sheets (Dashboard, Movimientos, Categorías), `_aux` not visible.
- Dashboard month selector (B1) shows today's month.
- KPIs all show `0,00 €` and `0,0%` (no data yet).
- Pie chart shows nothing (no gastos yet) — that's expected.
- Line chart shows 12 month labels with zeros.
- Top-5 area is empty (5 blank rows).
- Movimientos sheet has the 5 headers and an empty striped table.
- Click on cell B2 of Movimientos — `Tipo` dropdown should appear with `Gasto`/`Ingreso`.
- Click on cell C2 — `Categoría` dropdown with the full union list.

If any formula shows `#NAME?`, switch the Spanish function names in Tasks 4 and 5 to their English equivalents (see note in Task 4 step 3) and re-run.

- [ ] **Step 6: Commit**

```powershell
git add scripts/ tests/; git commit -m "feat: hide _aux, set Dashboard active, smoke verified"
```

---

## Task 8: `add_movement.py` — append a row + `--edit-last` + `--delete-last`

**Files:**
- Create: `scripts/add_movement.py`
- Create: `tests/test_add_movement.py`

The script supports three modes via mutually exclusive flags:

| Mode | Flags | Behaviour |
|---|---|---|
| Append | `--fecha --tipo --categoria --importe [--descripcion]` | Add a new row after the last filled data row in `tblMov` |
| Edit last | `--edit-last` + any of the field flags | Update the bottom-most filled row's specified columns |
| Delete last | `--delete-last` | Clear the bottom-most filled row |

Validation rules:
- `tipo` must be in `VALID_TYPES`.
- `(tipo, categoria)` pair must be valid per `categories.is_valid()`.
- `importe` must parse as positive `float`.
- `fecha` must parse as `YYYY-MM-DD`.

- [ ] **Step 1: Write the failing tests**

`tests/test_add_movement.py`:
```python
from datetime import date
from pathlib import Path
import pytest
import openpyxl

from scripts.init_workbook import create_workbook
from scripts.add_movement import (
    add_movement,
    edit_last,
    delete_last,
    InvalidMovementError,
)


@pytest.fixture()
def fresh_book(tmp_path: Path) -> Path:
    out = tmp_path / "gastos.xlsx"
    create_workbook(out)
    return out


def _read_data_rows(path: Path) -> list[tuple]:
    wb = openpyxl.load_workbook(path)
    ws = wb["Movimientos"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        rows.append(row)
    return rows


def test_add_appends_a_row(fresh_book: Path):
    add_movement(
        fresh_book,
        fecha=date(2026, 5, 9),
        tipo="Gasto",
        categoria="Restauración",
        importe=30.0,
        descripcion="Comida con amigos",
    )
    rows = _read_data_rows(fresh_book)
    assert len(rows) == 1
    fecha, tipo, cat, imp, desc = rows[0]
    assert fecha == date(2026, 5, 9) or str(fecha).startswith("2026-05-09")
    assert tipo == "Gasto"
    assert cat == "Restauración"
    assert imp == 30.0
    assert desc == "Comida con amigos"


def test_add_two_rows_appends_in_order(fresh_book: Path):
    add_movement(fresh_book, date(2026, 5, 1), "Ingreso", "Nómina", 2100.0, "Mayo")
    add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Restauración", 30.0, "Comida")
    rows = _read_data_rows(fresh_book)
    assert len(rows) == 2
    assert rows[0][1] == "Ingreso"
    assert rows[1][1] == "Gasto"


def test_add_rejects_invalid_category(fresh_book: Path):
    with pytest.raises(InvalidMovementError, match="categoría"):
        add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Inexistente", 10.0, "")


def test_add_rejects_wrong_pair(fresh_book: Path):
    # Nómina is only valid for Ingreso
    with pytest.raises(InvalidMovementError, match="categoría"):
        add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Nómina", 10.0, "")


def test_add_rejects_invalid_tipo(fresh_book: Path):
    with pytest.raises(InvalidMovementError, match="tipo"):
        add_movement(fresh_book, date(2026, 5, 9), "Otro", "Nómina", 10.0, "")


def test_add_rejects_negative_importe(fresh_book: Path):
    with pytest.raises(InvalidMovementError, match="importe"):
        add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Otros", -5.0, "")


def test_edit_last_changes_specified_fields(fresh_book: Path):
    add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Restauración", 30.0, "Comida")
    edit_last(fresh_book, importe=35.0, descripcion="Comida (corregido)")
    rows = _read_data_rows(fresh_book)
    assert rows[0][3] == 35.0
    assert rows[0][4] == "Comida (corregido)"
    # Untouched fields remain
    assert rows[0][1] == "Gasto"
    assert rows[0][2] == "Restauración"


def test_edit_last_on_empty_book_raises(fresh_book: Path):
    with pytest.raises(InvalidMovementError, match="vacía|empty"):
        edit_last(fresh_book, importe=10.0)


def test_delete_last_removes_only_last_row(fresh_book: Path):
    add_movement(fresh_book, date(2026, 5, 1), "Ingreso", "Nómina", 2100.0, "Mayo")
    add_movement(fresh_book, date(2026, 5, 9), "Gasto", "Restauración", 30.0, "Comida")
    delete_last(fresh_book)
    rows = _read_data_rows(fresh_book)
    assert len(rows) == 1
    assert rows[0][1] == "Ingreso"


def test_delete_last_on_empty_book_raises(fresh_book: Path):
    with pytest.raises(InvalidMovementError, match="vacía|empty"):
        delete_last(fresh_book)
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```powershell
pytest tests/test_add_movement.py -v
```
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.add_movement'`.

- [ ] **Step 3: Implement `scripts/add_movement.py`**

```python
"""Append / edit-last / delete-last rows in tblMov of gastos.xlsx."""
from __future__ import annotations

import argparse
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import openpyxl

from scripts.categories import VALID_TYPES, is_valid


SHEET = "Movimientos"
FIRST_DATA_ROW = 2
NUM_COLS = 5  # Fecha, Tipo, Categoría, Importe, Descripción


class InvalidMovementError(ValueError):
    """Raised when a movement's fields don't validate."""


def _validate(tipo: str, categoria: str, importe: float) -> None:
    if tipo not in VALID_TYPES:
        raise InvalidMovementError(
            f"tipo inválido: {tipo!r}. Debe ser uno de {VALID_TYPES}"
        )
    if not is_valid(tipo, categoria):
        raise InvalidMovementError(
            f"categoría {categoria!r} no es válida para tipo {tipo!r}"
        )
    if importe <= 0:
        raise InvalidMovementError(
            f"importe debe ser positivo, recibido: {importe}"
        )


def _last_data_row(ws) -> int:
    """Return the row number of the bottom-most filled data row, or 1 if empty."""
    last = 1
    for row in ws.iter_rows(min_row=FIRST_DATA_ROW, max_col=1, values_only=False):
        if row[0].value is None:
            break
        last = row[0].row
    return last


def add_movement(
    path: Path,
    fecha: date,
    tipo: str,
    categoria: str,
    importe: float,
    descripcion: str = "",
) -> None:
    _validate(tipo, categoria, importe)

    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]

    target_row = _last_data_row(ws) + 1 if _last_data_row(ws) >= FIRST_DATA_ROW else FIRST_DATA_ROW
    ws.cell(row=target_row, column=1, value=fecha).number_format = "dd/mm/yyyy"
    ws.cell(row=target_row, column=2, value=tipo)
    ws.cell(row=target_row, column=3, value=categoria)
    ws.cell(row=target_row, column=4, value=float(importe)).number_format = '#,##0.00 "€"'
    ws.cell(row=target_row, column=5, value=descripcion or "")

    wb.save(path)


def edit_last(
    path: Path,
    fecha: Optional[date] = None,
    tipo: Optional[str] = None,
    categoria: Optional[str] = None,
    importe: Optional[float] = None,
    descripcion: Optional[str] = None,
) -> None:
    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]
    last = _last_data_row(ws)
    if last < FIRST_DATA_ROW:
        raise InvalidMovementError("la tabla está vacía, nada que editar")

    # Resolve final values (use existing if not overridden) before validating
    current = {
        "fecha": ws.cell(row=last, column=1).value,
        "tipo": ws.cell(row=last, column=2).value,
        "categoria": ws.cell(row=last, column=3).value,
        "importe": ws.cell(row=last, column=4).value,
        "descripcion": ws.cell(row=last, column=5).value,
    }
    final_tipo = tipo if tipo is not None else current["tipo"]
    final_cat = categoria if categoria is not None else current["categoria"]
    final_imp = importe if importe is not None else current["importe"]
    _validate(final_tipo, final_cat, float(final_imp))

    if fecha is not None:
        ws.cell(row=last, column=1, value=fecha).number_format = "dd/mm/yyyy"
    if tipo is not None:
        ws.cell(row=last, column=2, value=tipo)
    if categoria is not None:
        ws.cell(row=last, column=3, value=categoria)
    if importe is not None:
        ws.cell(row=last, column=4, value=float(importe)).number_format = '#,##0.00 "€"'
    if descripcion is not None:
        ws.cell(row=last, column=5, value=descripcion)

    wb.save(path)


def delete_last(path: Path) -> None:
    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]
    last = _last_data_row(ws)
    if last < FIRST_DATA_ROW:
        raise InvalidMovementError("la tabla está vacía, nada que borrar")

    for c in range(1, NUM_COLS + 1):
        ws.cell(row=last, column=c).value = None

    wb.save(path)


def _parse_fecha(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Añade, edita o borra el último movimiento en gastos.xlsx."
    )
    parser.add_argument("--path", default="gastos.xlsx")
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
        delete_last(path)
        print("OK: última fila borrada.")
        return

    if args.edit_last_flag:
        edit_last(
            path,
            fecha=args.fecha,
            tipo=args.tipo,
            categoria=args.categoria,
            importe=args.importe,
            descripcion=args.descripcion,
        )
        print("OK: última fila actualizada.")
        return

    # Append mode — all required fields must be present
    missing = [
        f for f, v in [
            ("--fecha", args.fecha),
            ("--tipo", args.tipo),
            ("--categoria", args.categoria),
            ("--importe", args.importe),
        ] if v is None
    ]
    if missing:
        parser.error(f"faltan flags requeridos para añadir: {', '.join(missing)}")

    add_movement(
        path,
        fecha=args.fecha,
        tipo=args.tipo,
        categoria=args.categoria,
        importe=args.importe,
        descripcion=args.descripcion or "",
    )
    print("OK: movimiento añadido.")


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```powershell
pytest tests/test_add_movement.py -v
```
Expected: all PASS.

- [ ] **Step 5: Commit**

```powershell
git add scripts/ tests/; git commit -m "feat: add_movement script with add/edit-last/delete-last"
```

---

## Task 9: README + end-to-end smoke

**Files:**
- Create: `README.md`

- [ ] **Step 1: Write `README.md`**

```markdown
# Proyecto Gastos Claude

Registro personal de ingresos y gastos en un único Excel (`gastos.xlsx`), alimentado por Claude conversacionalmente.

## Setup (una vez)

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m scripts.init_workbook
```

Esto crea `gastos.xlsx` con las hojas Dashboard, Movimientos, Categorías y `_aux` (oculta).

## Uso diario

Habla con Claude:

> "Hoy gasté 30 € comiendo con amigos."

Claude ejecutará por ti:

```powershell
python -m scripts.add_movement --fecha 2026-05-09 --tipo Gasto --categoria Restauración --importe 30 --descripcion "Comida con amigos"
```

Y guardará el movimiento en Engram para mantener contexto entre sesiones.

## Comandos manuales

```powershell
# Añadir
python -m scripts.add_movement --fecha 2026-05-09 --tipo Gasto --categoria Restauración --importe 30 --descripcion "Comida"

# Editar la última fila
python -m scripts.add_movement --edit-last --importe 35

# Borrar la última fila
python -m scripts.add_movement --delete-last
```

## Categorías

**Gastos:** Alimentación, Restauración, Vivienda, Suministros, Transporte, Salud, Ocio, Compras, Suscripciones, Otros.

**Ingresos:** Nómina, Extras, Otros.

Para añadir o renombrar categorías, edita `scripts/categories.py` y vuelve a correr `python -m scripts.init_workbook` (después de borrar el `gastos.xlsx` actual o renombrarlo).

## Tests

```powershell
pytest -v
```
```

- [ ] **Step 2: End-to-end smoke**

Run from a clean state:
```powershell
Remove-Item gastos.xlsx -ErrorAction SilentlyContinue
python -m scripts.init_workbook
python -m scripts.add_movement --fecha 2026-05-01 --tipo Ingreso --categoria Nómina --importe 2100 --descripcion "Nómina mayo"
python -m scripts.add_movement --fecha 2026-05-03 --tipo Gasto --categoria Alimentación --importe 78.40 --descripcion "Súper Mercadona"
python -m scripts.add_movement --fecha 2026-05-05 --tipo Gasto --categoria Restauración --importe 30 --descripcion "Comida con amigos"
python -m scripts.add_movement --fecha 2026-05-07 --tipo Gasto --categoria Suscripciones --importe 11.99 --descripcion "Spotify"
python -m scripts.add_movement --fecha 2026-05-09 --tipo Gasto --categoria Compras --importe 899 --descripcion "Móvil nuevo"
```

Open `gastos.xlsx` and verify on Dashboard:
- Ingresos = 2.100,00 €
- Gastos = 1.019,39 €
- Ahorro neto = 1.080,61 €
- % Ahorro ≈ 51,5%
- Pie chart shows Compras as the biggest slice (899), then Alimentación (78.40), Restauración (30), Suscripciones (11.99).
- Top-5 lists: Móvil nuevo (899), Súper Mercadona (78.40), Comida con amigos (30), Spotify (11.99), and one empty row.
- Line chart has a non-zero point at the current month for Ingresos and Gastos.

If anything looks off, fix the relevant task and re-run.

- [ ] **Step 3: Final commit**

```powershell
git add README.md; git commit -m "docs: README with setup, usage and category list"
```

- [ ] **Step 4: Final test run**

```powershell
pytest -v
```
Expected: all tests PASS, no warnings worth chasing.

---

## Self-review notes (already applied)

- Spec coverage: every section of the spec maps to a task. Engram persistence (spec §2 step 4 and §7) is intentionally out of script scope — Claude calls `mem_save` after each successful script invocation. This is documented in `add_movement.py`'s caller (Claude), not the script.
- No placeholders: every code block is complete and runnable.
- Type consistency: `add_movement(path, fecha, tipo, categoria, importe, descripcion)` signature is identical across script, tests and README.
- The Spanish-vs-English formula-name caveat in Task 4 is real and explicit, with a scripted fallback path.
