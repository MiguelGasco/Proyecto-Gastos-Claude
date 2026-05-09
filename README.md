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
