# Proyecto Gastos Claude

Registro personal de ingresos y gastos con dashboard web autocontenido. Datos en `data/movements.json`, web en `dashboard.html` (regenerada automáticamente al añadir movimientos).

## Setup (una vez)

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m scripts.init_data
```

Esto crea `data/movements.json` (importando los movimientos del legacy `gastos.xlsx` si existe) y la primera versión de `dashboard.html`.

## Uso diario

Hablás con Claude:

> "Hoy gasté 30 € comiendo con amigos."

Claude ejecuta:

```powershell
python -m scripts.add_movement --fecha 2026-05-09 --tipo Gasto --categoria Restauración --importe 30 --descripcion "Comida con amigos"
```

Esto: (1) actualiza `data/movements.json`, (2) regenera `dashboard.html`. Vos refrescás (F5) y ves el cambio.

## Comandos manuales

```powershell
python -m scripts.add_movement --fecha 2026-05-09 --tipo Gasto --categoria Restauración --importe 30 --descripcion "Comida"
python -m scripts.add_movement --edit-last --importe 35
python -m scripts.add_movement --delete-last
python -m scripts.build_dashboard          # solo regenera dashboard.html
```

## Tests

```powershell
pytest -v
```
