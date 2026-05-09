"""JSON-backed investments store."""
from __future__ import annotations

import json
from datetime import date
from pathlib import Path
from typing import Optional


class InvalidOperationError(ValueError):
    """Raised when investment operation fields don't validate."""


_VALID_TIPOS = {"compra", "venta"}


def _validate(
    tipo: str,
    ticker: str,
    unidades: float,
    precio: float,
    comision: float,
    fecha: str,
) -> None:
    if tipo not in _VALID_TIPOS:
        raise InvalidOperationError(
            f"tipo inválido: {tipo!r}. Debe ser uno de {sorted(_VALID_TIPOS)}"
        )
    if not ticker or not ticker.strip():
        raise InvalidOperationError("ticker no puede estar vacío")
    if unidades <= 0:
        raise InvalidOperationError(
            f"unidades debe ser positivo, recibido: {unidades}"
        )
    if precio <= 0:
        raise InvalidOperationError(
            f"precio debe ser positivo, recibido: {precio}"
        )
    if comision < 0:
        raise InvalidOperationError(
            f"comision no puede ser negativa, recibido: {comision}"
        )
    # validate fecha as YYYY-MM-DD
    try:
        date.fromisoformat(str(fecha))
    except (ValueError, TypeError):
        raise InvalidOperationError(
            f"fecha no es una fecha válida (YYYY-MM-DD): {fecha!r}"
        )


def load(path: Path) -> dict:
    """Load store from JSON. Returns {'version': 1, 'operations': [...]}."""
    if not path.exists():
        raise FileNotFoundError(
            f"no existe {path}. Ejecuta primero `python -m scripts.init_data`."
        )
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save(path: Path, store: dict) -> None:
    """Atomic write via .tmp + replace."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=2)
    tmp.replace(path)


def _next_id(store: dict) -> int:
    if not store["operations"]:
        return 1
    return max(op["id"] for op in store["operations"]) + 1


def append(
    path: Path,
    fecha: str,
    tipo: str,
    ticker: str,
    nombre: str,
    unidades: float,
    precio: float,
    comision: float = 0.0,
    moneda: str = "EUR",
) -> int:
    """Append an investment operation. Returns the new id."""
    fecha_str = fecha.isoformat() if hasattr(fecha, "isoformat") else str(fecha)
    _validate(tipo, ticker, float(unidades), float(precio), float(comision), fecha_str)
    store = load(path)
    new_id = _next_id(store)
    store["operations"].append(
        {
            "id": new_id,
            "fecha": fecha_str,
            "tipo": tipo,
            "ticker": ticker,
            "nombre": nombre,
            "unidades": float(unidades),
            "precio": float(precio),
            "comision": float(comision),
            "moneda": moneda,
        }
    )
    save(path, store)
    return new_id


def edit_last(
    path: Path,
    *,
    fecha: Optional[str] = None,
    tipo: Optional[str] = None,
    ticker: Optional[str] = None,
    nombre: Optional[str] = None,
    unidades: Optional[float] = None,
    precio: Optional[float] = None,
    comision: Optional[float] = None,
    moneda: Optional[str] = None,
) -> None:
    """Edit the last operation's fields in place."""
    store = load(path)
    if not store["operations"]:
        raise InvalidOperationError("la lista está vacía, nada que editar")
    last = store["operations"][-1]

    final_fecha = (
        (fecha.isoformat() if hasattr(fecha, "isoformat") else str(fecha))
        if fecha is not None
        else last["fecha"]
    )
    final_tipo = tipo if tipo is not None else last["tipo"]
    final_ticker = ticker if ticker is not None else last["ticker"]
    final_unidades = float(unidades) if unidades is not None else last["unidades"]
    final_precio = float(precio) if precio is not None else last["precio"]
    final_comision = float(comision) if comision is not None else last["comision"]

    _validate(final_tipo, final_ticker, final_unidades, final_precio, final_comision, final_fecha)

    last["fecha"] = final_fecha
    last["tipo"] = final_tipo
    last["ticker"] = final_ticker
    if nombre is not None:
        last["nombre"] = nombre
    last["unidades"] = final_unidades
    last["precio"] = final_precio
    last["comision"] = final_comision
    if moneda is not None:
        last["moneda"] = moneda

    save(path, store)


def delete_last(path: Path) -> None:
    """Remove the last operation from the store."""
    store = load(path)
    if not store["operations"]:
        raise InvalidOperationError("la lista está vacía, nada que borrar")
    store["operations"].pop()
    save(path, store)
