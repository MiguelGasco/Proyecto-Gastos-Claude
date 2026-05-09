"""JSON-backed movements store."""
from datetime import date
from pathlib import Path
from typing import Optional
import json

from scripts.categories import VALID_TYPES, is_valid


class InvalidMovementError(ValueError):
    """Raised when movement fields don't validate."""


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


def load(path: Path) -> dict:
    """Load store from JSON. Returns {'version': 1, 'movements': [...]}."""
    if not path.exists():
        raise FileNotFoundError(
            f"no existe {path}. Ejecuta primero `python -m scripts.init_data`."
        )
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save(path: Path, store: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=2)
    tmp.replace(path)


def _next_id(store: dict) -> int:
    if not store["movements"]:
        return 1
    return max(m["id"] for m in store["movements"]) + 1


def append(
    path: Path,
    fecha: date,
    tipo: str,
    categoria: str,
    importe: float,
    descripcion: str = "",
) -> int:
    _validate(tipo, categoria, importe)
    store = load(path)
    new_id = _next_id(store)
    store["movements"].append(
        {
            "id": new_id,
            "fecha": fecha.isoformat() if hasattr(fecha, "isoformat") else str(fecha),
            "tipo": tipo,
            "categoria": categoria,
            "importe": float(importe),
            "descripcion": descripcion or "",
        }
    )
    save(path, store)
    return new_id


def edit_last(
    path: Path,
    *,
    fecha: Optional[date] = None,
    tipo: Optional[str] = None,
    categoria: Optional[str] = None,
    importe: Optional[float] = None,
    descripcion: Optional[str] = None,
) -> None:
    store = load(path)
    if not store["movements"]:
        raise InvalidMovementError("la lista está vacía, nada que editar")
    last = store["movements"][-1]
    final_tipo = tipo if tipo is not None else last["tipo"]
    final_cat = categoria if categoria is not None else last["categoria"]
    final_imp = importe if importe is not None else last["importe"]
    _validate(final_tipo, final_cat, float(final_imp))
    if fecha is not None:
        last["fecha"] = fecha.isoformat() if hasattr(fecha, "isoformat") else str(fecha)
    if tipo is not None:
        last["tipo"] = tipo
    if categoria is not None:
        last["categoria"] = categoria
    if importe is not None:
        last["importe"] = float(importe)
    if descripcion is not None:
        last["descripcion"] = descripcion
    save(path, store)


def delete_last(path: Path) -> None:
    store = load(path)
    if not store["movements"]:
        raise InvalidMovementError("la lista está vacía, nada que borrar")
    store["movements"].pop()
    save(path, store)
