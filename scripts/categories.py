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
