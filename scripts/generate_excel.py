"""Generate a quotation Excel file from a template and JSON data."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

TEMPLATE_PATH = Path("templates/Formato de cotizaciones Antartida y Altavolt.xlsx")
DATA_PATH = Path("data/catalogo_conceptos.json")
OUTPUT_PATH = Path("dist/Cotizacion_Antartida.xlsx")

FIELDS = [
    "partida",
    "clave",
    "concepto",
    "descripcion",
    "unidad",
    "cantidad",
    "precio_unitario",
]


def load_concepts() -> Iterable[dict]:
    with DATA_PATH.open("r", encoding="utf-8") as file:
        payload = json.load(file)
    return payload.get("conceptos", [])


def find_first_empty_row(sheet, start_row: int = 1) -> int:
    """Return the first row where all cells in the first 7 columns are empty."""
    row = start_row
    while True:
        if all(sheet.cell(row=row, column=col).value is None for col in range(1, len(FIELDS) + 1)):
            return row
        row += 1


def write_concepts(sheet, concepts: Iterable[dict]) -> None:
    start_row = find_first_empty_row(sheet, start_row=1)
    for offset, concept in enumerate(concepts):
        row = start_row + offset
        for col, key in enumerate(FIELDS, start=1):
            sheet.cell(row=row, column=col).value = concept.get(key)


def main() -> None:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"No se encontró la plantilla: {TEMPLATE_PATH}")
    if not DATA_PATH.exists():
        raise FileNotFoundError(f"No se encontró el catálogo: {DATA_PATH}")

    workbook = load_workbook(TEMPLATE_PATH)
    sheet = workbook.active
    concepts = load_concepts()
    write_concepts(sheet, concepts)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(OUTPUT_PATH)


if __name__ == "__main__":
    main()
