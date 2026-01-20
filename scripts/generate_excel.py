"""Generate quotation Excel file from template and JSON data."""

from __future__ import annotations

import json
from copy import copy
from pathlib import Path
from typing import Dict, Iterable, Tuple

from openpyxl import load_workbook

TEMPLATE_PATH = Path("templates/Formato de cotizaciones Antartida y Altavolt.xlsx")
DATA_PATH = Path("data/catalogo_conceptos.json")
OUTPUT_PATH = Path("dist/Cotizacion_Antartida.xlsx")

HEADER_KEYS = {
    "partida": {"partida"},
    "clave": {"clave"},
    "concepto": {"concepto"},
    "descripcion": {"descripcion", "descripción"},
    "unidad": {"unidad", "u"},
    "cantidad": {"cantidad", "cant"},
    "precio_unitario": {"precio unitario", "precio", "p.u."},
    "importe": {"importe", "total"},
}


def load_concepts() -> Iterable[dict]:
    with DATA_PATH.open("r", encoding="utf-8") as file:
        payload = json.load(file)
    return payload.get("conceptos", [])


def normalize(value: str) -> str:
    return " ".join(value.strip().lower().split())


def find_main_sheet(workbook) -> object:
    for sheet in workbook.worksheets:
        title = normalize(sheet.title)
        if "cotiz" in title:
            return sheet
    return workbook.active


def detect_header_row(sheet, max_rows: int = 200, max_columns: int = 20) -> Tuple[int, Dict[str, int]]:
    for row in range(1, max_rows + 1):
        header_map: Dict[str, int] = {}
        for col in range(1, max_columns + 1):
            value = sheet.cell(row=row, column=col).value
            if not isinstance(value, str):
                continue
            normalized = normalize(value)
            for key, aliases in HEADER_KEYS.items():
                if normalized in aliases:
                    header_map[key] = col
        if len(header_map) >= 3:
            return row, header_map
    raise ValueError("No se encontró la fila de encabezados de conceptos.")


def copy_row_format(sheet, source_row: int, target_row: int, max_columns: int = 60) -> None:
    for col in range(1, max_columns + 1):
        source_cell = sheet.cell(row=source_row, column=col)
        target_cell = sheet.cell(row=target_row, column=col)
        if source_cell.has_style:
            target_cell._style = copy(source_cell._style)
        target_cell.number_format = source_cell.number_format
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)


def write_concepts(sheet, header_row: int, header_map: Dict[str, int], concepts: Iterable[dict]) -> None:
    data_row = header_row + 1
    concepts = list(concepts)
    if not concepts:
        return

    template_last_row = data_row
    if len(concepts) > 1:
        sheet.insert_rows(data_row + 1, amount=len(concepts) - 1)

    for offset, concept in enumerate(concepts):
        current_row = data_row + offset
        if offset > 0:
            copy_row_format(sheet, template_last_row, current_row)
        for key, column in header_map.items():
            value = concept.get(key)
            cell = sheet.cell(row=current_row, column=column)
            if key == "importe" and cell.data_type == "f":
                continue
            if key in {"cantidad", "precio_unitario", "importe"} and value is not None:
                try:
                    value = float(value)
                except (TypeError, ValueError):
                    pass
            cell.value = value


def main() -> None:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"No se encontró la plantilla: {TEMPLATE_PATH}")
    if not DATA_PATH.exists():
        raise FileNotFoundError(f"No se encontró el catálogo: {DATA_PATH}")

    workbook = load_workbook(TEMPLATE_PATH)
    sheet = find_main_sheet(workbook)
    header_row, header_map = detect_header_row(sheet)
    concepts = load_concepts()
    write_concepts(sheet, header_row, header_map, concepts)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(OUTPUT_PATH)


if __name__ == "__main__":
    main()
