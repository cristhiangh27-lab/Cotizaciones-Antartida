"""Generate a new quotation Excel file from a template and JSON data."""

from __future__ import annotations

import json
import unicodedata
from pathlib import Path
from typing import Iterable, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet
from copy import copy

TEMPLATE_PATH = Path("templates/Formato de cotizaciones Antartida y Altavolt.xlsx")
DATA_PATH = Path("data/cotizacion.json")
OUTPUT_PATH = Path("dist/Cotizacion_Generada.xlsx")
TEMPLATE_SHEET_NAME = "Lomas Country Temixco"

HEADER_ANCHORS = {
    "cliente": ("cliente",),
    "direccion": ("direccion", "dirección"),
    "telefono": ("telefono", "teléfono"),
    "fecha": ("fecha del presupuesto", "fecha"),
    "validez_dias": ("validez",),
}

TABLE_ANCHORS = ("descripcion", "descripción")
TABLE_FIELDS = ("unidad", "cantidad", "precio", "total")


def load_payload() -> dict:
    with DATA_PATH.open("r", encoding="utf-8") as file:
        return json.load(file)


def normalize(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    without_accents = "".join(char for char in normalized if not unicodedata.combining(char))
    return " ".join(without_accents.lower().strip().split())


def find_anchor_cell(sheet: Worksheet, anchors: Iterable[str]) -> Optional[Cell]:
    normalized_anchors = {normalize(anchor) for anchor in anchors}
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                if normalize(cell.value) in normalized_anchors:
                    return cell
    return None


def write_anchor_value(sheet: Worksheet, anchors: Iterable[str], value: str) -> None:
    cell = find_anchor_cell(sheet, anchors)
    if not cell:
        return
    target = cell.offset(column=1)
    if target.value is None or target.value == "":
        target.value = value
    else:
        cell.value = value


def find_table_header(sheet: Worksheet) -> Tuple[int, dict]:
    for row in sheet.iter_rows():
        header_map = {}
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            normalized = normalize(cell.value)
            if normalized in TABLE_ANCHORS:
                header_map["descripcion"] = cell.column
            if "unidad" in normalized:
                header_map["unidad"] = cell.column
            if "cantidad" in normalized:
                header_map["cantidad"] = cell.column
            if "precio" in normalized:
                header_map["precio_unitario"] = cell.column
            if "total" in normalized:
                header_map["total"] = cell.column
        if "descripcion" in header_map and any(key in header_map for key in TABLE_FIELDS):
            return row[0].row, header_map
    raise ValueError("No se encontró la fila de encabezado de la tabla de conceptos.")


def clear_existing_concepts(sheet: Worksheet, start_row: int, columns: Iterable[int]) -> None:
    row = start_row
    while True:
        values = [sheet.cell(row=row, column=col).value for col in columns]
        if all(value is None for value in values):
            break
        for col in columns:
            cell = sheet.cell(row=row, column=col)
            if cell.data_type != "f":
                cell.value = None
        row += 1


def copy_row_style(sheet: Worksheet, source_row: int, target_row: int) -> None:
    sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height
    for col in range(1, sheet.max_column + 1):
        source_cell = sheet.cell(row=source_row, column=col)
        target_cell = sheet.cell(row=target_row, column=col)
        if source_cell.has_style:
            target_cell._style = copy(source_cell._style)
        target_cell.number_format = source_cell.number_format
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)


def write_concepts(sheet: Worksheet, concepts: Iterable[dict], header_row: int, header_map: dict) -> None:
    data_row = header_row + 1
    target_columns = [header_map[key] for key in header_map]
    clear_existing_concepts(sheet, data_row, target_columns)

    base_row = data_row
    for index, concept in enumerate(concepts):
        row = data_row + index
        if row > base_row:
            sheet.insert_rows(row)
            copy_row_style(sheet, base_row, row)
        descripcion_col = header_map["descripcion"]
        descripcion_cell = sheet.cell(row=row, column=descripcion_col)
        descripcion_cell.value = concept.get("descripcion")
        alignment = copy(descripcion_cell.alignment)
        alignment.wrap_text = True
        descripcion_cell.alignment = alignment

        for field in ("unidad", "cantidad", "precio_unitario", "total"):
            if field not in header_map:
                continue
            cell = sheet.cell(row=row, column=header_map[field])
            if field == "total" and cell.data_type == "f":
                continue
            if field == "total":
                cantidad = float(concept.get("cantidad", 0) or 0)
                precio = float(concept.get("precio_unitario", 0) or 0)
                cell.value = cantidad * precio
            else:
                value = concept.get(field)
                cell.value = value


def main() -> None:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"No se encontró la plantilla: {TEMPLATE_PATH}")
    if not DATA_PATH.exists():
        raise FileNotFoundError(f"No se encontró el catálogo: {DATA_PATH}")

    payload = load_payload()
    proyecto = payload.get("proyecto", {})
    conceptos = payload.get("conceptos", [])

    workbook = load_workbook(TEMPLATE_PATH)
    if TEMPLATE_SHEET_NAME not in workbook.sheetnames:
        raise ValueError(f"No se encontró la hoja plantilla: {TEMPLATE_SHEET_NAME}")

    template_sheet = workbook[TEMPLATE_SHEET_NAME]
    new_sheet = workbook.copy_worksheet(template_sheet)
    new_sheet.title = proyecto.get("nombre_hoja", "Cotizacion")
    template_sheet.sheet_state = "hidden"

    write_anchor_value(new_sheet, HEADER_ANCHORS["cliente"], proyecto.get("cliente", ""))
    write_anchor_value(new_sheet, HEADER_ANCHORS["direccion"], proyecto.get("direccion", ""))
    write_anchor_value(new_sheet, HEADER_ANCHORS["telefono"], proyecto.get("telefono", ""))
    write_anchor_value(new_sheet, HEADER_ANCHORS["fecha"], proyecto.get("fecha", ""))
    write_anchor_value(new_sheet, HEADER_ANCHORS["validez_dias"], str(proyecto.get("validez_dias", "")))

    header_row, header_map = find_table_header(new_sheet)
    write_concepts(new_sheet, conceptos, header_row, header_map)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(OUTPUT_PATH)


if __name__ == "__main__":
    main()
