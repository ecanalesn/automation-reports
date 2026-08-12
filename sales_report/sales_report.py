#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sales_report.py — Generador de informe de ventas (exports de WooCommerce → Excel)
=================================================================================

Lee dos exports de pedidos de WooCommerce (por ejemplo, una tienda por región),
los consolida en un único conjunto de datos, clasifica cada pedido en una
categoría de producto mediante reglas configurables con expresiones regulares y
escribe una hoja de Excel con formato por cada categoría, sobre una plantilla.

Esta es una versión GENÉRICA y anonimizada de un script de producción. Todos los
nombres de productos, reglas de categorías y nombres de archivo son marcadores de
ejemplo: sustitúyelos por los tuyos en la sección de CONFIGURACIÓN.

Uso
---
    python sales_report.py orders_region_a.xlsx orders_region_b.xlsx \
        --template template.xlsx --output weekly_report.xlsx

Requisitos
----------
    pip install pandas openpyxl
"""
from __future__ import annotations

import argparse
import re
import unicodedata
from copy import copy
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

# ---------------------------------------------------------------------------
# CONFIGURACIÓN  (sustituye por tu propia taxonomía de productos)
# ---------------------------------------------------------------------------
HEADER_ROW = 4          # fila de la plantilla que contiene las cabeceras
MODEL_ROW = 5           # fila cuyo estilo se copia en cada fila de datos
DATA_START_ROW = 5      # primera fila donde se escriben los datos
DATE_CELL = (1, 2)      # (fila, columna) donde se escribe la fecha del informe
TIME_CELL = (2, 2)      # (fila, columna) donde se escribe la hora del informe

# Columnas a incluir en la salida, en orden. Las que falten se crean vacías.
OUTPUT_COLUMNS = [
    "Customer name",
    "Order date",
    "Order total",
    "Product name",
    "Customer email",
    "Billing country",
    "Coupon code",
    "Discount amount",
    "Store",
]

# Lista ordenada de hojas (categorías) a generar.
SHEET_ORDER = [
    "Course A",
    "Course B",
    "Course C",
    "Welcome pack",
]

# (regex compilada sobre el nombre normalizado del producto) -> hoja destino.
# Gana el primer patrón que coincide.
CATEGORY_PATTERNS = [
    (re.compile(r"welcome pack", re.IGNORECASE), "Welcome pack"),
    (re.compile(r"course a|program a", re.IGNORECASE), "Course A"),
    (re.compile(r"course b|program b", re.IGNORECASE), "Course B"),
    (re.compile(r"course c|program c", re.IGNORECASE), "Course C"),
]

HEADER_FILL = PatternFill(start_color="FFC6E2FF", end_color="FFC6E2FF", fill_type="solid")


# ---------------------------------------------------------------------------
# FUNCIONES AUXILIARES
# ---------------------------------------------------------------------------
def normalise(text: str) -> str:
    """Pasa a minúsculas, elimina acentos y espacios sobrantes."""
    text = "".join(
        c for c in unicodedata.normalize("NFD", str(text))
        if unicodedata.category(c) != "Mn"
    )
    return text.strip().lower()


def safe_sheet_name(name: str) -> str:
    """Devuelve un nombre de hoja válido para Excel (máx. 31 caracteres, sin símbolos inválidos)."""
    name = re.sub(r"[\\/*?:\[\]]", "", str(name)).strip()
    return (name or "Sheet")[:31]


def map_to_sheet(product_name) -> str | None:
    """Clasifica un producto en una de las hojas del informe, o None si es irrelevante."""
    if pd.isna(product_name):
        return None
    norm = normalise(product_name)
    for pattern, sheet in CATEGORY_PATTERNS:
        if pattern.search(norm):
            return sheet
    return None


def copy_cell_style(src, dst) -> None:
    """Copia el estilo visual completo de una celda a otra."""
    try:
        if src.has_style:
            dst._style = copy(src._style)
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
    except Exception:
        pass


def copy_row_style(ws, src_row: int, dst_row: int, max_col: int) -> None:
    """Copia el estilo de una fila completa (fila modelo) a otra fila."""
    for col in range(1, max_col + 1):
        copy_cell_style(ws.cell(src_row, col), ws.cell(dst_row, col))


def header_column_map(ws, header_row: int) -> dict[str, int]:
    """Asocia el texto normalizado de cada cabecera con su índice de columna."""
    mapping = {}
    for col in range(1, max(ws.max_column, 30) + 1):
        value = ws.cell(header_row, col).value
        if value is not None:
            mapping[normalise(value)] = col
    return mapping


# ---------------------------------------------------------------------------
# PROCESO PRINCIPAL
# ---------------------------------------------------------------------------
def build_report(order_files: list[str], template_path: str, output_path: str) -> None:
    print("Cargando archivos de pedidos…")
    frames = []
    for i, path in enumerate(order_files):
        df = pd.read_excel(path)
        df["Store"] = f"Store {i + 1}"
        frames.append(df)
    df = pd.concat(frames, ignore_index=True)

    # Localizar la columna con el nombre del producto
    product_col = next(
        (c for c in df.columns if normalise(c) in {"product name", "nombre del articulo"}),
        None,
    )
    if not product_col:
        raise RuntimeError("No se encontró una columna 'Product name' en los datos de entrada.")

    df = df.dropna(subset=[product_col])
    print(f"  {len(df)} pedidos tras descartar filas sin nombre de producto.")

    # Cargar la plantilla y prepararla
    wb = load_workbook(template_path)
    template = wb[wb.sheetnames[0]]
    template.title = "_template_"

    col_map = header_column_map(template, HEADER_ROW)
    final_mapping: dict[str, int] = {}
    next_col = (max(col_map.values()) + 1) if col_map else 1
    for col in OUTPUT_COLUMNS:
        norm = normalise(col)
        if norm in col_map:
            final_mapping[col] = col_map[norm]
        else:
            template.cell(HEADER_ROW, next_col, value=col)
            final_mapping[col] = next_col
            next_col += 1
        if col not in df.columns:
            df[col] = ""

    max_col = max(final_mapping.values())
    for col_idx in range(1, max_col + 1):
        cell = template.cell(HEADER_ROW, col_idx)
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True)

    # Clasificar los pedidos
    df["_sheet"] = df[product_col].apply(map_to_sheet)
    relevant = df[df["_sheet"].notna()].copy()
    print(f"  {len(relevant)} pedidos encajan en una categoría del informe "
          f"({len(df) - len(relevant)} descartados).")

    now = datetime.now()
    for sheet_name in SHEET_ORDER:
        rows = relevant[relevant["_sheet"] == sheet_name]
        ws = wb.copy_worksheet(template)
        ws.title = safe_sheet_name(sheet_name)
        ws.cell(*DATE_CELL, value=now.date())
        ws.cell(*TIME_CELL, value=now.strftime("%H:%M"))
        for dst_row, (_, row) in enumerate(rows.iterrows(), start=DATA_START_ROW):
            copy_row_style(ws, MODEL_ROW, dst_row, max_col)
            for col_name, excel_col in final_mapping.items():
                value = row.get(col_name)
                if isinstance(value, pd.Timestamp):
                    value = value.to_pydatetime()
                if pd.isna(value):
                    value = None
                ws.cell(dst_row, excel_col, value=value)
        print(f"    {ws.title}: {len(rows)} filas")

    wb.remove(template)
    wb.save(output_path)
    print(f"Informe guardado en {output_path} ({now:%d/%m/%Y %H:%M}).")


def main() -> None:
    parser = argparse.ArgumentParser(description="Generador de informe de ventas semanal.")
    parser.add_argument("orders", nargs="+", help="Uno o más exports de pedidos de WooCommerce (.xlsx).")
    parser.add_argument("--template", default="template.xlsx", help="Libro de plantilla.")
    parser.add_argument("--output", default="weekly_report.xlsx", help="Archivo de salida.")
    args = parser.parse_args()
    build_report(args.orders, args.template, args.output)


if __name__ == "__main__":
    main()
