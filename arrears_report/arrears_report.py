#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
arrears_report.py — Informe de impagos por categoría de producto
================================================================

Procesa los exports completos de pedidos (una o varias tiendas) y detecta a los
clientes que han dejado de pagar algún plazo. Opcionalmente puede cruzar los
datos con un archivo de "impagos ya conocidos" mantenido a mano, de modo que la
salida separa:

  1. Impagos ya reportados   → resaltados en un color.
  2. Impagos nuevos detectados → resaltados en otro color.

Así el revisor ve de un vistazo lo que ya se estaba controlando y lo que el
script ha encontrado y no estaba en la lista manual.

Esta es una versión GENÉRICA y anonimizada de un script de producción. Los SKU,
las reglas de categorías, el tratamiento de divisas y los nombres de columnas son
marcadores de ejemplo.

Uso
---
    python arrears_report.py orders_region_a.xlsx orders_region_b.xlsx \
        --known known_arrears.xlsx --output arrears_report.xlsx

Requisitos
----------
    pip install pandas openpyxl
"""
from __future__ import annotations

import argparse
import re
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------
# Cuántos días antes/después de la fecha teórica de vencimiento se siguen
# considerando "en plazo".
TOLERANCE_BEFORE_DAYS = 7
TOLERANCE_AFTER_DAYS = 25
# Un plazo mensual por encima de este valor (en divisa base) se trata como probable
# desajuste de divisa (importe dejado en moneda local) y se marca, no se suma.
ANOMALY_INSTALMENT_THRESHOLD = 1500.0

CURRENCY_TO_BASE = {
    "EUR": 1.0, "USD": 0.93, "MXN": 0.0495, "CLP": 0.00097, "COP": 0.000242,
}

INSTALMENTS_RE = re.compile(r"(\d+)\s*plazos?", re.IGNORECASE)

# Prefijo de SKU → hoja de categoría. Sustituye por tu propia taxonomía.
SKU_CATEGORY = [
    (re.compile(r"^CAT-A", re.IGNORECASE), "Category A"),
    (re.compile(r"^CAT-B", re.IGNORECASE), "Category B"),
    (re.compile(r"^CAT-C", re.IGNORECASE), "Category C"),
]
DEFAULT_CATEGORY = "Other"

REPORTED_FILL = PatternFill("solid", fgColor="FFBDD7EE")       # banda azul
DETECTED_BAND_FILL = PatternFill("solid", fgColor="FFF8CBAD")  # banda naranja
DETECTED_ROW_FILL = PatternFill("solid", fgColor="FFFCE4D6")   # filas naranja claro
HEADER_FONT = Font(bold=True)
THIN = Side(style="thin", color="FFBFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ---------------------------------------------------------------------------
# FUNCIONES AUXILIARES
# ---------------------------------------------------------------------------
def to_base(amount: float, currency: str) -> float:
    return float(amount) * CURRENCY_TO_BASE.get(str(currency).upper(), 1.0)


def category_for_sku(sku: str) -> str:
    for pattern, category in SKU_CATEGORY:
        if pattern.search(str(sku)):
            return category
    return DEFAULT_CATEGORY


def parse_instalments(product_name: str) -> int:
    match = INSTALMENTS_RE.search(str(product_name))
    return int(match.group(1)) if match else 1


def load_known(path: str | None) -> set[tuple]:
    """Carga los impagos ya reportados como un conjunto de claves (email, sku)."""
    if not path:
        return set()
    df = pd.read_excel(path)
    df.columns = [str(c).strip().lower() for c in df.columns]
    email_col = next((c for c in df.columns if "email" in c or "correo" in c), None)
    sku_col = next((c for c in df.columns if "sku" in c), None)
    if not email_col:
        return set()
    keys = set()
    for _, row in df.iterrows():
        email = str(row.get(email_col, "")).strip().lower()
        sku = str(row.get(sku_col, "")).strip() if sku_col else ""
        if email:
            keys.add((email, sku))
    return keys


# ---------------------------------------------------------------------------
# NÚCLEO
# ---------------------------------------------------------------------------
def detect_arrears(order_files: list[str]) -> pd.DataFrame:
    frames = [pd.read_excel(p) for p in order_files]
    df = pd.concat(frames, ignore_index=True)
    df.columns = [str(c).strip() for c in df.columns]

    col_email = "Customer email"
    col_product = "Product name"
    col_status = "Order status"
    col_date = "Order date"
    col_total = "Order total"
    col_currency = "Currency"
    col_sku = "SKU"

    df[col_date] = pd.to_datetime(df[col_date], errors="coerce")

    records = []
    grouped = df.groupby([col_email, col_sku], dropna=False)
    for (email, sku), group in grouped:
        product = group[col_product].iloc[0]
        total_instalments = parse_instalments(product)
        if total_instalments <= 1:
            continue
        completed = (group[col_status].astype(str).str.lower()
                     .isin(["completed", "completado"]).sum())
        first_date = group[col_date].min()
        if pd.isna(first_date):
            continue
        instalment_amount = to_base(group[col_total].iloc[0] or 0,
                                    group[col_currency].iloc[0] if col_currency in group else "EUR")
        anomaly = instalment_amount > ANOMALY_INSTALMENT_THRESHOLD

        # Plazos que deberían haber vencido "hoy" según la ventana de tolerancia
        today = datetime.now()
        expected_due = 0
        for n in range(total_instalments):
            due = first_date + pd.DateOffset(months=n)
            if due <= today + pd.Timedelta(days=TOLERANCE_AFTER_DAYS):
                expected_due += 1
        missed = max(expected_due - completed, 0)
        if missed > 0:
            records.append({
                "email": str(email).strip().lower(),
                "sku": str(sku).strip(),
                "category": category_for_sku(sku),
                "product": product,
                "completed": completed,
                "expected": expected_due,
                "missed": missed,
                "instalment_amount": round(instalment_amount, 2),
                "anomaly": anomaly,
            })
    return pd.DataFrame(records)


# ---------------------------------------------------------------------------
# SALIDA A EXCEL
# ---------------------------------------------------------------------------
def write_excel(arrears: pd.DataFrame, known: set[tuple], output_path: str) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    headers = ["Email", "SKU", "Product", "Completed", "Expected",
               "Missed", "Instalment", "Anomaly"]

    categories = arrears["category"].unique() if not arrears.empty else [DEFAULT_CATEGORY]
    for category in categories:
        subset = arrears[arrears["category"] == category]
        ws = wb.create_sheet(re.sub(r"[\\/*?:\[\]]", "", str(category))[:31] or "Sheet")

        reported = subset[subset.apply(lambda r: (r["email"], r["sku"]) in known, axis=1)]
        detected = subset[subset.apply(lambda r: (r["email"], r["sku"]) not in known, axis=1)]

        row = 1

        def write_section(title, data, band_fill, row_fill):
            nonlocal row
            band = ws.cell(row, 1, title)
            band.font = Font(bold=True, color="FF000000")
            band.fill = band_fill
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
            row += 1
            for c, h in enumerate(headers, start=1):
                cell = ws.cell(row, c, h)
                cell.font = HEADER_FONT
                cell.border = BORDER
            row += 1
            for _, rec in data.iterrows():
                values = [rec["email"], rec["sku"], rec["product"], rec["completed"],
                          rec["expected"], rec["missed"], rec["instalment_amount"],
                          "Yes" if rec["anomaly"] else ""]
                for c, v in enumerate(values, start=1):
                    cell = ws.cell(row, c, v)
                    cell.border = BORDER
                    if row_fill:
                        cell.fill = row_fill
                row += 1
            row += 1  # fila en blanco de separación

        if not reported.empty:
            write_section("IMPAGOS REPORTADOS", reported, REPORTED_FILL, None)
        write_section("DETECTADOS POR EL SCRIPT", detected, DETECTED_BAND_FILL, DETECTED_ROW_FILL)

        for col, width in zip("ABCDEFGH", (30, 14, 34, 11, 10, 8, 12, 9)):
            ws.column_dimensions[col].width = width

    wb.save(output_path)
    print(f"Informe de impagos guardado en {output_path}. "
          f"{0 if arrears.empty else len(arrears)} filas de impagos en {len(categories)} categorías.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Informe de impagos (plazos vencidos).")
    parser.add_argument("orders", nargs="+", help="Archivos de export de pedidos (.xlsx).")
    parser.add_argument("--known", default=None, help="Archivo opcional de impagos conocidos (.xlsx).")
    parser.add_argument("--output", default="arrears_report.xlsx")
    args = parser.parse_args()

    arrears = detect_arrears(args.orders)
    known = load_known(args.known)
    write_excel(arrears, known, args.output)


if __name__ == "__main__":
    main()
