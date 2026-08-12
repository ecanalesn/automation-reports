#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
forecast_report.py — Previsión de ingresos para productos a plazos
==================================================================

Proyecta los ingresos futuros de productos vendidos a plazos. Para cada producto
que se paga en más de un plazo, calcula qué plazos quedan pendientes por cliente
y cuándo vencen, y agrega los ingresos esperados mes a mes hasta una fecha límite
configurable.

Excel de salida:
  • Una hoja "Summary" con los ingresos esperados por mes de todos los productos.
  • Una hoja por SKU de producto con los clientes activos y sus plazos pendientes
    por mes.

Esta es una versión GENÉRICA y anonimizada de un script de producción. Las tasas
de cambio, los SKU y los nombres de columnas son de ejemplo: adáptalos a tus datos.

Uso
---
    python forecast_report.py orders_region_a.xlsx orders_region_b.xlsx \
        --output forecast.xlsx --today 2026-06-10 --end 2027-06-30

Requisitos
----------
    pip install pandas openpyxl
"""
from __future__ import annotations

import argparse
import re
from collections import defaultdict
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------
END_DATE_DEFAULT = "2027-06-30"

# Tasas de cambio aproximadas a la divisa base. Sustituye por tu propia fuente.
CURRENCY_TO_BASE = {
    "EUR": 1.0,
    "USD": 0.93,
    "MXN": 0.0495,
    "CLP": 0.00097,
    "COP": 0.000242,
}

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")
INSTALMENTS_RE = re.compile(r"(\d+)\s*plazos?", re.IGNORECASE)  # "12 plazos"

HEADER_FILL = PatternFill("solid", fgColor="FF305496")
HEADER_FONT = Font(bold=True, color="FFFFFFFF")


# ---------------------------------------------------------------------------
# FUNCIONES AUXILIARES
# ---------------------------------------------------------------------------
def month_key(d: datetime) -> str:
    return d.strftime("%Y-%m")


def iter_months(start: datetime, end: datetime):
    """Genera el primer día de cada mes entre start y end, ambos incluidos."""
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        yield datetime(y, m, 1)
        m += 1
        if m > 12:
            m, y = 1, y + 1


def parse_instalments(product_name: str) -> int:
    """Número de plazos indicado en el nombre del producto (1 si es pago único)."""
    match = INSTALMENTS_RE.search(str(product_name))
    return int(match.group(1)) if match else 1


def to_base_currency(amount: float, currency: str) -> float:
    return float(amount) * CURRENCY_TO_BASE.get(str(currency).upper(), 1.0)


# ---------------------------------------------------------------------------
# LÓGICA PRINCIPAL
# ---------------------------------------------------------------------------
def build_forecast(order_files: list[str], today: datetime, end: datetime) -> dict:
    frames = [pd.read_excel(p) for p in order_files]
    df = pd.concat(frames, ignore_index=True)
    df.columns = [str(c).strip() for c in df.columns]

    # Nombres de columna esperados (de ejemplo): adáptalos a tu export.
    col_email = "Customer email"
    col_product = "Product name"
    col_status = "Order status"
    col_date = "Order date"
    col_total = "Order total"
    col_currency = "Currency"
    col_sku = "SKU"

    df[col_date] = pd.to_datetime(df[col_date], errors="coerce")

    # por (email, sku): total de plazos, completados, importe mensual, primera fecha
    plans: dict[tuple, dict] = {}
    for _, row in df.iterrows():
        email = str(row.get(col_email, "")).strip().lower()
        if not EMAIL_RE.match(email):
            continue
        product = row.get(col_product, "")
        total_instalments = parse_instalments(product)
        if total_instalments <= 1:
            continue  # pago único: no hay nada que prever
        sku = str(row.get(col_sku, product)).strip()
        key = (email, sku)
        plan = plans.setdefault(key, {
            "total": total_instalments,
            "completed": 0,
            "monthly": 0.0,
            "first_date": row.get(col_date),
        })
        status = str(row.get(col_status, "")).strip().lower()
        if status in {"completed", "completado"}:
            plan["completed"] += 1
        amount = to_base_currency(row.get(col_total, 0) or 0, row.get(col_currency, "EUR"))
        if amount and total_instalments:
            plan["monthly"] = amount  # importe por plazo
        if plan["first_date"] is None or (
            row.get(col_date) is not None and row[col_date] < plan["first_date"]
        ):
            plan["first_date"] = row.get(col_date)

    # Proyectar los plazos pendientes hacia meses futuros
    per_month_total: dict[str, float] = defaultdict(float)
    per_sku_rows: dict[str, list] = defaultdict(list)
    for (email, sku), plan in plans.items():
        pending = plan["total"] - plan["completed"]
        if pending <= 0 or plan["first_date"] is None:
            continue
        # El plazo n vence en primera_fecha + n meses.
        for n in range(plan["completed"], plan["total"]):
            due = plan["first_date"] + pd.DateOffset(months=n)
            due = due.to_pydatetime() if isinstance(due, pd.Timestamp) else due
            if today <= due <= end:
                key = month_key(due)
                per_month_total[key] += plan["monthly"]
                per_sku_rows[sku].append((email, key, round(plan["monthly"], 2)))

    return {"per_month_total": per_month_total, "per_sku_rows": per_sku_rows,
            "today": today, "end": end}


# ---------------------------------------------------------------------------
# SALIDA A EXCEL
# ---------------------------------------------------------------------------
def write_excel(forecast: dict, output_path: str) -> None:
    wb = Workbook()
    months = [month_key(d) for d in iter_months(forecast["today"], forecast["end"])]

    # Hoja resumen
    ws = wb.active
    ws.title = "Summary"
    ws.cell(1, 1, "Month").font = HEADER_FONT
    ws.cell(1, 1).fill = HEADER_FILL
    ws.cell(1, 2, "Expected income").font = HEADER_FONT
    ws.cell(1, 2).fill = HEADER_FILL
    for i, m in enumerate(months, start=2):
        ws.cell(i, 1, m)
        ws.cell(i, 2, round(forecast["per_month_total"].get(m, 0.0), 2))
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 18

    # Una hoja por SKU
    for sku, rows in forecast["per_sku_rows"].items():
        title = re.sub(r"[\\/*?:\[\]]", "", str(sku))[:31] or "SKU"
        sheet = wb.create_sheet(title)
        for c, header in enumerate(["Customer", "Due month", "Amount"], start=1):
            cell = sheet.cell(1, c, header)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center")
        for r, (email, m, amount) in enumerate(sorted(rows, key=lambda x: (x[1], x[0])), start=2):
            sheet.cell(r, 1, email)
            sheet.cell(r, 2, m)
            sheet.cell(r, 3, amount)
        for col in range(1, 4):
            sheet.column_dimensions[get_column_letter(col)].width = 24

    wb.save(output_path)
    total = sum(forecast["per_month_total"].values())
    print(f"Previsión guardada en {output_path}. Ingresos proyectados totales: {total:,.2f}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Previsión de ingresos de productos a plazos.")
    parser.add_argument("orders", nargs="+", help="Archivos de export de pedidos (.xlsx).")
    parser.add_argument("--output", default="forecast.xlsx")
    parser.add_argument("--today", default=None, help="Fecha de referencia AAAA-MM-DD (por defecto: hoy).")
    parser.add_argument("--end", default=END_DATE_DEFAULT, help="Horizonte de previsión AAAA-MM-DD.")
    args = parser.parse_args()

    today = datetime.strptime(args.today, "%Y-%m-%d") if args.today else datetime.now()
    end = datetime.strptime(args.end, "%Y-%m-%d")
    forecast = build_forecast(args.orders, today, end)
    write_excel(forecast, args.output)


if __name__ == "__main__":
    main()
