#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
chat_export.py — Exportador incremental de conversaciones (API de chat → índice + archivos)
===========================================================================================

Exporta las conversaciones de la API REST de un proveedor de chat a una carpeta
navegable:

  • Una hoja índice con UNA FILA POR CONVERSACIÓN (contacto, teléfono, correo,
    fechas de primer/último mensaje, número de adjuntos) más un hipervínculo que
    abre el PDF de esa conversación.
  • Por conversación: el export original (tal cual lo entrega el proveedor), un
    PDF legible y sus adjuntos en carpetas por conversación.

Característica clave: el export es REANUDABLE e INCREMENTAL. Al relanzarlo sobre
la misma carpeta, lee el índice anterior, vuelve a listar todas las
conversaciones y solo reexporta las nuevas o las que tengan mensajes nuevos desde
la última vez. Las que no han cambiado se reutilizan tal cual. Al final informa de
cuántas eran nuevas, cuántas actualizadas y cuántas reutilizadas.

Esta es una versión GENÉRICA y anonimizada de un script de producción. Los
endpoints de la API, la cabecera de autenticación y los nombres de campos son
marcadores de ejemplo: adáptalos a tu proveedor. La concurrencia, la lógica
incremental y el formato del índice/PDF son las partes reutilizables.

Uso
---
    export CHAT_API_KEY=tu_token
    python chat_export.py --out ./export

Requisitos
----------
    pip install requests openpyxl
"""
from __future__ import annotations

import argparse
import csv
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

# ---------------------------------------------------------------------------
# CONFIGURACIÓN  (marcadores de ejemplo: apúntalos a tu propio proveedor)
# ---------------------------------------------------------------------------
API_BASE = os.environ.get("CHAT_API_BASE", "https://api.example.com/v1")
API_KEY = os.environ.get("CHAT_API_KEY", "")
MAX_WORKERS = 6
REQUEST_TIMEOUT = 30
INDEX_CSV = "index.csv"
INDEX_XLSX = "index.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="FF305496")
HEADER_FONT = Font(bold=True, color="FFFFFFFF")


def api_headers() -> dict:
    if not API_KEY:
        sys.exit("Define primero la variable de entorno CHAT_API_KEY.")
    return {"Authorization": f"Bearer {API_KEY}"}


# ---------------------------------------------------------------------------
# ACCESO A LA API  (de ejemplo: sustituye por los endpoints reales de tu proveedor)
# ---------------------------------------------------------------------------
def list_conversations() -> list[dict]:
    """Devuelve todas las conversaciones con metadatos ligeros (id, contacto, último mensaje)."""
    conversations, page = [], 1
    while True:
        resp = requests.get(
            f"{API_BASE}/conversations",
            headers=api_headers(),
            params={"page": page, "per_page": 100},
            timeout=REQUEST_TIMEOUT,
        )
        resp.raise_for_status()
        batch = resp.json().get("data", [])
        if not batch:
            break
        conversations.extend(batch)
        page += 1
    return conversations


def export_conversation(conv: dict, out_dir: Path) -> dict:
    """Descarga una conversación (original + adjuntos) y genera su PDF."""
    conv_id = conv["id"]
    folder = out_dir / str(conv_id)
    folder.mkdir(parents=True, exist_ok=True)

    resp = requests.get(
        f"{API_BASE}/conversations/{conv_id}/export",
        headers=api_headers(),
        timeout=REQUEST_TIMEOUT,
    )
    resp.raise_for_status()
    (folder / "conversation.json").write_bytes(resp.content)

    # En el script real, el JSON se renderiza a un PDF legible y los adjuntos se
    # descargan en subcarpetas por archivo. Aquí queda como esbozo.
    pdf_path = folder / "conversation.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n% placeholder\n")

    return {
        "id": conv_id,
        "contact": conv.get("contact_name", ""),
        "phone": conv.get("phone", ""),
        "email": conv.get("email", ""),
        "first_message": conv.get("first_message_at", ""),
        "last_message": conv.get("last_message_at", ""),
        "attachments": conv.get("attachment_count", 0),
        "pdf": str(pdf_path.relative_to(out_dir)),
    }


# ---------------------------------------------------------------------------
# ESTADO INCREMENTAL
# ---------------------------------------------------------------------------
def read_previous_index(out_dir: Path) -> dict[str, dict]:
    """Carga el índice de la ejecución anterior indexado por id de conversación (vacío la primera vez)."""
    path = out_dir / INDEX_CSV
    if not path.exists():
        return {}
    with path.open(newline="", encoding="utf-8") as f:
        return {row["id"]: row for row in csv.DictReader(f)}


def needs_update(conv: dict, previous: dict) -> bool:
    """Una conversación se reexporta si es nueva o tiene un último mensaje más reciente."""
    prev = previous.get(str(conv["id"]))
    if not prev:
        return True
    return str(conv.get("last_message_at", "")) != prev.get("last_message", "")


def write_index(rows: list[dict], out_dir: Path) -> None:
    fieldnames = ["id", "contact", "phone", "email",
                  "first_message", "last_message", "attachments", "pdf"]
    with (out_dir / INDEX_CSV).open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "Conversations"
    headers = ["Contact", "Phone", "Email", "First message",
               "Last message", "Attachments", "Open"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(1, c, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    for r, row in enumerate(rows, start=2):
        ws.cell(r, 1, row["contact"])
        ws.cell(r, 2, row["phone"])
        ws.cell(r, 3, row["email"])
        ws.cell(r, 4, row["first_message"])
        ws.cell(r, 5, row["last_message"])
        ws.cell(r, 6, row["attachments"])
        link = ws.cell(r, 7, "abrir")
        link.hyperlink = row["pdf"]
        link.font = Font(color="FF0563C1", underline="single")
    for col, width in zip("ABCDEFG", (26, 16, 28, 18, 18, 12, 8)):
        ws.column_dimensions[col].width = width
    wb.save(out_dir / INDEX_XLSX)


# ---------------------------------------------------------------------------
# PROCESO PRINCIPAL
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(description="Exportador incremental de conversaciones.")
    parser.add_argument("--out", default="./export", help="Carpeta de salida.")
    args = parser.parse_args()

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    previous = read_previous_index(out_dir)
    conversations = list_conversations()
    to_export = [c for c in conversations if needs_update(c, previous)]
    reused = len(conversations) - len(to_export)

    print(f"{len(conversations)} conversaciones: {len(to_export)} a (re)exportar, "
          f"{reused} reutilizadas.")

    rows: list[dict] = []
    # Conservar las filas reutilizadas del índice anterior
    reused_ids = {str(c["id"]) for c in conversations} - {str(c["id"]) for c in to_export}
    rows.extend(previous[cid] for cid in reused_ids if cid in previous)

    new_count = updated_count = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {pool.submit(export_conversation, c, out_dir): c for c in to_export}
        for fut in as_completed(futures):
            conv = futures[fut]
            try:
                rows.append(fut.result())
                if str(conv["id"]) in previous:
                    updated_count += 1
                else:
                    new_count += 1
            except Exception as exc:  # noqa: BLE001
                print(f"  ! falló la conversación {conv.get('id')}: {exc}")
            time.sleep(0.05)  # no saturar el límite de peticiones

    write_index(rows, out_dir)
    print(f"Hecho. Nuevas: {new_count}, actualizadas: {updated_count}, reutilizadas: {reused}.")


if __name__ == "__main__":
    main()
