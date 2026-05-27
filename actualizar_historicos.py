"""
Script para actualizar historicos.xlsx con precios de cierre del día.
Se ejecuta automáticamente via GitHub Actions al cierre del mercado.
"""

import requests
import openpyxl
from openpyxl import load_workbook
from datetime import date
import os
import re
import time

# ── CONFIGURACIÓN ─────────────────────────────────────────────
ECO_BASE = "https://bonos.ecovalores.com.ar/eco/ticker.php"
HISTORICOS_FILE = "historicos.xlsx"

TICKERS = [
    "S29Y6", "S12J6", "T30J6", "S17L6", "S31L6", "S14G6", "S31G6", "S30S6",
    "S30O6", "S30N6", "T15E7", "T30A7", "T31Y7", "T30J7",
    "TO26", "TY30P",
    "TZX26", "TZXO6", "TZXD6", "TZXM7", "TZXA7", "TZX27", "TZXS7", "TZXD7",
    "TZXM8", "TZX28", "TZXS8", "TZXM9",
    "M31G6", "TMF27", "TMG27", "TMF28", "TMG28",
    "TZV26", "D30S6", "TZV27", "TZV28",
    "AO27D", "AO28D", "AL29D", "AN29D", "AL30D", "AL35D", "AE38D", "AL41D",
    "GD29D", "GD30D", "GD35D", "GD38D", "GD41D", "GD46D",
    "BPC7D", "BPD7D", "BPA8D", "BPB8D",
    "TXMJ8", "TXMJ9", "TTJ26", "TTS26",
]

# ── FETCH PRECIO ──────────────────────────────────────────────
def fetch_precio(ticker):
    try:
        url = f"{ECO_BASE}?t={ticker}"
        resp = requests.get(url, timeout=10, headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "text/html",
            "Referer": "https://bonos.ecovalores.com.ar"
        })
        html = resp.text
        match = re.search(r'<td class="precioticker">\s*([\d.,]+)\s*</td>', html)
        if match:
            price_str = match.group(1).replace(".", "").replace(",", ".")
            return float(price_str)
    except Exception as e:
        print(f"  Error fetching {ticker}: {e}")
    return None

# ── ACTUALIZAR EXCEL ──────────────────────────────────────────
def actualizar_historicos():
    hoy = date.today()
    fecha_str = hoy.strftime("%Y-%m-%d")
    print(f"Actualizando historicos para {fecha_str}...")

    # Cargar o crear el Excel
    if os.path.exists(HISTORICOS_FILE):
        wb = load_workbook(HISTORICOS_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Historicos"
        # Crear header con todos los tickers
        ws.cell(row=1, column=1, value="Fecha")
        for col, ticker in enumerate(TICKERS, start=2):
            ws.cell(row=1, column=col, value=ticker)
        print("Archivo historicos.xlsx creado desde cero.")

    # Verificar si ya existe la fila de hoy
    fecha_col = None
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0] and str(row[0])[:10] == fecha_str:
            print(f"Ya existe fila para {fecha_str}, saliendo.")
            return

    # Encontrar próxima fila vacía
    next_row = ws.max_row + 1

    # Escribir fecha
    ws.cell(row=next_row, column=1, value=fecha_str)

    # Asegurar que todos los tickers estén en el header
    header = {ws.cell(row=1, column=c).value: c for c in range(2, ws.max_column + 1)}
    for ticker in TICKERS:
        if ticker not in header:
            new_col = ws.max_column + 1
            ws.cell(row=1, column=new_col, value=ticker)
            header[ticker] = new_col

    # Fetchear precios
    ok = 0
    err = 0
    for ticker in TICKERS:
        print(f"  Fetching {ticker}...", end=" ")
        precio = fetch_precio(ticker)
        if precio:
            col = header[ticker]
            ws.cell(row=next_row, column=col, value=precio)
            print(f"${precio}")
            ok += 1
        else:
            print("sin precio")
            err += 1
        time.sleep(0.4)

    wb.save(HISTORICOS_FILE)
    print(f"\nListo: {ok} precios guardados, {err} sin precio.")
    print(f"Archivo guardado: {HISTORICOS_FILE}")

if __name__ == "__main__":
    actualizar_historicos()
