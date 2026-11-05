"""
Armado del Excel del estado de cuenta.

Razon de ser: todo lo que tiene que ver con como se ve la planilla que se le envia a ISL
(colores, fuente, anchos de columna, formulas de saldo, hoja RESUMEN) y donde se guarda.
Recibe el estado de cuenta ya calculado, asi que un cambio de formato no puede alterar los
saldos.
"""

import ctypes
import uuid
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import HOJA_RESUMEN, HOJAS
from .errores import ErrorUsuario

ENCABEZADOS = ["Factura", "Fecha", "Total", "Saldo", "Proyecto", "REF.", "COMENTARIO"]

# Colores (tema de la planilla original)
AZUL_ENCABEZADO = "96DCF8"
VERDE_CERRADO = "D9F2D0"
NARANJA_RECIBO = "F6C6AD"
ROJO_NC = "FFFF0000"

FORMATO_USD = '[$USD]\\ #,##0.00;\\-[$USD]\\ #,##0.00'
FORMATO_FECHA = "dd/mm/yyyy"
FUENTE = "Aptos Narrow"

BORDE = Border(*(Side(style="thin"),) * 4)
CENTRO = Alignment(horizontal="center", vertical="center")
DERECHA = Alignment(horizontal="right", vertical="center")


def carpeta_descargas():
    """Carpeta Descargas configurada en Windows (puede estar movida, por ejemplo a OneDrive)."""
    try:
        guid = ctypes.create_string_buffer(uuid.UUID("374DE290-123F-4565-9164-39C4925E467B").bytes_le)
        ruta = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(guid, 0, None, ctypes.byref(ruta)) == 0:
            carpeta = Path(ruta.value)
            ctypes.windll.ole32.CoTaskMemFree(ruta)
            return carpeta
    except Exception:
        pass
    return Path.home() / "Downloads"


def relleno(color):
    return PatternFill("solid", fgColor=color)


def escribir_hoja(wb, hoja, filas):
    cfg = HOJAS[hoja]
    ws = wb.create_sheet(hoja)
    ws.sheet_properties.tabColor = cfg["tab"]

    for j, ancho in enumerate(cfg["anchos"], start=1):
        ws.column_dimensions[get_column_letter(j)].width = ancho

    for j, titulo in enumerate(ENCABEZADOS, start=1):
        c = ws.cell(row=1, column=j, value=titulo)
        c.font = Font(name=FUENTE, size=11)
        c.fill = relleno(AZUL_ENCABEZADO)
        c.alignment = CENTRO
        c.border = BORDE

    for i, f in enumerate(filas, start=2):
        valores = [f["nro"], f["fecha"].to_pydatetime(), f["total"],
                   "=ROUND(C2,2)" if i == 2 else f"=ROUND(D{i-1}+C{i},2)",
                   f["proyecto"], f["ref"], f["estado"]]
        es_recibo = f["tipo"] == "recibo"
        for j, v in enumerate(valores, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.border = BORDE
            c.alignment = DERECHA if j in (3, 4) else CENTRO
            color_fuente = ROJO_NC if (f["tipo"] == "nc" and j <= 3) else None
            c.font = Font(name=FUENTE, size=11, bold=es_recibo, color=color_fuente)
            if j == 2:
                c.number_format = FORMATO_FECHA
            elif j in (3, 4):
                c.number_format = FORMATO_USD
            if es_recibo:
                c.fill = relleno(NARANJA_RECIBO)
            elif j == 7 and f["estado"] in ("PAGO", "CANJEADO", "ANULADA"):
                c.fill = relleno(VERDE_CERRADO)

    ultima = len(filas) + 1
    if ultima >= 2:
        c = ws.cell(row=ultima, column=4)
        c.fill = relleno(NARANJA_RECIBO)
        c.font = Font(name=FUENTE, size=11, bold=True)

    ws.auto_filter.ref = f"A1:G{ultima}"
    ws.freeze_panes = "A2"


def escribir_resumen(wb, ultimas):
    """Hoja con el saldo final de cada hoja (formula a su ultima fila) y el total."""
    ws = wb.create_sheet(HOJA_RESUMEN)
    ws.sheet_properties.tabColor = AZUL_ENCABEZADO
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 20

    for j, titulo in enumerate(["Hoja", "Saldo"], start=1):
        c = ws.cell(row=1, column=j, value=titulo)
        c.font = Font(name=FUENTE, size=11)
        c.fill = relleno(AZUL_ENCABEZADO)
        c.alignment = CENTRO
        c.border = BORDE

    for i, (hoja, ultima) in enumerate(ultimas.items(), start=2):
        saldo = f"='{hoja}'!D{ultima}" if ultima >= 2 else 0
        for j, v in enumerate([hoja, saldo], start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = Font(name=FUENTE, size=11)
            c.border = BORDE
            c.alignment = DERECHA if j == 2 else CENTRO
        ws.cell(row=i, column=2).number_format = FORMATO_USD
        ws.cell(row=i, column=1).fill = relleno(HOJAS[hoja]["tab"])

    fin = len(ultimas) + 2
    for j, v in enumerate(["TOTAL", f"=ROUND(SUM(B2:B{fin - 1}),2)"], start=1):
        c = ws.cell(row=fin, column=j, value=v)
        c.font = Font(name=FUENTE, size=11, bold=True)
        c.fill = relleno(NARANJA_RECIBO)
        c.border = BORDE
        c.alignment = DERECHA if j == 2 else CENTRO
    ws.cell(row=fin, column=2).number_format = FORMATO_USD


def guardar(estado, carpeta=None):
    """Escribe el Excel del estado de cuenta (por defecto en Descargas) y devuelve su ruta."""
    wb = Workbook()
    wb.remove(wb.active)
    for hoja in estado.hojas:
        escribir_hoja(wb, hoja.nombre, hoja.filas)
    escribir_resumen(wb, {h.nombre: len(h.filas) + 1 for h in estado.hojas})

    carpeta = Path(carpeta) if carpeta else carpeta_descargas()
    salida = carpeta / f"EC ISL al {date.today():%d.%m.%Y}.xlsx"
    try:
        carpeta.mkdir(parents=True, exist_ok=True)
        wb.save(salida)
    except PermissionError:
        raise ErrorUsuario(f"No se pudo guardar {salida.name} en {carpeta}. "
                           "Si esta abierto en Excel, cerralo y proba de nuevo.")
    return salida
