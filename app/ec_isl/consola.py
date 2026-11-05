"""
Uso por linea de comandos.

Razon de ser: generar el estado de cuenta sin abrir la ventana (python app\\iniciar.py --cli
archivo) y mostrar los controles como texto. La ventana tambien usa informar() para dejar el
detalle completo en la consola, que queda minimizada.
"""

import tkinter as tk
from pathlib import Path
from tkinter import filedialog

from .calculo import calcular
from .excel import guardar
from .lectura import TIPOS_ARCHIVO, leer_archivo


def elegir_archivo():
    """Abre un selector de archivos si no se paso la ruta por parametro."""
    try:
        raiz = tk.Tk()
        raiz.withdraw()
        ruta = filedialog.askopenfilename(
            title="Elegir el resultado de la consulta EC ISL",
            filetypes=TIPOS_ARCHIVO,
        )
        raiz.destroy()
        return ruta or None
    except Exception:
        return None


def informar(estado, salida):
    """Imprime el saldo de cada hoja y los controles contra SAP."""
    print("\nSaldo por hoja:")
    for h in estado.hojas:
        marca = "OK" if h.ok else f"DIFERENCIA (SAP: {h.esperado:,.2f})"
        print(f"  {h.nombre:<22} USD {h.saldo:>14,.2f}   {marca}")
    print(f"  {'TOTAL HOJAS':<22} USD {estado.total_hojas:>14,.2f}")
    print(f"  {'SALDO EN SAP':<22} USD {estado.total_sap:>14,.2f}   "
          f"{'OK' if estado.totales_ok else 'REVISAR'}")

    if estado.sin_hoja:
        print("\nATENCION: documentos abiertos sin hoja asignada (revisar Seccion en la consulta):")
        for d in estado.sin_hoja:
            print(f"  {d.nro}  {d.fecha:%d/%m/%Y}  USD {d.saldo:,.2f}  {d.seccion}")

    print(f"\nListo: {salida}")
    if not estado.todo_ok:
        print("Hay controles para revisar (ver arriba).")


def main_cli(args):
    ruta = args[0] if args else elegir_archivo()
    if not ruta:
        print("No se eligio ningun archivo.")
        return 1
    ruta = Path(ruta)
    print(f"Leyendo {ruta.name} ...")
    estado = calcular(leer_archivo(ruta))
    informar(estado, guardar(estado))
    return 0
