"""
Lectura de la consulta de SAP.

Razon de ser: la consulta puede llegar de tres formas (pegada desde el portapapeles, guardada
en un .txt o pegada en un Excel) y con los importes en formato de SAP (7.635,73). Este modulo
las convierte todas en la misma tabla, con fechas como fechas e importes como numeros, para
que el calculo no tenga que saber de donde vinieron los datos. Si un importe no se puede
leer, avisa en lugar de tomarlo como cero.
"""

import csv
import io
import re
from pathlib import Path

import pandas as pd

from .config import COLUMNAS_IMPORTE, COLUMNAS_REQUERIDAS
from .errores import ErrorUsuario

TIPOS_ARCHIVO = [("Consulta EC ISL", "*.xlsx *.xls *.txt"), ("Excel", "*.xlsx *.xls"), ("Texto", "*.txt")]


def importe_sap(valor):
    """Pasa '7.635,73' (formato de SAP) a '7635.73'; tambien acepta '7,635.73'."""
    if not isinstance(valor, str):
        return valor
    t = valor.replace("USD", "").replace(" ", "").strip()
    if "," in t:
        if "." in t and t.rfind(".") > t.rfind(","):
            return t.replace(",", "")
        return t.replace(".", "").replace(",", ".")
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t):
        return t.replace(".", "")
    return t


def normalizar(df):
    """Fechas como fechas e importes como numeros, vengan de Excel o de texto."""
    df = df.copy()
    df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True)
    df["Fecha de cobro"] = pd.to_datetime(df["Fecha de cobro"], dayfirst=True, errors="coerce")
    for c in COLUMNAS_IMPORTE:
        valores = pd.to_numeric(df[c].map(importe_sap), errors="coerce")
        ilegibles = df[c].notna() & df[c].astype(str).str.strip().ne("") & valores.isna()
        if ilegibles.any():
            ejemplos = ", ".join(df.loc[ilegibles, c].astype(str).head(3))
            raise ErrorUsuario(f"Hay {ilegibles.sum()} importe(s) en la columna {c} que no se "
                               f"pudieron leer (por ejemplo: {ejemplos}).")
        df[c] = valores.fillna(0.0)
    return df


def leer_excel(ruta):
    """Busca, en cualquier hoja y en las primeras filas, la tabla con las columnas de la consulta."""
    try:
        hojas = pd.read_excel(ruta, sheet_name=None, header=None, nrows=20)
    except PermissionError:
        raise ErrorUsuario(f"No se pudo abrir {ruta.name}. Si esta abierto en Excel, cerralo y proba de nuevo.")
    except Exception as e:
        raise ErrorUsuario(f"No se pudo leer {ruta.name} como Excel ({e}).")

    requeridas = set(COLUMNAS_REQUERIDAS)
    for nombre, muestra in hojas.items():
        for fila in range(len(muestra)):
            valores = {str(v).strip() for v in muestra.iloc[fila].tolist()}
            if requeridas <= valores:
                df = pd.read_excel(ruta, sheet_name=nombre, header=fila)
                df.columns = [str(c).strip() for c in df.columns]
                return normalizar(df.dropna(how="all"))

    raise ErrorUsuario(
        f"{ruta.name} no parece ser el resultado de la consulta EC ISL.\n"
        "  El programa necesita el Excel que sale de la consulta en SAP (Ejecutar > Copiar datos >\n"
        "  pegar en un Excel nuevo), no la planilla de ISL con las hojas ya armadas.\n"
        f"  Columnas que tiene que tener: {', '.join(COLUMNAS_REQUERIDAS)}"
    )


def leer_texto(texto, origen):
    """Lee la consulta copiada de SAP: texto separado por tabulaciones, con la fila de encabezados."""
    lineas = texto.splitlines()
    requeridas = set(COLUMNAS_REQUERIDAS)
    for i, linea in enumerate(lineas[:20]):
        if requeridas <= {c.strip() for c in linea.split("\t")}:
            try:
                df = pd.read_csv(io.StringIO("\n".join(lineas[i:])), sep="\t", dtype=str,
                                 quoting=csv.QUOTE_NONE, keep_default_na=False, na_values=[""])
            except Exception as e:
                raise ErrorUsuario(f"No se pudieron leer los datos de {origen} ({e}).")
            df.columns = [str(c).strip() for c in df.columns]
            df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
            return normalizar(df.dropna(how="all"))

    raise ErrorUsuario(
        f"{origen} no tiene el resultado de la consulta EC ISL.\n"
        "  En SAP, ejecuta la consulta y copia los datos con los encabezados.\n"
        f"  Columnas que tiene que tener: {', '.join(COLUMNAS_REQUERIDAS)}"
    )


def leer_archivo(ruta):
    """Lee la consulta desde un Excel o desde un .txt con lo copiado de SAP."""
    ruta = Path(ruta)
    if ruta.suffix.lower() not in (".txt", ".tsv"):
        return leer_excel(ruta)
    try:
        datos = ruta.read_bytes()
    except OSError as e:
        raise ErrorUsuario(f"No se pudo abrir {ruta.name} ({e}).")
    if datos[:2] in (b"\xff\xfe", b"\xfe\xff"):
        texto = datos.decode("utf-16")
    else:
        try:
            texto = datos.decode("utf-8-sig")
        except UnicodeDecodeError:
            texto = datos.decode("cp1252")
    return leer_texto(texto, ruta.name)
