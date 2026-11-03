"""
EC ISL - Generador del estado de cuenta de Independencia Shipping Lines (v1.1)

Toma el Excel que sale de la consulta SQL "EC ISL - Estado de cuenta con secciones"
(SAP Business One > Copiar datos > pegar en Excel) y genera el estado de cuenta
separado en hojas, con el mismo formato que la planilla que se le envia a ISL.

Uso:
    Doble clic en EC_ISL_v1.1.bat (o arrastrar el Excel sobre el .bat)
    python ec_isl_v1.1.py
    python ec_isl_v1.1.py "ruta\\al\\resultado_consulta.xlsx"
    python ec_isl_v1.1.py --cli "ruta\\al\\resultado_consulta.xlsx"
"""

import re
import sys
import threading
import tkinter as tk
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Configuracion
# --------------------------------------------------------------------------

# Movimientos anteriores a esta fecha no se incluyen (la planilla arranca en 2024).
FECHA_DESDE = date(2024, 1, 1)

# Hojas en el orden en que aparecen, con el ancho de columnas de la planilla original.
HOJAS = {
    "ESCALAS":             {"anchos": [20.4, 13.7, 17.0, 17.0, 20.0, 63.4, 20.0], "tab": "83CBEB"},
    "OPERATIVA":           {"anchos": [20.0, 13.7, 17.0, 17.0, 24.4, 58.3, 20.0], "tab": "F6C6AD"},
    "REEMBARCOS":          {"anchos": [20.6, 13.7, 17.7, 17.0, 16.0, 64.4, 20.0], "tab": "F6C6AD"},
    "OPERATIVA MURCHISON": {"anchos": [24.4, 13.7, 17.0, 17.0, 20.0, 85.6, 20.0], "tab": "D9F2D0"},
}

ENCABEZADOS = ["Factura", "Fecha", "Total", "Saldo", "Proyecto", "REF.", "COMENTARIO"]

# Hoja donde figuran los pagos a cuenta (sobrante de un recibo que no se aplico a ninguna factura).
HOJA_PAGOS_A_CUENTA = "OPERATIVA"

# Hoja final con el saldo de cada hoja y el total.
HOJA_RESUMEN = "RESUMEN"

COLUMNAS_REQUERIDAS = [
    "Fecha", "Tipo", "Factura", "Total", "Saldo pendiente", "Proyecto",
    "Comentarios", "Seccion", "Estado", "Cobrado", "Fecha de cobro", "Detalle de pago",
]

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


# --------------------------------------------------------------------------
# Lectura y preparacion de datos
# --------------------------------------------------------------------------

def elegir_archivo():
    """Abre un selector de archivos si no se paso la ruta por parametro."""
    try:
        raiz = tk.Tk()
        raiz.withdraw()
        ruta = filedialog.askopenfilename(
            title="Elegir el Excel con el resultado de la consulta EC ISL",
            filetypes=[("Excel", "*.xlsx *.xls")],
        )
        raiz.destroy()
        return ruta or None
    except Exception:
        return None


def numero_factura(texto, es_nc):
    """Devuelve 'A-11742' a partir de 'eTicket A-11742 -'; si no hay numero, el texto original."""
    m = re.search(r"A-\s?(\d+)", str(texto))
    if m:
        return f"A-{m.group(1)}"
    return str(texto).strip()


def recibo_de(detalle, fecha_cobro):
    """Busca en 'Detalle de pago' el numero del recibo cobrado en esa fecha."""
    if not isinstance(detalle, str) or pd.isna(fecha_cobro):
        return ""
    f = fecha_cobro.strftime("%d/%m/%Y")
    for parte in detalle.split(" | "):
        if parte.startswith(f"Cobro {f}"):
            m = re.search(r"A-\s?(\d+)", parte)
            if m:
                return f"A-{m.group(1)}"
    return ""


def fecha_canje(detalle):
    """Fecha de la primera compensacion por canje que aparece en 'Detalle de pago'."""
    if not isinstance(detalle, str):
        return None
    m = re.search(r"Canje[^|]*?(\d{2}/\d{2}/\d{4})", detalle)
    return datetime.strptime(m.group(1), "%d/%m/%Y") if m else None


class ErrorUsuario(Exception):
    """Error con un mensaje pensado para el usuario (sin detalle tecnico)."""


def leer_consulta(ruta):
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
                return df.dropna(how="all")

    raise ErrorUsuario(
        f"{ruta.name} no parece ser el resultado de la consulta EC ISL.\n"
        "  El programa necesita el Excel que sale de la consulta en SAP (Ejecutar > Copiar datos >\n"
        "  pegar en un Excel nuevo), no la planilla de ISL con las hojas ya armadas.\n"
        f"  Columnas que tiene que tener: {', '.join(COLUMNAS_REQUERIDAS)}"
    )


def preparar(df):
    df = df.copy()
    df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True)
    df["Fecha de cobro"] = pd.to_datetime(df["Fecha de cobro"], dayfirst=True, errors="coerce")
    for c in ["Total", "Saldo pendiente", "Cobrado"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    docs = df[df["Tipo"].isin(["Factura", "Nota de credito"])].copy()
    docs = docs[docs["Fecha"].dt.date >= FECHA_DESDE]

    docs["es_nc"] = docs["Tipo"].eq("Nota de credito")
    docs["signo"] = docs["Total"].apply(lambda x: -1 if x < 0 else 1)
    docs["nro"] = [numero_factura(t, nc) for t, nc in zip(docs["Factura"], docs["es_nc"])]
    docs["tiene_canje"] = docs["Detalle de pago"].fillna("").str.contains("Canje")

    # Importe cerrado por canje: lo cerrado que no fue cobro (solo si hubo canje; lo
    # cerrado con NC no se cuenta porque la NC ya figura como fila en la hoja).
    cerrado = docs["Total"].abs() - docs["Saldo pendiente"].abs()
    docs["canjeado"] = (cerrado - docs["Cobrado"]).where(
        docs["tiene_canje"] & docs["Estado"].ne("ANULADA"), 0.0
    ).clip(lower=0).round(2)

    # Estado como en la planilla: lo cerrado sin cobro y por canje se muestra CANJEADO.
    docs["estado_ec"] = docs["Estado"].fillna("")
    docs.loc[
        docs["estado_ec"].eq("PAGO") & docs["Cobrado"].le(0.005) & docs["tiene_canje"],
        "estado_ec",
    ] = "CANJEADO"
    return df, docs


def pagos_a_cuenta(df):
    """Cobros con parte sin aplicar: el sobrante es credito de ISL y va al EC como PAGO A CUENTA."""
    cob = df[df["Tipo"].eq("Cobro") & (df["Saldo pendiente"].abs() > 0.005)]
    return [
        {"orden": 2, "fecha": r["Fecha"], "nro": numero_factura(r["Factura"], False),
         "total": round(r["Saldo pendiente"], 2), "proyecto": "", "ref": "PAGO A CUENTA",
         "estado": "A CUENTA", "tipo": "recibo"}
        for _, r in cob.iterrows()
    ]


def movimientos_de_hoja(docs, hoja, extra=()):
    """Lista ordenada de filas de una hoja: documentos, recibos, canjes y los renglones extra."""
    d = docs[docs["Seccion"] == hoja]
    filas = list(extra)

    for _, r in d.iterrows():
        filas.append({
            "orden": 0, "fecha": r["Fecha"], "nro": r["nro"], "total": round(r["Total"], 2),
            "proyecto": r["Proyecto"] if pd.notna(r["Proyecto"]) else "",
            "ref": r["Comentarios"] if pd.notna(r["Comentarios"]) else "",
            "estado": r["estado_ec"], "tipo": "nc" if r["es_nc"] else "doc",
        })

    # Recibos: un renglon por cobro, con la parte que corresponde a esta hoja.
    cob = d[d["Cobrado"].abs() > 0.005].copy()
    if not cob.empty:
        cob["recibo"] = [recibo_de(x, f) for x, f in zip(cob["Detalle de pago"], cob["Fecha de cobro"])]
        cob["parte"] = cob["signo"] * cob["Cobrado"]
        for (f, rec), g in cob.groupby(["Fecha de cobro", "recibo"]):
            importe = round(-g["parte"].sum(), 2)
            if abs(importe) > 0.005:
                filas.append({"orden": 1, "fecha": f, "nro": rec, "total": importe,
                              "proyecto": "", "ref": "RECIBO", "estado": "", "tipo": "recibo"})

    # Canjes: compensaciones sin dinero (asientos de canje, cruces con otro socio).
    can = d[d["canjeado"] > 0.005].copy()
    if not can.empty:
        can["f_canje"] = [fecha_canje(x) or f for x, f in zip(can["Detalle de pago"], can["Fecha"])]
        can["parte"] = can["signo"] * can["canjeado"]
        for f, g in can.groupby("f_canje"):
            importe = round(-g["parte"].sum(), 2)
            if abs(importe) > 0.005:
                filas.append({"orden": 1, "fecha": pd.Timestamp(f), "nro": "", "total": importe,
                              "proyecto": "", "ref": "RECIBO CANJE", "estado": "", "tipo": "recibo"})

    def clave(x):
        m = re.search(r"(\d+)", x["nro"] or "")
        return (x["fecha"], x["orden"], int(m.group(1)) if m else 0)

    return sorted(filas, key=clave)


# --------------------------------------------------------------------------
# Escritura del Excel
# --------------------------------------------------------------------------

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

    saldo = 0.0
    for i, f in enumerate(filas, start=2):
        saldo += f["total"]
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
    return round(saldo, 2)


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


@dataclass
class ControlHoja:
    nombre: str
    saldo: float
    esperado: float

    @property
    def ok(self):
        return abs(self.saldo - self.esperado) < 0.05


@dataclass
class DocSinHoja:
    nro: str
    fecha: object
    saldo: float
    seccion: str


@dataclass
class Resultado:
    salida: Path
    hojas: list = field(default_factory=list)
    total_hojas: float = 0.0
    total_sap: float = 0.0
    sin_hoja: list = field(default_factory=list)

    @property
    def totales_ok(self):
        return abs(self.total_hojas - self.total_sap) < 0.05


def generar(ruta):
    """Arma el Excel del estado de cuenta y devuelve los controles de saldo."""
    ruta = Path(ruta)
    df, docs = preparar(leer_consulta(ruta))

    a_cuenta = pagos_a_cuenta(df)

    wb = Workbook()
    wb.remove(wb.active)
    hojas = []
    ultimas = {}
    for hoja in HOJAS:
        extra = a_cuenta if hoja == HOJA_PAGOS_A_CUENTA else []
        filas = movimientos_de_hoja(docs, hoja, extra)
        saldo = escribir_hoja(wb, hoja, filas)
        ultimas[hoja] = len(filas) + 1
        esperado = docs.loc[docs["Seccion"] == hoja, "Saldo pendiente"].sum()
        esperado = round(esperado + sum(f["total"] for f in extra), 2)
        hojas.append(ControlHoja(hoja, saldo, esperado))
    escribir_resumen(wb, ultimas)

    otras = docs[~docs["Seccion"].isin(HOJAS)]
    otras = otras[otras["Saldo pendiente"].abs() > 0.005]
    sin_hoja = []
    for _, r in otras.iterrows():
        seccion = r["Seccion"] if pd.notna(r["Seccion"]) else ""
        sin_hoja.append(DocSinHoja(r["nro"], r["Fecha"], r["Saldo pendiente"], str(seccion)))

    salida = ruta.with_name(f"EC ISL al {date.today():%d.%m.%Y}.xlsx")
    try:
        wb.save(salida)
    except PermissionError:
        raise ErrorUsuario(f"No se pudo guardar {salida.name}. Si esta abierto en Excel, cerralo y proba de nuevo.")

    return Resultado(
        salida=salida,
        hojas=hojas,
        total_hojas=round(sum(h.saldo for h in hojas), 2),
        total_sap=round(df["Saldo pendiente"].sum(), 2),
        sin_hoja=sin_hoja,
    )


def informar(res):
    """Imprime los mismos controles que la version de consola."""
    print("\nSaldo por hoja:")
    ok = True
    for h in res.hojas:
        marca = "OK" if h.ok else f"DIFERENCIA (SAP: {h.esperado:,.2f})"
        ok &= h.ok
        print(f"  {h.nombre:<22} USD {h.saldo:>14,.2f}   {marca}")
    print(f"  {'TOTAL HOJAS':<22} USD {res.total_hojas:>14,.2f}")
    print(f"  {'SALDO EN SAP':<22} USD {res.total_sap:>14,.2f}   "
          f"{'OK' if res.totales_ok else 'REVISAR'}")

    if res.sin_hoja:
        ok = False
        print("\nATENCION: documentos abiertos sin hoja asignada (revisar Seccion en la consulta):")
        for d in res.sin_hoja:
            print(f"  {d.nro}  {d.fecha:%d/%m/%Y}  USD {d.saldo:,.2f}  {d.seccion}")

    print(f"\nListo: {res.salida}")
    if not ok:
        print("Hay controles para revisar (ver arriba).")


def main_cli():
    args = [a for a in sys.argv[1:] if a != "--cli"]
    ruta = args[0] if args else elegir_archivo()
    if not ruta:
        print("No se eligio ningun archivo.")
        return 1
    ruta = Path(ruta)
    print(f"Leyendo {ruta.name} ...")
    informar(generar(ruta))
    return 0


# --------------------------------------------------------------------------
# Interfaz
# --------------------------------------------------------------------------

FONDO = "#FFFFFF"
TINTA = "#1C1C1E"
SUAVE = "#6B6B70"
LINEA = "#E6E6E8"
CAMPO = "#F6F6F7"
ACENTO = "#1F4E79"
ACENTO_ACTIVO = "#163A56"
OK_COLOR = "#157A3A"
ERROR = "#B42318"

FUENTE = "Segoe UI"
ANCHO = 420


def preparar_dpi():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            from ctypes import windll
            windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class Aplicacion(tk.Tk):
    def __init__(self, archivo_inicial=None):
        super().__init__()
        self.title("EC ISL")
        self.configure(bg=FONDO)
        self.ruta = None
        self.ocupado = False
        self._armar()
        self._ajustar(centrar=True)
        inicial = Path(archivo_inicial) if archivo_inicial else None
        if inicial and inicial.is_file():
            self._poner_archivo(inicial)
            self.after(200, self._generar)
        elif archivo_inicial:
            self._mostrar_estado(f"No se encontró el archivo:\n{archivo_inicial}", ERROR)

    def _ajustar(self, centrar=False):
        """Ajusta la ventana al contenido. El usuario no la redimensiona."""
        self.resizable(True, True)
        self.update_idletasks()
        ancho = max(self.winfo_reqwidth(), ANCHO + 56)
        alto = self.winfo_reqheight()
        if centrar:
            x = max((self.winfo_screenwidth() - ancho) // 2, 0)
            y = max((self.winfo_screenheight() - alto) // 3, 0)
        else:
            x, y = self.winfo_x(), self.winfo_y()
        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.resizable(False, False)

    def _armar(self):
        marco = tk.Frame(self, bg=FONDO, padx=28, pady=26)
        marco.pack(fill="both", expand=True)

        tk.Label(marco, text="EC ISL", bg=FONDO, fg=TINTA,
                 font=(FUENTE, 22), anchor="w").pack(fill="x")
        tk.Label(marco, text="Estado de cuenta", bg=FONDO, fg=SUAVE,
                 font=(FUENTE, 11), anchor="w").pack(fill="x", pady=(2, 22))

        tk.Label(marco, text="Excel de la consulta", bg=FONDO, fg=SUAVE,
                 font=(FUENTE, 9), anchor="w").pack(fill="x")

        self.zona_archivo = tk.Frame(marco, bg=CAMPO, highlightbackground=LINEA,
                                      highlightthickness=1, cursor="hand2")
        self.zona_archivo.pack(fill="x", pady=(6, 0))
        self.lbl_nombre = tk.Label(self.zona_archivo, text="Elegí un archivo",
                                    bg=CAMPO, fg=TINTA, font=(FUENTE, 11),
                                    anchor="w", justify="left", padx=12, pady=8,
                                    wraplength=ANCHO)
        self.lbl_nombre.pack(fill="x")
        self.lbl_carpeta = tk.Label(self.zona_archivo, text="SAP  ·  Ejecutar, copiar datos, pegar en Excel",
                                     bg=CAMPO, fg=SUAVE, font=(FUENTE, 9),
                                     anchor="w", justify="left", padx=12,
                                     wraplength=ANCHO)
        self.lbl_carpeta.pack(fill="x", pady=(0, 10))
        for w in (self.zona_archivo, self.lbl_nombre, self.lbl_carpeta):
            w.bind("<Button-1>", lambda e: self._elegir())

        self.boton = tk.Button(
            marco, text="Generar estado de cuenta", command=self._generar,
            bg="#E4E4E7", fg="#A1A1AA", activebackground="#E4E4E7", activeforeground="#A1A1AA",
            relief="flat", bd=0, font=(FUENTE, 11), pady=10, cursor="hand2",
            state="disabled",
        )
        self.boton.pack(fill="x", pady=(16, 0))
        self.boton.bind("<Enter>", self._hover_on)
        self.boton.bind("<Leave>", self._hover_off)
        tk.Frame(marco, bg=FONDO, width=ANCHO, height=1).pack(fill="x")

        self.estado = tk.Label(marco, text="", bg=FONDO, fg=SUAVE,
                                font=(FUENTE, 9), anchor="w", justify="left", wraplength=ANCHO)
        self.estado.pack(fill="x", pady=(12, 0))

    def _poner_archivo(self, ruta):
        self.ruta = Path(ruta)
        self.lbl_nombre.config(text=self.ruta.name)
        self.lbl_carpeta.config(text=str(self.ruta.parent))
        self._habilitar(True)
        self._mostrar_estado("")

    def _elegir(self):
        if self.ocupado:
            return
        inicial = str(self.ruta.parent) if self.ruta else str(Path.cwd())
        ruta = filedialog.askopenfilename(
            title="Elegir el Excel con el resultado de la consulta EC ISL",
            filetypes=[("Excel", "*.xlsx *.xls")],
            initialdir=inicial,
        )
        if ruta:
            self._poner_archivo(ruta)
            self._mostrar_estado("")

    def _habilitar(self, activo):
        if activo and not self.ocupado:
            self.boton.config(state="normal", bg=ACENTO, fg="#FFFFFF",
                              activebackground=ACENTO_ACTIVO, activeforeground="#FFFFFF",
                              cursor="hand2")
        else:
            self.boton.config(state="disabled", bg="#E4E4E7", fg="#A1A1AA",
                              activebackground="#E4E4E7", activeforeground="#A1A1AA",
                              cursor="arrow")

    def _hover_on(self, _evento):
        if str(self.boton.cget("state")) != "disabled":
            self.boton.config(bg=ACENTO_ACTIVO)

    def _hover_off(self, _evento):
        if str(self.boton.cget("state")) != "disabled":
            self.boton.config(bg=ACENTO)

    def _mostrar_estado(self, texto, color=SUAVE):
        self.estado.config(text=texto, fg=color)

    def _generar(self):
        if self.ocupado or not self.ruta:
            return
        self.ocupado = True
        self._habilitar(False)
        self._mostrar_estado("Generando…")
        threading.Thread(target=self._correr, daemon=True).start()

    def _correr(self):
        try:
            print(f"Leyendo {self.ruta.name} ...")
            resultado = generar(self.ruta)
            informar(resultado)
        except ErrorUsuario as e:
            print(f"\nERROR: {e}")
            self._en_ui(lambda m=str(e): self._fallo(m))
        except Exception as e:
            print(f"\nERROR: {e}")
            self._en_ui(lambda m=str(e): self._fallo(f"Error inesperado: {m}"))
        else:
            self._en_ui(lambda n=resultado.salida.name: self._listo(n))

    def _en_ui(self, fn):
        try:
            self.after(0, fn)
        except tk.TclError:
            pass

    def _fallo(self, mensaje):
        self.ocupado = False
        self._habilitar(True)
        self._mostrar_estado(mensaje, ERROR)
        self._ajustar()

    def _listo(self, nombre):
        self.ocupado = False
        self._habilitar(True)
        self._mostrar_estado(f"Listo: {nombre}", OK_COLOR)
        self._ajustar()


def mostrar_error_fatal(exc):
    import traceback
    texto = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    try:
        raiz = tk.Tk()
        raiz.withdraw()
        messagebox.showerror("EC ISL", texto)
        raiz.destroy()
    except Exception:
        Path(__file__).with_name("ec_isl_error.log").write_text(texto, encoding="utf-8")


if __name__ == "__main__":
    if "--cli" in sys.argv:
        try:
            sys.exit(main_cli())
        except ErrorUsuario as e:
            print(f"\nERROR: {e}")
            sys.exit(1)
    else:
        try:
            preparar_dpi()
            inicial = sys.argv[1] if len(sys.argv) > 1 else None
            Aplicacion(inicial).mainloop()
        except Exception as e:
            mostrar_error_fatal(e)
            sys.exit(1)
