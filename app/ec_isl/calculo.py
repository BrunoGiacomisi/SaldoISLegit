"""
Calculo del estado de cuenta.

Razon de ser: es el corazon del negocio. Decide que movimientos van en cada hoja, arma los
renglones de recibos, canjes y pagos a cuenta, y controla que el saldo de cada hoja coincida
con el saldo pendiente de SAP. No sabe nada de Excel ni de la ventana, asi que se puede
probar solo con datos. Si cambia la forma de registrar cobros o canjes en SAP, se toca aca.
"""

import re
from dataclasses import dataclass
from datetime import datetime

import pandas as pd

from .config import FECHA_DESDE, HOJA_PAGOS_A_CUENTA, HOJAS


@dataclass
class HojaCalculada:
    nombre: str
    filas: list
    esperado: float

    @property
    def saldo(self):
        return round(sum(f["total"] for f in self.filas), 2)

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
class EstadoCuenta:
    hojas: list
    total_sap: float
    sin_hoja: list

    @property
    def total_hojas(self):
        return round(sum(h.saldo for h in self.hojas), 2)

    @property
    def totales_ok(self):
        return abs(self.total_hojas - self.total_sap) < 0.05

    @property
    def todo_ok(self):
        return all(h.ok for h in self.hojas) and self.totales_ok and not self.sin_hoja


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


def preparar(df):
    """Facturas y notas de credito desde FECHA_DESDE, con los datos que necesita el estado de cuenta."""
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
    return docs


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


def calcular(consulta):
    """Arma las hojas del estado de cuenta a partir de la consulta leida y sus controles contra SAP."""
    docs = preparar(consulta)
    a_cuenta = pagos_a_cuenta(consulta)

    hojas = []
    for hoja in HOJAS:
        extra = a_cuenta if hoja == HOJA_PAGOS_A_CUENTA else []
        esperado = docs.loc[docs["Seccion"] == hoja, "Saldo pendiente"].sum()
        esperado = round(esperado + sum(f["total"] for f in extra), 2)
        hojas.append(HojaCalculada(hoja, movimientos_de_hoja(docs, hoja, extra), esperado))

    otras = docs[~docs["Seccion"].isin(HOJAS)]
    otras = otras[otras["Saldo pendiente"].abs() > 0.005]
    sin_hoja = []
    for _, r in otras.iterrows():
        seccion = r["Seccion"] if pd.notna(r["Seccion"]) else ""
        sin_hoja.append(DocSinHoja(r["nro"], r["Fecha"], r["Saldo pendiente"], str(seccion)))

    return EstadoCuenta(
        hojas=hojas,
        total_sap=round(consulta["Saldo pendiente"].sum(), 2),
        sin_hoja=sin_hoja,
    )
