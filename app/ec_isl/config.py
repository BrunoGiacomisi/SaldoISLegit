"""
Configuracion del estado de cuenta.

Razon de ser: reunir las reglas que cambian por decision del negocio y no por cambios en el
programa: desde que fecha se toman los movimientos, que hojas tiene la planilla que se le
envia a ISL y que columnas tiene que traer la consulta de SAP. Si ISL pide una hoja nueva o
se modifica la consulta, se cambia aca.
"""

from datetime import date

VERSION = "1.1"
EMPRESA = "Independencia Shipping Lines"

# Movimientos anteriores a esta fecha no se incluyen (la planilla arranca en 2024).
FECHA_DESDE = date(2024, 1, 1)

# Hojas en el orden en que aparecen, con el ancho de columnas y el color de pestana de la planilla original.
HOJAS = {
    "ESCALAS":             {"anchos": [20.4, 13.7, 17.0, 17.0, 20.0, 63.4, 20.0], "tab": "83CBEB"},
    "OPERATIVA":           {"anchos": [20.0, 13.7, 17.0, 17.0, 24.4, 58.3, 20.0], "tab": "F6C6AD"},
    "REEMBARCOS":          {"anchos": [20.6, 13.7, 17.7, 17.0, 16.0, 64.4, 20.0], "tab": "F6C6AD"},
    "OPERATIVA MURCHISON": {"anchos": [24.4, 13.7, 17.0, 17.0, 20.0, 85.6, 20.0], "tab": "D9F2D0"},
}

# Hoja donde figuran los pagos a cuenta (sobrante de un recibo que no se aplico a ninguna factura).
HOJA_PAGOS_A_CUENTA = "OPERATIVA"

# Hoja final con el saldo de cada hoja y el total.
HOJA_RESUMEN = "RESUMEN"

# Columnas que tiene que traer la consulta "EC ISL - Estado de cuenta con secciones".
COLUMNAS_REQUERIDAS = [
    "Fecha", "Tipo", "Factura", "Total", "Saldo pendiente", "Proyecto",
    "Comentarios", "Seccion", "Estado", "Cobrado", "Fecha de cobro", "Detalle de pago",
]
COLUMNAS_IMPORTE = ["Total", "Saldo pendiente", "Cobrado"]
