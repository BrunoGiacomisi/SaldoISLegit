"""
Pruebas automaticas del estado de cuenta.

Razon de ser: verificar en segundos que un cambio no rompio los saldos. Usan una consulta de
SAP inventada (no datos reales de clientes) que cubre cada caso del negocio: cobro, nota de
credito, canje, pago a cuenta, movimiento anterior a FECHA_DESDE y documento sin hoja.

Se corren desde la carpeta del proyecto con:
    python -m unittest discover tests
"""

import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "app"))

from ec_isl.calculo import calcular  # noqa: E402
from ec_isl.errores import ErrorUsuario  # noqa: E402
from ec_isl.excel import guardar  # noqa: E402
from ec_isl.lectura import importe_sap, leer_texto  # noqa: E402

ENCABEZADO = ("Fecha\tTipo\tFactura\tTotal\tSaldo acumulado\tSaldo pendiente\tProyecto\tComentarios\t"
              "Seccion\tEstado\tCobrado\tFecha de cobro\tDetalle de pago\t")

# Mismo formato que copia SAP: tabulaciones, importes 7.635,73 y una tabulacion al final.
FILAS = [
    "10/12/2023\tFactura\teTicket A-100 -\t500,00\t500,00\t0,00\tESC01\t\tESCALAS\tPAGO\t500,00\t20/12/2023\t"
    "Cobro 20/12/2023: eTicket  (cob. propia) A-899 - (aplicado USD 500.00)\t",
    "15/01/2024\tFactura\teTicket A-101 -\t1.000,00\t1.000,00\t0,00\tESC01\tESCALA ENERO\tESCALAS\tPAGO\t1.000,00\t"
    "20/02/2024\tCobro 20/02/2024: eTicket  (cob. propia) A-900 - (aplicado USD 1000.00)\t",
    "01/03/2024\tFactura\teTicket A-102 -\t2.500,50\t3.500,50\t2.500,50\tESC02\tESCALA MARZO\tESCALAS\tPENDIENTE\t"
    "0,00\t\tSin aplicar\t",
    "02/03/2024\tFactura\teTicket A-103 -\t1.200,00\t4.700,50\t1.200,00\tTALLER\tMANTENIMIENTO\tOPERATIVA\t"
    "PENDIENTE\t0,00\t\tSin aplicar\t",
    "05/03/2024\tNota de credito\tNC de eTicket A-50 -\t-200,00\t4.500,50\t-200,00\tTALLER\tBONIFICACION\t"
    "OPERATIVA\tPENDIENTE\t0,00\t\tSin aplicar\t",
    "10/03/2024\tCobro\teTicket  (cob. propia) A-901 -\t-150,00\t4.350,50\t-150,00\tGENERI\t\t\t\t0,00\t\t\t",
    "01/04/2024\tFactura\teTicket A-105 -\t700,00\t5.050,50\t0,00\tREEMB\tREEMBARCO ABRIL\tREEMBARCOS\tPAGO\t"
    "0,00\t\tCanje 15/04/2024: asiento 123 (aplicado USD 700.00)\t",
]
SIN_HOJA = ("20/04/2024\tFactura\teTicket A-104 -\t400,00\t4.750,50\t400,00\tOTRO\t\tOTRA SECCION\t"
            "PENDIENTE\t0,00\t\tSin aplicar\t")


def consulta(*extra):
    return leer_texto("\n".join([ENCABEZADO, *FILAS, *extra]), "La prueba")


def hoja(estado, nombre):
    return next(h for h in estado.hojas if h.nombre == nombre)


class TestImportes(unittest.TestCase):
    def test_formatos(self):
        casos = {"7.635,73": 7635.73, "-238.364,60": -238364.60, "0,00": 0.0, "14,09": 14.09,
                 "800.00": 800.0, "7,635.73": 7635.73, "1.500": 1500.0, "USD 1.234,56": 1234.56}
        for texto, esperado in casos.items():
            with self.subTest(texto=texto):
                self.assertAlmostEqual(float(importe_sap(texto)), esperado)

    def test_importe_ilegible_avisa(self):
        with self.assertRaises(ErrorUsuario) as ctx:
            consulta(SIN_HOJA.replace("400,00", "4OO,00", 1))
        self.assertIn("4OO,00", str(ctx.exception))


class TestLectura(unittest.TestCase):
    def test_lee_lo_copiado_de_sap(self):
        df = consulta()
        self.assertEqual(len(df), len(FILAS))
        self.assertFalse(any(c.startswith("Unnamed") for c in df.columns))
        self.assertEqual(df["Total"].sum(), 5550.50)

    def test_texto_sin_la_consulta(self):
        with self.assertRaises(ErrorUsuario):
            leer_texto("hola\tmundo", "El portapapeles")


class TestCalculo(unittest.TestCase):
    def setUp(self):
        self.estado = calcular(consulta())

    def test_saldos_por_hoja_y_total(self):
        saldos = {h.nombre: h.saldo for h in self.estado.hojas}
        self.assertEqual(saldos, {"ESCALAS": 2500.50, "OPERATIVA": 850.00,
                                  "REEMBARCOS": 0.00, "OPERATIVA MURCHISON": 0.00})
        self.assertEqual(self.estado.total_sap, 3350.50)
        self.assertTrue(self.estado.todo_ok)

    def test_cobro_genera_recibo(self):
        recibos = [f for f in hoja(self.estado, "ESCALAS").filas if f["ref"] == "RECIBO"]
        self.assertEqual([(f["nro"], f["total"]) for f in recibos], [("A-900", -1000.00)])

    def test_pago_a_cuenta_va_a_operativa(self):
        a_cuenta = [f for f in hoja(self.estado, "OPERATIVA").filas if f["ref"] == "PAGO A CUENTA"]
        self.assertEqual([(f["nro"], f["total"]) for f in a_cuenta], [("A-901", -150.00)])

    def test_canje(self):
        filas = hoja(self.estado, "REEMBARCOS").filas
        self.assertEqual([f["estado"] for f in filas if f["tipo"] == "doc"], ["CANJEADO"])
        canjes = [f for f in filas if f["ref"] == "RECIBO CANJE"]
        self.assertEqual([(f["fecha"].strftime("%d/%m/%Y"), f["total"]) for f in canjes],
                         [("15/04/2024", -700.00)])

    def test_excluye_movimientos_anteriores_a_fecha_desde(self):
        nros = {f["nro"] for h in self.estado.hojas for f in h.filas}
        self.assertNotIn("A-100", nros)
        self.assertNotIn("A-899", nros)

    def test_documento_sin_hoja_se_avisa(self):
        estado = calcular(consulta(SIN_HOJA))
        self.assertEqual([d.nro for d in estado.sin_hoja], ["A-104"])
        self.assertFalse(estado.totales_ok)
        self.assertFalse(estado.todo_ok)


class TestExcel(unittest.TestCase):
    def test_guarda_hojas_y_resumen(self):
        estado = calcular(consulta())
        with tempfile.TemporaryDirectory() as carpeta:
            salida = guardar(estado, carpeta)
            wb = load_workbook(salida)
            self.assertEqual(wb.sheetnames, ["ESCALAS", "OPERATIVA", "REEMBARCOS",
                                             "OPERATIVA MURCHISON", "RESUMEN"])
            ws = wb["ESCALAS"]
            self.assertEqual([c.value for c in ws[1]],
                             ["Factura", "Fecha", "Total", "Saldo", "Proyecto", "REF.", "COMENTARIO"])
            self.assertEqual(ws.max_row, len(hoja(estado, "ESCALAS").filas) + 1)
            self.assertEqual(wb["RESUMEN"]["B2"].value, f"='ESCALAS'!D{ws.max_row}")
            wb.close()


if __name__ == "__main__":
    unittest.main()
