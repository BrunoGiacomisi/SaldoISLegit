"""
EC ISL - Generador del estado de cuenta de Independencia Shipping Lines.

Razon de ser: agrupar la logica del programa separada por responsabilidad, para que cada
cambio se haga en un solo lugar:

    config.py     reglas que se ajustan a mano: desde que fecha, que hojas, que columnas
    errores.py    errores con mensajes pensados para el usuario
    lectura.py    leer la consulta de SAP (portapapeles, .txt o Excel) y dejarla con tipos correctos
    calculo.py    armar los movimientos de cada hoja y controlar los saldos contra SAP
    excel.py      escribir la planilla con el formato de ISL y guardarla en Descargas
    consola.py    uso por linea de comandos
    interfaz.py   la ventana

El programa se inicia desde app/iniciar.py.
"""
