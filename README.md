# SaldoISLegit

Generador del estado de cuenta de Independencia Shipping Lines (ISL) a partir de SAP Business One.

Toma el resultado de la consulta SQL, lo separa en hojas (ESCALAS, OPERATIVA, REEMBARCOS y OPERATIVA MURCHISON) con el formato de la planilla que se le envía a ISL, agrega una hoja RESUMEN con el saldo de cada una y controla que el total coincida con el saldo de SAP.

## Requisitos

- Python 3 con `pandas` y `openpyxl`:

  ```
  pip install -r requirements.txt
  ```

## Uso

1. En SAP Business One, ejecutar la consulta de `sql/consulta_ec_isl_v1.1.sql` y copiar los datos.
2. Hacer doble clic en `EC_ISL_v1.1.bat` y tocar **Pegar desde SAP**. También se puede elegir un Excel o `.txt` con el resultado de la consulta (o arrastrarlo sobre el `.bat`).
3. Se genera `EC ISL al DD.MM.AAAA.xlsx` en la carpeta **Descargas** de Windows.

Los importes se aceptan con el formato de SAP (`7.635,73`). Si alguno no se puede leer, el programa avisa en lugar de tomarlo como cero.

También se puede correr desde la terminal:

```
python app\iniciar.py "ruta\al\resultado_consulta.xlsx"
python app\iniciar.py --cli "ruta\al\resultado_consulta.txt"
```

## Estructura

```
SaldoISLegit/
├── EC_ISL_v1.1.bat            Acceso directo para abrir la ventana con doble clic
├── app/
│   ├── iniciar.py             Punto de entrada: abre la ventana o corre por consola (--cli)
│   └── ec_isl/
│       ├── config.py          Reglas del negocio: fecha desde, hojas, columnas de la consulta
│       ├── errores.py         Errores con mensajes pensados para el usuario
│       ├── lectura.py         Lee la consulta (portapapeles, .txt o Excel) e interpreta importes y fechas
│       ├── calculo.py         Arma los movimientos de cada hoja y controla los saldos contra SAP
│       ├── excel.py           Escribe la planilla con el formato de ISL y la guarda en Descargas
│       ├── consola.py         Uso por línea de comandos y resumen de controles en texto
│       └── interfaz.py        La ventana
├── sql/
│   └── consulta_ec_isl_v1.1.sql   Consulta para SAP B1 con los movimientos, su sección y estado de cobro
├── tests/
│   └── test_ec_isl.py         Pruebas automáticas con una consulta inventada
├── requirements.txt           Dependencias de Python
└── README.md
```

Cada módulo explica al principio su razón de ser. Como regla general: un cambio en las reglas del negocio va en `config.py` o `calculo.py`, un cambio en cómo se ve la planilla va en `excel.py` y un cambio en la pantalla va en `interfaz.py`.

## Pruebas

Antes de subir un cambio, correr desde la carpeta del proyecto:

```
python -m unittest discover tests
```

Las pruebas usan datos inventados que cubren cobros, notas de crédito, canjes, pagos a cuenta, movimientos anteriores a 2024 y documentos sin hoja asignada.
