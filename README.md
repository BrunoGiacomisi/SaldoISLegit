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
python app\ec_isl_v1.1.py "ruta\al\resultado_consulta.xlsx"
python app\ec_isl_v1.1.py --cli "ruta\al\resultado_consulta.txt"
```

## Estructura

```
SaldoISLegit/
├── EC_ISL_v1.1.bat          Acceso directo para abrir la ventana con doble clic
├── app/
│   └── ec_isl_v1.1.py       Script que genera el Excel (con ventana o por línea de comandos)
├── sql/
│   └── consulta_ec_isl_v1.1.sql   Consulta para SAP B1 con los movimientos, su sección y estado de cobro
├── requirements.txt         Dependencias de Python
└── README.md
```
