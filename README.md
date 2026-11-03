# SaldoISLegit

Generador del estado de cuenta de Independencia Shipping Lines (ISL) a partir de SAP Business One.

Toma el resultado de la consulta SQL, lo separa en hojas (ESCALAS, OPERATIVA, REEMBARCOS y OPERATIVA MURCHISON) con el formato de la planilla que se le envía a ISL, agrega una hoja RESUMEN con el saldo de cada una y controla que el total coincida con el saldo de SAP.

## Requisitos

- Python 3 con `pandas` y `openpyxl`:

  ```
  pip install pandas openpyxl
  ```

## Uso

1. En SAP Business One, ejecutar la consulta de `consulta_ec_isl_v1.1.sql`, copiar los datos y pegarlos en un Excel.
2. Hacer doble clic en `EC_ISL_v1.1.bat` y elegir ese Excel (o arrastrar el Excel sobre el `.bat`).
3. Se genera `EC ISL al DD.MM.AAAA.xlsx` en la misma carpeta que el Excel de origen.

También se puede correr desde la terminal:

```
python ec_isl_v1.1.py "ruta\al\resultado_consulta.xlsx"
python ec_isl_v1.1.py --cli "ruta\al\resultado_consulta.xlsx"
```

## Archivos

| Archivo | Descripción |
| --- | --- |
| `consulta_ec_isl_v1.1.sql` | Consulta para SAP B1 que arma los movimientos con su sección y estado de cobro. |
| `ec_isl_v1.1.py` | Script que genera el Excel del estado de cuenta (con ventana o por línea de comandos). |
| `EC_ISL_v1.1.bat` | Acceso directo para ejecutar el script con doble clic. |
