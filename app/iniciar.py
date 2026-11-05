"""
Punto de entrada del programa (lo ejecuta EC_ISL_v1.1.bat).

Razon de ser: decidir como arranca el programa, con la ventana (opcionalmente con un archivo
ya elegido) o por consola con --cli. No tiene logica propia: todo esta en el paquete ec_isl.

Uso:
    python app\\iniciar.py
    python app\\iniciar.py "ruta\\al\\resultado_consulta.xlsx"
    python app\\iniciar.py --cli "ruta\\al\\resultado_consulta.txt"
"""

import sys

from ec_isl.consola import main_cli
from ec_isl.errores import ErrorUsuario
from ec_isl.interfaz import abrir_ventana, mostrar_error_fatal

if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--cli"]
    if "--cli" in sys.argv:
        try:
            sys.exit(main_cli(args))
        except ErrorUsuario as e:
            print(f"\nERROR: {e}")
            sys.exit(1)
    else:
        try:
            abrir_ventana(args[0] if args else None)
        except Exception as e:
            mostrar_error_fatal(e)
            sys.exit(1)
