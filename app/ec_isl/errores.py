"""
Errores pensados para mostrarle al usuario.

Razon de ser: distinguir los problemas esperables (el Excel esta abierto, el portapapeles no
tiene la consulta, un importe no se puede leer) de las fallas del programa. Los primeros se
muestran tal cual en la ventana o la consola; las demas aparecen como "Error inesperado".
"""


class ErrorUsuario(Exception):
    """Error con un mensaje pensado para el usuario (sin detalle tecnico)."""
