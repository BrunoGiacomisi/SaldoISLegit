"""
Ventana del programa.

Razon de ser: es la forma de uso diaria: pegar desde SAP o elegir un archivo, generar y ver
los controles. Solo muestra y coordina; los datos los lee lectura.py, el calculo lo hace
calculo.py y el Excel lo arma excel.py. Si hay que cambiar algo de lo que se ve en pantalla,
se toca aca sin riesgo de alterar los saldos.
"""

import os
import subprocess
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox

from .calculo import calcular
from .config import EMPRESA, FECHA_DESDE, HOJA_RESUMEN, HOJAS, VERSION
from .consola import informar
from .errores import ErrorUsuario
from .excel import carpeta_descargas, guardar
from .lectura import TIPOS_ARCHIVO, leer_archivo, leer_texto

FONDO = "#FFFFFF"
TINTA = "#1C1C1E"
SUAVE = "#6B6B70"
LINEA = "#E6E6E8"
CAMPO = "#F6F6F7"
ACENTO = "#1F4E79"
ACENTO_ACTIVO = "#163A56"
PEGAR = "#E8EEF5"
PEGAR_ACTIVO = "#D5E0EC"
OK_COLOR = "#157A3A"
ERROR = "#B42318"

FUENTE_UI = "Segoe UI"
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
        self.title(f"Estado de cuenta ISL - {EMPRESA}")
        self.configure(bg=FONDO)
        self.ruta = None
        self.texto = None
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
            y = max(min(y, self.winfo_screenheight() - alto - 80), 0)
        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.resizable(False, False)

    def _armar(self):
        marco = tk.Frame(self, bg=FONDO, padx=28, pady=26)
        marco.pack(fill="both", expand=True)

        tk.Label(marco, text="Estado de cuenta ISL", bg=FONDO, fg=TINTA,
                 font=(FUENTE_UI, 20), anchor="w").pack(fill="x")
        tk.Label(marco, text=EMPRESA, bg=FONDO, fg=SUAVE,
                 font=(FUENTE_UI, 11), anchor="w").pack(fill="x", pady=(2, 0))
        tk.Label(marco, text=f"SAP Business One  ·  Movimientos desde {FECHA_DESDE:%d/%m/%Y}  ·  v{VERSION}",
                 bg=FONDO, fg=SUAVE, font=(FUENTE_UI, 9), anchor="w").pack(fill="x", pady=(2, 16))

        self.ayuda = tk.Frame(marco, bg=FONDO)
        self.ayuda.pack(fill="x")
        tk.Frame(self.ayuda, bg=LINEA, height=1).pack(fill="x", pady=(0, 14))
        pasos = tk.Frame(self.ayuda, bg=FONDO)
        pasos.pack(fill="x", pady=(0, 18))
        pasos.columnconfigure(1, weight=1)
        textos = [
            "En SAP, ejecutar la consulta EC ISL y copiar los datos.",
            "Tocar «Pegar desde SAP». También se puede elegir un Excel o .txt con la consulta.",
            f"Se crea «EC ISL al {date.today():%d.%m.%Y}.xlsx» en Descargas, "
            f"con una hoja por sección y el {HOJA_RESUMEN}.",
        ]
        for i, texto in enumerate(textos):
            tk.Label(pasos, text=str(i + 1), bg=FONDO, fg=ACENTO, font=(FUENTE_UI, 9, "bold"),
                     width=2, anchor="nw").grid(row=i, column=0, sticky="nw", pady=1)
            tk.Label(pasos, text=texto, bg=FONDO, fg=TINTA, font=(FUENTE_UI, 9), anchor="w",
                     justify="left", wraplength=ANCHO - 24).grid(row=i, column=1, sticky="w", pady=1)

        self.lbl_excel = tk.Label(marco, text="Datos de la consulta", bg=FONDO, fg=SUAVE,
                                  font=(FUENTE_UI, 9), anchor="w")
        self.lbl_excel.pack(fill="x")

        origen = tk.Frame(marco, bg=FONDO)
        origen.pack(fill="x", pady=(6, 0))
        self.boton_pegar = tk.Button(
            origen, text="Pegar\ndesde SAP", command=self._pegar,
            bg=PEGAR, fg=ACENTO, activebackground=PEGAR_ACTIVO, activeforeground=ACENTO_ACTIVO,
            relief="flat", bd=0, font=(FUENTE_UI, 10), padx=16, cursor="hand2",
        )
        self.boton_pegar.pack(side="right", fill="y", padx=(8, 0))

        self.zona_archivo = tk.Frame(origen, bg=CAMPO, highlightbackground=LINEA,
                                      highlightthickness=1, cursor="hand2")
        self.zona_archivo.pack(side="left", fill="both", expand=True)
        self.lbl_nombre = tk.Label(self.zona_archivo, text="Elegí un archivo",
                                    bg=CAMPO, fg=TINTA, font=(FUENTE_UI, 11),
                                    anchor="w", justify="left", padx=12, pady=8,
                                    wraplength=ANCHO - 130)
        self.lbl_nombre.pack(fill="x")
        self.lbl_carpeta = tk.Label(self.zona_archivo, text="Excel o .txt  ·  Hacé clic para buscarlo",
                                     bg=CAMPO, fg=SUAVE, font=(FUENTE_UI, 9),
                                     anchor="w", justify="left", padx=12,
                                     wraplength=ANCHO - 130)
        self.lbl_carpeta.pack(fill="x", pady=(0, 10))
        for w in (self.zona_archivo, self.lbl_nombre, self.lbl_carpeta):
            w.bind("<Button-1>", lambda e: self._elegir())

        self.boton = tk.Button(
            marco, text="Generar estado de cuenta", command=self._generar,
            bg="#E4E4E7", fg="#A1A1AA", activebackground="#E4E4E7", activeforeground="#A1A1AA",
            relief="flat", bd=0, font=(FUENTE_UI, 11), pady=10, cursor="hand2",
            state="disabled",
        )
        self.boton.pack(fill="x", pady=(16, 0))
        self.boton.bind("<Enter>", self._hover_on)
        self.boton.bind("<Leave>", self._hover_off)
        tk.Frame(marco, bg=FONDO, width=ANCHO, height=1).pack(fill="x")

        self.estado = tk.Label(marco, text="", bg=FONDO, fg=SUAVE,
                                font=(FUENTE_UI, 9), anchor="w", justify="left", wraplength=ANCHO)
        self.estado.pack(fill="x", pady=(12, 0))

        self.panel = tk.Frame(marco, bg=FONDO)

    def _pegar(self):
        if self.ocupado:
            return
        try:
            texto = self.clipboard_get()
        except tk.TclError:
            texto = ""
        if not texto.strip():
            self._mostrar_estado("El portapapeles está vacío. En SAP, copiá los datos de la consulta "
                                 "y probá de nuevo.", ERROR)
            self._ajustar()
            return
        self.ruta, self.texto = None, texto
        filas = len(texto.strip().splitlines()) - 1
        self.lbl_nombre.config(text="Datos pegados desde SAP")
        self.lbl_carpeta.config(text=f"{filas:,} filas del portapapeles".replace(",", "."))
        self._generar()

    def _poner_archivo(self, ruta):
        self.ruta, self.texto = Path(ruta), None
        self.lbl_nombre.config(text=self.ruta.name)
        self.lbl_carpeta.config(text=str(self.ruta.parent))
        self._habilitar(True)
        self._mostrar_estado("")
        self._ocultar_resultado()

    def _ocultar_resultado(self):
        if self.panel.winfo_ismapped():
            self.panel.pack_forget()
            self.ayuda.pack(fill="x", before=self.lbl_excel)
            self._ajustar()

    def _mostrar_resultado(self, estado, salida):
        for w in self.panel.winfo_children():
            w.destroy()

        tk.Frame(self.panel, bg=LINEA, height=1).pack(fill="x", pady=(14, 12))
        tk.Label(self.panel, text="Saldo por hoja", bg=FONDO, fg=SUAVE,
                 font=(FUENTE_UI, 9), anchor="w").pack(fill="x")

        tabla = tk.Frame(self.panel, bg=FONDO)
        tabla.pack(fill="x", pady=(6, 0))
        tabla.columnconfigure(1, weight=1)

        def fila(i, nombre, saldo, marca, ok, color=None, negrita=False):
            peso = "bold" if negrita else "normal"
            if color:
                tk.Frame(tabla, bg=f"#{color}", width=10, height=10).grid(row=i, column=0, padx=(0, 8))
            tk.Label(tabla, text=nombre, bg=FONDO, fg=TINTA, font=(FUENTE_UI, 9, peso),
                     anchor="w").grid(row=i, column=1, sticky="w", pady=1)
            tk.Label(tabla, text=f"USD {saldo:,.2f}", bg=FONDO, fg=TINTA, font=(FUENTE_UI, 9, peso),
                     anchor="e").grid(row=i, column=2, sticky="e", padx=(12, 0))
            tk.Label(tabla, text=marca, bg=FONDO, fg=OK_COLOR if ok else ERROR,
                     font=(FUENTE_UI, 9), anchor="w").grid(row=i, column=3, sticky="w", padx=(12, 0))

        for i, h in enumerate(estado.hojas):
            fila(i, h.nombre.title(), h.saldo,
                 "OK" if h.ok else f"SAP: {h.esperado:,.2f}", h.ok, color=HOJAS[h.nombre]["tab"])

        n = len(estado.hojas)
        tk.Frame(tabla, bg=LINEA, height=1).grid(row=n, column=0, columnspan=4, sticky="ew", pady=6)
        fila(n + 1, "Total hojas", estado.total_hojas, "", True, negrita=True)
        fila(n + 2, "Saldo en SAP", estado.total_sap,
             "OK" if estado.totales_ok else "Revisar", estado.totales_ok, negrita=True)

        if estado.sin_hoja:
            lineas = [f"{len(estado.sin_hoja)} documento(s) abierto(s) sin hoja asignada "
                      "(revisar Sección en la consulta):"]
            for d in estado.sin_hoja[:5]:
                lineas.append(f"   {d.nro}  ·  {d.fecha:%d/%m/%Y}  ·  USD {d.saldo:,.2f}  ·  {d.seccion or 'sin sección'}")
            if len(estado.sin_hoja) > 5:
                lineas.append(f"   y {len(estado.sin_hoja) - 5} más (ver la consola).")
            tk.Label(self.panel, text="\n".join(lineas), bg=FONDO, fg=ERROR, font=(FUENTE_UI, 9),
                     anchor="w", justify="left", wraplength=ANCHO).pack(fill="x", pady=(10, 0))

        acciones = tk.Frame(self.panel, bg=FONDO)
        acciones.pack(fill="x", pady=(14, 0))
        for texto, accion in (("Abrir Excel", lambda: os.startfile(salida)),
                              ("Mostrar en carpeta", lambda: subprocess.Popen(f'explorer /select,"{salida}"'))):
            tk.Button(acciones, text=texto, command=accion, bg=CAMPO, fg=ACENTO,
                      activebackground=LINEA, activeforeground=ACENTO_ACTIVO, relief="flat", bd=0,
                      font=(FUENTE_UI, 9), padx=12, pady=6, cursor="hand2").pack(side="left", padx=(0, 8))

        self.ayuda.pack_forget()
        self.panel.pack(fill="x")

    def _elegir(self):
        if self.ocupado:
            return
        inicial = str(self.ruta.parent) if self.ruta else str(carpeta_descargas())
        ruta = filedialog.askopenfilename(
            title="Elegir el resultado de la consulta EC ISL",
            filetypes=TIPOS_ARCHIVO,
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
        if self.ocupado or not (self.ruta or self.texto):
            return
        self.ocupado = True
        self._habilitar(False)
        self._mostrar_estado("Generando…")
        self._ocultar_resultado()
        threading.Thread(target=self._correr, daemon=True).start()

    def _correr(self):
        try:
            if self.texto is not None:
                print("Leyendo los datos pegados desde SAP ...")
                consulta = leer_texto(self.texto, "El portapapeles")
            else:
                print(f"Leyendo {self.ruta.name} ...")
                consulta = leer_archivo(self.ruta)
            estado = calcular(consulta)
            salida = guardar(estado)
            informar(estado, salida)
        except ErrorUsuario as e:
            print(f"\nERROR: {e}")
            self._en_ui(lambda m=str(e): self._fallo(m))
        except Exception as e:
            print(f"\nERROR: {e}")
            self._en_ui(lambda m=str(e): self._fallo(f"Error inesperado: {m}"))
        else:
            self._en_ui(lambda e=estado, s=salida: self._listo(e, s))

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

    def _listo(self, estado, salida):
        self.ocupado = False
        self._habilitar(True)
        if estado.todo_ok:
            self._mostrar_estado(f"Listo: {salida.name}\nGuardado en {salida.parent}", OK_COLOR)
        else:
            self._mostrar_estado(f"Generado {salida.name} en {salida.parent}, "
                                 "pero hay controles para revisar.", ERROR)
        self._mostrar_resultado(estado, salida)
        self._ajustar()


def abrir_ventana(archivo_inicial=None):
    preparar_dpi()
    Aplicacion(archivo_inicial).mainloop()


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
