# -*- coding: utf-8 -*-
"""
Prueba de portabilidad. NO procesa datos de nadie.

Comprueba lo que necesitara el programa real:
  - que el programa abra una ventana,
  - que pueda escribir en carpetas del usuario,
  - que pueda crear y leer un Excel (openpyxl) con estilos y formulas,
  - que pueda abrir el selector de archivos y leer un Excel del equipo
    (solo muestra nombres de hojas y tamano, nunca el contenido).
"""
import datetime
import os
import platform
import sys
import tempfile

VERSION = "1.0"


def carpeta_programa():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def carpeta_documentos():
    casa = os.path.expanduser("~")
    docs = os.path.join(casa, "Documents")
    return docs if os.path.isdir(docs) else casa


def _probar_escritura(carpeta):
    ruta = os.path.join(carpeta, "_prueba_escritura.tmp")
    with open(ruta, "w", encoding="utf-8") as f:
        f.write("ok")
    with open(ruta, encoding="utf-8") as f:
        assert f.read() == "ok"
    os.remove(ruta)
    return carpeta


def _probar_excel():
    import openpyxl
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill

    ruta = os.path.join(tempfile.gettempdir(), "_prueba_portable.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Prueba"
    ws["A1"] = "Hola"
    ws["A1"].font = Font(name="Arial", bold=True)
    ws["A1"].fill = PatternFill("solid", fgColor="1F3864")
    ws["A2"] = "=1+1"
    wb.save(ruta)
    wb2 = load_workbook(ruta, read_only=True)
    primera = next(wb2["Prueba"].iter_rows(values_only=True))
    wb2.close()
    os.remove(ruta)
    assert primera[0] == "Hola"
    return "openpyxl " + openpyxl.__version__


def _correr(nombre, funcion):
    try:
        detalle = funcion()
        return (nombre, True, str(detalle or ""))
    except Exception as e:  # se informa, no se oculta
        return (nombre, False, "%s: %s" % (type(e).__name__, e))


def ejecutar_pruebas():
    programa = carpeta_programa()
    info = [
        ("Fecha y hora", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Sistema", "%s %s" % (platform.system(), platform.version())),
        ("Python", sys.version.split()[0]),
        ("Empaquetado (.exe)", "Si" if getattr(sys, "frozen", False) else "No (se ejecuta como script)"),
        ("Carpeta del programa", programa),
        ("Corre desde unidad de red", "Si" if programa.startswith("\\\\") else "No"),
        ("Carpeta temporal", tempfile.gettempdir()),
    ]
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        info.append(("Carpeta de extraccion (modo un archivo)", sys._MEIPASS))
    filas = [
        _correr("Escribir junto al programa", lambda: _probar_escritura(programa)),
        _correr("Escribir en Documentos del usuario", lambda: _probar_escritura(carpeta_documentos())),
        _correr("Escribir en carpeta temporal", lambda: _probar_escritura(tempfile.gettempdir())),
        _correr("Crear y leer un Excel (openpyxl)", _probar_excel),
    ]
    return info, filas


def probar_excel_del_usuario(ruta):
    from openpyxl import load_workbook
    wb = load_workbook(ruta, read_only=True)
    try:
        partes = []
        for ws in wb.worksheets:
            partes.append("%s (%s filas x %s columnas)" % (ws.title, ws.max_row, ws.max_column))
        return "; ".join(partes)
    finally:
        wb.close()


def informe_texto(info, filas):
    lineas = ["INFORME DE PRUEBA DE PORTABILIDAD v%s" % VERSION, ""]
    lineas += ["%s: %s" % (k, v) for k, v in info]
    lineas += ["", "Resultados:"]
    for nombre, ok, detalle in filas:
        lineas.append("[%s] %s%s" % ("OK" if ok else "FALLO", nombre, (" - " + detalle) if detalle else ""))
    return "\n".join(lineas) + "\n"


def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    info, filas = ejecutar_pruebas()
    filas.insert(0, ("Abrir la ventana del programa", True, "si ve este texto, funciono"))

    root = tk.Tk()
    root.title("Prueba de portabilidad v" + VERSION)
    root.geometry("720x540")

    caja = tk.Text(root, wrap="word", font=("Segoe UI", 10), padx=10, pady=10)
    caja.pack(fill="both", expand=True)
    caja.tag_configure("ok", foreground="#006100", font=("Segoe UI", 10, "bold"))
    caja.tag_configure("mal", foreground="#9C0006", font=("Segoe UI", 10, "bold"))

    def pintar():
        caja.configure(state="normal")
        caja.delete("1.0", "end")
        caja.insert("end", "Prueba de portabilidad\n\n")
        for k, v in info:
            caja.insert("end", "%s: %s\n" % (k, v))
        caja.insert("end", "\nResultados\n")
        for nombre, ok, detalle in filas:
            caja.insert("end", "[%s] " % ("OK" if ok else "FALLO"), "ok" if ok else "mal")
            caja.insert("end", nombre + ((" - " + detalle) if detalle else "") + "\n")
        caja.configure(state="disabled")

    def probar_excel():
        ruta = filedialog.askopenfilename(
            title="Elija un Excel (.xlsx) de prueba",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if not ruta:
            return
        resultado = _correr("Abrir un Excel del equipo (solo estructura)",
                            lambda: probar_excel_del_usuario(ruta))
        filas.append(resultado)
        pintar()

    def guardar():
        ruta = os.path.join(carpeta_documentos(), "informe_prueba_portable.txt")
        try:
            with open(ruta, "w", encoding="utf-8-sig") as f:
                f.write(informe_texto(info, filas))
            messagebox.showinfo("Informe guardado", "Se guardo en:\n" + ruta)
        except Exception as e:
            messagebox.showerror("No se pudo guardar", "%s: %s" % (type(e).__name__, e))

    barra = tk.Frame(root)
    barra.pack(fill="x", padx=10, pady=8)
    tk.Button(barra, text="Probar con un Excel del equipo", command=probar_excel).pack(side="left")
    tk.Button(barra, text="Guardar informe", command=guardar).pack(side="left", padx=8)
    tk.Button(barra, text="Cerrar", command=root.destroy).pack(side="right")

    pintar()
    root.mainloop()


if __name__ == "__main__":
    main()
