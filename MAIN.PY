import os
import pandas as pd
import tkinter as tk
from tkinter import Label, filedialog, messagebox, PhotoImage
import tkinter.font as tkFont

SUBCARPETAS = ["1.Planeación", "2.Ingreso", "3.Ejecución", "4.Certificación"]
DIRECTORIO_BASE = os.path.join(os.path.expanduser("~"), "Documents", "Instructores")
RUTA_DESCARGAS = os.path.join(os.path.expanduser("~"), "Downloads", "plantilla_instructores.xlsx")

def crear_carpeta_si_no_existe(ruta):
    if not os.path.exists(ruta):
        os.makedirs(ruta)

def generar_plantilla_excel():
    df = pd.DataFrame(columns=["Instructor"])
    df.to_excel(RUTA_DESCARGAS, index=False)
    messagebox.showinfo("Plantilla Creada", f"La plantilla se ha guardado en: {RUTA_DESCARGAS}")

def procesar_excel(ruta_excel):
    try:
        crear_carpeta_si_no_existe(DIRECTORIO_BASE)
        df = pd.read_excel(ruta_excel)
        if "Instructor" not in df.columns:
            raise ValueError("El archivo Excel debe contener una columna llamada 'Instructor'.")
        for nombre in df["Instructor"].dropna():
            nombre = str(nombre).strip().upper()
            ruta_instructor = os.path.join(DIRECTORIO_BASE, nombre)
            crear_carpeta_si_no_existe(ruta_instructor)
            ruta_ficha_programa = os.path.join(ruta_instructor, "FICHA+PROGRAMA")
            crear_carpeta_si_no_existe(ruta_ficha_programa)
            for sub in SUBCARPETAS:
                crear_carpeta_si_no_existe(os.path.join(ruta_ficha_programa, sub))
        messagebox.showinfo("Proceso completado", f"Carpetas creadas en:\n{DIRECTORIO_BASE}")
    except Exception as e:
        messagebox.showerror("Error", f"Se produjo un error: {e}")

def seleccionar_archivo():
    ruta_excel = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
    if ruta_excel:
        procesar_excel(ruta_excel)

root = tk.Tk()
root.title("Gestor de Carpetas para Instructores")
root.geometry("400x250")
root.resizable(False, False)

# Cargar imágenes con rutas absolutas corregidas y dobles barras
icon_descarga = PhotoImage(file="C:\\CREATE-FOLDERS\\descarga.png").subsample(10, 12)
icon_select = PhotoImage(file="C:\\CREATE-FOLDERS\\select.png").subsample(10, 12)

btn_plantilla = tk.Button(root, text="Descargar Plantilla Excel", command=generar_plantilla_excel, padx=10, pady=5, image=icon_descarga, compound="left")
btn_plantilla.pack(pady=20)

btn_seleccionar = tk.Button(root, text="Seleccionar archivo Excel", command=seleccionar_archivo, padx=10, pady=5, image=icon_select, compound="left")
btn_seleccionar.pack(pady=20)

font_footer = tkFont.Font(family="Arial", size=10, weight="bold")
footer = Label(root, text="Desarrollado por Marlon Mosquera ADSO 2671143", font=font_footer, bg='#f4f4f9', fg='#555', cursor="hand2")
footer.pack(side="bottom", pady=5)

root.iconbitmap("C:\\CREATE-FOLDERS\\carpeta.ico")
root.mainloop()
