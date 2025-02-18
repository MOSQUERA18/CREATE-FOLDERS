import os
import pandas as pd
import tkinter as tk
from tkinter import Label, filedialog, messagebox, PhotoImage
import tkinter.font as tkFont

SUBCARPETAS = ["1.Planeación", "2.Ingreso", "3.Ejecución", "4.Certificación"]
DIRECTORIO_BASE_FICHAS = os.path.join(os.path.expanduser("~"), "Documents", "Fichas_Programas")
DIRECTORIO_BASE_INSTRUCTORES = os.path.join(os.path.expanduser("~"), "Documents", "Instructores")
RUTA_DESCARGAS_FICHAS = os.path.join(os.path.expanduser("~"), "Downloads", "plantilla_fichas_programas.xlsx")
RUTA_DESCARGAS_INSTRUCTORES = os.path.join(os.path.expanduser("~"), "Downloads", "plantilla_instructores.xlsx")

def crear_carpeta_si_no_existe(ruta):
    if not os.path.exists(ruta):
        os.makedirs(ruta)

def generar_plantilla_fichas_programas():
    df = pd.DataFrame(columns=["Ficha", "Programa"])
    df.to_excel(RUTA_DESCARGAS_FICHAS, index=False)
    messagebox.showinfo("Plantilla Creada", f"La plantilla se ha guardado en: {RUTA_DESCARGAS_FICHAS}")

def generar_plantilla_instructores():
    df = pd.DataFrame(columns=["Instructor"])
    df.to_excel(RUTA_DESCARGAS_INSTRUCTORES, index=False)
    messagebox.showinfo("Plantilla Creada", f"La plantilla se ha guardado en: {RUTA_DESCARGAS_INSTRUCTORES}")

def procesar_excel_fichas_programas(ruta_excel):
    try:
        crear_carpeta_si_no_existe(DIRECTORIO_BASE_FICHAS)
        df = pd.read_excel(ruta_excel)
        if "Ficha" not in df.columns or "Programa" not in df.columns:
            raise ValueError("El archivo Excel debe contener columnas llamadas 'Ficha' y 'Programa'.")
        for _, fila in df.dropna(subset=["Ficha", "Programa"]).iterrows():
            ficha = str(fila["Ficha"]).strip().upper()
            programa = str(fila["Programa"]).strip().upper()
            ruta_ficha_programa = os.path.join(DIRECTORIO_BASE_FICHAS, f"{ficha}_{programa}")
            crear_carpeta_si_no_existe(ruta_ficha_programa)
            for sub in SUBCARPETAS:
                crear_carpeta_si_no_existe(os.path.join(ruta_ficha_programa, sub))
        messagebox.showinfo("Proceso completado", f"Carpetas creadas en:\n{DIRECTORIO_BASE_FICHAS}")
    except Exception as e:
        messagebox.showerror("Error", f"Se produjo un error: {e}")

def procesar_excel_instructores(ruta_excel):
    try:
        crear_carpeta_si_no_existe(DIRECTORIO_BASE_INSTRUCTORES)
        df = pd.read_excel(ruta_excel)
        if "Instructor" not in df.columns:
            raise ValueError("El archivo Excel debe contener una columna llamada 'Instructor'.")
        for nombre in df["Instructor"].dropna():
            nombre = str(nombre).strip().upper()
            ruta_instructor = os.path.join(DIRECTORIO_BASE_INSTRUCTORES, nombre)
            crear_carpeta_si_no_existe(ruta_instructor)
            ruta_ficha_programa = os.path.join(ruta_instructor, "FICHA+PROGRAMA")
            crear_carpeta_si_no_existe(ruta_ficha_programa)
            for sub in SUBCARPETAS:
                crear_carpeta_si_no_existe(os.path.join(ruta_ficha_programa, sub))
        messagebox.showinfo("Proceso completado", f"Carpetas creadas en:\n{DIRECTORIO_BASE_INSTRUCTORES}")
    except Exception as e:
        messagebox.showerror("Error", f"Se produjo un error: {e}")

def seleccionar_archivo_fichas_programas():
    ruta_excel = filedialog.askopenfilename(title="Seleccionar archivo Excel para Fichas y Programas", filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
    if ruta_excel:
        procesar_excel_fichas_programas(ruta_excel)

def seleccionar_archivo_instructores():
    ruta_excel = filedialog.askopenfilename(title="Seleccionar archivo Excel para Instructores", filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
    if ruta_excel:
        procesar_excel_instructores(ruta_excel)

def mostrar_opciones():
    def opcion_fichas_programas():
        btn_seleccionar.config(command=seleccionar_archivo_fichas_programas)
        btn_plantilla.config(command=generar_plantilla_fichas_programas)
        mensaje.config(text="Selecciona un archivo Excel con Fichas y Programas")

    def opcion_instructores():
        btn_seleccionar.config(command=seleccionar_archivo_instructores)
        btn_plantilla.config(command=generar_plantilla_instructores)
        mensaje.config(text="Selecciona un archivo Excel con Instructores")

    root = tk.Tk()
    root.title("Seleccionar Tipo de Carpeta")
    root.geometry("600x350")
    root.resizable(False, False)

    mensaje = Label(root, text="Elige un tipo de carpeta:", font=("Arial", 12))
    mensaje.pack(pady=20)

    btn_fichas_programas = tk.Button(root, text="Fichas y Programas", command=opcion_fichas_programas)
    btn_fichas_programas.pack(pady=10)

    btn_instructores = tk.Button(root, text="Instructores", command=opcion_instructores)
    btn_instructores.pack(pady=10)

    icon_descarga = PhotoImage(file="C:\\CREATE-FOLDERS\\descarga.png").subsample(10, 12)
    icon_select = PhotoImage(file="C:\\CREATE-FOLDERS\\select.png").subsample(10, 12)

    btn_plantilla = tk.Button(root, text="Descargar Plantilla Excel", command=generar_plantilla_fichas_programas, padx=10, pady=5, image=icon_descarga, compound="left")
    btn_plantilla.pack(pady=20)

    btn_seleccionar = tk.Button(root, text="Seleccionar archivo Excel", command=seleccionar_archivo_fichas_programas, padx=10, pady=5, image=icon_select, compound="left")
    btn_seleccionar.pack(pady=20)

    root.iconbitmap("C:\\CREATE-FOLDERS\\carpeta.ico")
    root.mainloop()

mostrar_opciones()