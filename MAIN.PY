import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Definir nombres de subcarpetas
SUBCARPETAS = ["1.Planeación", "2.Ingreso", "3.Ejecución", "4.Certificación"]

# Ruta principal dentro de Documentos
DIRECTORIO_BASE = os.path.join(os.path.expanduser("~"), "Documents", "Instructores")

def crear_carpeta_si_no_existe(ruta):
    """Crea una carpeta si no existe"""
    if not os.path.exists(ruta):
        os.makedirs(ruta)

def procesar_excel(ruta_excel):
    """Lee el Excel y crea carpetas para cada instructor"""
    try:
        # Crear la carpeta principal en Documentos si no existe
        crear_carpeta_si_no_existe(DIRECTORIO_BASE)

        # Leer el archivo Excel
        df = pd.read_excel(ruta_excel)

        # Verificar si la columna con los nombres de instructores existe
        if "Instructor" not in df.columns:
            raise ValueError("El archivo Excel debe contener una columna llamada 'Instructor'.")

        # Recorrer los nombres de instructores y crear carpetas
        for nombre in df["Instructor"].dropna():
            nombre = str(nombre).strip().upper()  # Convertir a string, eliminar espacios y poner en mayúsculas
            ruta_instructor = os.path.join(DIRECTORIO_BASE, nombre)

            # Crear carpeta del instructor
            crear_carpeta_si_no_existe(ruta_instructor)

            # Crear la carpeta "FICHA+PROGRAMA" dentro del instructor
            ruta_ficha_programa = os.path.join(ruta_instructor, "FICHA+PROGRAMA")
            crear_carpeta_si_no_existe(ruta_ficha_programa)

            # Crear subcarpetas dentro de "FICHA+PROGRAMA"
            for sub in SUBCARPETAS:
                crear_carpeta_si_no_existe(os.path.join(ruta_ficha_programa, sub))

        # Mostrar mensaje de éxito
        messagebox.showinfo("Proceso completado", f"Carpetas creadas en:\n{DIRECTORIO_BASE}")

    except Exception as e:
        messagebox.showerror("Error", f"Se produjo un error: {e}")

def seleccionar_archivo():
    """Abre el explorador de archivos y permite seleccionar el archivo Excel"""
    ruta_excel = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos de Excel", "*.xlsx;*.xls")]
    )
    
    if ruta_excel:  # Si el usuario selecciona un archivo
        procesar_excel(ruta_excel)

# Crear la ventana principal
root = tk.Tk()
root.title("Gestor de Carpetas para Instructores")
root.geometry("400x200")
root.resizable(False, False)

# Crear botón para seleccionar archivo
btn_seleccionar = tk.Button(root, text="Seleccionar archivo Excel", command=seleccionar_archivo, padx=10, pady=5)
btn_seleccionar.pack(pady=50)

# Iniciar la aplicación
root.mainloop()
