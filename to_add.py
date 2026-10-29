import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd


def procesar_carpeta():
    # Abrir el selector de carpetas
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los archivos")

    if not carpeta:
        return  # El usuario canceló la selección

    # Limpiar la tabla previa
    for row in tabla.get_children():
        tabla.delete(row)

    total_filas_carpeta = 0
    archivos_encontrados = False

    # Listar y filtrar archivos .xlsx y .csv
    for archivo in os.listdir(carpeta):
        ruta_completa = os.path.join(carpeta, archivo)

        # Saltar si es una carpeta
        if os.path.isdir(ruta_completa):
            continue

        ext = archivo.split(".")[-1].lower()
        if ext in ["xlsx", "csv"]:
            archivos_encontrados = True
            try:
                # Leer archivo según su extensión
                if ext == "csv":
                    # Se lee el CSV (Pandas asume la primera fila como encabezado por defecto)
                    df = pd.read_csv(ruta_completa)
                else:
                    # Se lee el Excel (.xlsx)
                    df = pd.read_excel(ruta_completa)

                # len(df) da el conteo de filas EXCLUYENDO el encabezado
                filas_archivo = len(df)
                total_filas_carpeta += filas_archivo

                # Insertar resultado en la tabla visual
                tabla.insert(
                    "", tk.END, values=(archivo, ext.upper(), filas_archivo)
                )

            except Exception as e:
                # En caso de error (archivo corrupto, bloqueado, etc.)
                tabla.insert(
                    "", tk.END, values=(archivo, ext.upper(), f"Error: {e}")
                )

    if not archivos_encontrados:
        messagebox.showinfo(
            "Información", "No se encontraron archivos .xlsx o .csv en esta carpeta."
        )
    else:
        # Actualizar la etiqueta del gran total
        lbl_total.config(
            text=f"Sumatoria Total de Filas: {total_filas_carpeta:,}"
        )


# --- Configuración de la Ventana Principal de Tkinter ---
ventana = tk.Tk()
ventana.title("Contador de Filas - Excel & CSV")
ventana.geometry("600x450")
ventana.resizable(False, False)

# Botón para seleccionar carpeta
btn_seleccionar = tk.Button(
    ventana,
    text="📁 Seleccionar Carpeta y Contar",
    font=("Arial", 12, "bold"),
    bg="#2ecc71",
    fg="white",
    command=procesar_carpeta,
    pady=10,
)
btn_seleccionar.pack(pady=15, fill=tk.X, padx=20)

# Contenedor para la tabla de resultados
frame_tabla = tk.Frame(ventana)
frame_tabla.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)

# Definición de la Tabla (Treeview)
columnas = ("archivo", "tipo", "filas")
tabla = ttk.Treeview(frame_tabla, columns=columnas, show="headings")
tabla.heading("archivo", text="Nombre del Archivo")
tabla.heading("tipo", text="Tipo")
tabla.heading("filas", text="Filas (sin encabezado)")

tabla.column("archivo", width=350, anchor=tk.W)
tabla.column("tipo", width=80, anchor=tk.CENTER)
tabla.column("filas", width=120, anchor=tk.E)

# Barra de desplazamiento para la tabla
scrollbar = ttk.Scrollbar(
    frame_tabla, orient=tk.VERTICAL, command=tabla.yview
)
tabla.configure(yscrollcommand=scrollbar.set)
tabla.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Etiqueta para el Gran Total al final
lbl_total = tk.Label(
    ventana,
    text="Sumatoria Total de Filas: 0",
    font=("Arial", 13, "bold"),
    fg="#2c3e50",
    anchor=tk.E,
)
lbl_total.pack(pady=15, padx=20, fill=tk.X)

# Iniciar la aplicación
ventana.mainloop()
