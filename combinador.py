import os
import glob
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class SheetSelectorApp:
    def __init__(self, root, file_name, sheet_list):
        self.root = root
        self.root.title("Selecciona la hoja a importar")
        self.root.geometry("400x180")
        self.root.resizable(False, False)
        
        self.selected_sheet = None
        
        # Estilos y etiquetas
        label_info = ttk.Label(root, text=f"Elige la hoja para el archivo:", font=("Arial", 10))
        label_info.pack(pady=10)
        
        label_file = ttk.Label(root, text=file_name, font=("Arial", 10, "bold"), foreground="blue")
        label_file.pack(pady=2)
        
        # ComboBox Scrolleable para las hojas
        self.combobox = ttk.Combobox(root, values=sheet_list, state="readonly", width=40)
        self.combobox.pack(pady=10)
        self.combobox.current(0) # Selecciona la primera hoja por defecto
        
        # Botón Aceptar
        btn_aceptar = ttk.Button(root, text="Aceptar", command=self.obtener_seleccion)
        btn_aceptar.pack(pady=10)
        
    def obtener_seleccion(self):
        self.selected_sheet = self.combobox.get()
        self.root.destroy()

def seleccionar_hoja_interactivo(file_name, sheet_list):
    """Abre la interfaz gráfica de Tkinter para seleccionar una hoja."""
    root = tk.Tk()
    app = SheetSelectorApp(root, file_name, sheet_list)
    root.mainloop()
    return app.selected_sheet

def combinar_archivos():
    # Ocultar la ventana principal de Tkinter temporalmente para usar los diálogos
    root = tk.Tk()
    root.withdraw()
    
    # 1. SELECCIONAR LA CARPETA
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con tus archivos Excel y CSV")
    if not carpeta:
        return

    # Buscar todos los archivos Excel y CSV
    archivos_excel = glob.glob(os.path.join(carpeta, "*.xl*"))
    archivos_csv = glob.glob(os.path.join(carpeta, "*.csv"))
    all_files = archivos_excel + archivos_csv
    
    if not all_files:
        messagebox.showwarning("Atención", "No se encontraron archivos de Excel o CSV en la carpeta seleccionada.")
        return
        
    dataframes_list = []
    contador = 0

    for ruta_completa in all_files:
        archivo = os.path.basename(ruta_completa)
        _, ext = os.path.splitext(archivo)
        ext = ext.lower()
        
        try:
            # 2. PROCESAR ARCHIVOS DE EXCEL
            if ext.startswith(".xl"):
                excel_file = pd.ExcelFile(ruta_completa)
                hojas = excel_file.sheet_names
                
                # Si tiene más de una hoja, mostramos el selector interactivo
                if len(hojas) > 1:
                    hoja_elegida = seleccionar_hoja_interactivo(archivo, hojas)
                    if not hoja_elegida: # Por si cierran la ventana sin elegir
                        hoja_elegida = hojas[0]
                else:
                    hoja_elegida = hojas[0]
                
                # Leer el Excel (Pandas conserva los tipos de datos automáticamente)
                df = pd.read_excel(ruta_completa, sheet_name=hoja_elegida)
                dataframes_list.append(df)
                contador += 1
                
            # 3. PROCESAR ARCHIVOS CSV
            elif ext == ".csv":
                # Detecta automáticamente separadores comunes de forma nativa (Local:=True equivalente)
                try:
                    df = pd.read_csv(ruta_completa, sep=None, engine='python')
                except Exception:
                    df = pd.read_csv(ruta_completa, sep=';') # Fallback común en sistemas en español
                
                dataframes_list.append(df)
                contador += 1
                
        except Exception as e:
            print(f"Error procesando el archivo {archivo}: {e}")

    # 4. COMBINAR FILA SOBRE FILA Y GUARDAR
    if dataframes_list:
        # concat une todo fila sobre fila conservando un único encabezado automáticamente
        resultado_final = pd.concat(dataframes_list, ignore_index=True)
        
        # Ruta para guardar el archivo consolidado
        ruta_salida = os.path.join(carpeta, "CONSOLIDADO_FINAL.xlsx")
        
        # Guardar a un nuevo archivo Excel
        resultado_final.to_excel(ruta_salida, index=False)
        
        messagebox.showinfo("Hecho", f"¡De puta madre, mano! Se combinaron {contador} archivos.\nGuardado en: {ruta_salida}")
    else:
        messagebox.showerror("Error", "No se pudo procesar ningún archivo.")

if __name__ == "__main__":
    combinar_archivos()

