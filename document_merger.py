import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import pandas as pd
from datetime import datetime
import threading
import platform
import subprocess
import re

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

class ExcelCsvMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🧩 Combinador de Archivos Excel/CSV/PDF")
        
        # 1. Centrar la ventana al iniciar
        window_width = 650
        window_height = 650
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        
        # 2. Configurar pesos para que los elementos se expandan
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.resizable(False, False)
        
        # Variables
        self.selected_files = []
        self.output_file = tk.StringVar()
        
        # Configurar estilo
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_styles()
        
        self.create_widgets()
        
    def configure_styles(self):
        self.style.configure('Title.TLabel', font=('Segoe UI', 15, 'bold'))
        self.style.configure('Custom.TButton', font=('Segoe UI', 9), padding=8)
        self.style.configure('Danger.TButton', font=('Segoe UI', 9), padding=8)
        self.style.configure('Success.TButton', font=('Segoe UI', 10, 'bold'), padding=10)
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Título
        title_lbl = ttk.Label(main_frame, text="🧩 Consolidador Masivo de Nóminas", style='Title.TLabel')
        title_lbl.grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky=tk.W)
        
        # Sección Selección de Archivos
        files_frame = ttk.LabelFrame(main_frame, text=" Archivos Fuente seleccionados ", padding="10")
        files_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        main_frame.rowconfigure(1, weight=3)
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(0, weight=1)
        
        self.files_listbox = tk.Listbox(files_frame, height=8, selectmode=tk.EXTENDED, font=('Consolas', 10))
        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, command=self.files_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.files_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Botones de archivos
        btn_frame = ttk.Frame(files_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0), sticky=tk.W)
        
        add_btn = ttk.Button(btn_frame, text="➕ Agregar Archivos", command=self.add_files, style='Custom.TButton')
        add_btn.grid(row=0, column=0, padx=(0, 5))
        
        remove_btn = ttk.Button(btn_frame, text="❌ Quitar Seleccionados", command=self.remove_files, style='Danger.TButton')
        remove_btn.grid(row=0, column=1, padx=5)
        
        clear_btn = ttk.Button(btn_frame, text="🧹 Limpiar Todo", command=self.clear_all, style='Danger.TButton')
        clear_btn.grid(row=0, column=2, padx=5)
        
        # Sección Destino
        dest_frame = ttk.LabelFrame(main_frame, text=" Archivo de Salida (Consolidado) ", padding="10")
        dest_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        dest_frame.columnconfigure(1, weight=1)
        
        dest_lbl = ttk.Label(dest_frame, text="Ruta:")
        dest_lbl.grid(row=0, column=0, padx=(0, 5))
        
        self.dest_entry = ttk.Entry(dest_frame, textvariable=self.output_file, state='readonly')
        self.dest_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        search_dest_btn = ttk.Button(dest_frame, text="🔍 Guardar como...", command=self.select_output_path)
        search_dest_btn.grid(row=0, column=2, padx=(5, 0))
        
        # Consola de Registros (Logs)
        log_frame = ttk.LabelFrame(main_frame, text=" Historial de Ejecución / Consola ", padding="10")
        log_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        main_frame.rowconfigure(3, weight=2)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=6, font=('Consolas', 9), bg='#1e1e1e', fg='#d4d4d4')
        self.log_area.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_area.config(state=tk.DISABLED)
        
        # Botón Central Procesar
        self.merge_btn = ttk.Button(main_frame, text="⚙️ INICIAR CONSOLIDACIÓN COMPLETA", command=self.start_merge, style='Success.TButton')
        self.merge_btn.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        
        if pdfplumber is None:
            self.log_message("⚠️ ADVERTENCIA: La librería 'pdfplumber' no está instalada. Ejecuta: pip install pdfplumber")
        else:
            self.log_message("🚀 Sistema listo. Por favor, selecciona los archivos PDF/CSV/Excel a unificar.")

    def log_message(self, message):
        self.log_area.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)

    def add_files(self):
        file_types = [
            ("Formatos Admitidos", "*.pdf;*.csv;*.xlsx;*.xls"),
            ("Archivos PDF (*.pdf)", "*.pdf"),
            ("Archivos CSV (*.csv)", "*.csv"),
            ("Archivos Excel (*.xlsx;*.xls)", "*.xlsx;*.xls")
        ]
        files = filedialog.askopenfilenames(title="Seleccionar archivos de nóminas", filetypes=file_types)
        if files:
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    self.files_listbox.insert(tk.END, os.path.basename(file))
            self.log_message(f"📥 Se agregaron {len(files)} archivo(s) a la lista.")
            self.auto_suggest_output()

    def remove_files(self):
        selected_indices = list(self.files_listbox.curselection())
        if not selected_indices:
            messagebox.showwarning("Atención", "Por favor, selecciona los archivos que deseas remover de la lista.")
            return
            
        for index in reversed(selected_indices):
            removed_file = self.selected_files.pop(index)
            self.files_listbox.delete(index)
            self.log_message(f"🗑️ Se quitó de la lista: {os.path.basename(removed_file)}")
        
        self.auto_suggest_output()

    def clear_all(self):
        if self.selected_files:
            self.selected_files.clear()
            self.files_listbox.delete(0, tk.END)
            self.output_file.set("")
            self.log_message("🧹 Lista desocupada por completo.")

    def select_output_path(self):
        file_types = [("Libro de Excel (*.xlsx)", "*.xlsx"), ("Archivo CSV (*.csv)", "*.csv")]
        default_name = f"Nomina_Consolidada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = filedialog.asksaveasfilename(title="Guardar archivo consolidado", defaultextension=".xlsx", filetypes=file_types, initialfile=default_name)
        if path:
            self.output_file.set(path)
            self.log_message(f"📍 Destino fijado: {os.path.basename(path)}")

    def auto_suggest_output(self):
        if self.selected_files:
            first_file_dir = os.path.dirname(self.selected_files[0])
            suggested_name = f"Nomina_Consolidada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            self.output_file.set(os.path.join(first_file_dir, suggested_name))
        else:
            self.output_file.set("")

    def start_merge(self):
        if not self.selected_files:
            messagebox.showerror("Error", "Debes añadir por lo menos un archivo a la lista.")
            return
        if not self.output_file.get():
            messagebox.showerror("Error", "Debes definir un destino para el archivo consolidado.")
            return
            
        self.merge_btn.config(state=tk.DISABLED)
        
        self.loading_window = tk.Toplevel(self.root)
        self.loading_window.title("Por favor espere...")
        self.loading_window.geometry("300x120")
        self.loading_window.resizable(False, False)
        self.loading_window.transient(self.root)
        self.loading_window.grab_set()
        
        lbl = ttk.Label(self.loading_window, text="⚙️ Consolidando información...\nPor favor, no cierres la aplicación.", justify=tk.CENTER, font=('Segoe UI', 10))
        lbl.pack(pady=20)
        
        progress = ttk.Progressbar(self.loading_window, mode='indeterminate')
        progress.pack(fill=tk.X, padx=30)
        progress.start(10)
        
        threading.Thread(target=self._execute_merge, daemon=True).start()

    def _execute_merge(self):
        output_path = self.output_file.get()
        dfs_validos = []
        
        try:
            for file_path in self.selected_files:
                ext = os.path.splitext(file_path)[1].lower()
                try:
                    # 1. PARSEO DE PDF ROBUSTO (Diseñado específicamente para el formato real de tus PDFs)
                    if ext == '.pdf':
                        if pdfplumber is None:
                            self.log_message(f"❌ Error: pdfplumber no instalado. Saltando {os.path.basename(file_path)}")
                            continue
                            
                        data = []
                        with pdfplumber.open(file_path) as pdf:
                            for page in pdf.pages:
                                text = page.extract_text()
                                if not text:
                                    continue
                                
                                # Quitamos comillas molestas que introduce el formateador del PDF
                                text_clean = text.replace('"', '')
                                
                                # Patrón maestro: Encuentra Cédula (5-9 dígitos) seguido por el texto hasta la cuenta (20 dígitos) y el total monetario
                                pattern = re.compile(
                                    r'(\b\d{5,9}\b)\s*,\s*(.*?)\s*(\b\d{20}\b)\s*([AC]?)\s*,\s*([\d.,]+)', 
                                    re.DOTALL
                                )
                                
                                matches = pattern.findall(text_clean)
                                for match in matches:
                                    cedula, nombre_sucio, cuenta, tipo, total = match
                                    
                                    # Limpiar saltos de línea internos en el nombre
                                    nombre = re.sub(r'\s+', ' ', nombre_sucio).strip()
                                    
                                    # Limpiar posibles códigos o residuos numéricos previos al nombre real
                                    nombre = re.sub(r'^\d+\s*', '', nombre).strip()
                                    
                                    data.append([
                                        "",  # Nro vacío (se auto-generará si lo deseas)
                                        cedula.strip(),
                                        nombre.upper(),
                                        tipo.strip() if tipo else "A",
                                        cuenta.strip(),
                                        total.strip()
                                    ])
                                    
                        if data:
                            df = pd.DataFrame(data, columns=['NRO', 'CEDULA', 'APELLIDOS Y NOMBRES', 'TIPO', 'CUENTA', 'TOTAL'])
                            dfs_validos.append(df)
                            self.log_message(f"✅ PDF procesado con éxito: {os.path.basename(file_path)}")
                        else:
                            self.log_message(f"⚠️ No se pudo extraer registros válidos del PDF: {os.path.basename(file_path)}")

                    # 2. PROCESAMIENTO CSV
                    elif ext == '.csv':
                        try:
                            df = pd.read_csv(file_path, sep=None, engine='python', encoding='utf-8-sig')
                        except:
                            df = pd.read_csv(file_path, sep=None, engine='python', encoding='latin1')
                        
                        df.columns = df.columns.astype(str).str.strip().str.upper()
                        if 'CEDULA' in df.columns or 'CÉDULA' in df.columns:
                            dfs_validos.append(df)
                            self.log_message(f"✅ CSV procesado: {os.path.basename(file_path)}")
                    
                    # 3. PROCESAMIENTO EXCEL
                    elif ext in ['.xlsx', '.xls']:
                        xl = pd.ExcelFile(file_path)
                        hoja_encontrada = False
                        for sheet_name in xl.sheet_names:
                            df_header = pd.read_excel(file_path, sheet_name=sheet_name, nrows=1)
                            df_header.columns = df_header.columns.astype(str).str.strip().str.upper()
                            
                            if 'CEDULA' in df_header.columns or 'CÉDULA' in df_header.columns:
                                df = pd.read_excel(file_path, sheet_name=sheet_name)
                                df.columns = df.columns.astype(str).str.strip().str.upper()
                                dfs_validos.append(df)
                                hoja_encontrada = True
                                self.log_message(f"✅ Excel procesado: {os.path.basename(file_path)} (Hoja: '{sheet_name}')")
                                break 
                        if not hoja_encontrada:
                            self.log_message(f"ℹ️ Excel saltado (Falta 'CEDULA'): {os.path.basename(file_path)}")

                except Exception as inner_e:
                    self.log_message(f"❌ Error en {os.path.basename(file_path)}: {str(inner_e)}")

            # 4. COMPILACIÓN FINAL Y GENERACIÓN DE EXCEL
            if dfs_validos:
                final_df = pd.concat(dfs_validos, ignore_index=True)
                
                # Normalizar nombres de columnas principales
                columnas_mapeo = {
                    'CÉDULA': 'CEDULA',
                    'NOMBRES': 'APELLIDOS Y NOMBRES',
                    'NOMBRE': 'APELLIDOS Y NOMBRES'
                }
                final_df.rename(columns=columnas_mapeo, inplace=True)
                
                # Asegurar secuencialidad limpia de la columna NRO
                final_df['NRO'] = range(1, len(final_df) + 1)
                
                if output_path.endswith('.csv'):
                    final_df.to_csv(output_path, index=False, sep=';', encoding='utf-8-sig')
                else:
                    final_df.to_excel(output_path, index=False)
                    
                self.root.after(0, self._merge_success, output_path, len(final_df))
            else:
                self.root.after(0, self._merge_finish_empty)
                
        except Exception as e:
            self.root.after(0, self._merge_error, str(e))

    def _merge_finish_empty(self):
        self.loading_window.destroy()
        self.merge_btn.config(state=tk.NORMAL)
        messagebox.showwarning("Advertencia", "No se encontraron datos de nómina válidos en los archivos seleccionados.")
        self.log_message("❌ Proceso abortado: Sin datos.")

    def _merge_success(self, output_path, total_rows):
        self.loading_window.destroy()
        self.merge_btn.config(state=tk.NORMAL)
        messagebox.showinfo("Éxito", f"Datos consolidados correctamente.\nTotal de filas: {total_rows}")
        self.log_message(f"✅ ¡Proceso finalizado! Se guardaron {total_rows} registros perfectamente organizados.")
        self.open_file(output_path)

    def _merge_error(self, error_msg):
        self.loading_window.destroy()
        self.merge_btn.config(state=tk.NORMAL)
        messagebox.showerror("Error", f"Ocurrió un error al procesar:\n{error_msg}")
        self.log_message(f"❌ Error en la consolidación: {error_msg}")

    def open_file(self, filepath):
        if platform.system() == 'Darwin':
            subprocess.call(('open', filepath))
        elif platform.system() == 'Windows':
            os.startfile(filepath)
        else:
            subprocess.call(('xdg-open', filepath))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCsvMergerApp(root)
    root.mainloop()
