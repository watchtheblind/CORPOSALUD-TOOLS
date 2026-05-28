import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import pandas as pd
from datetime import datetime
import threading
import platform
import subprocess

class ExcelCsvMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🧩 Combinador de Archivos Excel/CSV")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # Variables
        self.selected_files = []
        self.output_file = tk.StringVar()
        
        # Configurar estilo
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_styles()
        
        self.create_widgets()
        
    def configure_styles(self):
        self.style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'))
        self.style.configure('Custom.TButton', font=('Segoe UI', 10), padding=10)
        self.style.configure('Danger.TButton', font=('Segoe UI', 10), padding=10)
        self.style.configure('Success.TButton', font=('Segoe UI', 10), padding=10)
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        title_label = ttk.Label(main_frame, text="📊 Combinador de Datos (Excel/CSV)", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 30))
        
        # Frame para selección de archivos
        files_frame = ttk.LabelFrame(main_frame, text="📁 Archivos seleccionados", padding="15")
        files_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        files_frame.columnconfigure(0, weight=1)
        
        self.files_listbox = tk.Listbox(files_frame, height=8, font=('Consolas', 9))
        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        files_scrollbar = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, command=self.files_listbox.yview)
        files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.files_listbox.configure(yscrollcommand=files_scrollbar.set)
        
        btn_frame = ttk.Frame(files_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(btn_frame, text="➕ Agregar Archivos", style='Custom.TButton', command=self.add_files).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="➖ Remover Seleccionado", style='Danger.TButton', command=self.remove_file).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="🗑️ Limpiar Todo", style='Danger.TButton', command=self.clear_files).pack(side=tk.LEFT, padx=(10, 0))
        
        # Frame para archivo de salida
        output_frame = ttk.LabelFrame(main_frame, text="💾 Archivo de Salida", padding="15")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        output_frame.columnconfigure(0, weight=1)
        
        output_entry_frame = ttk.Frame(output_frame)
        output_entry_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        output_entry_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(output_entry_frame, textvariable=self.output_file, font=('Segoe UI', 10)).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(output_entry_frame, text="📂 Guardar como...", style='Custom.TButton', command=self.browse_output).grid(row=0, column=1)
        
        # Botón principal
        self.merge_btn = ttk.Button(main_frame, text="🔗 COMBINAR ARCHIVOS", style='Success.TButton', command=self.start_merge_thread, state=tk.DISABLED)
        self.merge_btn.grid(row=3, column=0, columnspan=3, pady=20)
        
        # Log
        log_frame = ttk.LabelFrame(main_frame, text="📋 Registro", padding="15")
        log_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=('Consolas', 9), state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)

    # --- Funciones de Interfaz y Logs ---

    def log_message(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def safe_log(self, message):
        """Usa esto para escribir en el log desde el hilo secundario de forma segura"""
        self.root.after(0, self.log_message, message)

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos Excel/CSV", 
            filetypes=[("Archivos de datos", "*.xlsx *.xls *.csv"), ("Todos los archivos", "*.*")]
        )
        for f in files:
            if f not in self.selected_files:
                self.selected_files.append(f)
                self.files_listbox.insert(tk.END, f"📄 {os.path.basename(f)}")
        self.update_merge_button()

    def remove_file(self):
        selection = self.files_listbox.curselection()
        if selection:
            self.selected_files.pop(selection[0])
            self.files_listbox.delete(selection[0])
            self.update_merge_button()

    def clear_files(self):
        self.selected_files.clear()
        self.files_listbox.delete(0, tk.END)
        self.update_merge_button()

    def browse_output(self):
        fn = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Archivo Excel", "*.xlsx"), ("Archivo CSV", "*.csv")]
        )
        if fn: self.output_file.set(fn)

    def normalize_dates(self, df):
        """Detecta columnas de fecha y las formatea a DD/MM/AAAA, ignorando números normales"""
        for col in df.columns:
            # 1. Si la columna ya es numérica pura (int o float), la ignoramos por completo
            if pd.api.types.is_numeric_dtype(df[col]):
                continue
                
            # 2. Si Pandas ya la detectó como fecha nativa (muy común al leer desde Excel)
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%d/%m/%Y')
                continue
                
            # 3. Si es texto, verificamos que realmente parezca una fecha
            if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
                # Comprobamos si la columna tiene caracteres típicos de fechas (/ o - o :)
                # Esto evita convertir cédulas como "27000123" o valores que sean solo texto
                has_date_chars = df[col].astype(str).str.contains(r'[/:-]', na=False).sum()
                valid_rows = df[col].notna().sum()
                
                # Si una buena cantidad de celdas tienen esos caracteres, intentamos convertir
                if valid_rows > 0 and has_date_chars >= (valid_rows * 0.5):
                    try:
                        converted = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
                        # Si la conversión fue exitosa para la mayoría, aplicamos el cambio
                        if converted.notna().sum() >= (valid_rows * 0.5):
                            df[col] = converted.dt.strftime('%d/%m/%Y')
                    except Exception:
                        continue
        return df

    def update_merge_button(self):
        self.merge_btn.config(state=tk.NORMAL if self.selected_files else tk.DISABLED)

    def open_file(self, filepath):
        """Abre el archivo generado usando el programa predeterminado del SO"""
        try:
            if platform.system() == 'Windows':
                os.startfile(filepath)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(('open', filepath))
            else:  # Linux
                subprocess.call(('xdg-open', filepath))
            self.log_message("📂 Archivo abierto exitosamente.")
        except Exception as e:
            self.log_message(f"⚠️ No se pudo abrir el archivo automáticamente: {e}")

    # --- Lógica de Extracción (Pandas) ---

    def extract_valid_data(self, df):
        target_words = {"CÉDULA", "CEDULA", "CI", "C.I", "C.I."}
        for i in range(min(30, len(df))):
            row_values = df.iloc[i].astype(str).str.upper().str.strip().tolist()
            if any(word in row_values for word in target_words):
                df.columns = df.iloc[i]
                df = df.iloc[i+1:].copy()
                df = df.dropna(how='all')
                
                # APLICAR NORMALIZACIÓN DE FECHAS AQUÍ
                df = self.normalize_dates(df)
                return df
        return None

    def process_file(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == '.csv':
            try:
                df = pd.read_csv(path, header=None, encoding='utf-8')
                return self.extract_valid_data(df)
            except Exception:
                df = pd.read_csv(path, header=None, encoding='latin1', sep=';')
                return self.extract_valid_data(df)
        elif ext in ['.xlsx', '.xls']:
            xls = pd.ExcelFile(path)
            if "DATOS" in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name="DATOS", header=None)
                result = self.extract_valid_data(df)
                if result is not None:
                    return result
            for sheet in xls.sheet_names:
                if sheet == "DATOS": 
                    continue
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                result = self.extract_valid_data(df)
                if result is not None:
                    return result
        return None

    # --- Gestión de Hilos y Ventana de Carga ---

    def start_merge_thread(self):
        output_path = self.output_file.get()
        if not output_path:
            self.browse_output()
            output_path = self.output_file.get()
            if not output_path: return

        # 1. Crear y mostrar ventana de "Cargando"
        self.loading_window = tk.Toplevel(self.root)
        self.loading_window.title("Procesando...")
        self.loading_window.geometry("300x120")
        self.loading_window.resizable(False, False)
        
        # Hacer que se comporte como un modal (bloquea la ventana principal)
        self.loading_window.transient(self.root)
        self.loading_window.grab_set() 
        
        # Centrar la ventanita respecto a la principal
        x = self.root.winfo_rootx() + 200
        y = self.root.winfo_rooty() + 200
        self.loading_window.geometry(f"+{x}+{y}")

        ttk.Label(self.loading_window, text="⏳ Procesando archivos...\nPor favor, no cierres el programa.", 
                  font=('Segoe UI', 10), justify="center").pack(pady=(20, 10))
        
        progress = ttk.Progressbar(self.loading_window, mode='indeterminate')
        progress.pack(fill=tk.X, padx=20)
        progress.start(15) # Velocidad de la animación

        self.merge_btn.config(state=tk.DISABLED)

        # 2. Iniciar el hilo en segundo plano (daemon=True hace que muera si cierras la app)
        threading.Thread(target=self._merge_files_worker, args=(output_path,), daemon=True).start()

    def _merge_files_worker(self, output_path):
        """Esta función se ejecuta en el fondo. No debe llamar a MessageBox directamente."""
        try:
            self.safe_log("🚀 Iniciando extracción y consolidación de datos...")
            all_dataframes = []
            
            for i, path in enumerate(self.selected_files, 1):
                filename = os.path.basename(path)
                self.safe_log(f"[{i}/{len(self.selected_files)}] Procesando: {filename}")
                
                df_extracted = self.process_file(path)
                
                if df_extracted is not None and not df_extracted.empty:
                    all_dataframes.append(df_extracted)
                    self.safe_log("  └─ Datos extraídos correctamente.")
                else:
                    self.safe_log("  └─ ⚠️ No se encontró cédula o no hay datos.")

            if not all_dataframes:
                # Volvemos al hilo principal para mostrar errores o finalizar
                self.root.after(0, self._merge_finish_empty)
                return

            final_df = pd.concat(all_dataframes, ignore_index=True)

            if output_path.endswith('.csv'):
                final_df.to_csv(output_path, index=False, encoding='utf-8-sig')
            else:
                final_df.to_excel(output_path, index=False)

            # Volvemos al hilo principal para avisar que todo salió bien
            self.root.after(0, self._merge_success, output_path, len(final_df))
            
        except Exception as e:
            # Volvemos al hilo principal para mostrar el error
            self.root.after(0, self._merge_error, str(e))

    # --- Callbacks para volver al Hilo Principal (GUI Segura) ---

    def _merge_finish_empty(self):
        self.loading_window.destroy()
        self.merge_btn.config(state=tk.NORMAL)
        messagebox.showwarning("Advertencia", "No se encontraron datos válidos en los archivos seleccionados.")
        self.log_message("❌ Proceso abortado: Sin datos.")

    def _merge_success(self, output_path, total_rows):
        self.loading_window.destroy()
        self.merge_btn.config(state=tk.NORMAL)
        messagebox.showinfo("Éxito", f"Datos consolidados correctamente.\nTotal de filas: {total_rows}")
        self.log_message(f"✅ ¡Proceso finalizado! Se guardaron {total_rows} registros.")
        self.open_file(output_path) # Llama a la función que abre el archivo automáticamente

    def _merge_error(self, error_msg):
        self.loading_window.destroy()
        self.merge_btn.config(state=tk.NORMAL)
        messagebox.showerror("Error", f"Ocurrió un error al procesar:\n{error_msg}")
        self.log_message(f"❌ Error crítico: {error_msg}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCsvMergerApp(root)
    root.mainloop()