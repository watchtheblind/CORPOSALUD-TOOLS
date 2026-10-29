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
        self.root.geometry("780x680")
        self.root.resizable(True, True)
        
        # Variables
        self.selected_files = []
        self.output_file = tk.StringVar()

        # Hojas de Excel: ruta -> [hojas disponibles] y ruta -> hoja elegida
        self.available_sheets = {}
        self.sheet_choice = {}
        self.sheets_loading = set()
        self._sheets_lock = threading.Lock()
        self._sheet_dialog_open = False
        self.row_counter_window = None
        
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
        
        self.files_listbox = tk.Listbox(files_frame, height=9, font=('Consolas', 9))
        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.files_listbox.bind(
            "<Double-Button-1>", lambda _event: self.open_sheet_selector()
        )
        
        files_scrollbar = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, command=self.files_listbox.yview)
        files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.files_listbox.configure(yscrollcommand=files_scrollbar.set)
        
        btn_frame = ttk.Frame(files_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0), sticky=(tk.W, tk.E))
        
        ttk.Button(btn_frame, text="➕ Agregar Archivos", style='Custom.TButton', command=self.add_files).grid(row=0, column=0, padx=(0, 8))
        self.sheets_btn = ttk.Button(btn_frame, text="🗂️ Elegir Hojas", style='Custom.TButton', command=self.open_sheet_selector, state=tk.DISABLED)
        self.sheets_btn.grid(row=0, column=1, padx=(0, 8))
        ttk.Button(btn_frame, text="➖ Remover Seleccionado", style='Danger.TButton', command=self.remove_file).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(btn_frame, text="🗑️ Limpiar Todo", style='Danger.TButton', command=self.clear_files).grid(row=0, column=3)
        
        # Frame para archivo de salida
        output_frame = ttk.LabelFrame(main_frame, text="💾 Archivo de Salida", padding="15")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        output_frame.columnconfigure(0, weight=1)
        
        output_entry_frame = ttk.Frame(output_frame)
        output_entry_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        output_entry_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(output_entry_frame, textvariable=self.output_file, font=('Segoe UI', 10)).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(output_entry_frame, text="📂 Guardar como...", style='Custom.TButton', command=self.browse_output).grid(row=0, column=1)
        
        # Botones de acción
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=3, column=0, columnspan=3, pady=20)

        self.merge_btn = ttk.Button(action_frame, text="🔗 COMBINAR ARCHIVOS", style='Success.TButton', command=self.start_merge_thread, state=tk.DISABLED)
        self.merge_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Herramienta opcional: no participa del flujo de combinación.
        ttk.Button(
            action_frame,
            text="🧮 Contador de filas (opcional)",
            style='Custom.TButton',
            command=self.open_row_counter,
        ).pack(side=tk.LEFT)
        
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

    @staticmethod
    def _is_excel(path):
        return os.path.splitext(path)[1].lower() in ('.xlsx', '.xls', '.xlsm')

    @staticmethod
    def _acortar_texto(texto, limite=80, desde_inicio=False):
        """Recorta textos largos para que no desborden la ventana."""
        if len(texto) <= limite:
            return texto
        if desde_inicio:
            return "…" + texto[-(limite - 1):]
        return texto[:limite - 1] + "…"

    def _refresh_files_listbox(self):
        """Redibuja la lista mostrando la hoja elegida para cada archivo."""
        self.files_listbox.delete(0, tk.END)
        for path in self.selected_files:
            name = os.path.basename(path)
            if not self._is_excel(path):
                detalle = "CSV (sin hojas)"
            else:
                with self._sheets_lock:
                    cargando = path in self.sheets_loading
                    hojas = list(self.available_sheets.get(path, []))
                    elegida = self.sheet_choice.get(path)

                if cargando:
                    detalle = "leyendo hojas..."
                elif not hojas:
                    detalle = "⚠️ no se pudieron leer las hojas"
                elif elegida:
                    detalle = f"Hoja: {elegida}"
                    if len(hojas) > 1:
                        detalle += f"  ({len(hojas)} disponibles)"
                else:
                    detalle = "⚠️ hoja sin elegir"

            self.files_listbox.insert(tk.END, f"📄 {name}   →   {detalle}")

    def _scan_sheets_async(self, paths):
        """Lee los nombres de las hojas en segundo plano para no congelar la GUI."""
        excel_paths = [p for p in paths if self._is_excel(p)]
        if not excel_paths:
            return

        with self._sheets_lock:
            self.sheets_loading.update(excel_paths)
            for path in excel_paths:
                self.available_sheets.pop(path, None)
                self.sheet_choice.pop(path, None)

        self._refresh_files_listbox()
        self.update_merge_button()
        threading.Thread(
            target=self._scan_sheets_worker, args=(excel_paths,), daemon=True
        ).start()

    def _scan_sheets_worker(self, paths):
        leidos = 0
        for path in paths:
            try:
                with pd.ExcelFile(path) as libro:
                    hojas = list(libro.sheet_names)
            except Exception as exc:
                hojas = []
                self.safe_log(
                    f"⚠️ No se pudieron leer las hojas de {os.path.basename(path)}: {exc}"
                )

            with self._sheets_lock:
                self.sheets_loading.discard(path)
                if path not in self.selected_files:
                    continue
                if hojas:
                    self.available_sheets[path] = hojas
                    actual = self.sheet_choice.get(path)
                    if actual not in hojas:
                        # Mantiene el comportamiento anterior: prioriza "DATOS".
                        self.sheet_choice[path] = "DATOS" if "DATOS" in hojas else hojas[0]
                    leidos += 1
                else:
                    self.available_sheets.pop(path, None)
                    self.sheet_choice.pop(path, None)

            self.root.after(0, self._refresh_files_listbox)

        if leidos:
            self.safe_log(f"🗂️ Hojas detectadas en {leidos} archivo(s).")
        self.root.after(0, self._after_sheets_scanned)

    def _after_sheets_scanned(self):
        self.update_merge_button()
        with self._sheets_lock:
            cargando = bool(self.sheets_loading)
        if cargando or self._sheet_dialog_open:
            return
        if any(self._is_excel(p) for p in self.selected_files):
            # Se abre solo para que confirmes la hoja de cada archivo.
            self.open_sheet_selector()

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos Excel/CSV", 
            filetypes=[
                ("Archivos de datos", "*.xlsx *.xls *.xlsm *.csv"),
                ("Todos los archivos", "*.*"),
            ]
        )
        nuevos = [f for f in files if f not in self.selected_files]
        if not nuevos:
            return

        self.selected_files.extend(nuevos)
        with self._sheets_lock:
            for path in nuevos:
                if not self._is_excel(path):
                    self.available_sheets[path] = []

        self.log_message(f"📁 {len(nuevos)} archivo(s) agregado(s).")
        self._scan_sheets_async(nuevos)
        self._refresh_files_listbox()
        self.update_merge_button()

    def remove_file(self):
        selection = self.files_listbox.curselection()
        if not selection:
            return

        path = self.selected_files.pop(selection[0])
        with self._sheets_lock:
            self.available_sheets.pop(path, None)
            self.sheet_choice.pop(path, None)
            self.sheets_loading.discard(path)

        self._refresh_files_listbox()
        self.update_merge_button()

    def clear_files(self):
        self.selected_files.clear()
        with self._sheets_lock:
            self.available_sheets.clear()
            self.sheet_choice.clear()
            self.sheets_loading.clear()

        self._refresh_files_listbox()
        self.update_merge_button()

    def open_sheet_selector(self):
        """Ventana con scroll para elegir la hoja de cada archivo Excel."""
        if self._sheet_dialog_open:
            return

        excel_files = [p for p in self.selected_files if self._is_excel(p)]
        if not excel_files:
            messagebox.showinfo(
                "Sin hojas para elegir",
                "Los archivos CSV no tienen hojas. Se procesan completos.",
            )
            return

        with self._sheets_lock:
            cargando = bool(self.sheets_loading)
        if cargando:
            messagebox.showinfo(
                "Leyendo hojas",
                "Todavía se están leyendo las hojas de los archivos.\n"
                "Esperá un momento y volvé a intentar.",
            )
            return

        self._sheet_dialog_open = True
        win = tk.Toplevel(self.root)
        win.title("🗂️ Elegir hoja de cada archivo")
        win.geometry("780x560")
        win.minsize(620, 380)
        win.transient(self.root)
        win.grab_set()
        win.columnconfigure(0, weight=1)
        win.rowconfigure(1, weight=1)

        encabezado = ttk.Frame(win, padding=(18, 16, 18, 8))
        encabezado.grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Label(
            encabezado,
            text="Elegí la hoja que se usará en cada archivo",
            style='Title.TLabel',
        ).pack(anchor=tk.W)
        ttk.Label(
            encabezado,
            text="Solo se leerá la hoja seleccionada. Los CSV no tienen hojas y se procesan completos.",
            foreground="#555555",
        ).pack(anchor=tk.W, pady=(4, 0))

        contenedor = ttk.Frame(win)
        contenedor.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=18)
        contenedor.columnconfigure(0, weight=1)
        contenedor.rowconfigure(0, weight=1)

        canvas = tk.Canvas(contenedor, highlightthickness=0, borderwidth=0)
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar = ttk.Scrollbar(contenedor, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        canvas.configure(yscrollcommand=scrollbar.set)

        interior = ttk.Frame(canvas)
        ventana_id = canvas.create_window((0, 0), window=interior, anchor="nw")

        def ajustar_scroll(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def ajustar_ancho(event):
            canvas.itemconfigure(ventana_id, width=event.width)

        interior.bind("<Configure>", ajustar_scroll)
        canvas.bind("<Configure>", ajustar_ancho)

        def mover_rueda(event):
            if getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(1, "units")
            else:
                delta = int(-1 * (event.delta / 120)) if event.delta else 0
                if delta:
                    canvas.yview_scroll(delta, "units")
            return "break"

        variables = {}
        for i, path in enumerate(excel_files, 1):
            fila = ttk.Frame(interior, padding=(0, 9))
            fila.pack(fill=tk.X)
            fila.columnconfigure(1, weight=1)

            ttk.Label(fila, text=f"{i}.", width=3).grid(
                row=0, column=0, rowspan=2, sticky=tk.NW
            )
            ttk.Label(
                fila,
                text=self._acortar_texto(os.path.basename(path), 70),
                font=('Segoe UI', 10, 'bold'),
            ).grid(row=0, column=1, sticky=tk.W)
            ttk.Label(
                fila,
                text=self._acortar_texto(os.path.dirname(path), 85, desde_inicio=True),
                foreground="#666666",
                font=('Segoe UI', 8),
            ).grid(row=1, column=1, sticky=tk.W)

            with self._sheets_lock:
                hojas = list(self.available_sheets.get(path, []))
                elegida = self.sheet_choice.get(path)

            inicial = elegida if elegida in hojas else (hojas[0] if hojas else "")
            var = tk.StringVar(value=inicial)
            variables[path] = var

            combo = ttk.Combobox(
                fila,
                textvariable=var,
                values=hojas,
                state="readonly",
                width=30,
                font=('Segoe UI', 9),
            )
            combo.grid(row=0, column=2, rowspan=2, sticky=tk.E, padx=(12, 6))
            combo.bind("<MouseWheel>", mover_rueda)
            combo.bind("<Button-4>", mover_rueda)
            combo.bind("<Button-5>", mover_rueda)

            if hojas:
                ttk.Label(
                    fila,
                    text=f"{len(hojas)} hoja(s) disponible(s)",
                    foreground="#666666",
                    font=('Segoe UI', 8),
                ).grid(row=2, column=1, sticky=tk.W, pady=(2, 0))
            else:
                combo.config(state=tk.DISABLED)
                ttk.Label(
                    fila,
                    text="⚠️ No se pudieron leer las hojas de este archivo.",
                    foreground="#b00020",
                    font=('Segoe UI', 8),
                ).grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=(2, 0))

            ttk.Separator(interior, orient=tk.HORIZONTAL).pack(fill=tk.X)

        # Rueda del mouse sobre cualquier parte de la ventana.
        win.bind_all("<MouseWheel>", mover_rueda)
        win.bind_all("<Button-4>", mover_rueda)
        win.bind_all("<Button-5>", mover_rueda)

        pie = ttk.Frame(win, padding=(18, 10, 18, 16))
        pie.grid(row=2, column=0, sticky=(tk.W, tk.E))
        pie.columnconfigure(0, weight=1)

        def primera_para_todos():
            for path, var in variables.items():
                with self._sheets_lock:
                    hojas = self.available_sheets.get(path, [])
                if hojas:
                    var.set(hojas[0])

        def cerrar():
            win.unbind_all("<MouseWheel>")
            win.unbind_all("<Button-4>")
            win.unbind_all("<Button-5>")
            self._sheet_dialog_open = False
            win.destroy()

        def aplicar():
            aplicadas = 0
            for path, var in variables.items():
                hoja = var.get()
                if hoja:
                    self.sheet_choice[path] = hoja
                    aplicadas += 1
            self._refresh_files_listbox()
            self.log_message(f"🗂️ Hojas confirmadas para {aplicadas} archivo(s).")
            cerrar()

        ttk.Button(
            pie,
            text="☑️ Primera hoja para todos",
            command=primera_para_todos,
        ).grid(row=0, column=0, sticky=tk.W)
        ttk.Button(pie, text="Cancelar", command=cerrar).grid(
            row=0, column=1, padx=(8, 0)
        )
        ttk.Button(
            pie,
            text="✅ Aplicar",
            style='Success.TButton',
            command=aplicar,
        ).grid(row=0, column=2, padx=(8, 0))

        win.protocol("WM_DELETE_WINDOW", cerrar)
        win.update_idletasks()
        x = self.root.winfo_rootx() + max(
            (self.root.winfo_width() - win.winfo_width()) // 2, 0
        )
        y = self.root.winfo_rooty() + max(
            (self.root.winfo_height() - win.winfo_height()) // 2, 0
        )
        win.geometry(f"+{x}+{y}")
        win.focus_set()

    def browse_output(self):
        fn = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Archivo Excel", "*.xlsx"), ("Archivo CSV", "*.csv")]
        )
        if fn: self.output_file.set(fn)

    def open_row_counter(self):
        """Abre la herramienta opcional de conteo de filas por carpeta."""
        if self.row_counter_window is not None:
            try:
                self.row_counter_window.win.deiconify()
                self.row_counter_window.win.lift()
                self.row_counter_window.win.focus_force()
                return
            except tk.TclError:
                self.row_counter_window = None

        self.row_counter_window = RowCounterWindow(
            self.root, on_close=self._row_counter_closed
        )

    def _row_counter_closed(self):
        self.row_counter_window = None

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
        with self._sheets_lock:
            cargando = bool(self.sheets_loading)
        hay_excel = any(self._is_excel(p) for p in self.selected_files)

        self.merge_btn.config(
            state=tk.NORMAL if self.selected_files and not cargando else tk.DISABLED
        )
        self.sheets_btn.config(
            state=tk.NORMAL if hay_excel and not cargando else tk.DISABLED
        )

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
                # Guardamos el encabezado original para el archivo de salida.
                headers = [
                    "" if pd.isna(valor) else str(valor).strip()
                    for valor in df.iloc[i].tolist()
                ]

                datos = df.iloc[i + 1:].copy()
                datos = datos.dropna(how='all')

                # Columnas posicionales: si se conservan los nombres originales,
                # pandas alinea por nombre y los encabezados que cambian de un
                # archivo a otro (por ejemplo una coma por un punto) hacen que
                # los valores se pierdan o queden en columnas extra.
                datos.columns = range(datos.shape[1])

                # APLICAR NORMALIZACIÓN DE FECHAS AQUÍ
                datos = self.normalize_dates(datos)
                datos.attrs["headers"] = headers
                return datos
        return None

    def process_file(self, path, sheet=None):
        ext = os.path.splitext(path)[1].lower()
        if ext == '.csv':
            try:
                df = pd.read_csv(path, header=None, encoding='utf-8')
                return self.extract_valid_data(df)
            except Exception:
                df = pd.read_csv(path, header=None, encoding='latin1', sep=';')
                return self.extract_valid_data(df)
        elif ext in ['.xlsx', '.xls', '.xlsm']:
            with pd.ExcelFile(path) as xls:
                sheet_names = list(xls.sheet_names)

                # Si el usuario eligió una hoja, se respeta esa elección.
                if sheet and sheet in sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet, header=None)
                    return self.extract_valid_data(df)

                # Sin elección (o si la hoja ya no existe): detección automática.
                if "DATOS" in sheet_names:
                    df = pd.read_excel(xls, sheet_name="DATOS", header=None)
                    result = self.extract_valid_data(df)
                    if result is not None:
                        return result
                for sheet_name in sheet_names:
                    if sheet_name == "DATOS": 
                        continue
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
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
            headers = None
            ancho_referencia = None

            for i, path in enumerate(self.selected_files, 1):
                filename = os.path.basename(path)
                hoja = self.sheet_choice.get(path)
                if hoja:
                    self.safe_log(
                        f"[{i}/{len(self.selected_files)}] Procesando: {filename} "
                        f"(hoja: {hoja})"
                    )
                else:
                    self.safe_log(
                        f"[{i}/{len(self.selected_files)}] Procesando: {filename}"
                    )

                try:
                    df_extracted = self.process_file(path, hoja)
                except Exception as exc:
                    self.safe_log(f"  └─ ❌ Error al leer el archivo: {exc}")
                    continue
                
                if df_extracted is not None:
                    if headers is None:
                        headers = list(df_extracted.attrs.get("headers") or [])
                        ancho_referencia = df_extracted.shape[1]
                    elif (
                        ancho_referencia is not None
                        and df_extracted.shape[1] != ancho_referencia
                    ):
                        self.safe_log(
                            f"  └─ ⚠️ Tiene {df_extracted.shape[1]} columnas; "
                            f"la referencia tiene {ancho_referencia}."
                        )

                    if not df_extracted.empty:
                        all_dataframes.append(df_extracted)
                        self.safe_log("  └─ Datos extraídos correctamente.")
                    else:
                        self.safe_log("  └─ ⚠️ La hoja no tiene filas de datos.")
                elif hoja:
                    self.safe_log(
                        f"  └─ ⚠️ La hoja '{hoja}' no tiene datos válidos "
                        "(encabezado con CÉDULA)."
                    )
                else:
                    self.safe_log("  └─ ⚠️ No se encontró cédula o no hay datos.")

            if not all_dataframes:
                # Volvemos al hilo principal para mostrar errores o finalizar
                self.root.after(0, self._merge_finish_empty)
                return

            # Los datos se apilan por posición de columna, no por nombre.
            final_df = pd.concat(all_dataframes, ignore_index=True)

            # Restauramos el encabezado del primer archivo válido.
            if headers:
                ancho = final_df.shape[1]
                headers = headers[:ancho]
                if len(headers) < ancho:
                    headers.extend(
                        f"Columna {n}" for n in range(len(headers) + 1, ancho + 1)
                    )
                final_df.columns = headers

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

class RowCounterWindow:
    """Herramienta opcional: cuenta filas de los Excel/CSV de una carpeta.

    No participa del flujo de combinación; se abre desde el botón
    "Contador de filas (opcional)".
    """

    EXTENSIONES = {'.xlsx', '.xls', '.xlsm', '.csv'}

    def __init__(self, parent, on_close=None):
        self.parent = parent
        self.on_close = on_close
        self._closed = False
        self._counting = False

        self.win = tk.Toplevel(parent)
        self.win.title("🧮 Contador de filas por carpeta")
        self.win.geometry("700x500")
        self.win.minsize(560, 380)
        self.win.transient(parent)
        self.win.columnconfigure(0, weight=1)
        self.win.rowconfigure(2, weight=1)
        self.win.protocol("WM_DELETE_WINDOW", self._close)

        self._build()
        self._center()

    def _build(self):
        encabezado = ttk.Frame(self.win, padding=(16, 14, 16, 6))
        encabezado.grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Label(
            encabezado,
            text="Contador de filas por carpeta",
            style='Title.TLabel',
        ).pack(anchor=tk.W)
        ttk.Label(
            encabezado,
            text="Herramienta opcional: cuenta las filas de los Excel/CSV de una "
                 "carpeta. No modifica el combinador.",
            foreground="#555555",
        ).pack(anchor=tk.W, pady=(4, 0))

        barra = ttk.Frame(self.win, padding=(16, 0, 16, 8))
        barra.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.btn_carpeta = ttk.Button(
            barra,
            text="📁 Seleccionar carpeta y contar",
            style='Custom.TButton',
            command=self._elegir_carpeta,
        )
        self.btn_carpeta.pack(side=tk.LEFT)
        self.lbl_estado = ttk.Label(barra, text="", foreground="#555555")
        self.lbl_estado.pack(side=tk.LEFT, padx=(12, 0))

        tabla_frame = ttk.Frame(self.win, padding=(16, 0, 16, 0))
        tabla_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tabla_frame.columnconfigure(0, weight=1)
        tabla_frame.rowconfigure(0, weight=1)

        columnas = ("archivo", "tipo", "hoja", "filas")
        self.tabla = ttk.Treeview(
            tabla_frame, columns=columnas, show="headings", height=12
        )
        self.tabla.heading("archivo", text="Nombre del archivo")
        self.tabla.heading("tipo", text="Tipo")
        self.tabla.heading("hoja", text="Hoja con CÉDULA")
        self.tabla.heading("filas", text="Filas (sin encabezado)")
        self.tabla.column("archivo", width=300, anchor=tk.W)
        self.tabla.column("tipo", width=70, anchor=tk.CENTER, stretch=False)
        self.tabla.column("hoja", width=150, anchor=tk.W, stretch=False)
        self.tabla.column("filas", width=130, anchor=tk.E, stretch=False)
        self.tabla.tag_configure("error", foreground="#b00020")
        self.tabla.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        scrollbar = ttk.Scrollbar(
            tabla_frame, orient=tk.VERTICAL, command=self.tabla.yview
        )
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.tabla.configure(yscrollcommand=scrollbar.set)

        pie = ttk.Frame(self.win, padding=(16, 10, 16, 14))
        pie.grid(row=3, column=0, sticky=(tk.W, tk.E))
        pie.columnconfigure(0, weight=1)

        self.lbl_total = ttk.Label(
            pie,
            text="Sumatoria total de filas: 0",
            font=('Segoe UI', 12, 'bold'),
            anchor=tk.E,
        )
        self.lbl_total.grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(pie, text="Cerrar", command=self._close).grid(
            row=0, column=1, padx=(12, 0)
        )

        ttk.Label(
            self.win,
            text="Nota: en Excel se busca la hoja cuyo encabezado contenga "
                 "CÉDULA; en CSV, la primera fila se toma como encabezado.",
            foreground="#666666",
            font=('Segoe UI', 8),
        ).grid(row=4, column=0, sticky=tk.W, padx=16, pady=(0, 10))

    def _center(self):
        self.win.update_idletasks()
        x = self.parent.winfo_rootx() + max(
            (self.parent.winfo_width() - self.win.winfo_width()) // 2, 0
        )
        y = self.parent.winfo_rooty() + max(
            (self.parent.winfo_height() - self.win.winfo_height()) // 2, 0
        )
        self.win.geometry(f"+{x}+{y}")

    def _close(self):
        self._closed = True
        if self.on_close is not None:
            self.on_close()
        self.win.destroy()

    def _elegir_carpeta(self):
        if self._counting:
            return

        carpeta = filedialog.askdirectory(
            title="Seleccioná la carpeta con los archivos", parent=self.win
        )
        if not carpeta:
            return

        try:
            with os.scandir(carpeta) as entradas:
                archivos = sorted(
                    (
                        entrada
                        for entrada in entradas
                        if entrada.is_file()
                        and not entrada.name.startswith("~$")
                        and os.path.splitext(entrada.name)[1].lower()
                        in self.EXTENSIONES
                    ),
                    key=lambda entrada: entrada.name.lower(),
                )
                rutas = [entrada.path for entrada in archivos]
        except OSError as exc:
            messagebox.showerror(
                "Error",
                f"No se pudo leer la carpeta:\n{exc}",
                parent=self.win,
            )
            return

        self._limpiar()
        self.lbl_total.config(text="Sumatoria total de filas: 0")

        if not rutas:
            self.lbl_estado.config(text="")
            messagebox.showinfo(
                "Sin archivos",
                "No se encontraron archivos .xlsx, .xls, .xlsm o .csv "
                "en esa carpeta.",
                parent=self.win,
            )
            return

        self._counting = True
        self.btn_carpeta.config(state=tk.DISABLED)
        self.lbl_estado.config(text=f"Contando 0/{len(rutas)}...")

        threading.Thread(
            target=self._contar_worker, args=(rutas,), daemon=True
        ).start()

    def _limpiar(self):
        for row in self.tabla.get_children():
            self.tabla.delete(row)

    @staticmethod
    def _fila_tiene_cedula(fila):
        objetivo = {"CÉDULA", "CEDULA", "CI", "C.I", "C.I."}
        valores = fila.astype(str).str.upper().str.strip().tolist()
        return any(valor in objetivo for valor in valores)

    @classmethod
    def _buscar_hoja_cedula(cls, libro):
        """Devuelve (nombre_hoja, indice_encabezado) de la hoja con CÉDULA."""
        for hoja in libro.sheet_names:
            muestra = pd.read_excel(
                libro, sheet_name=hoja, header=None, nrows=30
            )
            for i in range(min(30, len(muestra))):
                if cls._fila_tiene_cedula(muestra.iloc[i]):
                    return hoja, i
        return None, None

    @classmethod
    def _contar_archivo(cls, ruta, ext):
        if ext == ".csv":
            try:
                df = pd.read_csv(ruta)
            except UnicodeDecodeError:
                df = pd.read_csv(ruta, encoding="latin1")
            return "—", len(df)

        with pd.ExcelFile(ruta) as libro:
            hoja, indice = cls._buscar_hoja_cedula(libro)
            if hoja is None:
                raise ValueError(
                    "No se encontró CÉDULA en el encabezado de ninguna hoja"
                )

            df = pd.read_excel(libro, sheet_name=hoja, header=None)
            datos = df.iloc[indice + 1:].dropna(how="all")
            return hoja, len(datos)

    def _contar_worker(self, rutas):
        total = 0
        for i, ruta in enumerate(rutas, 1):
            nombre = os.path.basename(ruta)
            ext = os.path.splitext(nombre)[1].lower()
            tipo = ext.lstrip(".").upper()

            try:
                hoja, filas = self._contar_archivo(ruta, ext)
                total += filas
                self._ui(
                    self._agregar_fila,
                    nombre,
                    tipo,
                    hoja,
                    f"{filas:,}",
                    False,
                )
            except Exception as exc:
                self._ui(
                    self._agregar_fila,
                    nombre,
                    tipo,
                    "—",
                    f"Error: {exc}",
                    True,
                )

            self._ui(self._actualizar_estado, i, len(rutas))

        self._ui(self._finalizar, total, len(rutas))

    def _ui(self, func, *args):
        """Envía una actualización a la GUI desde el hilo de conteo."""
        if self._closed:
            return
        try:
            self.win.after(0, func, *args)
        except (tk.TclError, RuntimeError):
            pass

    def _agregar_fila(self, nombre, tipo, hoja, valor, error):
        self.tabla.insert(
            "",
            tk.END,
            values=(nombre, tipo, hoja, valor),
            tags=("error",) if error else (),
        )

    def _actualizar_estado(self, i, total):
        self.lbl_estado.config(text=f"Contando {i}/{total}...")

    def _finalizar(self, total, cantidad):
        self._counting = False
        self.btn_carpeta.config(state=tk.NORMAL)
        self.lbl_estado.config(text=f"{cantidad} archivo(s) contado(s).")
        self.lbl_total.config(text=f"Sumatoria total de filas: {total:,}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCsvMergerApp(root)
    root.mainloop()