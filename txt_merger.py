import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
from pathlib import Path
from datetime import datetime

class TextMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🧩 Combinador de Archivos TXT")
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
        """Configura los estilos personalizados"""
        self.style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'))
        self.style.configure('Custom.TButton', font=('Segoe UI', 10), padding=10)
        self.style.configure('Danger.TButton', font=('Segoe UI', 10), padding=10)
        self.style.configure('Success.TButton', font=('Segoe UI', 10), padding=10)
        
    def create_widgets(self):
        """Crea todos los widgets de la interfaz"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        title_label = ttk.Label(main_frame, text="🧩 Combinador de Archivos TXT", style='Title.TLabel')
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
        self.merge_btn = ttk.Button(main_frame, text="🔗 COMBINAR ARCHIVOS", style='Success.TButton', command=self.merge_files, state=tk.DISABLED)
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

    def log_message(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def add_files(self):
        files = filedialog.askopenfilenames(title="Seleccionar archivos TXT", filetypes=[("Archivos de texto", "*.txt"), ("Todos los archivos", "*.*")])
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
        fn = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Archivo de texto", "*.txt")])
        if fn: self.output_file.set(fn)

    def update_merge_button(self):
        self.merge_btn.config(state=tk.NORMAL if self.selected_files else tk.DISABLED)

    def merge_files(self):
        output_path = self.output_file.get()
        if not output_path:
            self.browse_output()
            output_path = self.output_file.get()
            if not output_path: return

        try:
            self.log_message("🚀 Iniciando mezcla limpia...")
            with open(output_path, 'w', encoding='utf-8') as outfile:
                for i, path in enumerate(self.selected_files, 1):
                    filename = os.path.basename(path)
                    content = None
                    
                    # Intentar diferentes codificaciones para evitar errores de lectura
                    for enc in ['utf-8', 'latin-1', 'cp1252', 'utf-8-sig']:
                        try:
                            with open(path, 'r', encoding=enc) as infile:
                                lines = infile.readlines()
                                if lines:
                                    # [1:] quita la primera línea, ''.join() los vuelve a unir
                                    content = "".join(lines[1:])
                                else:
                                    content = ""
                                break
                        except (UnicodeDecodeError, TypeError):
                            continue
                    
                    if content is None:
                        # Fallback por si todo falla
                        with open(path, 'r', encoding='utf-8', errors='ignore') as infile:
                            lines = infile.readlines()
                            content = "".join(lines[1:]) if lines else ""

                    # Escribir el contenido directamente sin decoraciones ni espacios extra
                    outfile.write(content)
                    
                    # Aseguramos que si el archivo no termina en salto de línea, 
                    # el siguiente no se pegue en la misma fila (opcional, pero recomendado)
                    if content and not content.endswith('\n'):
                        outfile.write('\n')
                    
                    self.log_message(f"({i}/{len(self.selected_files)}) Sin cabecera: {filename}")

            messagebox.showinfo("Éxito", "Archivos mezclados (sin 1ra línea y sin espacios).")
            self.log_message("✅ ¡Proceso finalizado!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {e}")
            self.log_message(f"❌ Error: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TextMergerApp(root)
    root.mainloop()