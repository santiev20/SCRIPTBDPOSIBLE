import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
import os
from datetime import datetime
import locale

# --- Configuración ---
SHEET_POSIBLES = "POSIBLES"
SHEET_ELIMINADOS = "ELIMINADOS"
SHEET_RESUMEN = "RESUMEN"
SHEET_MATRIZ = "MATRIZ" 
HOJAS_A_CONSERVAR_BASE = ["POSIBLES", "ENVIADOS", "METAS", "ASI VAMOS", "APROVECHABLES", SHEET_MATRIZ]
COLUMNAS_ELIMINADOS = ["Cliente", "Residuo", "Tipología", "Línea", "Comercial", "Fecha CC", "CC", "Subtotal", "Fecha de Eliminación"]
COLUMNAS_POSIBLES_ESPERADAS = []
COLUMNAS_MAPEADAS_CSV = ["Cliente", "Nit Cliente", "Peso CC", "Vlr Unit", "Subtotal", "Fecha CC"]
COLUMNAS_REQUERIDAS_CSV = ['client', 'nit', 'weight', 'price', 'collection_date']
MATRIZ_NIT_COL = 'NIT'; MATRIZ_COMERCIAL_COL = 'NOMBRE CO'
FECHA_COLUMNA = "Fecha CC"; CLIENTE_COLUMNA = "Cliente"; COMERCIAL_COLUMNA = "Comercial"; SUBTOTAL_COLUMNA = "Subtotal"; PESO_COLUMNA = 'Peso CC'

# --- Funciones Auxiliares ---
# (cargar_datos_excel, generar_resumen, guardar_cambios_excel_completo, procesar_y_cargar_csv SIN CAMBIOS)
def cargar_datos_excel(filepath, sheet_name):
    if not filepath or not os.path.exists(filepath): messagebox.showerror("Error", f"Archivo no válido: {filepath}"); return None
    critical_columns_map = {
        SHEET_POSIBLES: [FECHA_COLUMNA, CLIENTE_COLUMNA, COMERCIAL_COLUMNA, SUBTOTAL_COLUMNA, PESO_COLUMNA],
        SHEET_MATRIZ: [MATRIZ_NIT_COL, MATRIZ_COMERCIAL_COL]
    }
    critical_columns = critical_columns_map.get(sheet_name, [])
    try:
        df = pd.read_excel(filepath, sheet_name=sheet_name)
        missing_cols = [col for col in critical_columns if col not in df.columns]
        if missing_cols: messagebox.showwarning("Advertencia Columnas", f"Faltan columnas esperadas en '{sheet_name}': {', '.join(missing_cols)}.")
        if sheet_name == SHEET_POSIBLES:
            global COLUMNAS_POSIBLES_ESPERADAS; COLUMNAS_POSIBLES_ESPERADAS = df.columns.tolist()
            if FECHA_COLUMNA in df.columns: df[FECHA_COLUMNA] = pd.to_datetime(df[FECHA_COLUMNA], errors='coerce')
            if SUBTOTAL_COLUMNA in df.columns: df[SUBTOTAL_COLUMNA] = pd.to_numeric(df[SUBTOTAL_COLUMNA], errors='coerce')
            if PESO_COLUMNA in df.columns: df[PESO_COLUMNA] = pd.to_numeric(df[PESO_COLUMNA], errors='coerce')
            if 'Vlr Unit' in df.columns: df['Vlr Unit'] = pd.to_numeric(df['Vlr Unit'], errors='coerce')
            if 'Nit Cliente' in df.columns: df['Nit Cliente'] = df['Nit Cliente'].astype(str).str.strip()
        elif sheet_name == SHEET_MATRIZ:
             if MATRIZ_NIT_COL in df.columns: df[MATRIZ_NIT_COL] = df[MATRIZ_NIT_COL].astype(str).str.strip()
        return df
    except FileNotFoundError: messagebox.showerror("Error", f"Archivo no encontrado: {filepath}"); return None
    except ValueError as e:
         if "Worksheet" in str(e) and "not found" in str(e):
              if sheet_name in HOJAS_A_CONSERVAR_BASE or sheet_name == SHEET_ELIMINADOS or sheet_name == SHEET_RESUMEN: return pd.DataFrame()
              else: messagebox.showerror("Error", f"No se encontró la hoja '{sheet_name}'."); return None
         else: messagebox.showerror("Error", f"Error al leer hoja '{sheet_name}': {e}"); return None
    except Exception as e: messagebox.showerror("Error", f"Error inesperado al cargar {os.path.basename(filepath)} ({sheet_name}): {e}"); return None
    
def generar_resumen(df, fecha_columna="Fecha CC"):  # Añadimos parámetro para la columna de fecha
    if df is None or df.empty:
        return pd.DataFrame({"Mensaje": ["No hay datos"]})
    required_cols = [fecha_columna, COMERCIAL_COLUMNA, SUBTOTAL_COLUMNA]
    missing_cols = [col for col in required_cols if col not in df.columns and col is not None] # Modificado para manejar None
    if missing_cols:
        return pd.DataFrame({"Error": [f"Faltan columnas: {', '.join(missing_cols)}"]})
    
    if fecha_columna is not None and fecha_columna in df.columns:  # Añadimos comprobación antes de la conversión
        if not pd.api.types.is_datetime64_any_dtype(df[fecha_columna]):
            df[fecha_columna] = pd.to_datetime(df[fecha_columna], errors='coerce')
    if SUBTOTAL_COLUMNA is not None and SUBTOTAL_COLUMNA in df.columns:  # Añadimos comprobación antes de la conversión
        if not pd.api.types.is_numeric_dtype(df[SUBTOTAL_COLUMNA]):
            df[SUBTOTAL_COLUMNA] = pd.to_numeric(df[SUBTOTAL_COLUMNA], errors='coerce')
    
    df_valid = df.dropna(subset=[col for col in [fecha_columna, SUBTOTAL_COLUMNA] if col is not None]).copy() # Modificado para manejar None
    if df_valid.empty:
        return pd.DataFrame({"Mensaje": ["No hay datos válidos"]})
    
    df_valid = df_valid.copy()
    if fecha_columna is not None and fecha_columna in df_valid.columns:  # Añadimos comprobación antes de crear 'Mes'
        df_valid['Mes'] = df_valid[fecha_columna].dt.strftime('%Y-%m')
    if COMERCIAL_COLUMNA is not None and COMERCIAL_COLUMNA in df_valid.columns:  # Añadimos comprobación antes de llenar 'Sin Asignar'
        df_valid[COMERCIAL_COLUMNA] = df_valid[COMERCIAL_COLUMNA].fillna('Sin Asignar')
    
    try:
        pivot_resumen = pd.pivot_table(df_valid, values=SUBTOTAL_COLUMNA, index='Mes',
                                      columns=COMERCIAL_COLUMNA, aggfunc='sum',
                                      fill_value=0, margins=True, margins_name='Total General')
        pivot_resumen = pivot_resumen.rename(index={'Total General': 'Total Comercial'},
                                            columns={'Total General': 'Total Mes'})
        print("Tabla pivote de resumen generada.")
        return pivot_resumen
    except Exception as e:
        return pd.DataFrame({"Error": [f"Error pivote: {e}"]})

def guardar_cambios_excel_completo(filepath, df_posibles_actualizado, df_nuevos_eliminados):
    if not filepath: messagebox.showerror("Error", "No hay archivo seleccionado."); return False
    try:
        df_eliminados_total = None
        try:
            df_eliminados_existente = pd.read_excel(filepath, sheet_name=SHEET_ELIMINADOS);
            if FECHA_COLUMNA in df_eliminados_existente.columns: df_eliminados_existente[FECHA_COLUMNA] = pd.to_datetime(df_eliminados_existente[FECHA_COLUMNA], errors='coerce')
        except ValueError: df_eliminados_existente = pd.DataFrame()
        except Exception as e: messagebox.showwarning("Advertencia", f"No se pudo leer '{SHEET_ELIMINADOS}': {e}."); df_eliminados_existente = pd.DataFrame()
        if df_nuevos_eliminados is None: df_nuevos_eliminados = pd.DataFrame()
        if not df_nuevos_eliminados.empty:
            if FECHA_COLUMNA in df_nuevos_eliminados.columns: df_nuevos_eliminados[FECHA_COLUMNA] = pd.to_datetime(df_nuevos_eliminados[FECHA_COLUMNA], errors='coerce')
            df_eliminados_total = pd.concat([df_eliminados_existente, df_nuevos_eliminados], ignore_index=True)
        else: df_eliminados_total = df_eliminados_existente
        df_resumen = generar_resumen(df_posibles_actualizado, FECHA_COLUMNA); # Resumen de POSIBLES
        if df_resumen is None: messagebox.showerror("Error", "Fallo Resumen."); return False
        
        # --- NUEVO: Generar resumen de eliminados ---
        df_eliminados_resumen = None
        if df_eliminados_total is not None and not df_eliminados_total.empty:
            df_eliminados_resumen = generar_resumen(df_eliminados_total.copy(), "Fecha de Eliminación")  # Resumen de ELIMINADOS
            if df_eliminados_resumen is None:
                messagebox.showwarning("Advertencia", "No se pudo generar resumen de eliminados.")
                df_eliminados_resumen = pd.DataFrame({"Mensaje": ["No se pudo generar resumen de eliminados"]})
        # --- FIN NUEVO ---

        hojas_a_escribir = {}
        if df_posibles_actualizado is not None and not df_posibles_actualizado.empty: hojas_a_escribir[SHEET_POSIBLES] = df_posibles_actualizado
        if df_resumen is not None and not df_resumen.empty and 'Error' not in df_resumen.columns and 'Mensaje' not in df_resumen.columns: hojas_a_escribir[SHEET_RESUMEN] = df_resumen
        if df_eliminados_total is not None and not df_eliminados_total.empty:
             if FECHA_COLUMNA in df_eliminados_total.columns: df_eliminados_total[FECHA_COLUMNA] = df_eliminados_total[FECHA_COLUMNA].dt.strftime('%Y/%m/%d')
             hojas_a_escribir[SHEET_ELIMINADOS] = df_eliminados_total
        
        # --- NUEVO: Añadir resumen de eliminados a 'ENVIADOS' ---
        if df_eliminados_resumen is not None and not df_eliminados_resumen.empty and 'Error' not in df_eliminados_resumen.columns and 'Mensaje' not in df_eliminados_resumen.columns:
            hojas_a_escribir["ENVIADOS"] = df_eliminados_resumen
        # --- FIN NUEVO ---

        for nombre_hoja in HOJAS_A_CONSERVAR_BASE:
            if nombre_hoja not in hojas_a_escribir:
                try:
                    df_hoja_conservar = pd.read_excel(filepath, sheet_name=nombre_hoja)
                    if df_hoja_conservar is not None: hojas_a_escribir[nombre_hoja] = df_hoja_conservar; print(f"Conservando: '{nombre_hoja}'")
                except ValueError as e:
                    if "Worksheet" in str(e) and "not found" in str(e): print(f"Nota: Hoja '{nombre_hoja}' no encontrada.")
                    else: messagebox.showwarning("Advertencia", f"No se pudo leer '{nombre_hoja}': {e}.")
                except Exception as e: messagebox.showwarning("Advertencia", f"Error leer '{nombre_hoja}': {e}.")
        if not hojas_a_escribir: messagebox.showerror("Error Guardado", "No hay datos válidos."); return False
        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            for nombre_hoja, df_hoja in hojas_a_escribir.items():
                write_index = (nombre_hoja == SHEET_RESUMEN or nombre_hoja == "ENVIADOS")  # Modificado para ENVIADOS
                df_hoja.to_excel(writer, sheet_name=nombre_hoja, index=write_index)
            if SHEET_RESUMEN in hojas_a_escribir or "ENVIADOS" in hojas_a_escribir:
                workbook = writer.book
                for sheet_name in [SHEET_RESUMEN, "ENVIADOS"]:
                    if sheet_name in hojas_a_escribir:
                        worksheet = writer.sheets[sheet_name]
                        money_format = workbook.add_format({'num_format': '$ #,##0'})
                        start_col = 1
                        end_col = hojas_a_escribir[sheet_name].shape[1]
                        worksheet.set_column(start_col, end_col, None, money_format)
                        print(f"Formato moneda aplicado a '{sheet_name}'.")
    except ImportError: messagebox.showerror("Error Dependencia", "Instala 'xlsxwriter': pip install xlsxwriter"); return False
    except PermissionError: messagebox.showerror("Error Permiso", f"No se pudo guardar '{os.path.basename(filepath)}'.\n¡Archivo abierto?"); return False
    except Exception as e: messagebox.showerror("Error", f"Error al guardar '{os.path.basename(filepath)}': {e}"); return False
    return True

def procesar_y_cargar_csv(csv_filepath, df_principal):
    # (Función sin cambios)
    if df_principal is None: messagebox.showerror("Error", "Carga Excel principal primero."); return None
    columnas_requeridas_csv_local = COLUMNAS_REQUERIDAS_CSV
    try:
        try: df_csv = pd.read_csv(csv_filepath, delimiter=';', decimal=',', encoding='utf-8', dtype={'nit': str})
        except UnicodeDecodeError: df_csv = pd.read_csv(csv_filepath, delimiter=';', decimal=',', encoding='latin1', dtype={'nit': str})
        except Exception as e: raise ValueError(f"Error al leer CSV: {e}")
        missing_csv_cols = [col for col in columnas_requeridas_csv_local if col not in df_csv.columns]
        if missing_csv_cols: raise ValueError(f"Faltan columnas en CSV: {', '.join(missing_csv_cols)}")
        df_import = pd.DataFrame()
        df_import['CC'] = df_csv['manifest'].astype(str).str.strip()
        df_import['Cliente'] = df_csv['client']
        df_import['Nit Cliente'] = df_csv['nit'].astype(str).str[:-1].str.strip()
        df_import['Peso CC'] = pd.to_numeric(df_csv['weight'], errors='coerce')
        df_import['Vlr Unit'] = pd.to_numeric(df_csv['price'], errors='coerce')
        df_import['Fecha CC'] = pd.to_datetime(df_csv['collection_date'], errors='coerce')
        nan_dates = df_import['Fecha CC'].isna().sum()
        if nan_dates > 0: messagebox.showwarning("Advertencia Fechas CSV", f"{nan_dates} filas CSV tenían fecha inválida.")
        df_import['Subtotal'] = df_import['Peso CC'] * df_import['Vlr Unit']
        if not COLUMNAS_POSIBLES_ESPERADAS: messagebox.showerror("Error Interno", "No se conocen columnas Excel."); return None
        for col in COLUMNAS_POSIBLES_ESPERADAS:
            if col not in df_import.columns: df_import[col] = pd.NA
        df_import = df_import.reindex(columns=COLUMNAS_POSIBLES_ESPERADAS)
        df_actualizado = pd.concat([df_principal, df_import], ignore_index=True)
        added_indices = df_actualizado.index[-len(df_import):]
        nan_rows_peso = df_actualizado.loc[added_indices, 'Peso CC'].isna().sum(); nan_rows_vlr = df_actualizado.loc[added_indices, 'Vlr Unit'].isna().sum()
        nan_rows_subtotal = df_actualizado.loc[added_indices, 'Subtotal'].isna().sum(); nan_dates_import = df_actualizado.loc[added_indices, 'Fecha CC'].isna().sum()
        if nan_rows_peso > 0 or nan_rows_vlr > 0 or nan_dates_import > 0:
             msg = f"Importadas {len(df_import)} filas.\n"
             if nan_rows_peso > 0: msg += f"* 'weight' no numérico/vacío: {nan_rows_peso}\n"
             if nan_rows_vlr > 0: msg += f"* 'price' no numérico/vacío: {nan_rows_vlr}\n"
             if nan_dates_import > 0: msg += f"* 'collection_date' inválida: {nan_dates_import}\n"
             if nan_rows_subtotal > 0: msg += f"=> 'Subtotal' inválido en {nan_rows_subtotal} filas."
             messagebox.showwarning("Advertencia Importación", msg)
        print(f"Procesadas {len(df_csv)} filas CSV. Añadidas {len(df_import)} filas.")
        return df_actualizado
    except FileNotFoundError: messagebox.showerror("Error", f"CSV no encontrado: {csv_filepath}"); return None
    except ValueError as ve: messagebox.showerror("Error Formato CSV", str(ve)); return None
    except Exception as e: messagebox.showerror("Error Procesar CSV", f"Error: {e}"); return None

# --- Aplicación GUI --- 
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestor de Datos Excel (v1.14 - Layout Tabla)")
        self.geometry("1000x780+50+50")
        self.selected_excel_file = None; self.df_original = None; self.df_filtrado = None
        self.original_indices = {}; self.lista_clientes_unicos = []
        # --- Frames ---
        frame_archivo = ttk.Frame(self, padding="10"); frame_archivo.pack(fill=tk.X, pady=2)
        frame_filtros = ttk.Frame(self, padding="10"); frame_filtros.pack(fill=tk.X, pady=2)
        # Usar fill=tk.BOTH y expand=True para que frame_tabla use el espacio vertical
        frame_tabla = ttk.Frame(self, padding="10"); frame_tabla.pack(fill=tk.BOTH, expand=True, pady=2)
        frame_resumen_sel = ttk.Frame(self, padding="5"); frame_resumen_sel.pack(fill=tk.X, pady=0)
        frame_acciones = ttk.Frame(self, padding="10"); frame_acciones.pack(fill=tk.X, pady=5)
        # --- Widgets Archivo ---
        self.select_button = ttk.Button(frame_archivo, text="1. Seleccionar Excel", command=self.seleccionar_archivo); self.select_button.grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.selected_file_label = ttk.Label(frame_archivo, text="Ningún archivo", foreground="grey", width=40); self.selected_file_label.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        self.import_button = ttk.Button(frame_archivo, text="2. Importar CSV", command=self.importar_reporte_csv, state='disabled'); self.import_button.grid(row=0, column=2, padx=10, pady=5, sticky='w')
        frame_archivo.grid_columnconfigure(1, weight=1)
        # --- Widgets Filtros ---
        ttk.Label(frame_filtros, text="Fecha Inicio (Filtro):").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.fecha_inicio_entry = DateEntry(frame_filtros, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd', state='disabled'); self.fecha_inicio_entry.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        ttk.Label(frame_filtros, text="Fecha Fin (Filtro):").grid(row=0, column=2, padx=5, pady=2, sticky="w")
        self.fecha_fin_entry = DateEntry(frame_filtros, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd', state='disabled'); self.fecha_fin_entry.grid(row=0, column=3, padx=5, pady=2, sticky="w")
        ttk.Label(frame_filtros, text="Buscar Cliente (Filtro):").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.cliente_var = tk.StringVar()
        self.cliente_entry = ttk.Entry(frame_filtros, width=50, state='disabled', textvariable=self.cliente_var); self.cliente_entry.grid(row=1, column=1, columnspan=3, padx=5, pady=2, sticky="ew")
        # Listbox con altura aumentada
        self.cliente_listbox = tk.Listbox(frame_filtros, height=6, width=50); # Height = 6
        self.cliente_entry.bind("<KeyRelease>", self.on_keyrelease_cliente)
        self.cliente_entry.bind("<FocusOut>", self.on_cliente_entry_focus_out) # Nuevo handler
        self.cliente_listbox.bind("<<ListboxSelect>>", self.on_select_cliente)
        self.cliente_listbox.bind("<FocusOut>", self.on_cliente_listbox_focus_out) # Nuevo handler
        self.filtrar_button = ttk.Button(frame_filtros, text="Aplicar Filtros", command=self.aplicar_filtros_vista, state='disabled'); self.filtrar_button.grid(row=2, column=1, columnspan=1, padx=5, pady=5, sticky="ew")
        self.mostrar_todo_button = ttk.Button(frame_filtros, text="Mostrar Todo", command=self.mostrar_todo, state='disabled'); self.mostrar_todo_button.grid(row=2, column=2, columnspan=2, padx=5, pady=5, sticky="ew")
        frame_filtros.grid_columnconfigure(1, weight=1); frame_filtros.grid_columnconfigure(2, weight=1)

        # --- Widget Tabla (MODIFICADO A GRID) ---
        self.tree = ttk.Treeview(frame_tabla, selectmode='extended');
        self.tree.bind("<<TreeviewSelect>>", self.update_selection_summary)
        # Scrollbars
        scrollbar_y = ttk.Scrollbar(frame_tabla, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(frame_tabla, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        # Configurar Grid para frame_tabla
        frame_tabla.grid_rowconfigure(0, weight=1)
        frame_tabla.grid_columnconfigure(0, weight=1)
        # Posicionar con Grid
        self.tree.grid(row=0, column=0, sticky='nsew')
        scrollbar_y.grid(row=0, column=1, sticky='ns')
        scrollbar_x.grid(row=1, column=0, sticky='ew')
        # --- Fin Modificación Tabla ---

        # --- Widget Resumen Selección ---
        self.summary_label = ttk.Label(frame_resumen_sel, text="Selección: 0 filas | Suma Subtotal: 0,00 | Suma Peso CC: 0,00", anchor=tk.W)
        self.summary_label.pack(fill=tk.X, padx=10, pady=2)
        # --- Widgets Acciones ---
        self.guardar_button = ttk.Button(frame_acciones, text="Guardar Cambios Actuales", command=self.guardar_sin_mover, state='disabled'); self.guardar_button.pack(side=tk.LEFT, padx=10, pady=5)
        self.mover_button = ttk.Button(frame_acciones, text=f"Mover Seleccionados y Guardar", command=self.mover_seleccionados, state='disabled'); self.mover_button.pack(side=tk.LEFT, padx=10, pady=5)

    # --- Métodos Autocompletar (incluye nuevos handlers focus out) ---
    def on_keyrelease_cliente(self, event=None):
        value = self.cliente_var.get(); self.cliente_listbox.delete(0, tk.END);
        listbox_displayed = False
        if value:
            data = [item for item in self.lista_clientes_unicos if value.lower() in item.lower()]
            if data:
                for item in data: self.cliente_listbox.insert(tk.END, item)
                self.cliente_entry.update_idletasks(); entry_x = self.cliente_entry.winfo_x(); entry_y = self.cliente_entry.winfo_y()
                entry_height = self.cliente_entry.winfo_height(); entry_width = self.cliente_entry.winfo_width()
                self.cliente_listbox.place(in_=self.cliente_entry.master, x=entry_x, y=entry_y + entry_height, width=entry_width); self.cliente_listbox.lift()
                listbox_displayed = True
        if not listbox_displayed: self.hide_cliente_listbox()

    def on_select_cliente(self, event=None):
        widget = event.widget; selection = widget.curselection()
        if selection: value = widget.get(selection[0]); self.cliente_var.set(value); self.hide_cliente_listbox(); self.cliente_entry.focus_set(); self.cliente_entry.icursor(tk.END)

    def hide_cliente_listbox(self): self.cliente_listbox.place_forget()

    def on_cliente_entry_focus_out(self, event=None):
        self.after(150, self._check_focus_and_hide_listbox)

    def on_cliente_listbox_focus_out(self, event=None):
        self.after(150, self._check_focus_and_hide_listbox)

    def _check_focus_and_hide_listbox(self):
        try:
            focused_widget = self.focus_get()
            if focused_widget != self.cliente_listbox and focused_widget != self.cliente_entry:
                self.hide_cliente_listbox()
        except tk.TclError: pass

    # --- Métodos Principales (sin cambios) ---
    def seleccionar_archivo(self):
        filepath = filedialog.askopenfilename(title="Selecciona archivo Excel PRINCIPAL", filetypes=[("Excel", "*.xlsx")])
        if filepath:
            self.selected_excel_file = filepath; filename = os.path.basename(filepath)
            self.selected_file_label.config(text=filename, foreground="black")
            self.df_original = cargar_datos_excel(self.selected_excel_file, SHEET_POSIBLES)
            self.limpiar_treeview(); self.lista_clientes_unicos = []; self.cliente_var.set("")
            self.fecha_inicio_entry.config(state='disabled'); self.fecha_fin_entry.config(state='disabled')
            self.cliente_entry.config(state='disabled'); self.filtrar_button.config(state='disabled')
            self.mostrar_todo_button.config(state='disabled'); self.mover_button.config(state='disabled')
            self.import_button.config(state='disabled'); self.hide_cliente_listbox(); self.guardar_button.config(state='disabled')
            if self.df_original is None: messagebox.showerror("Error Carga", f"No se pudo cargar '{SHEET_POSIBLES}'."); return
            messagebox.showinfo("Carga Completa", f"Cargados {len(self.df_original)} registros de '{filename}'.")
            if CLIENTE_COLUMNA in self.df_original.columns:
                clientes_raw = self.df_original[CLIENTE_COLUMNA].astype(str).unique()
                self.lista_clientes_unicos = sorted([c for c in clientes_raw if c.lower() != 'nan'])
            else: messagebox.showwarning("Advertencia", f"Columna '{CLIENTE_COLUMNA}' no encontrada.")
            self.df_filtrado = self.df_original.copy(); self.mostrar_en_treeview()
            self.fecha_inicio_entry.config(state='normal'); self.fecha_fin_entry.config(state='normal')
            self.fecha_fin_entry.set_date(datetime.now()); self.cliente_entry.config(state='normal')
            self.filtrar_button.config(state='normal'); self.mostrar_todo_button.config(state='normal');
            self.mover_button.config(state='normal'); self.import_button.config(state='normal'); self.guardar_button.config(state='normal')
        else:
            if not self.selected_excel_file: self.selected_file_label.config(text="Ningún archivo seleccionado", foreground="grey")

    def importar_reporte_csv(self):
        if self.df_original is None: messagebox.showwarning("Acción Requerida", "Selecciona Excel principal primero."); return
        if not self.selected_excel_file: messagebox.showerror("Error", "No se ha seleccionado el archivo Excel principal."); return
        csv_filepath = filedialog.askopenfilename(title="Selecciona CSV Reporte Cantidades", filetypes=[("CSV", "*.csv"), ("Todos", "*.*")])
        if not csv_filepath: return
        df_actualizado = procesar_y_cargar_csv(csv_filepath, self.df_original)
        if df_actualizado is not None:
            self.df_original = df_actualizado
            print(f"Datos de '{os.path.basename(csv_filepath)}' añadidos. Registros: {len(self.df_original)}")
            print("Actualizando 'Comercial' desde MATRIZ...")
            df_matriz = cargar_datos_excel(self.selected_excel_file, SHEET_MATRIZ)
            if df_matriz is None: messagebox.showwarning("Advertencia Mapeo", f"No se pudo cargar '{SHEET_MATRIZ}'.")
            elif MATRIZ_NIT_COL not in df_matriz.columns or MATRIZ_COMERCIAL_COL not in df_matriz.columns: messagebox.showwarning("Advertencia Mapeo", f"Faltan '{MATRIZ_NIT_COL}' o '{MATRIZ_COMERCIAL_COL}' en '{SHEET_MATRIZ}'.")
            elif 'Nit Cliente' not in self.df_original.columns: messagebox.showwarning("Advertencia Mapeo", f"Falta 'Nit Cliente'.")
            elif COMERCIAL_COLUMNA not in self.df_original.columns:
                 messagebox.showwarning("Advertencia Mapeo", f"Falta '{COMERCIAL_COLUMNA}'. Se intentará crear."); self.df_original[COMERCIAL_COLUMNA] = pd.NA
                 try:
                    df_matriz[MATRIZ_NIT_COL] = df_matriz[MATRIZ_NIT_COL].astype(str).str.strip()
                    self.df_original['Nit Cliente'] = self.df_original['Nit Cliente'].astype(str).str.strip()
                    df_matriz_unique = df_matriz.dropna(subset=[MATRIZ_NIT_COL, MATRIZ_COMERCIAL_COL]).drop_duplicates(subset=[MATRIZ_NIT_COL], keep='first')
                    nit_to_comercial_map = pd.Series(df_matriz_unique[MATRIZ_COMERCIAL_COL].values, index=df_matriz_unique[MATRIZ_NIT_COL]).to_dict()
                    self.df_original[COMERCIAL_COLUMNA] = self.df_original['Nit Cliente'].map(nit_to_comercial_map)
                    num_actualizados = self.df_original[COMERCIAL_COLUMNA].notna().sum()
                    print(f"Mapeo completado. Asignados {num_actualizados} a '{COMERCIAL_COLUMNA}'.")
                    if num_actualizados > 0: messagebox.showinfo("Mapeo Comercial", f"Asignados {num_actualizados} a '{COMERCIAL_COLUMNA}' desde '{SHEET_MATRIZ}'.")
                 except Exception as e: messagebox.showerror("Error Mapeo", f"Error al mapear 'Comercial': {e}")
            else:
                try:
                    df_matriz[MATRIZ_NIT_COL] = df_matriz[MATRIZ_NIT_COL].astype(str).str.strip()
                    self.df_original['Nit Cliente'] = self.df_original['Nit Cliente'].astype(str).str.strip()
                    df_matriz_unique = df_matriz.dropna(subset=[MATRIZ_NIT_COL, MATRIZ_COMERCIAL_COL]).drop_duplicates(subset=[MATRIZ_NIT_COL], keep='first')
                    nit_to_comercial_map = pd.Series(df_matriz_unique[MATRIZ_COMERCIAL_COL].values, index=df_matriz_unique[MATRIZ_NIT_COL]).to_dict()
                    nuevos_comerciales = self.df_original['Nit Cliente'].map(nit_to_comercial_map)
                    comerciales_antes = self.df_original[COMERCIAL_COLUMNA].copy()
                    self.df_original[COMERCIAL_COLUMNA] = nuevos_comerciales.combine_first(self.df_original[COMERCIAL_COLUMNA])
                    actualizados = (self.df_original[COMERCIAL_COLUMNA] != comerciales_antes) & (nuevos_comerciales.notna())
                    num_actualizados = actualizados.sum()
                    print(f"Mapeo completado. '{COMERCIAL_COLUMNA}' actualizados: {num_actualizados}.")
                    if num_actualizados > 0: messagebox.showinfo("Mapeo Comercial", f"'{COMERCIAL_COLUMNA}' actualizado para {num_actualizados} registros.")
                except Exception as e: messagebox.showerror("Error Mapeo", f"Error al actualizar 'Comercial': {e}")
            if CLIENTE_COLUMNA in self.df_original.columns:
                clientes_raw = self.df_original[CLIENTE_COLUMNA].astype(str).unique()
                self.lista_clientes_unicos = sorted([c for c in clientes_raw if c.lower() != 'nan'])
            self.mostrar_todo()
        else: pass

    def aplicar_filtros_vista(self):
        if self.df_original is None: messagebox.showwarning("Datos no Cargados", "Selecciona un archivo."); return
        try: fecha_inicio = pd.to_datetime(self.fecha_inicio_entry.get_date()); fecha_fin = pd.to_datetime(self.fecha_fin_entry.get_date())
        except ValueError: messagebox.showerror("Error", "Formato de fecha inválido."); return
        cliente_buscar = self.cliente_var.get().strip().lower()
        df_filtrado_temp = self.df_original.copy()
        if FECHA_COLUMNA in df_filtrado_temp.columns and pd.api.types.is_datetime64_any_dtype(df_filtrado_temp[FECHA_COLUMNA]):
             mask_fecha = (df_filtrado_temp[FECHA_COLUMNA].notna()) & (df_filtrado_temp[FECHA_COLUMNA] >= fecha_inicio) & (df_filtrado_temp[FECHA_COLUMNA] <= fecha_fin)
             df_filtrado_temp = df_filtrado_temp[mask_fecha]
        if cliente_buscar and CLIENTE_COLUMNA in df_filtrado_temp.columns:
              mask_cliente = df_filtrado_temp[CLIENTE_COLUMNA].notna() & df_filtrado_temp[CLIENTE_COLUMNA].astype(str).str.lower().str.contains(cliente_buscar, na=False)
              df_filtrado_temp = df_filtrado_temp[mask_cliente]
        self.df_filtrado = df_filtrado_temp; self.mostrar_en_treeview()
        messagebox.showinfo("Filtro Aplicado", f"Vista actualizada. Mostrando {len(self.df_filtrado)} registros.")

    def mostrar_todo(self):
        if self.df_original is None: messagebox.showwarning("Datos no Cargados", "Selecciona un archivo."); return
        self.df_filtrado = self.df_original.copy(); self.mostrar_en_treeview()
        self.cliente_var.set(""); self.hide_cliente_listbox()
        messagebox.showinfo("Vista Restaurada", f"Mostrando los {len(self.df_filtrado)} registros originales.")

    def limpiar_treeview(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        self.tree["columns"] = []; self.tree["show"] = "headings"
        if hasattr(self, 'summary_label'):
             self.summary_label.config(text="Selección: 0 filas | Suma Subtotal: 0,00 | Suma Peso CC: 0,00")

    def mostrar_en_treeview(self):
         self.limpiar_treeview()
         if self.df_filtrado is None or self.df_filtrado.empty: return
         columnas_con_coma = {SUBTOTAL_COLUMNA, PESO_COLUMNA, 'Vlr Unit'}
         columnas_mostrar = COLUMNAS_POSIBLES_ESPERADAS if COLUMNAS_POSIBLES_ESPERADAS else list(self.df_filtrado.columns)
         columnas_mostrar_existentes = [col for col in columnas_mostrar if col in self.df_filtrado.columns]
         self.tree["columns"] = columnas_mostrar_existentes; self.tree["show"] = "headings"
         for col in columnas_mostrar_existentes:
             self.tree.heading(col, text=col); col_width = max(len(col)*10, 100)
             self.tree.column(col, anchor=tk.W, width=col_width, stretch=tk.NO)
         self.original_indices.clear()
         for index_filtrado, row in self.df_filtrado.iterrows():
             valores_formateados = []
             for col_name in columnas_mostrar_existentes:
                 val = row[col_name]
                 if pd.isna(val): valores_formateados.append("")
                 elif col_name in columnas_con_coma and isinstance(val, (int, float)):
                     decimals = 2
                     if col_name == PESO_COLUMNA: decimals = 2
                     try: formatted_val = f"{val:.{decimals}f}".replace('.', ',')
                     except (ValueError, TypeError): formatted_val = str(val)
                     valores_formateados.append(formatted_val)
                 elif isinstance(val, (datetime, pd.Timestamp)): valores_formateados.append(val.strftime('%Y/%m/%d'))
                 else: valores_formateados.append(str(val))
             item_id = self.tree.insert("", tk.END, values=valores_formateados)
             self.original_indices[item_id] = index_filtrado
         self.update_selection_summary()

    def update_selection_summary(self, event=None):
        count = 0; total_subtotal = 0.0; total_peso = 0.0
        selected_ids = self.tree.selection()
        if selected_ids and self.df_original is not None:
            original_indices_seleccionados = [self.original_indices[item_id] for item_id in selected_ids if item_id in self.original_indices]
            if original_indices_seleccionados:
                selected_rows = self.df_original.loc[original_indices_seleccionados]
                count = len(selected_rows)
                if SUBTOTAL_COLUMNA in selected_rows.columns:
                    subtotal_numeric = pd.to_numeric(selected_rows[SUBTOTAL_COLUMNA], errors='coerce')
                    total_subtotal = subtotal_numeric.sum(skipna=True)
                if PESO_COLUMNA in selected_rows.columns:
                    peso_numeric = pd.to_numeric(selected_rows[PESO_COLUMNA], errors='coerce')
                    total_peso = peso_numeric.sum(skipna=True)
        try: subtotal_str = f"{total_subtotal:.2f}".replace('.', ',')
        except: subtotal_str = "Error"
        try: peso_str = f"{total_peso:.2f}".replace('.', ',')
        except: peso_str = "Error"
        summary_text = f"Selección: {count} filas | Suma Subtotal: {subtotal_str} | Suma Peso CC: {peso_str}"
        self.summary_label.config(text=summary_text)

    def guardar_sin_mover(self):
        if not self.selected_excel_file: messagebox.showerror("Error", "No hay archivo seleccionado."); return
        if self.df_original is None: messagebox.showerror("Error", "No hay datos cargados."); return
        confirm = messagebox.askyesno("Confirmar Guardado", f"Guardar estado actual en:\n'{os.path.basename(self.selected_excel_file)}'?\n(NO se moverán filas).\n\n¡CERRAR ARCHIVO ANTES!")
        if not confirm: return
        if guardar_cambios_excel_completo(self.selected_excel_file, self.df_original, pd.DataFrame()):
             messagebox.showinfo("Éxito", f"Cambios guardados en:\n'{os.path.basename(self.selected_excel_file)}'.")
        else: messagebox.showerror("Error Guardado", "Error al intentar guardar.")

    def mover_seleccionados(self):
            if not self.selected_excel_file: messagebox.showerror("Error", "No hay archivo seleccionado."); return
            selected_item_ids = self.tree.selection()
            if not selected_item_ids: messagebox.showwarning("Advertencia", "No hay filas seleccionadas."); return
            if self.df_original is None: messagebox.showerror("Error", "No hay datos originales cargados."); return
            confirm = messagebox.askyesno("Confirmar Acción", f"Modificar:\n'{os.path.basename(self.selected_excel_file)}'?\n(Mover {len(selected_item_ids)} filas, etc.)\n¡CERRAR ARCHIVO ANTES!")
            if not confirm: return
            indices_a_mover = [self.original_indices[item_id] for item_id in selected_item_ids if item_id in self.original_indices]
            if not indices_a_mover: messagebox.showerror("Error", "No se identificaron filas."); return
            columnas_para_elim_existentes = [col for col in COLUMNAS_ELIMINADOS if col in self.df_original.columns]
            df_a_mover = self.df_original.loc[indices_a_mover, columnas_para_elim_existentes].copy()
            # --- NUEVO: Añadir Fecha de Eliminación ---
            df_a_mover['Fecha de Eliminación'] = datetime.now().strftime('%Y/%m/%d')
            # --- FIN NUEVO ---
            df_posibles_actualizado = self.df_original.drop(indices_a_mover).copy()
            if guardar_cambios_excel_completo(self.selected_excel_file, df_posibles_actualizado, df_a_mover):
                messagebox.showinfo("Éxito", f"Proceso completado en:\n'{os.path.basename(self.selected_excel_file)}'.")
                self.df_original = df_posibles_actualizado.copy()
                if CLIENTE_COLUMNA in self.df_original.columns:
                    clientes_raw = self.df_original[CLIENTE_COLUMNA].astype(str).unique()
                    self.lista_clientes_unicos = sorted([c for c in clientes_raw if c.lower() != 'nan'])
                else:
                    self.lista_clientes_unicos = []
                self.mostrar_todo()
            else:
                messagebox.showwarning("Advertencia Guardado", "Error al guardar.")
                self.mover_button.config(state='disabled')

# --- Fin de la clase App ---

# --- Ejecutar la aplicación ---
if __name__ == "__main__":
    app = App()
    app.mainloop()