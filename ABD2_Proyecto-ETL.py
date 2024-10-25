import pyodbc
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import json  # Para manejar el archivo de configuracion

# Funcion para la conexion a la base de datos con servidor y base de datos dinamicos
def connect_to_db(server, database):
    try:
        conn_str = (
            rf"DRIVER={{ODBC Driver 17 for SQL Server}};"
            rf"SERVER={server};"
            rf"DATABASE={database};"
            r"Trusted_Connection=yes;"
        )
        conn = pyodbc.connect(conn_str)
        messagebox.showinfo("Conexion", "Conexion a la base de datos establecida correctamente")
        return conn
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo conectar a la base de datos: {str(e)}")
        return None

class ETLApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema ETL: Excel a SQL Server")
        self.root.geometry("900x600")

        # Inicializa la conexion como None
        self.conn = None
        self.cursor = None
        self.ruta_excel = None
        self.hojas = []
        self.tablas_db = []
        self.config_file = "db_config.json"  # Archivo de configuración

        # Establecer estilos
        self.set_styles()

        # Crear menubar
        self.create_menu()

        # Crear frame principal donde se mostraran las opciones
        self.main_frame = ttk.Frame(self.root, padding=10, style="Main.TFrame")
        self.main_frame.pack(fill='both', expand=True)

        # Mostrar los campos de entrada de servidor y base de datos
        self.display_connection_options()

    def set_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        primary_color = "#2c3e50"
        secondary_color = "#ecf0f1"
        accent_color = "#3498db"
        button_color = "#2980b9"
        bg_color = "#34495e"
        self.style.configure("TFrame", background=primary_color)
        self.style.configure("Main.TFrame", background=bg_color)
        self.style.configure("TLabel", background=primary_color, foreground=secondary_color, font=("Arial", 12))
        self.style.configure("TButton", background=button_color, foreground=secondary_color, font=("Arial", 12, "bold"), padding=10)
        self.style.configure("Accent.TButton", background=accent_color, foreground=secondary_color, font=("Arial", 12, "bold"))
        self.root.option_add("*TButton*highlightThickness", 0)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Menu de ETL (solo accesible si hay conexion)
        etl_menu = tk.Menu(menubar, tearoff=0)
        etl_menu.add_command(label="Conectar Base de Datos", command=self.display_connection_options)
        etl_menu.add_command(label="Seleccionar Excel", command=self.seleccionar_archivo_excel)
        etl_menu.add_command(label="Configurar y Crear Tabla", command=self.configurar_tabla)
        etl_menu.add_command(label="Iniciar ETL", command=self.iniciar_etl)
        menubar.add_cascade(label="ETL", menu=etl_menu)

        # Menu de ayuda
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Acerca de", command=self.about_message)
        menubar.add_cascade(label="Ayuda", menu=help_menu)

    def display_connection_options(self):
        self.clear_frame()

        # Cargar credenciales si estan guardadas
        config = self.cargar_configuracion()

        ttk.Label(self.main_frame, text="Servidor:", style="TLabel").pack(pady=5)
        self.server_entry = ttk.Entry(self.main_frame)
        self.server_entry.insert(0, config.get("server", ""))  # Prellenar con valor guardado si existe
        self.server_entry.pack(fill='x', padx=20, pady=5)

        ttk.Label(self.main_frame, text="Base de Datos:", style="TLabel").pack(pady=5)
        self.database_entry = ttk.Entry(self.main_frame)
        self.database_entry.insert(0, config.get("database", ""))  # Prellenar con valor guardado si existe
        self.database_entry.pack(fill='x', padx=20, pady=5)

        # Botones de conectar y guardar credenciales
        ttk.Button(self.main_frame, text="Conectar", style="Accent.TButton", command=self.connect_db).pack(pady=10)
        ttk.Button(self.main_frame, text="Guardar Credenciales", style="Accent.TButton", command=self.guardar_credenciales).pack(pady=10)

    def cargar_configuracion(self):
        """Carga los parámetros de conexión desde un archivo de configuración JSON."""
        try:
            with open(self.config_file, "r") as file:
                return json.load(file)
        except FileNotFoundError:
            return {}

    def guardar_credenciales(self):
        """Guarda las credenciales en un archivo JSON."""
        config = {
            "server": self.server_entry.get().strip(),
            "database": self.database_entry.get().strip()
        }
        with open(self.config_file, "w") as file:
            json.dump(config, file)
        messagebox.showinfo("Credenciales Guardadas", "Los parametros de conexión han sido guardados correctamente.")

    def seleccionar_archivo_excel(self):
        """Funcion para seleccionar un archivo Excel."""
        self.ruta_excel = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.ruta_excel:
            messagebox.showinfo("Archivo Excel", f"Archivo seleccionado: {self.ruta_excel}")
            # Cargar las hojas disponibles en el Excel
            self.hojas = self.cargar_hojas_excel(self.ruta_excel)
            if self.hojas:
                self.vista_previa_hoja()

    def cargar_hojas_excel(self, ruta):
        """Carga las hojas del archivo Excel."""
        try:
            xls = pd.ExcelFile(ruta)
            return xls.sheet_names
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo Excel: {str(e)}")
            return []

    def vista_previa_hoja(self):
        """Muestra una vista previa de la primera hoja del Excel."""
        if self.hojas:
            df = pd.read_excel(self.ruta_excel, sheet_name=self.hojas[0], nrows=100)  # Leer solo las primeras 5 filas
            vista_previa = df.to_string(index=False)

            # Mostrar la vista previa en un cuadro de texto
            self.clear_frame()
            ttk.Label(self.main_frame, text="Vista previa de la hoja seleccionada:", style="TLabel").pack(pady=5)
            text_area = tk.Text(self.main_frame, height=10, wrap="none")
            text_area.insert("1.0", vista_previa)
            text_area.config(state="disabled")  # Hacer que el texto sea de solo lectura
            text_area.pack(fill="both", padx=20, pady=10)

            ttk.Button(self.main_frame, text="Configurar y Crear Tabla", style="Accent.TButton", command=self.configurar_tabla).pack(pady=10)

    def connect_db(self):
        server = self.server_entry.get().strip()
        database = self.database_entry.get().strip()
        if not server or not database:
            messagebox.showwarning("Advertencia", "Por favor, introduce el servidor y la base de datos.")
            return

        if not self.conn:
            self.conn = connect_to_db(server, database)
            if self.conn:
                self.cursor = self.conn.cursor()

    def configurar_tabla(self):
        """Configura la tabla a partir de la primera fila del Excel y selecciona tipos de datos."""
        if not self.conn or not self.ruta_excel:
            messagebox.showwarning("Advertencia", "Primero debes conectar a la base de datos y seleccionar un archivo Excel.")
            return

        hoja_origen = self.hojas[0]  # Se puede extender para seleccionar otras hojas
        df = pd.read_excel(self.ruta_excel, sheet_name=hoja_origen, nrows=0)  # Solo leer la fila de encabezado
        columnas_excel = df.columns.tolist()

        # Mostrar opciones de vinculacion y tipos de datos
        self.clear_frame()

        ttk.Label(self.main_frame, text=f"Configurar tabla para la hoja '{hoja_origen}':", style="TLabel").pack(pady=5)

        self.table_name_entry = ttk.Entry(self.main_frame)
        self.table_name_entry.insert(0, hoja_origen)  # El nombre de la tabla sera igual al nombre de la hoja
        self.table_name_entry.pack(pady=5)

        self.column_types = []
        self.column_configs_frame = ttk.Frame(self.main_frame)
        self.column_configs_frame.pack(pady=10)

        # Combobox para tipos de datos
        tipos_de_datos = ["INT", "VARCHAR(50)", "FLOAT", "DATETIME", "TEXT"]

        for i, col in enumerate(columnas_excel):
            frame = ttk.Frame(self.column_configs_frame)
            frame.pack(fill="x", padx=10, pady=5)

            ttk.Label(frame, text=col).pack(side="left")

            combobox = ttk.Combobox(frame, values=tipos_de_datos, width=20)
            combobox.pack(side="right")
            self.column_types.append((col, combobox))

        ttk.Button(self.main_frame, text="Crear Tabla en SQL", style="Accent.TButton", command=self.crear_tabla_sql).pack(pady=10)

    def crear_tabla_sql(self):
        """Crea la tabla en SQL Server con los tipos de datos seleccionados."""
        table_name = self.table_name_entry.get().strip()
        if not table_name:
            messagebox.showwarning("Advertencia", "Debes proporcionar un nombre para la tabla.")
            return

        columnas_sql = []
        for col, combobox in self.column_types:
            tipo_dato = combobox.get()
            if not tipo_dato:
                messagebox.showwarning("Advertencia", f"Selecciona un tipo de dato para la columna '{col}'.")
                return
            columnas_sql.append(f"{col} {tipo_dato}")

        columnas_sql_str = ", ".join(columnas_sql)

        # Crear tabla en SQL Server
        try:
            create_table_sql = f"CREATE TABLE {table_name} ({columnas_sql_str})"
            self.cursor.execute(create_table_sql)
            self.conn.commit()
            messagebox.showinfo("Tabla Creada", f"Tabla '{table_name}' creada con exito.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear la tabla: {str(e)}")

    def iniciar_etl(self):
        """Funcion para iniciar el proceso ETL."""
        if not self.conn or not self.ruta_excel:
            messagebox.showwarning("Advertencia", "Primero debes conectar a la base de datos y seleccionar un archivo Excel.")
            return

        hoja_origen = self.hojas[0]
        df = pd.read_excel(self.ruta_excel, sheet_name=hoja_origen)

        # Insertar datos en la tabla recien creada
        self.insertar_datos(df, hoja_origen)

    def insertar_datos(self, df, tabla_destino):
        """Inserta los datos del Excel en la tabla SQL recien creada."""
        cursor = self.conn.cursor()

        columnas = ', '.join(df.columns)
        valores = ', '.join(['?' for _ in df.columns])

        for _, row in df.iterrows():
            data = tuple(row)
            sql_insert = f"INSERT INTO {tabla_destino} ({columnas}) VALUES ({valores})"
            cursor.execute(sql_insert, data)

        self.conn.commit()
        messagebox.showinfo("Proceso completado", "Datos insertados correctamente en la tabla.")

    def about_message(self):
        messagebox.showinfo("Acerca de", "Sistema ETL desarrollado para la migracion de datos de Excel a SQL Server.")

    def clear_frame(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = ETLApp(root)
    root.mainloop()