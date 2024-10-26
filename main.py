from conexion_db import connect_to_db, guardar_credenciales, cargar_credenciales
from carga_excel import cargar_excel, seleccionar_hoja
from limpieza_datos import eliminar_duplicados, eliminar_nulos
from transformacion_datos import transformar_datos, almacenar_estado_original, revertir_estado
from carga_sql import crear_tabla, cargar_datos_sql
from cargar_exclusiones import guardar_exclusiones
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox

class ETLApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema ETL: Excel a SQL Server")

        # Variables para conexión y datos
        self.server = tk.StringVar()
        self.database = tk.StringVar()
        self.excel_path = ""
        self.sheet_name = tk.StringVar()
        self.df = None
        self.estado_original = None  # Variable para almacenar el estado original

        # Crear interfaz de conexión
        self.create_connection_frame()

        # Crear botones para cada etapa del ETL
        self.create_buttons()

    def create_connection_frame(self):
        connection_frame = ttk.LabelFrame(self.root, text="Conexión a la Base de Datos")
        connection_frame.pack(padx=10, pady=10, fill="x")

        ttk.Label(connection_frame, text="Servidor:").grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(connection_frame, textvariable=self.server).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(connection_frame, text="Base de Datos:").grid(row=1, column=0, padx=5, pady=5)
        ttk.Entry(connection_frame, textvariable=self.database).grid(row=1, column=1, padx=5, pady=5)

        ttk.Button(connection_frame, text="Conectar", command=self.connect_db).grid(row=2, column=0, columnspan=2, pady=10)

    def create_buttons(self):
        button_frame = ttk.LabelFrame(self.root, text="Operaciones ETL")
        button_frame.pack(padx=10, pady=10, fill="x")

        ttk.Button(button_frame, text="Cargar Archivo Excel", command=self.load_excel).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(button_frame, text="Seleccionar Hoja", command=self.select_sheet).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(button_frame, text="Eliminar Duplicados", command=self.remove_duplicates).grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(button_frame, text="Eliminar Nulos", command=self.remove_nulls).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(button_frame, text="Transformar Datos", command=self.transform_data).grid(row=2, column=0, padx=5, pady=5)
        ttk.Button(button_frame, text="Revertir Cambios", command=self.revert_changes).grid(row=2, column=1, padx=5, pady=5)  # Nuevo botón para revertir
        ttk.Button(button_frame, text="Crear Tabla SQL", command=self.create_table).grid(row=3, column=0, padx=5, pady=5)
        ttk.Button(button_frame, text="Cargar Datos a SQL", command=self.load_data_to_sql).grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(button_frame, text="Guardar Exclusiones", command=self.save_exclusions).grid(row=4, column=0, columnspan=2, padx=5, pady=5)

    # Funciones para cada botón

    def connect_db(self):
        server = self.server.get()
        database = self.database.get()
        self.conn = connect_to_db(server, database)
        if self.conn:
            guardar_credenciales(server, database)
            messagebox.showinfo("Éxito", "Conexión establecida correctamente")

    def load_excel(self):
        self.xls = cargar_excel()
        if self.xls:
            messagebox.showinfo("Éxito", "Archivo Excel cargado correctamente")

    def select_sheet(self):
        if self.xls:
            sheet_name = self.sheet_name.get() or self.xls.sheet_names[0]
            self.df = seleccionar_hoja(self.xls, sheet_name)
            if self.df is not None:
                self.estado_original = almacenar_estado_original(self.df)  # Almacena el estado original
                messagebox.showinfo("Éxito", f"Hoja '{sheet_name}' cargada correctamente")

    def remove_duplicates(self):
        if self.df is not None:
            self.df = eliminar_duplicados(self.df, "ID")
            messagebox.showinfo("Éxito", "Duplicados eliminados correctamente")

    def remove_nulls(self):
        if self.df is not None:
            self.df = eliminar_nulos(self.df, ["Columna1", "Columna2"])
            messagebox.showinfo("Éxito", "Valores nulos eliminados correctamente")

    def transform_data(self):
        if self.df is not None:
            self.df = transformar_datos(self.df, "Columna1", lambda x: x.upper())
            messagebox.showinfo("Éxito", "Datos transformados correctamente")

    def revert_changes(self):
        """
        Revertir al estado original del DataFrame.
        """
        if self.estado_original is not None:
            self.df = revertir_estado(self.df, self.estado_original)
            messagebox.showinfo("Éxito", "Cambios revertidos al estado original")

    def create_table(self):
        if self.df is not None and self.conn:
            crear_tabla(self.conn, "nombre_tabla", self.df.columns)
            messagebox.showinfo("Éxito", "Tabla SQL creada correctamente")

    def load_data_to_sql(self):
        if self.df is not None and self.conn:
            cargar_datos_sql(self.df, self.conn, "nombre_tabla")
            messagebox.showinfo("Éxito", "Datos cargados en SQL Server correctamente")

    def save_exclusions(self):
        exclusiones = self.df[self.df.isnull().any(axis=1)]
        guardar_exclusiones(exclusiones)
        messagebox.showinfo("Éxito", "Archivo de exclusiones guardado correctamente")

# Ejecución de la aplicación
if __name__ == "__main__":
    root = tk.Tk()
    app = ETLApp(root)
    root.mainloop()
