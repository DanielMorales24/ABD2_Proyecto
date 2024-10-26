import pyodbc
from tkinter import messagebox
import json

def connect_to_db(server, database):
    """
    Conecta a SQL Server usando los parámetros de servidor y base de datos.
    """
    try:
        conn_str = (
            rf"DRIVER={{ODBC Driver 17 for SQL Server}};"
            rf"SERVER={server};"
            rf"DATABASE={database};"
            rf"Trusted_Connection=yes;"
        )
        conn = pyodbc.connect(conn_str)
        messagebox.showinfo("Conexión", "Conexión a la base de datos establecida correctamente")
        return conn
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo conectar a la base de datos: {str(e)}")
        return None

def guardar_credenciales(server, database, file_path='config.json'):
    """
    Guarda las credenciales de conexión en un archivo JSON.
    """
    try:
        config = {'server': server, 'database': database}
        with open(file_path, 'w') as f:
            json.dump(config, f)
        print("Credenciales guardadas en config.json")
    except Exception as e:
        print(f"Error al guardar credenciales: {e}")

def cargar_credenciales(file_path='config.json'):
    """
    Carga las credenciales de conexión desde un archivo JSON.
    """
    try:
        with open(file_path, 'r') as f:
            config = json.load(f)
        return config['server'], config['database']
    except Exception as e:
        print(f"Error al cargar credenciales: {e}")
        return None, None
