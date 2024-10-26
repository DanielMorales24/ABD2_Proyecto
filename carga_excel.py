import pandas as pd
from tkinter import filedialog
from tkinter import messagebox

def cargar_excel():
    """
    Abre un diálogo para seleccionar el archivo Excel y lo carga.
    """
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
        return None

    try:
        xls = pd.ExcelFile(file_path)
        return xls
    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar el archivo: {e}")
        return None

def seleccionar_hoja(xls, nombre_hoja):
    """
    Selecciona la hoja especificada de un archivo Excel.
    """
    try:
        df = pd.read_excel(xls, sheet_name=nombre_hoja)
        return df
    except Exception as e:
        print(f"Error al seleccionar la hoja: {e}")
        return None
