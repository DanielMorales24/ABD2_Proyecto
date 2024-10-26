import pandas as pd

def guardar_exclusiones(df_exclusiones, ruta_archivo='exclusiones.xlsx'):
    """
    Guarda las filas excluidas en un archivo Excel.
    """
    try:
        df_exclusiones.to_excel(ruta_archivo, index=False)
        print(f"Archivo de exclusiones guardado en {ruta_archivo}")
    except Exception as e:
        print(f"Error al guardar exclusiones: {e}")
