# transformacion_datos.py

def almacenar_estado_original(df):
    """
    Almacena una copia del DataFrame original antes de aplicar transformaciones.
    """
    return df.copy()

def revertir_estado(df, estado_original):
    """
    Restaura el DataFrame al estado original almacenado.
    """
    return estado_original.copy()

def transformar_datos(df, columna, transformacion_funcion):
    """
    Aplica una función de transformación a una columna en el DataFrame.
    """
    try:
        df[columna] = df[columna].apply(transformacion_funcion)
        return df
    except Exception as e:
        print(f"Error al transformar datos: {e}")
        return df
