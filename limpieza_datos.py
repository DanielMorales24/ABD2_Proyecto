def eliminar_duplicados(df, columna_id):
    """
    Elimina duplicados en el DataFrame basado en una columna específica.
    """
    df_clean = df.drop_duplicates(subset=[columna_id])
    return df_clean

def eliminar_nulos(df, columnas):
    """
    Elimina filas con valores nulos en las columnas especificadas.
    """
    df_clean = df.dropna(subset=columnas)
    return df_clean
