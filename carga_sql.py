def crear_tabla(conexion, nombre_tabla, columnas):
    """
    Crea una tabla en SQL Server con el nombre y columnas especificadas.
    """
    cursor = conexion.cursor()
    columnas_str = ", ".join([f"{columna} NVARCHAR(MAX)" for columna in columnas])  # Modificar si necesitas otros tipos de datos
    query = f"CREATE TABLE {nombre_tabla} ({columnas_str})"
    
    try:
        cursor.execute(query)
        conexion.commit()
        print(f"Tabla '{nombre_tabla}' creada exitosamente.")
    except Exception as e:
        print(f"Error al crear la tabla: {e}")

def cargar_datos_sql(df, conexion, nombre_tabla):
    """
    Inserta los datos de un DataFrame en la tabla SQL Server especificada.
    """
    cursor = conexion.cursor()
    columnas = ", ".join(df.columns)
    placeholders = ", ".join(["?" for _ in df.columns])
    query = f"INSERT INTO {nombre_tabla} ({columnas}) VALUES ({placeholders})"
    
    for _, row in df.iterrows():
        try:
            cursor.execute(query, tuple(row))
        except Exception as e:
            print(f"Error al insertar datos: {e}")
    
    conexion.commit()
    print("Datos insertados exitosamente.")
