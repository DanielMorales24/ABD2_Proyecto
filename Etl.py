import tkinter as tk  # Biblioteca para crear interfaces gráficas (GUI)
from tkinter import filedialog, messagebox, ttk  # Herramientas de Tkinter para diálogos de archivos, mensajes y widgets
import pandas as pd  # Biblioteca para manipulación de datos
import pyodbc  # Biblioteca para conectarse a bases de datos ODBC, como SQL Server

# Variables globales para almacenar el DataFrame actual y el original
df = None
df_original = None

# Función para cargar el archivo Excel y seleccionar una hoja
def cargar_excel():
    global df, df_original
    archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])  # Abre diálogo para seleccionar un archivo de Excel
    if archivo:
        try:
            # Obtener las hojas disponibles en el archivo de Excel
            hojas = pd.ExcelFile(archivo).sheet_names
            
            # Crear una ventana para seleccionar la hoja a cargar
            hoja_ventana = tk.Toplevel(root)
            hoja_ventana.title("Selecciona la Hoja")
            tk.Label(hoja_ventana, text="Selecciona la hoja que quieres cargar:").pack(pady=5)
            
            # Crear lista desplegable para elegir la hoja
            var_hoja = tk.StringVar(value=hojas[0])  # Variable que guarda la hoja seleccionada
            hoja_menu = tk.OptionMenu(hoja_ventana, var_hoja, *hojas)
            hoja_menu.pack(pady=5)
            
            # Función para cargar los datos de la hoja seleccionada
            def cargar_hoja():
                global df, df_original
                hoja_seleccionada = var_hoja.get()  # Obtener la hoja seleccionada
                df = pd.read_excel(archivo, sheet_name=hoja_seleccionada)  # Cargar los datos de la hoja
                df_original = df.copy()  # Guardar una copia original del DataFrame
                actualizar_tabla(df)  # Actualizar la vista previa de los datos
                actualizar_columnas(df.columns)  # Actualizar la lista de columnas
                hoja_ventana.destroy()  # Cerrar la ventana de selección de hoja
            
            # Botón para confirmar la selección de la hoja
            btn_confirmar = tk.Button(hoja_ventana, text="Cargar Hoja", command=cargar_hoja)
            btn_confirmar.pack(pady=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error cargando archivo: {e}")  # Mostrar error si no se puede cargar el archivo

# Función para actualizar las columnas seleccionables
def actualizar_columnas(columnas):
    # Limpiar los checkboxes previos
    for widget in frame_columnas.winfo_children():
        widget.destroy()

    global checkbox_vars
    checkbox_vars = {}
    
    # Crear un checkbox para cada columna
    for col in columnas:
        var = tk.BooleanVar(value=True)  # Variable para guardar el estado del checkbox (True = seleccionado)
        checkbox = tk.Checkbutton(frame_columnas, text=col, variable=var)  # Checkbox para cada columna
        checkbox.pack(anchor='w')  # Colocar el checkbox en el frame
        checkbox_vars[col] = var  # Guardar la variable de estado para cada columna

# Función para actualizar la tabla de vista previa
def actualizar_tabla(dataframe):
    # Limpiar la tabla actual
    for row in tree.get_children():
        tree.delete(row)
    
    # Configurar las columnas de la tabla
    tree["columns"] = list(dataframe.columns)  # Asignar las columnas del DataFrame
    tree["show"] = "headings"  # Mostrar solo las cabeceras

    # Asignar los encabezados de las columnas
    for col in dataframe.columns:
        tree.heading(col, text=col)
        tree.column(col, minwidth=100, width=150, stretch=tk.NO)  # Configurar el ancho de cada columna

    # Insertar los datos del DataFrame en la tabla
    for index, row in dataframe.iterrows():
        tree.insert("", "end", values=list(row))

# Función para aplicar transformaciones en los datos
def transformar_datos():
    global df
    if df is not None:
        # Filtrar las columnas seleccionadas
        columnas_seleccionadas = [col for col, var in checkbox_vars.items() if var.get()]  # Solo columnas con checkbox seleccionado
        df = df[columnas_seleccionadas]  # Filtrar el DataFrame según las columnas seleccionadas

        # Aplicar las opciones de transformación
        if var_limpiar_nulos.get() == "Eliminar Filas":
            df.dropna(inplace=True)  # Eliminar filas con valores nulos
        elif var_limpiar_nulos.get() == "Rellenar con Cero":
            df.fillna(0, inplace=True)  # Rellenar valores nulos con 0
        elif var_limpiar_nulos.get() == "Rellenar con Vacío":
            df.fillna('', inplace=True)  # Rellenar valores nulos con cadenas vacías
        
        # Actualizar la vista previa de los datos
        actualizar_tabla(df)
    else:
        messagebox.showwarning("Advertencia", "No hay datos cargados")  # Mostrar advertencia si no hay datos cargados

# Función para deshacer cambios y restaurar el DataFrame original
def deshacer_cambios():
    global df, df_original
    if df_original is not None:
        df = df_original.copy()  # Restaurar el DataFrame original
        actualizar_tabla(df)  # Actualizar la vista previa con el DataFrame original
        messagebox.showinfo("Deshacer", "Los cambios han sido revertidos al estado original.")
    else:
        messagebox.showwarning("Advertencia", "No hay cambios para deshacer.")

# Función para cargar los datos en SQL Server
def cargar_sql_server():
    global df
    if df is not None:
        try:
            # Conectar a SQL Server
            conn = pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=;'#Nombre del servidor
                'DATABASE=y;'#Nombre de la base de datos
                
            )
            cursor = conn.cursor()

            # Crear tabla en SQL Server
            table_name = 'tu_tabla_etl'
            cursor.execute(f"IF OBJECT_ID('{table_name}', 'U') IS NOT NULL DROP TABLE {table_name}")  # Eliminar la tabla si ya existe
            columnas = ", ".join([f"{col} NVARCHAR(MAX)" for col in df.columns])  # Crear columnas de la tabla
            cursor.execute(f"CREATE TABLE {table_name} ({columnas})")

            # Insertar los datos del DataFrame en la tabla
            for index, row in df.iterrows():
                valores = "', '".join([str(val).replace("'", "''") for val in row])
                cursor.execute(f"INSERT INTO {table_name} VALUES ('{valores}')")

            conn.commit()  # Confirmar los cambios en la base de datos
            messagebox.showinfo("Éxito", "Datos cargados exitosamente en SQL Server.")
        except Exception as e:
            messagebox.showerror("Error", f"Error cargando en SQL Server: {e}")  # Mostrar mensaje de error si falla la carga
    else:
        messagebox.showwarning("Advertencia", "No hay datos para cargar")  # Mostrar advertencia si no hay datos cargados

# Configuración de la ventana principal de la interfaz gráfica
root = tk.Tk()
root.title("Simulador ETL con Transformación de Datos")
root.geometry("1200x700")  # Dimensiones de la ventana

# Frame para opciones de transformación y selección de archivo
frame_opciones = tk.Frame(root)
frame_opciones.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

# Botón para cargar un archivo de Excel
btn_cargar_excel = tk.Button(frame_opciones, text="Cargar Excel", command=cargar_excel)
btn_cargar_excel.pack(side=tk.LEFT, padx=5)

# Botón para aplicar las transformaciones
btn_transformar = tk.Button(frame_opciones, text="Transformar Datos", command=transformar_datos)
btn_transformar.pack(side=tk.LEFT, padx=5)

# Botón para deshacer los cambios realizados
btn_deshacer = tk.Button(frame_opciones, text="Deshacer Cambios", command=deshacer_cambios)
btn_deshacer.pack(side=tk.LEFT, padx=5)

# Botón para cargar los datos en SQL Server
btn_cargar_sql = tk.Button(frame_opciones, text="Cargar en SQL Server", command=cargar_sql_server)
btn_cargar_sql.pack(side=tk.LEFT, padx=5)

# Opciones para la limpieza de valores nulos en los datos
var_limpiar_nulos = tk.StringVar(value="Ninguno")
tk.Label(frame_opciones, text="Limpieza de Nulos:").pack(side=tk.LEFT, padx=5)
tk.OptionMenu(frame_opciones, var_limpiar_nulos, "Ninguno", "Eliminar Filas", "Rellenar con Cero", "Rellenar con Vacío").pack(side=tk.LEFT, padx=5)

# Frame para seleccionar las columnas a transformar
frame_columnas = tk.Frame(root)
frame_columnas.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
tk.Label(frame_columnas, text="Selecciona las columnas:").pack()

# Frame para la tabla de vista previa
frame_tabla = tk.Frame(root)
frame_tabla.pack(fill=tk.BOTH, expand=True)

# Tabla de vista previa de los datos
tree = ttk.Treeview(frame_tabla)
tree.pack(fill=tk.BOTH, expand=True)

# Barra de desplazamiento vertical para la tabla
scroll_y = tk.Scrollbar(frame_tabla, orient="vertical", command=tree.yview)
scroll_y.pack(side=tk.RIGHT, fill="y")
tree.configure(yscrollcommand=scroll_y.set)

# Barra de desplazamiento horizontal para la tabla
scroll_x = tk.Scrollbar(frame_tabla, orient="horizontal", command=tree.xview)
scroll_x.pack(side=tk.BOTTOM, fill="x")
tree.configure(xscrollcommand=scroll_x.set)

# Iniciar el loop de la interfaz gráfica
root.mainloop()
