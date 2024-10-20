import os
import pandas as pd
from datetime import datetime

def buscar_archivo_reciente(directorio, palabra_clave):
    try:
        archivos = [f for f in os.listdir(directorio) if palabra_clave in f and (f.endswith('.xls') or f.endswith('.xlsx'))]
        archivo_reciente = None
        fecha_reciente = datetime.min

        if not archivos:
            print(f"No se encontraron archivos que contengan '{palabra_clave}' en el directorio {directorio}.")
            return None

        for archivo in archivos:
            ruta_completa = os.path.join(directorio, archivo)
            fecha_modificacion = datetime.fromtimestamp(os.path.getmtime(ruta_completa))
            if fecha_modificacion > fecha_reciente:
                fecha_reciente = fecha_modificacion
                archivo_reciente = ruta_completa

        return archivo_reciente
    except Exception as e:
        print(f"Error al buscar archivos: {e}")
        return None

def buscar_matricula_en_archivo(archivo, matricula):
    try:
        extension = archivo.split('.')[-1]
        if extension == 'xlsx':
            print(f"Abriendo el archivo .xlsx con openpyxl: {archivo}")
            df = pd.read_excel(archivo, engine='openpyxl')
        elif extension == 'xls':
            print(f"Abriendo el archivo .xls con xlrd: {archivo}")
            df = pd.read_excel(archivo, engine='xlrd')
        else:
            print(f"Formato de archivo no soportado: {archivo}")
            return None

        if df.empty:
            print("El archivo está vacío.")
            return None

        # Convertimos la columna B a cadena y luego verificamos si la parte numérica está contenida en alguna celda
        df['ColumnaB'] = df.iloc[:, 1].astype(str)  # Columna B en index 1

        # Busca si la parte numérica está contenida en la columna B
        fila_encontrada = df[df['ColumnaB'].str.contains(matricula)]

        if not fila_encontrada.empty:
            # Extraer los datos de las columnas correspondientes
            return {
                'nombre_ep': fila_encontrada.iloc[0, 0],  # Columna A
                'armador': fila_encontrada.iloc[0, 22],  # Columna W (23 en 0-based index)
                'permiso': fila_encontrada.iloc[0, 16],  # Columna Q (17 en 0-based index)
                'estado_permiso': fila_encontrada.iloc[0, 18],  # Columna S (19 en 0-based index)
                'nominacion': fila_encontrada.iloc[0, 32],  # Columna AG (33 en 0-based index)
                'bodega_m3': fila_encontrada.iloc[0, 14],  # Columna O (15 en 0-based index)
                'bodega_tm': fila_encontrada.iloc[0, 15],  # Columna P (16 en 0-based index)
                'baliza': fila_encontrada.iloc[0, 21],  # Columna V (22 en 0-based index)
                'precinto': fila_encontrada.iloc[0, 24]  # Columna Y (25 en 0-based index)
            }
        else:
            print(f"La matrícula {matricula} no se encontró en el archivo.")
            return None
    except Exception as e:
        print(f"Error al leer o procesar el archivo: {e}")
        return None

# Configuración
directorio = "C:/Users/Alan/Desktop/Intertek/Herramientas/Verificar EP/listado/"  # Cambia esta ruta a la carpeta donde están los archivos Excel
palabra_clave = "allvessels"

# Pedir la matrícula al usuario
matricula = input("Ingresa la parte numérica de la matrícula: ")

# Buscar el archivo más reciente
archivo_reciente = buscar_archivo_reciente(directorio, palabra_clave)

if archivo_reciente:
    # Buscar los datos en el archivo
    datos = buscar_matricula_en_archivo(archivo_reciente, matricula)

    if datos:
        # Mostrar los datos encontrados
        print(f"Datos para la matrícula {matricula}:")
        print(f"Nombre EP: {datos['nombre_ep']}")
        print(f"Armador: {datos['armador']}")
        print(f"Permiso: {datos['permiso']}")
        print(f"Estado de Permiso: {datos['estado_permiso']}")
        print(f"Nominación: {datos['nominacion']}")
        print(f"Bodega m3: {datos['bodega_m3']}")
        print(f"Bodega tm: {datos['bodega_tm']}")
        print(f"Baliza: {datos['baliza']}")
        print(f"Precinto: {datos['precinto']}")
else:
    print(f"No se encontró ningún archivo que contenga '{palabra_clave}' en el nombre.")

# Evitar que la ventana se cierre automáticamente
input("\nPresiona Enter para salir...")
