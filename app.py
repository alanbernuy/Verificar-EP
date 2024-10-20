from flask import Flask, render_template, request
import os
import pandas as pd
from datetime import datetime

app = Flask(__name__)

def buscar_archivo_reciente(directorio, palabra_clave):
    try:
        archivos = [f for f in os.listdir(directorio) if palabra_clave in f and (f.endswith('.xls') or f.endswith('.xlsx'))]
        archivo_reciente = None
        fecha_reciente = datetime.min

        if not archivos:
            return None

        for archivo in archivos:
            ruta_completa = os.path.join(directorio, archivo)
            fecha_modificacion = datetime.fromtimestamp(os.path.getmtime(ruta_completa))
            if fecha_modificacion > fecha_reciente:
                fecha_reciente = fecha_modificacion
                archivo_reciente = ruta_completa

        return archivo_reciente
    except Exception as e:
        return None

def buscar_matricula_en_archivo(archivo, matricula):
    try:
        extension = archivo.split('.')[-1]
        if extension == 'xlsx':
            df = pd.read_excel(archivo, engine='openpyxl')
        elif extension == 'xls':
            df = pd.read_excel(archivo, engine='xlrd')
        else:
            return None

        if df.empty:
            return None

        df['ColumnaB'] = df.iloc[:, 1].astype(str)

        fila_encontrada = df[df['ColumnaB'].str.contains(matricula)]

        if not fila_encontrada.empty:
            return {
                'nombre_ep': fila_encontrada.iloc[0, 0],
                'armador': fila_encontrada.iloc[0, 22],
                'permiso': fila_encontrada.iloc[0, 16],
                'estado_permiso': fila_encontrada.iloc[0, 18],
                'nominacion': fila_encontrada.iloc[0, 32],
                'bodega_m3': fila_encontrada.iloc[0, 14],
                'bodega_tm': fila_encontrada.iloc[0, 15],
                'baliza': fila_encontrada.iloc[0, 21],
                'precinto': fila_encontrada.iloc[0, 24]
            }
        else:
            return None
    except Exception as e:
        return None

@app.route('/', methods=['GET', 'POST'])
def index():
    resultado = None
    if request.method == 'POST':
        matricula = request.form['matricula']
        directorio = "C:/Users/Alan/Desktop/Intertek/Herramientas/Verificar EP/listado/"  # Cambia esta ruta a la carpeta donde están tus archivos Excel
        palabra_clave = "allvessels"
        archivo_reciente = buscar_archivo_reciente(directorio, palabra_clave)

        if archivo_reciente:
            resultado = buscar_matricula_en_archivo(archivo_reciente, matricula)

    return render_template('index.html', resultado=resultado)

if __name__ == '__main__':
    app.run(debug=True)
