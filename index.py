import pandas as pd
import glob
import os

# Ruta de los archivos CSV
ruta_archivos_csv = "C:/Users/radea/OneDrive/Escritorio/Hugo/Input/*.csv"
archivos_csv = glob.glob(ruta_archivos_csv)

# Función para leer un CSV con múltiples intentos de codificación
def leer_csv_con_codificacion(archivo):
    for encoding in ['utf-8', 'latin1', 'iso-8859-1']:
        try:
            return pd.read_csv(archivo, encoding=encoding, dtype=str)
        except UnicodeDecodeError as e:
            print(f"Error al leer {archivo} con codificación {encoding}: {e}")
        except Exception as e:
            print(f"Otro error al leer {archivo} con codificación {encoding}: {e}")
    return None

# Función para normalizar los nombres de las columnas
def normalizar_columnas(df):
    columnas_renombradas = {
        'casilla': 'Casilla',
        'seccional': 'Sección',
        'consecutivo': 'Consecutivo',
        'nombre': 'Nombre (s)',
        'apellido paterno': 'Apellido paterno',
        'apellido materno': 'Apellido materno',
        'Nombre (s)': 'Nombre (s)',
        'Apellido paterno': 'Apellido paterno',
        'Apellido materno': 'Apellido materno'
    }
    df.rename(columns=columnas_renombradas, inplace=True)
    return df

# Convertir cada archivo CSV a Excel
for archivo_csv in archivos_csv:
    df = leer_csv_con_codificacion(archivo_csv)
    if df is not None:
        df = normalizar_columnas(df)
        # Limitar el tamaño de los archivos y limpiar datos si es necesario
        if len(df) > 100000:  # Por ejemplo, dividir archivos grandes en partes
            for i in range(0, len(df), 100000):
                parte_df = df.iloc[i:i+100000]
                nombre_archivo_excel = f"{os.path.splitext(archivo_csv)[0]}_parte_{i//100000}.xlsx"
                parte_df.to_excel(nombre_archivo_excel, index=False, engine='xlsxwriter')
                print(f"Convertido {archivo_csv} a {nombre_archivo_excel}")
        else:
            nombre_archivo_excel = os.path.splitext(archivo_csv)[0] + '.xlsx'
            df.to_excel(nombre_archivo_excel, index=False, engine='xlsxwriter')
            print(f"Convertido {archivo_csv} a {nombre_archivo_excel}")
    else:
        print(f"No se pudo leer el archivo {archivo_csv} con ninguna codificación.")
