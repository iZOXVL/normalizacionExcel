import pandas as pd
import glob
import os

# Ruta del archivo principal
archivo_principal = "C:/Users/radea/OneDrive/Escritorio/Hugo/Normal/seccionales_reforma_2024.xlsx"

# Verificar que el archivo principal exista
if not os.path.exists(archivo_principal):
    raise FileNotFoundError(f"No se encontró el archivo principal: {archivo_principal}")

df_principal = pd.read_excel(archivo_principal)

# Ruta de los archivos adicionales
ruta_archivos_adicionales = "C:/Users/radea/OneDrive/Escritorio/Hugo/Input/*.xlsx"

# Verificar la ruta especificada
print(f"Buscando archivos en la ruta: {ruta_archivos_adicionales}")

# Leer los archivos adicionales
archivos_adicionales = glob.glob(ruta_archivos_adicionales)

# Verificar si se encontraron archivos adicionales
if not archivos_adicionales:
    raise ValueError("No se encontraron archivos adicionales en la ruta especificada.")

# Imprimir los archivos encontrados para depuración
print("Archivos adicionales encontrados:", archivos_adicionales)

# Unir todos los archivos adicionales en un solo DataFrame
dataframes_adicionales = []
for archivo in archivos_adicionales:
    df_adicional = pd.read_excel(archivo)
    if not df_adicional.empty and not df_adicional.isna().all().all():  # Excluir DataFrames vacíos o con todos los valores NA
        dataframes_adicionales.append(df_adicional)

# Verificar si se pudieron leer archivos adicionales
if not dataframes_adicionales:
    raise ValueError("No se pudieron leer archivos adicionales correctamente.")

df_adicional = pd.concat(dataframes_adicionales, ignore_index=True)

# Verificar que las columnas necesarias existan en ambos DataFrames
columnas_necesarias = ['Nombre (s)', 'Apellido paterno', 'Apellido materno', 'Sección', 'Casilla', 'Consecutivo']
for columna in columnas_necesarias:
    if columna not in df_adicional.columns:
        raise ValueError(f"Falta la columna '{columna}' en los archivos adicionales.")
        
if 'NOMBRE' not in df_principal.columns or 'APELLIDO PATERNO' not in df_principal.columns or 'APELLIDO MATERNO' not in df_principal.columns:
    raise ValueError("Faltan columnas necesarias en el archivo principal.")

# Normalizar nombres y apellidos para evitar problemas de comparación
def normalizar_cadena(cadena):
    if pd.isna(cadena):
        return ''
    return str(cadena).strip().lower()

df_principal['NOMBRE_NORMALIZADO'] = df_principal['NOMBRE'].apply(normalizar_cadena)
df_principal['APELLIDO_PATERNO_NORMALIZADO'] = df_principal['APELLIDO PATERNO'].apply(normalizar_cadena)
df_adicional['NOMBRE_NORMALIZADO'] = df_adicional['Nombre (s)'].apply(normalizar_cadena)
df_adicional['APELLIDO_PATERNO_NORMALIZADO'] = df_adicional['Apellido paterno'].apply(normalizar_cadena)

# Eliminar duplicados en df_adicional
df_adicional = df_adicional.drop_duplicates(subset=['NOMBRE_NORMALIZADO', 'APELLIDO_PATERNO_NORMALIZADO'])

# Crear un diccionario con los datos adicionales
diccionario_datos_adicionales = df_adicional.set_index(['NOMBRE_NORMALIZADO', 'APELLIDO_PATERNO_NORMALIZADO'])[['Sección', 'Casilla', 'Consecutivo']].to_dict('index')

# Función para completar los datos faltantes
def completar_datos(row):
    nombre = row['NOMBRE_NORMALIZADO']
    apellido_paterno = row['APELLIDO_PATERNO_NORMALIZADO']
    clave = (nombre, apellido_paterno)
    if pd.isna(row['SECCION']) or pd.isna(row['CASILLA']) or pd.isna(row['CONSECUTIVO']):
        if clave in diccionario_datos_adicionales:
            datos = diccionario_datos_adicionales[clave]
            row['SECCION'] = datos['Sección']
            row['CASILLA'] = datos['Casilla']
            row['CONSECUTIVO'] = datos['Consecutivo']
    return row

# Aplicar la función a cada fila del DataFrame principal
df_principal = df_principal.apply(completar_datos, axis=1)

# Eliminar las columnas de nombres normalizados
df_principal.drop(columns=['NOMBRE_NORMALIZADO', 'APELLIDO_PATERNO_NORMALIZADO'], inplace=True)

# Guardar el DataFrame principal actualizado en un nuevo archivo
df_principal.to_excel("C:/Users/radea/OneDrive/Escritorio/Hugo/Output/archivo_principal_actualizado.xlsx", index=False)

print("Archivo principal actualizado correctamente.")
