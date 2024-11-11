import os
import pandas as pd
import datetime as dt
import numpy as np

# Lee el archivo de Excel
ruta_archivo = 'J:\\python\\Depuracion\\Original\\asigM.xlsx'
df = pd.read_excel(ruta_archivo, sheet_name='Base de Cuentas', header=1)

indices_columnas_innecesarias = [6, 7, 8, 10, 21, 22, 23, 24, 25, 26, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 44, 45, 53, 54, 55]

df = df.drop(df.columns[indices_columnas_innecesarias], axis=1)
df = df.drop(index=0)

df.insert(0, 'Orden', range(1, len(df) + 1))
df.insert(1,'Moneda Capital', None)
df.insert(1,'MONTO COSTAS', None)
df.insert(1,'TIPO INTERES', None)


# Encuentra los índices donde las cuentas tienen valores en colones y dolares
indices = df[(df['Saldo pantalla ¢'] > 0) & (df['Saldo pantalla $'] > 0)].index

print(indices)

print("ASIGNACIÓN CONTIENE "+str(len(indices))+"CUENTAS QUE SON EL COLONES Y DOLARES")


# Duplicar los registros donde ambas columnas tienen montos mayores a 0
dolares_colones = []

for idx in indices:
    row = df.iloc[[idx-1]]  # Seleccionar la fila que cumple la condición
    dolares_colones.append(row)

dolares_colones = pd.DataFrame(np.vstack(dolares_colones), columns=df.columns)
dolares_colones['Moneda Capital'] = 'DOLARES'
df = pd.concat([df, dolares_colones], ignore_index=True)    

df = df.sort_values(by='Orden').reset_index(drop=True)
duplicados = df.duplicated(subset=['Orden'], keep=False)

# Asignar valores basados en el índice de duplicación en cada grupo
df.loc[duplicados & (df.groupby('Orden').cumcount() == 0), 'Moneda Capital'] = 'COLONES'
df.loc[duplicados & (df.groupby('Orden').cumcount() == 1), 'Moneda Capital'] = 'DOLARES'

df['Cuenta'] = df['Cuenta'].astype(str)
df['Tarjeta'] = df['Tarjeta'].astype(str)

# Reemplazar valores en la columna 'Estado Civil'
df['Estado Civil'] = df['Estado Civil'].replace({
    'SOLTERA': 'SOLTERO',
    'CASADA': 'CASADO',
    'DIVORCIADA': 'DIVORCIADO',
    'VIUDA': 'VIUDO'
})

# Reemplazar valores en la columna 'Estado Civil'
df['Categoría'] = df['Categoría'].replace({
    'Principal': 'INCOBRABLE',
    'Rel Inc': 'INCOBRABLE',
    'Rel Act': 'ACTIVA'
})

# Reemplazar valores en la columna 'Estado Civil'
df['Tipo de Cuenta'] = df['Tipo de Cuenta'].replace({
    'Con demanda': 'SIN NOTIFICAR',
    'Sin demanda': 'SIN DEMANDA'
})

# Reemplazar valores en la columna 'Estado Civil'
df['Patrimonio'] = df['Patrimonio'].replace({
    'Con Patrimonio': 'CON PATRIMONIO'
})

df.insert(18,'CLIENTE', None)
df['CLIENTE'] = np.where(df['Producto'] == 'Walmart', 'SERVIVALORES', 'CREDOMATIC')


cyber_columns = ['# CYBER 1', '# CYBER 2', '# CYBER 3', '# CYBER 4']

# Aplica la extracción de los últimos 8 caracteres a cada columna
for col in cyber_columns:
    df[col] = df[col].astype(str).str[-8:]


df['Moneda Capital'] = np.where(
    df['Moneda Capital'].isnull() & (df['Saldo pantalla ¢'] > 0) & (df['Saldo pantalla $'] <= 0), 'COLONES',
    np.where(df['Moneda Capital'].isnull() & (df['Saldo pantalla ¢'] <= 0) & (df['Saldo pantalla $'] > 0), 'DOLARES', df['Moneda Capital'])
)

df.loc[df['Moneda Capital'] == 'DOLARES', 'Saldo real ¢.'] = df['Saldo real $.']
df.loc[df['Moneda Capital'] == 'DOLARES', 'Int. Virtual ¢.'] = df['Int. Virtual $.']
df.loc[df['Moneda Capital'] == 'DOLARES', 'Saldo - Int. V ¢'] = df['Saldo - Int. V $']

df.loc[(df['Categoría'] == 'ACTIVA') & (df['Moneda Capital'] == 'COLONES'), 'Saldo real ¢.'] = df['Saldo pantalla ¢']
df.loc[(df['Categoría'] == 'ACTIVA') & (df['Moneda Capital'] == 'COLONES'), 'Saldo - Int. V ¢'] = df['Saldo pantalla ¢']

df.loc[(df['Categoría'] == 'ACTIVA') & (df['Moneda Capital'] == 'DOLARES'), 'Saldo real ¢.'] = df['Saldo pantalla $']
df.loc[(df['Categoría'] == 'ACTIVA') & (df['Moneda Capital'] == 'DOLARES'), 'Saldo - Int. V ¢'] = df['Saldo pantalla $']


# Eliminar las columnas por nombre
df = df.drop(['Saldo real $.', 'Int. Virtual $.', 'Saldo - Int. V $'], axis=1)

df.insert(1,'GESTOR', '990')
df.insert(1,'CARTERA', 'CUATRO')

df.insert(1,'PORCENTAJE HONORARIOS', None)

df['PORCENTAJE HONORARIOS'] = df['Tipo de Cuenta'].apply(
    lambda x: 10 if x == 'SIN NOTIFICAR' else (5 if x == 'SIN DEMANDA' else None)
)

df = df.rename(columns={
    'Saldo real ¢.': 'CAPITAL CALCULO',
    'Saldo - Int. V ¢': 'CAPITAL',
    'Int. Virtual ¢.': 'INTERES ANTES CJ',
    'Tipo de Cuenta': 'OBSERVACION FASE 3',
    'Categoría': 'CLASIF CARTERA',
    'Moneda Capital': 'MONEDA CAPITAL',
    'Cuenta': 'CUENTA',
    'Probabilidad': 'TEXTO 5',
    'Cédula': 'IDENTIFICACION',
    'Estado Civil': 'ESTADO CIVIL'
})


def eliminar_numeros_duplicados(row):
    # Si el valor en #CYBER 1 es igual a #CYBER 2 y #CYBER 3, dejar solo #CYBER 1
    if row['# CYBER 1'] == row['# CYBER 2'] == row['# CYBER 3']:
        row['# CYBER 2'] = None
        row['# CYBER 3'] = None
    # Si #CYBER 1 y #CYBER 2 son iguales, eliminar #CYBER 2 y #CYBER 3
    elif row['# CYBER 1'] == row['# CYBER 2']:
        row['# CYBER 2'] = None
    # Si #CYBER 2 y #CYBER 3 son iguales, eliminar #CYBER 3
    elif row['# CYBER 2'] == row['# CYBER 3']:
        row['# CYBER 3'] = None
    return row


def ordenar_numeros(row):
    # Si el valor en #CYBER 1 es igual a #CYBER 2 y #CYBER 3, dejar solo #CYBER 1
    if row['# CYBER 1'] == None or row['# CYBER 1'] == '0' and row['# CYBER 2'] != None and row['# CYBER 2'] != '0':
        row['# CYBER 1'] = row['# CYBER 2']
        row['# CYBER 2'] = row['# CYBER 3']
    # Si #CYBER 1 y #CYBER 2 son iguales, eliminar #CYBER 2 y #CYBER 3
    elif row['# CYBER 1'] and row['# CYBER 2'] == None or row['# CYBER 1'] and row['# CYBER 2'] == '0':
        row['# CYBER 1'] = row['# CYBER 3']
    return row

def ordenar_numeros(row):
    # Si el valor en #CYBER 1 es igual a #CYBER 2 y #CYBER 3, dejar solo #CYBER 1
    if row['# CYBER 1'] == None or row['# CYBER 1'] == '0' and row['# CYBER 2'] != None and row['# CYBER 2'] != '0':
        row['# CYBER 1'] = row['# CYBER 2']
        row['# CYBER 2'] = row['# CYBER 3']
    # Si #CYBER 1 y #CYBER 2 son iguales, eliminar #CYBER 2 y #CYBER 3
    elif row['# CYBER 1'] and row['# CYBER 2'] == None or row['# CYBER 1'] and row['# CYBER 2'] == '0':
        row['# CYBER 1'] = row['# CYBER 3']
    return row

def celulares_primero(row):
    if str(row['# CYBER 1']).startswith('2'):
        # Verificar si # CYBER 2 no empieza con '2' y mover el valor
        if not str(row['# CYBER 2']).startswith('2'):
            row['# CYBER 2'] = row['# CYBER 1']
        # Verificar si # CYBER 3 no empieza con '2' y mover el valor
        elif not str(row['# CYBER 3']).startswith('2'):
            row['# CYBER 3'] = row['# CYBER 1']
        row['# CYBER 1'] = None  
    return row

def ordenar_numeros2(row):
    # Si el valor en #CYBER 1 es igual a #CYBER 2 y #CYBER 3, dejar solo #CYBER 1
    if row['# CYBER 1'] == None or row['# CYBER 1'] == '0':
        row['# CYBER 1'] = row['# CYBER 2']
        row['# CYBER 2'] = row['# CYBER 3']
        row['# CYBER 3'] = None
    return row

def ordenar_numeros3(row):
    # Si el valor en #CYBER 1 es igual a #CYBER 2 y #CYBER 3, dejar solo #CYBER 1
    if row['# CYBER 1'] == None and row['# CYBER 1'] == '0' and row['# CYBER 2'] == None and row['# CYBER 2'] == '0':
        row['# CYBER 1'] = row['# CYBER 3']
        row['# CYBER 2'] = None
        row['# CYBER 3'] = None
    return row


def asignar_monto_costas(row):
    # Verificar si existen valores 'COLONES' y 'DOLARES' en las filas relacionadas
    matching_rows = df[df['CUENTA'] == row['CUENTA']]  # Filtramos todas las filas con la misma cuenta
    if 'COLONES' in matching_rows['MONEDA CAPITAL'].values and 'DOLARES' in matching_rows['MONEDA CAPITAL'].values:
        if row['MONEDA CAPITAL'] == 'COLONES':
            row['MONTO COSTAS'] = 30000
        elif row['MONEDA CAPITAL'] == 'DOLARES':
            row['MONTO COSTAS'] = 0
    else:
        if row['MONEDA CAPITAL'] == 'COLONES':
            row['MONTO COSTAS'] = 30000
        elif row['MONEDA CAPITAL'] == 'DOLARES':
            row['MONTO COSTAS'] = -1
    return row

df = df.apply(asignar_monto_costas, axis=1)




df = df.apply(eliminar_numeros_duplicados, axis=1)
df = df.apply(ordenar_numeros, axis=1)
df = df.apply(celulares_primero, axis=1)
df = df.apply(ordenar_numeros2, axis=1)
df = df.apply(ordenar_numeros3, axis=1)
df = df.apply(celulares_primero, axis=1)
df = df.apply(ordenar_numeros2, axis=1)
df = df.apply(ordenar_numeros3, axis=1)
df = df.apply(celulares_primero, axis=1)

def asignar_tipo_interes(row):
    if row['MONEDA CAPITAL'] == 'DOLARES':
        row['TIPO INTERES'] = '2,6'
    else:
        row['TIPO INTERES'] = '3,3'  
    return row


df = df.apply(asignar_tipo_interes, axis=1)


df = df.rename(columns={
    '# CYBER 1': 'TELEFONO 1 DEUDOR',
    '# CYBER 2': 'TELEFONO 2 DEUDOR',
    '# CYBER 3': 'TELEFONO 3 DEUDOR',
    '# CYBER 4': 'TELEFONO 4 DEUDOR',
    'EMAIL': 'CORREO ELECTRONICO DEUDOR'
})


ruta_guardado = 'J:\\python\\Depuracion\\Original\\asigM_Depurado.xlsx'
df.to_excel(ruta_guardado, sheet_name='Base de Cuentas', index=False)
os.startfile(ruta_guardado)