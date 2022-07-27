# -*- coding: utf-8 -*-
"""
Created on Wed Jul 24 12:18:07 2022

@author: braviles
"""

import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np

#%% Archivos y datos de entrada
archivo_trimestral = 'sietel_savs_2t_2022.xls'
archivo_histórico_total = '4.1.1-Suscripciones-TV-Paga_Mar-2022.xlsx'
mes_1 = '2022-04'
mes_2 = '2022-05'
mes_3 = '2022-06'

#%% Procesamiento del archivo trimestral de Control
df_avs = pd.read_excel(archivo_trimestral)
df_avs.drop(index=df_avs.index[-1],axis=0,inplace=True)
df_avs.drop(['fechaCarga', 'fechaCarga.1', 'fechaCarga.2'], inplace=True, axis=1)
df_avs['mes_1'].fillna(value=0, inplace=True)
df_avs['mes_2'].fillna(value=0, inplace=True)
df_avs['mes_3'].fillna(value=0, inplace=True)
df_avs = df_avs[df_avs['CATEGORIA'].notna()]

# Se importa el archivo histórico
df_history = pd.read_excel(archivo_histórico_total)
df_history.drop(['No.'], inplace=True, axis=1)

# Cambio del formato de fechas
longitud_columnas = len(df_history.columns)
fechas_largas = (df_history.iloc[:, 5:longitud_columnas].columns.tolist())

fechas_columns = [datetime.strptime(str(x.date()), '%Y-%m-%d').strftime('%Y-%m') for x in fechas_largas]
demas_columns = (df_history.iloc[:, 0:5].columns.tolist())
columnas = demas_columns + fechas_columns
df_history.columns = columnas

# Creación de la columna clave para identificación
df_history['clave'] = df_history['Concesionario'] + df_history['Provincia'] + df_history['Cobertura']
df_avs['clave'] = df_avs['CONCESIONARIO'] + df_avs['PROVINCIA'] + df_avs['COBERTURA']

# Eliminación en el archivo df_avs las filas que tienen en la columna ESTADO valores CANCELADO y PRUEBAS
df_avs = df_avs.loc[(df_avs['ESTADO'] !='CANCELADO') & (df_avs['ESTADO'] !='PARA PRUEBAS DEL SIE')]
df_avs = df_avs.reset_index(drop=True)

# Cambiar las columnas clave a la posición 0
clave_df_history = df_history.pop('clave')
clave_df_avs = df_avs.pop('clave')
df_history.insert(0, 'clave', clave_df_history)
df_avs.insert(0, 'clave', clave_df_avs)

#%% Función para clasificar la información
def clasificar(mes_1, mes_2, mes_3):
    df_nuevos = pd.DataFrame(columns=['clave', 'Categoría', 'Concesionario', 'Nombre Estación', 'Provincia', 'Cobertura', mes_1, mes_2, mes_3])
    for i in range(len(df_avs)):
        for j in range(len(df_history)):
            if df_avs.loc[i, 'clave'] == df_history.loc[j, 'clave'] :
                df_history.loc[j, mes_1] = df_avs.loc[i, 'mes_1']
                df_history.loc[j, mes_2] = df_avs.loc[i, 'mes_2']
                df_history.loc[j, mes_3] = df_avs.loc[i, 'mes_3']
    
    for i in range(len(df_avs)):
        if df_avs.loc[i, 'clave'] not in list(df_history['clave']):
                df_nuevos.loc[len(df_nuevos.index)]=[
                    df_avs.loc[i, 'clave'], 
                    df_avs.loc[i, 'CATEGORIA'], 
                    df_avs.loc[i, 'CONCESIONARIO'], 
                    df_avs.loc[i, 'NOMBRE ESTACIÓN'],
                    df_avs.loc[i, 'PROVINCIA'],
                    df_avs.loc[i, 'COBERTURA'],
                    df_avs.loc[i, 'mes_1'],
                    df_avs.loc[i, 'mes_2'],
                    df_avs.loc[i, 'mes_3']
                    ]
    
    df_history_final = pd.concat([df_history, df_nuevos], axis=0, ignore_index=True)
    
    for j in range(len(df_history_final)):
        if df_history_final.iloc[j, -4] > 0 and df_history_final.iloc[j, -3] == 0:
            df_history_final.iloc[j, -3] = round(df_history_final.iloc[j, -4] * (1 + (((df_history_final.iloc[j, -4]/df_history_final.iloc[j, -6])**(1/3))-1))) 
        if df_history_final.iloc[j, -3] > 0 and df_history_final.iloc[j, -2] == 0:
            df_history_final.iloc[j, -2] = round(df_history_final.iloc[j, -3] * (1 + (((df_history_final.iloc[j, -3]/df_history_final.iloc[j, -5])**(1/3))-1)))
        if df_history_final.iloc[j, -2] > 0 and df_history_final.iloc[j, -1] == 0:
            df_history_final.iloc[j, -1] = round(df_history_final.iloc[j, -2] * (1 + (((df_history_final.iloc[j, -2]/df_history_final.iloc[j, -4])**(1/3))-1)))
        
        if df_history_final.iloc[j, -4] > 0 and str(df_history_final.iloc[j, -3]) == 'nan':
            df_history_final.iloc[j, -3] = round(df_history_final.iloc[j, -4] * (1 + (((df_history_final.iloc[j, -4]/df_history_final.iloc[j, -6])**(1/3))-1))) 
        if df_history_final.iloc[j, -3] > 0 and str(df_history_final.iloc[j, -2]) == 'nan':
            df_history_final.iloc[j, -2] = round(df_history_final.iloc[j, -3] * (1 + (((df_history_final.iloc[j, -3]/df_history_final.iloc[j, -5])**(1/3))-1)))
        if df_history_final.iloc[j, -2] > 0 and str(df_history_final.iloc[j, -1]) == 'nan':
            df_history_final.iloc[j, -1] = round(df_history_final.iloc[j, -2] * (1 + (((df_history_final.iloc[j, -2]/df_history_final.iloc[j, -4])**(1/3))-1)))
    
    print(f'Mes {mes_1}:', df_history_final[mes_1].sum())
    print(f'Mes {mes_2}:', df_history_final[mes_2].sum())
    print(f'Mes {mes_3}:', df_history_final[mes_3].sum())
    
    return df_history_final

#%% Ingresar los meses del trimestre que se quiere crear la estadística
df_history_final = clasificar(mes_1, mes_2, mes_3)

#%% Exportando resultados
df_history_final.to_excel(r'resultado_avs.xlsx', index=True, header=True)

#%% Graficar el histórico del servicio AVS
longitud_columns = len(df_history_final.columns)
fechas = df_history_final.iloc[:, 6:longitud_columns].columns.tolist()
history_lines = []
for mes in fechas:
    monthly_lines = round(df_history_final[mes].sum())
    history_lines.append(monthly_lines)

plt.figure()
plt.plot(fechas, history_lines, 'tab:red')
plt.xlabel('Tiempo')
plt.ylabel('Suscriptores')
ax = plt.gca()
plt.xticks(rotation=90)
for label in ax.get_xaxis().get_ticklabels()[::2]:
    label.set_visible(False)
plt.show()

