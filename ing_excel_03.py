import getpass
import sys
import subprocess
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import time

# Obtener usuario y rutas
user = getpass.getuser()
rutahome = f'/home/{user}/'
if user != 'max':
    rutahome = '/home/' + user + '/madmax/'
sys.path.append(rutahome + '00_funciones/')

# Importar módulos
from mis_paquetes import *
from funciones_aemetod import *
from arrays import *

# Cargar API Key
with open(rutahome + '00_funciones/aemet_api_key.txt') as f:
    api_key = f.readline().strip()

# Definir ruta del archivo Excel
ruta_archivo = rutahome + "/08_ingdirect/Excel_ing.xlsx"

# Descargar el archivo de Dropbox
subprocess.run([
    rutahome + "/scripts/dropboxup/dropbox_uploader.sh", "download",
    "/Mi Drop/0002 Excels/Excel_ing.xlsx", ruta_archivo
])

# Leer datos de las hojas 0, 2 y 3
df_ing = pd.read_excel(ruta_archivo, sheet_name=0)
df_naranja = pd.read_excel(ruta_archivo, sheet_name=2)
df_deposito = pd.read_excel(ruta_archivo, sheet_name=3)
df_tesoro = pd.read_excel(ruta_archivo, sheet_name=4)
# Identificar nombres de columnas relevantes
clm_fecha = df_ing.columns[0]  # Se asume que la columna de fecha es la misma en todas las hojas
clm_importe = 'IMPORTE (€)'  # Se asume que la columna de importes es la misma en todas las hojas

# Convertir fechas y valores en cada hoja, y asegurarnos de que solo se conserve la fecha (sin hora)
for df in [df_ing, df_naranja, df_deposito,df_tesoro]:
    df['fecha'] = pd.to_datetime(df[clm_fecha]).dt.date  # Eliminar la hora y conservar solo la fecha
    df[clm_importe] = pd.to_numeric(df[clm_importe], errors='coerce')

# Generar rango de fechas desde 1/1/2023 hasta hoy
fecha_inicio = datetime(2023,1, 1)
fecha_fin = datetime.now()
fechas = pd.date_range(fecha_inicio, fecha_fin)

# Crear DataFrame con fechas y saldos iniciales
df_saldos = pd.DataFrame({'fecha': fechas})
df_saldos['saldo_ing'] = 5063.16
df_saldos['saldo_naranja'] = 0
df_saldos['saldo_deposito'] = 0
df_saldos['saldo_tesoro']=0
df_saldos['saldo_total'] = 5063.16


# Asegurarse de que las fechas en df_saldos también están sin hora
df_saldos['fecha'] = df_saldos['fecha'].dt.date


# Función para actualizar saldos
def actualizar_saldo(df_movs, saldo_anterior):
    movimientos = df_movs.groupby('fecha')[clm_importe].sum()
    return saldo_anterior + movimientos.reindex(fechas, fill_value=0).values

# Calcular saldos acumulados
df_saldos['saldo_ing'] = actualizar_saldo(df_ing, df_saldos['saldo_ing'])
df_saldos['saldo_naranja'] = actualizar_saldo(df_naranja, df_saldos['saldo_naranja'])
df_saldos['saldo_deposito'] = actualizar_saldo(df_deposito, df_saldos['saldo_deposito'])
df_saldos['saldo_tesoro']=actualizar_saldo(df_tesoro,df_saldos['saldo_tesoro'])
for i in range(1, len(df_saldos)):
    movimientos_ing = df_ing[df_ing['fecha'] == df_saldos.at[i, 'fecha']][clm_importe].sum()
    movimientos_naranja = df_naranja[df_naranja['fecha'] == df_saldos.at[i, 'fecha']][clm_importe].sum()
    movimientos_deposito = df_deposito[df_deposito['fecha'] == df_saldos.at[i, 'fecha']][clm_importe].sum()
    movimientos_tesoro = df_tesoro[df_tesoro['fecha'] == df_saldos.at[i, 'fecha']][clm_importe].sum()
    
    df_saldos.at[i, 'saldo_ing'] = df_saldos.at[i-1, 'saldo_ing'] + movimientos_ing
    df_saldos.at[i, 'saldo_naranja'] = df_saldos.at[i-1, 'saldo_naranja'] + movimientos_naranja
    df_saldos.at[i, 'saldo_deposito'] = df_saldos.at[i-1, 'saldo_deposito'] + movimientos_deposito
    df_saldos.at[i, 'saldo_tesoro'] = df_saldos.at[i-1, 'saldo_tesoro'] + movimientos_tesoro
    df_saldos.at[i, 'saldo_total'] = df_saldos.at[i, 'saldo_ing'] + df_saldos.at[i, 'saldo_naranja'] + df_saldos.at[i, 'saldo_deposito']+df_saldos.at[i, 'saldo_tesoro']

# Guardar DataFrame en un archivo Excel
ruta_salida = rutahome + "/08_ingdirect/Saldo_Total.xlsx"
df_saldos.iloc[::-1].to_excel(ruta_salida, index=False)

# Subir archivo a Dropbox
subprocess.run([rutahome + "/scripts/dropboxup/dropbox_uploader.sh", "upload", ruta_salida,"/Mi Drop/0002 Excels/"])


saldo = np.array(df_saldos['saldo_total'])
fechas = np.array(df_saldos['fecha'])

saldo_aprox = mediamovil(saldo, 50)


# Plotear 'saldo_total' frente a 'fecha'
fig1, ax1 = plt.subplots(figsize=(15, 9), dpi=100)
ax1.plot(fechas, saldo, marker='', linestyle='-', alpha=0.2)
ax1.plot(fechas, saldo_aprox, linestyle='-', linewidth=2.0)

# Definir el rango de fechas para los límites
fecha_inicio = datetime(2023, 1, 1)
fecha_fin = datetime.now()


saldo_final = df_saldos['saldo_total'].iloc[-1]

# Configurar gráfico
ax1.set_title(f'Saldo Total: {saldo_final:,.2f} € — Updated: {datetime.now().strftime("%d/%m/%y %H:%M")}', fontsize=14)
ax1.set_xlabel('Fecha')
ax1.set_ylabel('Saldo (€)')
ax1.set_ylim((0,17000))
ax1.set_xlim([datetime(2023, 10, 1), fecha_fin])
ax1.grid(True, linestyle='--', linewidth=0.9, alpha=0.3)

hoy = datetime.now()

def restar_anio(fecha, n):
    try:
        return fecha.replace(year=fecha.year - n)
    except ValueError:
        return fecha.replace(month=2, day=28, year=fecha.year - n)

n_anios = 5  # ajusta aquí

for i in range(1, n_anios + 1):
    fecha_linea = restar_anio(hoy, i)
    ax1.axvline(fecha_linea, linestyle='--', linewidth=1, alpha=0.6)

# Guardar imagen
pwd_imagen = rutahome + "/08_ingdirect/Movs_ing_total.png"
fig1.savefig(pwd_imagen, bbox_inches='tight', dpi=100)
time.sleep(2)

#Subir imagen a Dropbox (descomentar si necesario)
subprocess.run([rutahome + "/scripts/dropboxup/dropbox_uploader.sh", "upload", pwd_imagen,"/Mi Drop/0002 Excels/" ])
