import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import getpass, sys, os, subprocess
from datetime import datetime, timedelta

# --- Cargar y preparar datos ---
# Obtener usuario y rutas
user = getpass.getuser()
rutahome = f'/home/{user}/'
if user != 'max':
    rutahome = '/home/' + user + '/madmax/'
sys.path.append(rutahome + '00_funciones/')
os.chdir(rutahome + '08_ingdirect/')

# Cargar Excel
df = pd.read_excel("Saldo_Total.xlsx", parse_dates=["fecha"]).iloc[::-1].reset_index(drop=True)


# Filtrar fechas a partir del 1 de octubre de 2023
df_saldos = df[df["fecha"] >= pd.Timestamp("2023-10-01")]


# Obtener el último día de cada mes
df_ultimos_dias = df_saldos[df_saldos['fecha'].dt.is_month_end]


#### ADAPTAR DESDE AQUÍ, HASTA AQUÍ MANTENER

# Crear columna de mes completado (abreviado tipo "Jun 2025", usando el mes real)
df_ultimos_dias["mes_completado"] = df_ultimos_dias["fecha"].dt.strftime("%b %Y").str.capitalize()

# Calcular ganancia mensual (diferencia del saldo_total)
df_ultimos_dias["ganancia"] = df_ultimos_dias["saldo_total"].diff()



# --- Crear gráfico ---

# Eliminar primera fila (sin datos de ganancia)
df_plot = df_ultimos_dias.iloc[1:].copy()

# Preparar datos para graficar
x_labels = df_plot["mes_completado"]
y = df_plot["ganancia"]
colors = ["green" if val >= 0 else "red" for val in y]
x_pos = np.arange(len(x_labels))

# Graficar
fig, ax = plt.subplots(figsize=(10, 6))
bars = ax.bar(x_pos, y, color=colors, edgecolor="black", width=0.8)

# Ticks centrados
ax.set_xticks(x_pos)
ax.set_xticklabels(x_labels, rotation=45)

# Desplazar solo los labels medio paso a la izquierda
for label in ax.get_xticklabels():
    label.set_horizontalalignment('right')
    label.set_x(label.get_position()[0] - 0.5)

# Estética
ax.set_title(f'Ganancia mensual por mes completado — Updated: {datetime.now().strftime("%d/%m/%y %H:%M")}', fontsize=14)
ax.set_xlabel("Mes completado")
ax.set_ylabel("Ganancia (€)")
ax.grid(axis='y', linestyle='--', alpha=0.7)

# Guardar y mostrar
plt.tight_layout()

# Guardar imagen PNG
pwdimagen = rutahome + '08_ingdirect/Movs_ing_mensual.png'
plt.savefig(pwdimagen, dpi=300)

# Subir imagen a Dropbox
subprocess.run([rutahome + "/scripts/dropboxup/dropbox_uploader.sh", "upload", pwdimagen, "/Mi Drop/0006 Plots bancarios/"])
    
# Guardar DataFrame en un archivo Excel
ruta_salida = rutahome + "/08_ingdirect/Saldo_Total_dia30.xlsx"
df_ultimos_dias.iloc[::-1].to_excel(ruta_salida, index=False)

from openpyxl import load_workbook

wb = load_workbook(ruta_salida)
ws = wb.active

for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter

    for cell in col:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))

    ws.column_dimensions[col_letter].width = max_length + 2

wb.save(ruta_salida)



# Subir archivo a Dropbox
subprocess.run([rutahome + "/scripts/dropboxup/dropbox_uploader.sh", "upload", ruta_salida,"/Mi Drop/0002 Excels/"])
