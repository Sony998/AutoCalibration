import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np

# Cargar el archivo Excel
archivo_excel = 'pulido.xlsx'
df = pd.read_excel(archivo_excel, sheet_name='TERMOMETRO', header=None)   
fila_actual = 5

def calcular_limites_grafica(datos):
    media = datos.mean()
    desviacion_estandar = datos.std()
    margen = desviacion_estandar * 0.75
    limite_superior = media + desviacion_estandar * 3 + margen
    limite_inferior = media - desviacion_estandar * 3 - margen
    return limite_inferior, limite_superior

while fila_actual < len(df):
    nombrecertificado = df.iat[fila_actual, 7]
    nombrecertificado = "".join([c if c.isalnum() else "_" for c in nombrecertificado])
    print(nombrecertificado)
    datospatron = df.iloc[fila_actual + 4, 1:6].astype(int)
    datos_seleccionados = df.iloc[fila_actual + 6, 1:6].astype(float)
    print(datos_seleccionados)
    fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
    ax.scatter(datospatron, datos_seleccionados, c='#3d9bff')
    ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 3.5))
    ax.set_ylim(calcular_limites_grafica(datos_seleccionados))
    ax.grid(True, which='both', linestyle='--', linewidth=0.5)
    ax.set_xlabel('PATRON')
    ax.set_ylabel('ERROR')
    ax.set_title(f'E.S.E CENTRO DE SALUD SANTA LUCIA \n{nombrecertificado}', fontsize=10, fontweight='bold')
    output_dir = "Graficos/Error"
    os.makedirs(output_dir, exist_ok=True)
    plt.savefig(f"{output_dir}/{nombrecertificado}.png", dpi=300, bbox_inches='tight')
    print(f"Guardado en  {nombrecertificado}.png")
    plt.close(fig)
    fila_actual += 19
else:
    print("Fin del archivo")
