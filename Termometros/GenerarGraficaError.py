import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
import argparse

def calcular_limites_grafica(datos):
    media = datos.mean()
    desviacion_estandar = datos.std()
    margen = desviacion_estandar * 0.75
    limite_superior = media + desviacion_estandar * 3 + margen
    limite_inferior = media - desviacion_estandar * 3 - margen
    return limite_inferior, limite_superior

def main(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name='TERMOMETRO', header=None)   
    dfdatos = pd.read_excel(archivo_excel, sheet_name='DATOS SOLICITANTE', header=None)
    fila_actual = 0

    while fila_actual < len(df):
        nombreEse = dfdatos.iat[3, 1]
        nocertificado = df.iat[fila_actual + 2 , 5]
        print(nocertificado)
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
        ax.set_title(f'{nombreEse} \n{nocertificado}', fontsize=10, fontweight='bold')
        output_dir = "OUTPUT/Graficos/Error"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nocertificado}.png", dpi=300, bbox_inches='tight')
        print(f"Guardado en  {nocertificado}.png")
        plt.close(fig)
        fila_actual += 13
    else:
        print("Fin del archivo")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generar gráfica de error a partir de un archivo Excel.')
    parser.add_argument('--f', type=str, required=True, help='Nombre del archivo Excel')
    parser.add_argument('--c', type=str, help='Carpeta')
    args = parser.parse_args()
    main(args.f)
