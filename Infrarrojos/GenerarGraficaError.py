import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
import argparse

def calcular_limites_grafica(datos):
    error_max = datos.max() + 1
    error_min = datos.min() - 1
    limite_superior = error_max
    limite_inferior = error_min
    return limite_superior, limite_inferior

def main(archivo_excel):
    sheetname = "TERMOMETRO INFRARROJO"
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None) 
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)  
    fila_inicial = 0

    while fila_inicial < len(df):
        nombreEse = dfdatos.iat[3, 1]
        nocertificado = df.iat[fila_inicial + 2, 5]
        datospatron = df.iloc[5, 1:12].astype(float) 
        datos_seleccionados = df.iloc[fila_inicial + 8, 1:12].astype(float)  # Fila de los datos seleccionados
        print("Creando", nocertificado)
        fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
        ax.scatter(datospatron, datos_seleccionados, c='#3d9bff')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 1))
        ax.set_ylim(calcular_limites_grafica(datos_seleccionados))
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        ax.set_xlabel('PATRON')
        ax.set_ylabel('ERROR')
        ax.set_title(
            str(nombreEse) + " - " + str(nocertificado),
            fontsize=10,
            fontweight="bold",
        )   
        output_dir = "OUTPUT/Graficos/Error"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nocertificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)
        fila_inicial += 13
    else:
        print("Fin del archivo")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generar gráfica de error.')
    parser.add_argument('--f', type=str, required=True, help='Nombre del archivo de Excel')
    parser.add_argument('--c', nargs="+", help='Nombre de la carpeta de salida')
    args = parser.parse_args()
    main(args.f)
