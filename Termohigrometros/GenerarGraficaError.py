import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
import argparse

sheetname = "TERMOHIGROMETRO"

def calcular_limites_grafica(datos):
    error_max = datos.max() + 1
    error_min = datos.min() - 1
    limite_superior = error_max
    limite_inferior = error_min
    return limite_superior, limite_inferior


def generar_grafico_error(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    fila_actual = 0
    while fila_actual < len(df):
        nombreEse = dfdatos.iat[3, 1]
        nocertificado = df.iat[fila_actual + 2 , 5]
        datospatron = df.iloc[fila_actual + 4, 1:7].astype(float)  # Fila de los datos patrón
        print(datospatron)
        datos_seleccionados = df.iloc[fila_actual + 7, 1:7].astype(float)  #
        fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
        ax.scatter(datospatron, datos_seleccionados, c='#3d9bff')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 4.2))
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
        print("Grafica de error generada para el certificado: ", nocertificado)
        fila_actual += 18
    else:
        print("Fin del archivo")

def main():
    parser = argparse.ArgumentParser(description="Generar certificado a partir de un archivo Excel.")
    parser.add_argument(
        "--f", 
        required=True, 
        help="Especifica el archivo Excel que se debe usar, por ejemplo: Tensiometros.xlsx"
    )
    parser.add_argument(
        "--c", 
        nargs="+", 
        help="Especifica el nombre de la nueva carpeta de drive"
    )
    args = parser.parse_args()
    if not args.f:
        print("Error: No se ha proporcionado un archivo Excel. Por favor, use el argumento --f para especificar el archivo.")
    else:
        generar_grafico_error(args.f)

if __name__ == "__main__":
    main()