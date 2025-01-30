import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import argparse
sheetname = "SERVOCUNA"

def calcular_limites_grafica(datos, error_promedio):
    error_promedio = abs(error_promedio)
    error_max = datos.max() + error_promedio
    error_min = datos.min() - error_promedio 
    limite_superior = error_max
    limite_inferior = error_min
    return limite_superior, limite_inferior



def generar_grafica_desviacion(archivo_excel):
    fila_inicial = 0
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    while fila_inicial < len(df):
        while fila_inicial < len(df):
            nombreEse = dfdatos.iat[3, 1]
            nocertificado = df.iat[fila_inicial + 2 , 5]
            datospatron = df.iloc[fila_inicial+4,1:11].astype(float)  # Fila de los datos patron
            datos_seleccionados = df.iloc[fila_inicial+6,1:11].astype(float)  # Fila de los datos seleccionados 
            error_promedio = df.iat[fila_inicial + 7, 1]
            desviacionestandar = df.iat[fila_inicial + 8, 1]
            errormaximo = max(datos_seleccionados) - desviacionestandar
            errorminimo = min(datos_seleccionados) + desviacionestandar
            errores_promedio = np.full(len(datospatron), error_promedio)
            fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
            ax.scatter(datospatron, datos_seleccionados, color="blue", label="Datos obtenidos")
            ax.errorbar(
                datospatron,
                errores_promedio,
                yerr=desviacionestandar,
                color="#f0d16c",
                ecolor="#f0d16c",
                alpha=0.5,  # Agregar transparencia
                capsize=15,  # Agregar remate a las barras de error
                elinewidth=25,
                label="Desviacion estandar",
            )

            ax.plot(
                datospatron,
                errores_promedio,
                "o-",
                color="red",
                markersize=5,
                label="Error promedio",
            )
            ax.set_ylim(calcular_limites_grafica(datos_seleccionados, error_promedio))
            ax.set_xticks(np.arange(int(min(datospatron)), int(max(datospatron)) + 1))
            ax.grid(True, which="both", linestyle="--", linewidth=0.5)
            ax.set_xlabel("PATRON")
            ax.set_ylabel("ERROR")
            ax.set_title(
                str(nombreEse) + " - " + str(nocertificado),
                fontsize=10,
                fontweight="bold",
            )
            ax.legend()
            output_dir = "OUTPUT/Graficos/Desviacion"
            os.makedirs(output_dir, exist_ok=True)
            plt.savefig(f"{output_dir}/{nocertificado}.png",dpi=300, bbox_inches='tight')
            print("Grafica de desviacion generada para el certificado: ", nocertificado)
            fila_inicial += 11

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
        generar_grafica_desviacion(args.f)

if __name__ == "__main__":
    main()