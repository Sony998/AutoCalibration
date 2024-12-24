import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt

sheetname = "PULSO OXIMETRO"
archivo_excel = '/home/raven/Quipama.xlsx'
fila_inicial = 0
df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)

def calcular_limites_grafica(datos, error_promedio):
    error_promedio = abs(error_promedio)
    error_max = datos.max() + error_promedio
    error_min = datos.min() - error_promedio 
    limite_superior = error_max + 1
    limite_inferior = error_min - 1
    return limite_superior, limite_inferior
nombreEse = dfdatos.iat[3,1]
def saturacion(fila_inicial):
    while fila_inicial < len(df):
        nocertificado = df.iat[fila_inicial + 2, 5]
        datospatron = df.iloc[6, 1:6].astype(int) 
        datos_seleccionados = df.iloc[fila_inicial + 8, 1:6].astype(float)
        error_promedio = df.iat[fila_inicial + 9, 1]
        print(nocertificado, error_promedio)
        desviacionestandar = df.iat[fila_inicial + 10, 1]
        print(desviacionestandar, error_promedio)
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
            markersize=3,
            label="Error promedio",
        )
        ax.set_ylim(calcular_limites_grafica(datos_seleccionados, error_promedio))
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 5))
        ax.grid(True, which="both", linestyle="--", linewidth=0.5)
        ax.set_xlabel("PATRON")
        ax.set_ylabel("ERROR")
        ax.set_title(
            str(nombreEse) + " - " + str(nocertificado),
            fontsize=10,
            fontweight="bold",
        )
        ax.legend()
        output_dir = "OUTPUT/Graficos/Saturacion/Desviacion"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nocertificado}.png", dpi=300, bbox_inches='tight')
        fila_inicial += 21
    else:
        print("Graficas de saturacion finalizadas")

def pulso(fila_inicial):
    while fila_inicial < len(df):
        nocertificado = df.iat[fila_inicial + 2, 5]
        datospatron = df.iloc[14, 1:6].astype(int) 
        datos_seleccionados = df.iloc[fila_inicial + 8, 1:6].astype(float) 
        error_promedio = df.iat[fila_inicial + 9, 1]
        print(nocertificado, error_promedio)
        desviacionestandar = df.iat[fila_inicial + 10, 1]
        print(desviacionestandar, error_promedio)
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
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 10))
        ax.grid(True, which="both", linestyle="--", linewidth=0.5)
        ax.set_xlabel("PATRON")
        ax.set_ylabel("ERROR")
        ax.set_title(
            str(nombreEse) + " - " + str(nocertificado),
            fontsize=10,
            fontweight="bold",
        )
        ax.legend()
        output_dir = "OUTPUT/Graficos/Pulso/Desviacion"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nocertificado}.png", dpi=300, bbox_inches='tight')
        fila_inicial += 21
    else:
        print("Graficas de f de pulso finalizadas")

if __name__ == "__main__":
    saturacion(fila_inicial)
    pulso(fila_inicial)