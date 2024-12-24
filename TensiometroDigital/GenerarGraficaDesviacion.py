import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import argparse

def calcular_limites_grafica(datos, error_promedio):
    error_promedio = abs(error_promedio)
    if error_promedio < 0.15:
        error_promedio += 1
    error_max = datos.max() + error_promedio
    error_min = datos.min() - error_promedio 
    limite_superior = error_max
    limite_inferior = error_min
    return limite_superior, limite_inferior

def sistolica(fila_actual, nombreEse, df, dfdatos):
    while fila_actual < len(df):
        certificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[6, 1:7].astype(int) 
        datos_seleccionados = df.iloc[fila_actual + 10, 1:7].astype(float) 
        error_promedio = float(df.iat[fila_actual + 11, 1])
        desviacionestandar = float(df.iloc[fila_actual + 12, 1])
        print(certificado, datos_seleccionados)
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
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 20))
        ax.grid(True, which="both", linestyle="--", linewidth=0.5)
        ax.set_xlabel("PATRON")
        ax.set_ylabel("ERROR")
        ax.set_title(
            nombreEse + " - SISTOLICA \n" + certificado,
            fontsize=10,
            fontweight="bold",
        )
        ax.legend()
        output_dir = "OUTPUT/Graficos/Desviacion/Sistolica"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{certificado}.png",dpi=300, bbox_inches='tight')
        fila_actual += 35

def diastolica(fila_actual, nombreEse, df, dfdatos):
    while fila_actual < len(df):
        certificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[16, 1:7].astype(int) 
        datos_seleccionados = df.iloc[fila_actual + 20, 1:7].astype(float) 
        error_promedio = df.iat[fila_actual + 21, 1]
        desviacionestandar = float(df.iloc[fila_actual + 22, 1])
        print(certificado, datos_seleccionados, desviacionestandar, error_promedio)
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
            nombreEse + " - DIASTOLICA\n" + certificado,
            fontsize=10,
            fontweight="bold",
        )
        ax.legend()
        output_dir = "OUTPUT/Graficos/Desviacion/Diastolica"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{certificado}.png",dpi=300, bbox_inches='tight')
        fila_actual += 35


def frecuencia(fila_actual, nombreEse, df, dfdatos):
    while fila_actual < len(df):
        certificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[26, 1:5].astype(int) 
        datos_seleccionados = df.iloc[fila_actual + 30, 1:5].astype(float)
        error_promedio = float(df.iat[fila_actual + 31, 1])
        desviacionestandar = float(df.iloc[fila_actual + 32, 1])
        print(certificado, datos_seleccionados, desviacionestandar, error_promedio)
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
            nombreEse + " - F. DE PULSO\n" + certificado,
            fontsize=10,
            fontweight="bold",
        )
        ax.legend()
        output_dir = "OUTPUT/Graficos/Desviacion/Frecuencia"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{certificado}.png",dpi=300, bbox_inches='tight')
        fila_actual += 35

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generar gráficos de desviación.')
    parser.add_argument('--f', type=str, required=True, help='Nombre del archivo de Excel')
    args = parser.parse_args()

    archivo_excel = args.f
    dfdatos = pd.read_excel(archivo_excel, sheet_name='DATOS SOLICITANTE', header=None)
    df = pd.read_excel(archivo_excel, sheet_name='TENSIOMETRO DIGITAL', header=None)   
    fila_actual = 0
    nombreEse = dfdatos.iat[3, 1]

    sistolica(fila_actual, nombreEse, df, dfdatos)
    diastolica(fila_actual, nombreEse, df, dfdatos)
    frecuencia(fila_actual, nombreEse, df, dfdatos)
