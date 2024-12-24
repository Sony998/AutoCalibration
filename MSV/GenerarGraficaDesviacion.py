import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt

archivo_excel = '/home/raven/Quipama.xlsx'
fila_actual = 0

df = pd.read_excel(archivo_excel, sheet_name="MONITORES MULTIPARAMETROS", header=None)
dfdatos = pd.read_excel(archivo_excel, sheet_name='DATOS SOLICITANTE', header=None)
nombreEse = dfdatos.iat[3, 1]

def calcular_limites_grafica(datos, error_promedio):
    error_promedio = abs(error_promedio)
    error_max = datos.max() + error_promedio
    error_min = datos.min() - error_promedio 
    limite_superior = error_max
    limite_inferior = error_min
    return limite_superior, limite_inferior

def sistolica(fila_actual):
    while fila_actual < len(df):
        nombrecertificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[6,1:7].values.flatten().astype(float)
        datos_seleccionados = df.iloc[fila_actual + 10, 1:7].values.flatten().astype(float) # Fila de los datos seleccionados
        error_promedio = df.iat[fila_actual + 11, 1]
        desviacionestandar = float(df.iloc[fila_actual + 12, 1])
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
            f"{nombreEse} - SISTOLICA \n" + nombrecertificado,
            fontsize=10,
            fontweight="bold",
        )
        ax.legend()
        output_dir = "OUTPUT/Graficos/Desviacion/Sistolica"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nombrecertificado}.png",dpi=300, bbox_inches='tight')
        fila_actual += 70

def diastolica(fila_actual):
    while fila_actual < len(df):
        nombrecertificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[16, 1:7].values.flatten().astype(float)   
        datos_seleccionados = df.iloc[fila_actual + 20, 1:7].values.flatten().astype(float)
        error_promedio = df.iat[fila_actual + 21, 1]
        desviacionestandar = float(df.iloc[fila_actual + 22, 1])
        print(nombrecertificado, datos_seleccionados, desviacionestandar, error_promedio)
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
            f"{nombreEse} - DIASTOLICA\n" + nombrecertificado,
            fontsize=10,
            fontweight="bold",
        )
        ax.legend()
        output_dir = "OUTPUT/Graficos/Desviacion/Diastolica"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nombrecertificado}.png",dpi=300, bbox_inches='tight')
        fila_actual += 70


def frecuencia(fila_actual):
    while fila_actual < len(df):
        nombrecertificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[26, 1:5].values.flatten().astype(float)
        datos_seleccionados = df.iloc[fila_actual + 30, 1:5].values.flatten().astype(float)
        error_promedio = df.iat[fila_actual + 31, 1]
        desviacionestandar = float(df.iloc[fila_actual + 32, 1])
        print(nombrecertificado, datos_seleccionados, desviacionestandar, error_promedio)
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
            f"{nombreEse} - F. DE PULSO\n" + nombrecertificado,
            fontsize=10,
            fontweight="bold",
        )
        ax.legend()
        output_dir = "OUTPUT/Graficos/Desviacion/Frecuencia"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nombrecertificado}.png",dpi=300, bbox_inches='tight')
        fila_actual += 70

def nueva_saturacion(fila_actual):
    while fila_actual < len(df):
        nombrecertificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[fila_actual + 54, 1:6]
        datos_seleccionados = df.iloc[fila_actual + 56, 1:6]
        error_promedio = float(df.iat[fila_actual + 57, 1])
        desviacionestandar = float(df.iloc[fila_actual + 58, 1])
        print(datos_seleccionados)
        print(error_promedio, desviacionestandar)
        errores_promedio = np.full(len(datospatron), error_promedio)
        fig, ax = plt.subplots(figsize=(7.04, 4.07))
        ax.scatter(datospatron, datos_seleccionados, color="blue", label="Datos obtenidos")
        ax.errorbar(
            datospatron,
            errores_promedio,
            yerr=desviacionestandar,
            color="#f0d16c",
            ecolor="#f0d16c",
            alpha=0.5,
            capsize=15,
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
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 5))
        ax.grid(True, which="both", linestyle="--", linewidth=0.5)
        ax.set_xlabel("PATRON")
        ax.set_ylabel("ERROR")
        ax.set_title(
            f"{nombreEse} - SPO2(%) \n{nombrecertificado}",
            fontsize=10,
            fontweight="bold",
        )
        output_dir = "OUTPUT/Graficos/Desviacion/Saturacion"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nombrecertificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)
        fila_actual += 70

def nuevo_pulso(fila_actual):
    while fila_actual < len(df):
        nombrecertificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[fila_actual + 63,1:6 ]
        datos_seleccionados = df.iloc[fila_actual + 65, 1:6]
        error_promedio = df.iat[fila_actual + 66, 1]
        desviacionestandar = float(df.iloc[fila_actual + 67, 1])
        print(nombrecertificado, datos_seleccionados)
        errores_promedio = np.full(len(datospatron), error_promedio)
        fig, ax = plt.subplots(figsize=(7.04, 4.07))
        ax.scatter(datospatron, datos_seleccionados, color="blue", label="Datos obtenidos")
        ax.errorbar(
            datospatron,
            errores_promedio,
            yerr=desviacionestandar,
            color="#f0d16c",
            ecolor="#f0d16c",
            alpha=0.5,
            capsize=15,
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
            f"{nombreEse}- SPO2(FP) \n{nombrecertificado}",
            fontsize=10,
            fontweight="bold",
        )
        output_dir = "OUTPUT/Graficos/Desviacion/Pulso/"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nombrecertificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)
        fila_actual += 70
if __name__ == "__main__":
    nueva_saturacion(fila_actual)
    nuevo_pulso(fila_actual)
    sistolica(fila_actual)
    diastolica(fila_actual)
    frecuencia(fila_actual)
