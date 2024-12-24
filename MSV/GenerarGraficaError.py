import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np

# Cargar el archivo Excel
archivo_excel = '/home/raven/Quipama.xlsx'
df = pd.read_excel(archivo_excel, sheet_name='MONITORES MULTIPARAMETROS', header=None)   
dfdatos = pd.read_excel(archivo_excel, sheet_name='DATOS SOLICITANTE', header=None)

fila_actual = 0

def calcular_limites_grafica(datos):
    error_max = datos.max() + 1
    error_min = datos.min() - 1
    limite_superior = error_max
    limite_inferior = error_min
    return limite_superior, limite_inferior

nombreEse = dfdatos.iat[3, 1]

def sistolica(fila_actual):
    while fila_actual < len(df):
        nombrecertificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[6,1:7].values.flatten().astype(float)
        datos_seleccionados = df.iloc[fila_actual + 10, 1:7].values.flatten().astype(float) # Fila de los datos seleccionados
        fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
        ax.scatter(datospatron, datos_seleccionados, c='#3d9bff')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 20))
        ax.set_ylim(calcular_limites_grafica(datos_seleccionados))
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        ax.set_xlabel('PATRON')
        ax.set_ylabel('ERROR')
        ax.set_title(f'{nombreEse} - SISTOLICA \n{nombrecertificado}', fontsize=10, fontweight='bold')
        output_dir = "OUTPUT/Graficos/Error/Sistolica/"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nombrecertificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)
        fila_actual += 70


def diastolica(fila_actual):
    while fila_actual < len(df):
        nombrecertificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[16, 1:7].values.flatten().astype(float)   
        datos_seleccionados = df.iloc[fila_actual + 20, 1:7].values.flatten().astype(float)
        fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
        ax.scatter(datospatron, datos_seleccionados, c='#3d9bff')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 10))
        ax.set_ylim(calcular_limites_grafica(datos_seleccionados))
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        ax.set_xlabel('PATRON')
        ax.set_ylabel('ERROR')
        ax.set_title(f'{nombreEse}- DIASTOLICA \n{nombrecertificado}', fontsize=10, fontweight='bold')
        output_dir = "OUTPUT/Graficos/Error/Diastolica/"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nombrecertificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)
        fila_actual += 70


def frecuencia(fila_actual):
    while fila_actual < len(df):
        nombrecertificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[26, 1:5].values.flatten().astype(float)
        datos_seleccionados = df.iloc[fila_actual + 30, 1:5].values.flatten().astype(float)  # Fila de los datos seleccionados
        fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
        ax.scatter(datospatron, datos_seleccionados, c='#3d9bff')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 10))
        ax.set_ylim(calcular_limites_grafica(datos_seleccionados))
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        ax.set_xlabel('PATRON')
        ax.set_ylabel('ERROR')
        ax.set_title(f'{nombreEse}- F. DE PULSO \n{nombrecertificado}', fontsize=10, fontweight='bold')
        output_dir = "OUTPUT/Graficos/Error/Frecuencia/"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nombrecertificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)
        fila_actual += 70

def saturacion(fila_actual):
    while fila_actual < len(df):
        nombrecertificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[fila_actual + 54, 1:6]
        datos_seleccionados = df.iloc[fila_actual + 56, 1:6]
        fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
        ax.scatter(datospatron, datos_seleccionados, c='#3d9bff')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 5))
        ax.set_ylim(calcular_limites_grafica(datos_seleccionados))
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        ax.set_xlabel('PATRON')
        ax.set_ylabel('ERROR')
        ax.set_title(f'{nombreEse} - SATURACION \n{nombrecertificado}', fontsize=10, fontweight='bold')
        output_dir = "OUTPUT/Graficos/Error/Saturacion/"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nombrecertificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)
        fila_actual += 70

def pulso(fila_actual):
    while fila_actual < len(df):
        nombrecertificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[fila_actual + 63,1:6 ]
        datos_seleccionados = df.iloc[fila_actual + 65, 1:6]
        print(nombrecertificado, datos_seleccionados)
        fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
        ax.scatter(datospatron, datos_seleccionados, c='#3d9bff')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 10))
        ax.set_ylim(calcular_limites_grafica(datos_seleccionados))
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        ax.set_xlabel('PATRON')
        ax.set_ylabel('ERROR')
        ax.set_title(f'{nombreEse}- F. PULSO \n{nombrecertificado}', fontsize=10, fontweight='bold')
        output_dir = "OUTPUT/Graficos/Error/Pulso/"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nombrecertificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)
        fila_actual += 70

if __name__ == "__main__":
    frecuencia(fila_actual)
    sistolica(fila_actual)
    diastolica(fila_actual)
    pulso(fila_actual)
    saturacion(fila_actual)
