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

def sistolica(df, dfdatos, fila_actual, nombreEse):
    while fila_actual < len(df):
        certificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[6, 1:7].astype(int) 
        datos_seleccionados = df.iloc[fila_actual + 10, 1:7].astype(float) 
        print(datospatron, datos_seleccionados)
        fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
        ax.scatter(datospatron, datos_seleccionados, c='#3d9bff')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 20))
        ax.set_ylim(calcular_limites_grafica(datos_seleccionados))
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        ax.set_xlabel('PATRON')
        ax.set_ylabel('ERROR')
        ax.set_title(f'{nombreEse} - SISTOLICA \n{certificado}', fontsize=10, fontweight='bold')
        output_dir = "OUTPUT/Graficos/Error/Sistolica/"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{certificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)
        fila_actual += 35

def diastolica(df, dfdatos, fila_actual, nombreEse):
    while fila_actual < len(df):
        certificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[16, 1:7].astype(int) 
        datos_seleccionados = df.iloc[fila_actual + 20, 1:7].astype(float) 
        fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
        ax.scatter(datospatron, datos_seleccionados, c='#3d9bff')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 10))
        ax.set_ylim(calcular_limites_grafica(datos_seleccionados))
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        ax.set_xlabel('PATRON')
        ax.set_ylabel('ERROR')
        ax.set_title(f'{nombreEse} - DIASTOLICA \n{certificado}', fontsize=10, fontweight='bold')
        output_dir = "OUTPUT/Graficos/Error/Diastolica"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{certificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)
        fila_actual += 35

def frecuencia(df, dfdatos, fila_actual, nombreEse):
    while fila_actual < len(df):
        certificado = df.iat[fila_actual + 2, 5]
        datospatron = df.iloc[26, 1:5].astype(int) 
        datos_seleccionados = df.iloc[fila_actual + 30, 1:5].astype(float)
        fig, ax = plt.subplots(figsize=(7.04, 4.07))  # Tamaño en pulgadas para obtener 2113x1220 píxeles a 300 DPI
        ax.scatter(datospatron, datos_seleccionados, c='#3d9bff')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 10))
        ax.set_ylim(calcular_limites_grafica(datos_seleccionados))
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        ax.set_xlabel('PATRON')
        ax.set_ylabel('ERROR')
        ax.set_title(f'{nombreEse} - F. DE PULSO \n{certificado}', fontsize=10, fontweight='bold')
        output_dir = "OUTPUT/Graficos/Error/Frecuencia"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{certificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)
        fila_actual += 35

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generar gráficos de error para tensiómetro digital.')
    parser.add_argument('--f', type=str, required=True, help='Nombre del archivo de Excel')
    args = parser.parse_args()

    archivo_excel = args.f
    dfdatos = pd.read_excel(archivo_excel, sheet_name='DATOS SOLICITANTE', header=None)
    df = pd.read_excel(archivo_excel, sheet_name='TENSIOMETRO DIGITAL', header=None)   
    fila_actual = 0
    nombreEse = dfdatos.iat[3, 1]

    sistolica(df, dfdatos, fila_actual, nombreEse)
    diastolica(df, dfdatos, fila_actual, nombreEse)
    frecuencia(df, dfdatos, fila_actual, nombreEse)