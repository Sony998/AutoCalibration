import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
import argparse
sheetname = "ELECTROESTIMULADOR"

def calcular_limites_grafica(datos):
    error_max = datos.max() + 0.5
    error_min = datos.min() - 0.5
    return error_max, error_min

def calcular_limites_grafica_corriente(datos):
    error_max = datos.max() + 100
    error_min = datos.min() - 100
    return error_max, error_min
def generar_grafico_corriente_2canales(archivo_excel):
    # Leer los datos del Excel
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    
    fila_inicial = 0
    while fila_inicial < len(df):
        # Verificar que hay suficientes filas para procesar
        if fila_inicial + 9 >= len(df):
            break
        
        # Extraer datos comunes
        nombreEse = dfdatos.iat[3, 1]
        nocertificado = df.iat[fila_inicial + 2, 5]
        if pd.isna(nocertificado):
            break
        for canal in range(1, 3):  # Canal 1 y 2
            if canal == 1:
                corriente = df.iloc[fila_inicial + 9, 1:5].astype(float)
            elif canal == 2:
                corriente = df.iloc[fila_inicial + 9, 7:11].astype(float)
            
            datospatron = df.iloc[fila_inicial + 7, 1:5].astype(float)
            fig, ax = plt.subplots(figsize=(7.04, 4.07))
            ax.plot(datospatron, corriente, marker='o', linestyle='--', color='b', label='CORRIENTE')
            ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 2))
            y_min, y_max = calcular_limites_grafica_corriente(corriente)
            ax.set_ylim(min(y_min, y_max), max(y_min, y_max))
            ax.set_yticks(np.arange(min(y_min, y_max), max(y_min, y_max) + 1, 50))
            ax.grid(True, which='both', linestyle='--', linewidth=0.5)
            ax.set_xlabel('INTENSIDAD')
            ax.set_ylabel('VALOR ENTREGADO')
            ax.set_title(f"{nombreEse} - {nocertificado} - Canal {canal} - CORRIENTE", fontsize=10, fontweight="bold")
            ax.legend()
            
            output_dir = f"OUTPUT/Graficos/Canal {canal}/Corriente"
            os.makedirs(output_dir, exist_ok=True)
            plt.savefig(f"{output_dir}/{nocertificado}.png", dpi=300, bbox_inches='tight')
            plt.close(fig)
            
            print(f"Gráfica de corriente generada para el certificado: {nocertificado}, Canal: {canal}")
        
        fila_inicial += 38
    
    print("Fin del archivo")
def genenerar_grafico_voltaje_2canales(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    fila_inicial = 0
    while fila_inicial < len(df):
        if fila_inicial + 9 >= len(df):
            break
        nombreEse = dfdatos.iat[3, 1]
        nocertificado = df.iat[fila_inicial + 2, 5]
        if pd.isna(nocertificado):
            break
        for canal in range(1, 3):  # Canal 1 y 2
            if canal == 1:
                corriente = df.iloc[fila_inicial + 8, 1:5].astype(float)
            elif canal == 2:
                corriente = df.iloc[fila_inicial + 8, 7:11].astype(float)
            datospatron = df.iloc[fila_inicial + 7, 1:5].astype(float)
            fig, ax = plt.subplots(figsize=(7.04, 4.07))
            ax.plot(datospatron, corriente, marker='o', linestyle='--', color='b', label='VOLTAJE')
            ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 2))
            y_min, y_max = calcular_limites_grafica(corriente)
            ax.set_ylim(min(y_min, y_max), max(y_min, y_max))
            ax.set_yticks(np.arange(min(y_min, y_max), max(y_min, y_max) + 1, 0.5))
            ax.grid(True, which='both', linestyle='--', linewidth=0.5)
            ax.set_xlabel('INTENSIDAD')
            ax.set_ylabel('VALOR ENTREGADO')
            ax.set_title(f"{nombreEse} - {nocertificado} - Canal {canal} - VOLTAJE", fontsize=10, fontweight="bold")
            ax.legend()
            output_dir = f"OUTPUT/Graficos/Canal {canal}/Voltaje"
            os.makedirs(output_dir, exist_ok=True)
            plt.savefig(f"{output_dir}/{nocertificado}.png", dpi=300, bbox_inches='tight')
            plt.close(fig)
            print(f"Gráfica de voltaje generada para el certificado: {nocertificado}, Canal: {canal}")

        fila_inicial += 38
    print("Fin del archivo")


def procesar_canal(archivo_excel):
    generar_grafico_corriente_2canales(archivo_excel)
    genenerar_grafico_voltaje_2canales(archivo_excel)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generar certificado a partir de un archivo Excel.")
    parser.add_argument(
        "--f", 
        required=True, 
        help="Especifica el archivo Excel que se debe usar, por ejemplo: Tensiometros.xlsx"
    )
    args = parser.parse_args()
    if not args.f:
        print("Error: No se ha proporcionado un archivo Excel. Por favor, use el argumento --f para especificar el archivo.")
    else:
        procesar_canal(args.f)