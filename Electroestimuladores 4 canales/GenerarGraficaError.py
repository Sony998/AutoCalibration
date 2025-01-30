import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
import argparse
sheetname = "ELECTROESTIMULADOR DE 4 CANALES"

def calcular_limites_grafica(datos):
    error_max = datos.max() + 0.5
    error_min = datos.min() - 0.5
    return error_max, error_min

def calcular_limites_grafica_corriente(datos):
    error_max = datos.max() + 200
    error_min = datos.min() - 200
    return error_max, error_min

def generar_grafico_error(archivo_excel, canal):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)  
    fila_inicial = 0

    while fila_inicial < len(df):
        if fila_inicial + 8 >= len(df):
            break
        nombreEse = dfdatos.iat[3, 1]
        nocertificado = df.iat[fila_inicial + 2, 5]
        datostest = df.iloc[fila_inicial + 9, 1:11].astype(float)
        print(datostest)
        if canal == 1:
            datospatron = df.iloc[fila_inicial + 7, 1:11].astype(float)
            datos_seleccionados = df.iloc[fila_inicial + 8, 1:11].astype(float)
        elif canal == 2:
            datospatron = df.iloc[fila_inicial + 7, 13:23].astype(float)
            datos_seleccionados = df.iloc[fila_inicial + 8, 13:23].astype(float)
        elif canal == 3:
            datospatron = df.iloc[fila_inicial + 7, 25:35].astype(float)
            datos_seleccionados = df.iloc[fila_inicial + 8, 25:35].astype(float)
        elif canal == 4:
            datospatron = df.iloc[fila_inicial + 7, 37:47].astype(float)
            datos_seleccionados = df.iloc[fila_inicial + 8, 37:47].astype(float)
        print(datospatron,datos_seleccionados)
        fig, ax = plt.subplots(figsize=(7.04, 4.07))
        ax.plot(datospatron, datos_seleccionados, marker='o', linestyle='-', color='b', label='VOLTAJE')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1))
        y_min, y_max = calcular_limites_grafica(datos_seleccionados)
        ax.set_ylim(min(y_min, y_max), max(y_min, y_max))
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        ax.set_xlabel('INTENSIDAD')
        ax.set_ylabel('VALOR OBTENIDO')
        ax.set_title(f"{nombreEse} - {nocertificado} - VOLTAJE - Canal {canal}", fontsize=10, fontweight="bold")
        ax.legend()

        output_dir = f"OUTPUT/Graficos/Canal {canal}/Voltaje"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nocertificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)

        print(f"Gráfica de voltaje generada para el certificado: {nocertificado}")
        fila_inicial += 26
    else:
        print("Fin del archivo")

def generar_grafico_corriente(archivo_excel, canal):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    fila_inicial = 0

    while fila_inicial < len(df):
        if fila_inicial + 8 >= len(df):
            break

        nombreEse = dfdatos.iat[3, 1]
        nocertificado = df.iat[fila_inicial + 2, 5]
        if canal == 1:
            datospatron = df.iloc[fila_inicial + 7, 1:11].astype(float)
            corriente = df.iloc[fila_inicial + 9, 1:11].astype(float)
            print(datospatron,corriente)
        elif canal == 2:
            datospatron = df.iloc[fila_inicial + 7, 13:23].astype(float)
            corriente = df.iloc[fila_inicial + 9, 13:23].astype(float)
            print(datospatron,corriente)
        elif canal == 3:
            datospatron = df.iloc[fila_inicial + 7, 25:35].astype(float)
            corriente = df.iloc[fila_inicial + 9, 25:35].astype(float)
            print(datospatron,corriente)
        elif canal == 4:
            datospatron = df.iloc[fila_inicial + 7, 37:47].astype(float)
            corriente = df.iloc[fila_inicial + 9, 37:47].astype(float)
            print(datospatron,corriente)

        fig, ax = plt.subplots(figsize=(7.04, 4.07))
        ax.plot(datospatron, corriente, marker='o', linestyle='--', color='b', label='CORRIENTE')
        ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1))
        y_min, y_max = calcular_limites_grafica_corriente(corriente)
        ax.set_ylim(min(y_min, y_max), max(y_min, y_max))
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        ax.set_xlabel('INTENSIDAD')
        ax.set_ylabel('VALOR ENTREGADO')
        ax.set_title(f"{nombreEse} - {nocertificado} - CORRIENTE - Canal {canal}", fontsize=10, fontweight="bold")
        ax.legend()

        output_dir = f"OUTPUT/Graficos/Canal {canal}/Corriente"
        os.makedirs(output_dir, exist_ok=True)
        plt.savefig(f"{output_dir}/{nocertificado}.png", dpi=300, bbox_inches='tight')
        plt.close(fig)

        print(f"Gráfica de corriente generada para el certificado: {nocertificado}")
        fila_inicial += 26
    else:
        print("Fin del archivo")

def procesar_canal(archivo_excel, canal):
    print(f"Procesando Canal {canal}...")
    generar_grafico_error(archivo_excel, canal)
    generar_grafico_corriente(archivo_excel, canal)

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
        for canal in range(1, 5):  # Cambia el rango según el número de canales
            procesar_canal(args.f, canal)
