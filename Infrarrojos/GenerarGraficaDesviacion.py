import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt

sheetname = "TERMOMETRO INFRARROJO"
archivo_excel = '/home/raven/SANTANA.xlsx'
fila_inicial = 0
df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
def calcular_limites_grafica(datos, error_promedio):
    error_promedio = abs(error_promedio)
    error_max = datos.max() + error_promedio
    error_min = datos.min() - error_promedio 
    limite_superior = error_max
    limite_inferior = error_min
    return limite_superior, limite_inferior

while True:
    if fila_inicial >= len(df):
        print("Fin del archivo")
        break
    nocertificado = df.iat[fila_inicial + 2 , 5]
    nombreEse = str(df.iat[3, 15])
    datospatron = df.iloc[5, 1:12].astype(int) 
    datos_seleccionados = df.iloc[fila_inicial + 8, 1:12].astype(float)  # Fila de los datos seleccionados
    error_promedio = df.iat[fila_inicial + 9, 1]
    print(nocertificado, error_promedio)
    desviacionestandar = df.iat[fila_inicial + 10, 1]
    print(desviacionestandar, error_promedio)
    errormaximo = max(datos_seleccionados) - desviacionestandar
    errorminimo = min(datos_seleccionados) + desviacionestandar
    errores_promedio = np.full(len(datospatron), error_promedio)
    errores_superiores = errormaximo + datos_seleccionados  # Distancia al límite máximo
    errores_inferiores = datos_seleccionados + errorminimo  # Distancia al límite mínimo
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
    ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 1))
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
    fila_inicial += 13
