import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt

archivo_excel = 'pulido.xlsx'
fila_actual = 5
df = pd.read_excel(archivo_excel, sheet_name="TERMOMETRO", header=None)

def calcular_limites_grafica(datos):
    media = datos.mean()
    desviacion_estandar = datos.std()
    margen = desviacion_estandar * 0.75
    limite_superior = media + desviacion_estandar * 3 + margen
    limite_inferior = media - desviacion_estandar * 3 - margen
    return limite_inferior, limite_superior

while True:
    if fila_actual >= len(df):
        print("Fin del archivo")
        break
    nombrecertificado = df.iat[fila_actual, 7]
    datospatron = df.iloc[9, 1:6].astype(float)
    datos_seleccionados = df.iloc[fila_actual + 6, 1:6].astype(float)
    error_promedio = df.iat[fila_actual + 7, 1]
    desviacionestandar = df.iat[fila_actual + 8, 1]
    print(nombrecertificado, desviacionestandar)
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
        elinewidth=18,
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
    ax.set_ylim(calcular_limites_grafica(datos_seleccionados))
    ax.set_xticks(np.arange(min(datospatron), max(datospatron) + 1, 5.5))
    ax.grid(True, which="both", linestyle="--", linewidth=0.5)
    ax.set_xlabel("PATRON")
    ax.set_ylabel("ERROR")
    ax.set_title(
        "E.S.E CENTRO DE SALUD SANTA LUCIA \n" + nombrecertificado,
        fontsize=10,
        fontweight="bold",
    )
    ax.legend()
    output_dir = "Graficos/Desviacion"
    os.makedirs(output_dir, exist_ok=True)
    plt.savefig(f"{output_dir}/{nombrecertificado}.png",dpi=300, bbox_inches='tight')
    fila_actual += 19
