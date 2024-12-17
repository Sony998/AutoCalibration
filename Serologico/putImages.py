from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
import argparse
sheetname = "BAÑO SEROLOGICO"
desviaciones = []
certificados = []
errores_list = []
errores_promedio = []
primeras = []
segundas = []   
incertidumbres_expandidas = []
incertidumbres = []
notas = []
img_fondo_path1 = "Formatos/partesReporte/Pagina1.png"
img_fondo_path2 = "Formatos/partesReporte/Pagina2.png"
img_fondo_path3 = "Formatos/partesReporte/Pagina3.png"
output_directory1 = "OUTPUT/Reportes/1"
output_directory2 = "OUTPUT/Reportes/2"
output_directory3 = "OUTPUT/Reportes/3"
def crear_paginas(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    fila_inicial = 0
    errorpos = 0
    while True:
        if fila_inicial >= len(df):
            break
        fecha = str(df.iat[1, 12])
        nocertificado = df.iat[fila_inicial + 2 , 5]
        nota = df.iat[fila_inicial + 1, 5]
        if pd.isna(nota):
            nota = "No se realizan observaciones"
        else:
            nota = str(nota)
        fecha = df.iat[5, 14]
        metrologo = df.iat[10, 14]
        nombreEse = df.iat[3, 14]
        if fila_inicial + 5 < len(df):
            primera = df.iloc[fila_inicial + 5,1:12].astype(float).tolist()
            print(primera)
        else:
            print(f"Index {fila_inicial + 5} is out of bounds for the DataFrame.")
            break
        incertidumbre = df.iat[fila_inicial + 6,1]
        incertidumbre_expandida = df.iat[fila_inicial +7 ,1]        
        if incertidumbre == "N.R":
            incertidumbre = str(incertidumbre)
        if incertidumbre_expandida == "N.R":
            incertidumbre_expandida = str(incertidumbre_expandida)
        primeras.append(primera)
        incertidumbres.append(incertidumbre)
        incertidumbres_expandidas.append(incertidumbre_expandida)
        certificados.append(nocertificado)
        notas.append(nota)
        fila_inicial += 13
    for certficado in certificados:
        agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha, metrologo, nombreEse)
        agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), certficado, incertidumbres[errorpos], incertidumbres_expandidas[errorpos], primeras[errorpos])
        agregar_imagenes_pdf3(img_fondo_path3, os.path.join(output_directory3, certficado + ".pdf"), notas[errorpos] )
        print("Se ha creado el reporte completo para el tensiometro con el certificado: ", certficado) 
        errorpos += 1
        if errorpos >= len(certificados):
            print("Se han creado todos los reportes.")
            os.system("python3 UnirPartes.py")
            break

def agregar_imagenes_pdf1(fondo_path, output_pdf_path, nombrecertificado, fecha, metrologo, nombreEse):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(360, 645, nombrecertificado)
    c.setFont("Arial", 13)
    c.drawString(300, 88, fecha)
    c.drawString(300, 68, fecha)
    c.drawString(300, 45, nombreEse)
    c.drawString(300, 20, metrologo)
    c.save()
def agregar_imagenes_pdf2(fondo_path, output_pdf_path, nombrecertificado, incertidumbre, incertidumbre_expandida, primera):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(345, 620, "12")
    c.drawString(438, 620, "24")
    c.drawString(390, 590, "1013")
    c.drawString(345, 565, "52")
    c.drawString(438, 565, "60")
    c.drawString(386, 80, "{:.2f}".format(float(f"{incertidumbre_expandida:.2f}")))
    c.drawString(386, 110, "{:.2f}".format(float(f"{incertidumbre:.2f}")))
    c.setFont("ArialI", 10)
    for i in range(11):
        c.drawString(155 + i * 31, 165, "{:.2f}".format(float(f"{primera[i]:.2f}")) if isinstance(primera[i], float) else str(primera[i]))
    c.save()


def agregar_imagenes_pdf3(fondo_path, output_pdf_path , nota):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    c.setFont("ArialBold", 14)
    c.drawString(220, 610, "OBSERVACIONES")
    c.setFont("Arial", 12)
    max_length = 75  # Maximum characters per line
    if len(nota) > max_length:
        nota_line1 = nota[:max_length]
        nota_line2 = nota[max_length:]
        c.drawString(70, 580, nota_line1)
        c.drawString(70, 610, nota_line2)
    else:
        c.drawString(85, 580, nota)
    c.save()

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
        crear_paginas(args.f)

if __name__ == "__main__":
    main()