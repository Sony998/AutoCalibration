from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
import argparse
sheetname = "ELECTROCARDIOGRAFO"
desviaciones = []
certificados = []
errores_list = []
errores_promedio = []
anchos_list = []
incertidumbres = []
patrones = []
patroncarga_list = []
segundos_list = []
amplitudes_list = []
respuestas_list = []
ondas_list = []
tiempopromedio_list = []
notas = []
img_fondo_path1 = "Formatos/partesReporte/Pagina1.png"
img_fondo_path2 = "Formatos/partesReporte/Pagina2.png"
img_fondo_path3 = "Formatos/partesReporte/Pagina3.png"
img_fondo_path4 = "Formatos/partesReporte/Pagina4.png"
img_fondo_path5 = "Formatos/partesReporte/Pagina5.png"
output_directory1 = "OUTPUT/Reportes/1"
output_directory2 = "OUTPUT/Reportes/2"
output_directory3 = "OUTPUT/Reportes/3"
output_directory4 = "OUTPUT/Reportes/4"
output_directory5 = "OUTPUT/Reportes/5"
img_directory = "OUTPUT/Graficos/Error"
def crear_paginas(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    fila_inicial = 0
    errorpos = 0
    while True:
        if fila_inicial >= len(df):
            break
        nombreEse = df.iat[3, 12]
        fecha = df.iat[5, 12]
        ubicacion = df.iat[7, 12]
        metrologo = df.iat[10,12]
        if fila_inicial + 2 < len(df) and fila_inicial + 5 < len(df):
            certificado = df.iat[fila_inicial + 2 , 5]
            patron = df.iloc[fila_inicial + 5, 1:10].astype(int).tolist()
        else:
            break
        patrones.append(patron)
        print(patron)
        error_promedio = df.iat[fila_inicial + 8, 1]
        desviacion = df.iat[fila_inicial + 9, 1]
        nota = df.iat[fila_inicial + 1, 5]
        if pd.isna(nota):
            nota = "No se realizan observaciones"
        else:
            nota = str(nota)
        ancho = df.iloc[fila_inicial + 12, 1:5].astype(str).tolist()
        onda = df.iloc[fila_inicial + 15, 1:4].astype(str).tolist()
        amplitudes = df.iloc[fila_inicial + 19, 1:4].astype(str).tolist()
        respuesta = df.iat[fila_inicial + 22, 1]
        ondas_list.append(onda)
        amplitudes_list.append(amplitudes)
        respuestas_list.append(respuesta)
        print(ancho)
        anchos_list.append(ancho)
        incertidumbre =  df.iat[fila_inicial + 22, 1]
        errores_promedio.append(error_promedio)
        errores = df.iloc[fila_inicial + 7, 1:10].astype(float).reset_index(drop=True)
        desviaciones.append(desviacion)        
        incertidumbres.append(incertidumbre)
        certificados.append(certificado)
        errores_list.append(errores)
        notas.append(nota)
        fila_inicial += 23
    for certficado in certificados:
        print(certificados)
        agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha, metrologo, ubicacion)
        if errorpos >= len(certificados):
            print("Se han creado todos los reportes.")
            #os.system("python3 UnirPartes.py")
            break
        img_superior_path1 = os.path.join(img_directory, certficado + ".png")
        output_pdf_path = os.path.join(output_directory3, certficado + ".pdf")
        agregar_imagenes_pdf2(img_fondo_path2, img_superior_path1 , os.path.join(output_directory2, certficado + ".pdf"), 230, errores_list[errorpos], desviaciones[errorpos], errores_promedio[errorpos], anchos_list[errorpos])
        agregar_imagenes_pdf3(img_fondo_path3, output_pdf_path, ondas_list[errorpos], amplitudes_list[errorpos], respuestas_list[errorpos], notas[errorpos])
        errorpos += 1
"""         agregar_imagenes_pdf3(img_fondo_path3, output_pdf_path, errores_promedio[errorpos], desviaciones[errorpos], incertidumbres[errorpos], patroncarga_list[errorpos], segundos_list[errorpos], tiempopromedio_list[errorpos])
        print("Se ha creado el reporte completo para el desfibrilador con el certificado: ", certficado) 
        agregar_imagenes_pdf5(img_fondo_path5, os.path.join(output_directory5, certficado + ".pdf"), notas[errorpos])
 """

def agregar_imagenes_pdf1(fondo_path, output_pdf_path, nombrecertificado, fecha, metrologo, ubicacion):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    pdfmetrics.registerFont(TTFont('Baskerville', 'Formatos/Fuentes/Baskerville-Bold.ttf'))
    c.setFont("Baskerville", 12)
    c.drawString(310, 665, nombrecertificado)
    c.setFont("Arial", 15)
    c.drawString(290, 290, fecha)  # desplazado 20 unidades hacia arriba
    c.drawString(290, 265, fecha)  # desplazado 20 unidades hacia arriba
    c.drawString(290, 240, ubicacion)  # desplazado 20 unidades hacia arriba
    c.drawString(290, 215, metrologo)  # desplazado 20 unidades hacia arriba
    c.setFont("ArialI", 14)
    c.drawString(310, 112, "22.5")
    c.drawString(438, 112, "26.8")
    c.drawString(360, 95, "1012")
    c.drawString(310, 74, "59")
    c.drawString(438, 74, "68")
    c.save()
def agregar_imagenes_pdf2(img_fondo_path, img_superior_path1, output_pdf_path, yinferior, errores, desviacion, error_promedio, ancho):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior = Image.open(img_superior_path1).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    image_superior = image_superior.resize((int(image_superior.width / 6), int(image_superior.height / 6)), Image.LANCZOS)
    x_fondo = 0  # Colocar el fondo en 0 para ocupar toda la página
    xsuperior = (carta_ancho - image_superior.width) // 2 - 30
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, xsuperior, yinferior, width=image_superior.width, height=image_superior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    for i in range(9):
        c.drawString(120 + i * 45,590 , "{:.2f}".format(float(f"{errores[i]:.2f}")))
    c.drawString(370, 520, "{:.2f}".format(float(f"{error_promedio:.2f}")))
    c.drawString(370, 505, "{:.2f}".format(float(f"{desviacion:.2f}")))
    c.setFont("Arial", 10)
    for i in range(4):
        c.drawString(210 + i * 90, 130 , ancho[i])
    c.save()

""" def agregar_imagenes_pdf2(fondo_path, output_pdf_path, nombrecertificado,  patron, errores):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 10)
    for i in range(10):
        c.drawString(260 + i * 48,110 , "{:.2f}".format(float(f"{errores[i]:.2f}")))
    c.save() """
def agregar_imagenes_pdf2(img_fondo_path, img_superior_path1, output_pdf_path, yinferior, errores, desviacion, error_promedio, ancho):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior = Image.open(img_superior_path1).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    image_superior = image_superior.resize((int(image_superior.width / 6), int(image_superior.height / 6)), Image.LANCZOS)
    x_fondo = 0  # Colocar el fondo en 0 para ocupar toda la página
    xsuperior = (carta_ancho - image_superior.width) // 2 - 30
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, xsuperior, yinferior, width=image_superior.width, height=image_superior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    for i in range(9):
        c.drawString(120 + i * 45,590 , "{:.2f}".format(float(f"{errores[i]:.2f}")))
    c.drawString(370, 520, "{:.2f}".format(float(f"{error_promedio:.2f}")))
    c.drawString(370, 505, "{:.2f}".format(float(f"{desviacion:.2f}")))
    c.setFont("Arial", 10)
    for i in range(4):
        c.drawString(210 + i * 90, 130 , ancho[i])
    c.save()

def agregar_imagenes_pdf3(img_fondo_path, output_pdf_path, ondas, amplitudes, respuesta, nota):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(img_fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 14)
    c.drawString(270, 510 , respuesta)
    for i in range(3):
        c.drawString(260 + i * 80, 685 , ondas[i])
    for i in range(3):
        c.drawString(260 + i * 85, 600 , amplitudes[i])
    c.setFont("ArialBold", 14)
    c.drawString(240, 370, "OBSERVACIONES")
    c.setFont("Arial", 12)
    max_length = 75  # Maximum characters per line
    if len(nota) > max_length:
        nota_line1 = nota[:max_length]
        nota_line2 = nota[max_length:]
        c.drawString(85, 340, nota_line1)
        c.drawString(85, 325, nota_line2)
    else:
        c.drawString(85, 340, nota)
    c.save()



def agregar_imagenes_pdf5(fondo_path, output_pdf_path , nota):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    c.setFont("ArialBold", 14)
    c.drawString(220, 630, "OBSERVACIONES")
    c.setFont("Arial", 12)
    max_length = 75  # Maximum characters per line
    if len(nota) > max_length:
        nota_line1 = nota[:max_length]
        nota_line2 = nota[max_length:]
        c.drawString(85, 600, nota_line1)
        c.drawString(85, 585, nota_line2)
    else:
        c.drawString(85, 600, nota)
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