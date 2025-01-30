from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
import argparse
sheetname = "LAMPARA DE FOTOTERAPIA"
desviaciones = []
certificados = []
errores_list = []
errores_promedio = []
primeras = []
terceras = []
cuartas = []
quintas = []
patrones_list = []
segundas = []   
incertidumbres_expandidas = []
incertidumbres = []
notas = []
img_fondo_path1 = "Formatos/partesReporte/Pagina1.png"
img_fondo_path2 = "Formatos/partesReporte/Pagina2.png"
img_fondo_path3 = "Formatos/partesReporte/Pagina3.png"
img_fondo_path4 = "Formatos/partesReporte/Pagina4.png"
output_directory1 = "OUTPUT/Reportes/1"
output_directory2 = "OUTPUT/Reportes/2"
output_directory3 = "OUTPUT/Reportes/3"
output_directory4 = "OUTPUT/Reportes/4"
inferior_directory = "OUTPUT/Graficos/LUX"
def crear_paginas(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    fila_inicial = 0
    errorpos = 0
    while True:
        if fila_inicial >= len(df):
            break
        nocertificado = str(df.iat[fila_inicial + 2 , 5])
        if fila_inicial + 9 < len(df):
            error_promedio = df.iat[fila_inicial + 9, 1]
        else:
            error_promedio = None
        nota = df.iat[fila_inicial + 1, 5]
        if pd.isna(nota):
            nota = "No se realizan observaciones"
        else:
            nota = str(nota)
        nombreEse = dfdatos.iat[3,1]
        fecha = dfdatos.iat[4, 1]
        metrologo = dfdatos.iat[7,1]
        temperatura_calibracion = dfdatos.iat[13,1]
        humedad_calibracion = dfdatos.iat[14,1]
        temperaturaminima =dfdatos.iat[10,1]
        temperaturamaxima = dfdatos.iat[10,2]
        humedadminima = dfdatos.iat[11,1]
        humedadmaxima = dfdatos.iat[11,2]
        presionbarometrica = dfdatos.iat[12,1]
        primera = df.iloc[fila_inicial+6:fila_inicial +12, 1].astype(float).to_list()
        segunda = df.iloc[fila_inicial+6:fila_inicial +12, 2].astype(float).to_list()
        tercera = df.iloc[fila_inicial+6:fila_inicial +12, 3].astype(float).to_list()
        cuarta = df.iloc[fila_inicial+6:fila_inicial +12, 4].astype(float).to_list()
        quinta = df.iloc[fila_inicial+6:fila_inicial +12, 5].astype(float).to_list()

        certificados.append(nocertificado)
        primeras.append(primera)
        segundas.append(segunda)
        terceras.append(tercera)
        cuartas.append(cuarta)
        quintas.append(quinta)

        notas.append(nota)
        fila_inicial += 11
    for certficado in certificados:
        agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha, metrologo, nombreEse)
        agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), temperaturaminima, temperaturamaxima, presionbarometrica, 
                              humedadminima, humedadmaxima, temperatura_calibracion, humedad_calibracion, primeras[errorpos], segundas[errorpos], terceras[errorpos], cuartas[errorpos], quintas[errorpos])
        img_superior_path1 = os.path.join(inferior_directory, certficado + ".png")
        # img_superior_path2 = os.path.join(superior_directory, certficado + ".png")
        agregar_imagenes_pdf3(img_fondo_path3, img_superior_path1, os.path.join(output_directory3, certficado + ".pdf"), 420, notas[errorpos])
        print("Se ha creado el reporte completo para la lampara de fotocurado con el certificado: ", certficado) 
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
    c.drawString(335, 675, nombrecertificado)
    c.setFont("Arial", 15)
    c.drawString(245, 250, fecha)
    c.drawString(245, 205, fecha)
    c.setFont("Arial", 12)
    c.drawString(245, 170, nombreEse)
    c.setFont("Arial", 15)
    c.drawString(245, 135, metrologo)
    c.save()
def agregar_imagenes_pdf2(fondo_path, output_pdf_path,temperaturaminima, temperaturamaxima, presionbarometrica, humedadminima, humedadmaxima, temperatura_calibracion, humedad_calibracion, primera, segunda, tercera, cuarta, quinta):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(325, 650, str(temperaturaminima))
    c.drawString(438, 650, str(temperaturamaxima))
    c.drawString(370, 627, str(presionbarometrica))
    c.drawString(325, 600, str(humedadminima))
    c.drawString(438, 600, str(humedadmaxima))
    c.drawString(400, 535, str(temperatura_calibracion))
    c.drawString(400, 508, str(humedad_calibracion))
    for i in range(5):
        c.drawString(160, 360 - i * 26, str(primera[i]))
        c.drawString(230, 360 - i * 26, str(segunda[i]))
        c.drawString(300, 360 - i * 26, str(tercera[i]))
        c.drawString(360, 360 - i * 26, str(cuarta[i]))
        c.drawString(420, 360 - i * 26, str(quinta[i]))
    
    c.save()

def agregar_imagenes_pdf3(img_fondo_path, img_superior_path, output_pdf_path, ysuperior, nota):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior = Image.open(img_superior_path).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    image_superior = image_superior.resize((int(image_superior.width / 5), int(image_superior.height / 5)), Image.LANCZOS)
    # Calcular la posición de la imagen
    x_fondo = 0  # Colocar el fondo en 0 para ocupar toda la página
    xsuperior = (carta_ancho - image_superior.width) // 2 - 30
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path, xsuperior, ysuperior, width=image_superior.width, height=image_superior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("ArialBold", 14)
    c.drawString(82, 310, "OBSERVACIONES")
    c.setFont("Arial", 12)
    max_length = 75  # Maximum characters per line
    if len(nota) > max_length:
        nota_line1 = nota[:max_length]
        nota_line2 = nota[max_length:]
        c.drawString(63, 300, nota_line1)
        c.drawString(63, 270, nota_line2)
    else:
        c.drawString(85, 285, nota)
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