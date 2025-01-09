from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
import argparse
sheetname = "INCUBADORA NEONATAL"
desviaciones = []
certificados = []
errores_list = []
errores_promedio = []
primeras_list = []
primerpatrones_list = []
segundopatron_list = []
patrones_list = []
segundas_list = []   
incertidumbres_expandidas = []
incertidumbres = []
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
inferior_directory = "OUTPUT/Graficos/Desviacion"
superior_directory = "OUTPUT/Graficos/Error"
def crear_paginas(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    fila_inicial = 0
    errorpos = 0
    while True:
        if fila_inicial >= len(df):
            break
        nocertificado = df.iat[fila_inicial + 2 , 5]
        error_promedio = df.iat[fila_inicial + 7, 1]
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
        desviacionestandar = df.iat[fila_inicial + 8, 1]
        primerpatron = df.iloc[fila_inicial + 4,1:9].apply(lambda x: str(x) if x == "N.R" else float(x)).tolist()
        print(primerpatron)
        segundopatron = df.iloc[fila_inicial + 4, 9:18].apply(lambda x: str(x) if x == "N.R" else float(x)).tolist()
        print(segundopatron)
        primera = df.iloc[fila_inicial + 5, 1:9].apply(lambda x: str(x) if x == "N.R" else float(x)).tolist()
        segunda = df.iloc[fila_inicial + 5, 9:18].apply(lambda x: str(x) if x == "N.R" else float(x)).tolist()

        incertidumbre = df.iat[fila_inicial + 9, 1]
        incertidumbre_expandida = df.iat[fila_inicial + 10, 1]
        if desviacionestandar == "N.R":
            desviacionestandar = str(desviacionestandar)
        if incertidumbre == "N.R":
            incertidumbre = str(incertidumbre)
        if incertidumbre_expandida == "N.R":
            incertidumbre_expandida = str(incertidumbre_expandida)
        primeras_list.append(primera)
        segundas_list.append(segunda)
        errores_promedio.append(error_promedio)
        primerpatrones_list.append(primerpatron)
        segundopatron_list.append(segundopatron)
        incertidumbres.append(incertidumbre)
        incertidumbres_expandidas.append(incertidumbre_expandida)
        certificados.append(nocertificado)
        desviaciones.append(desviacionestandar)
        notas.append(nota)
        fila_inicial += 11
    for certficado in certificados:
        agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha, metrologo, nombreEse)
        agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), temperaturaminima, temperaturamaxima, presionbarometrica, humedadminima, humedadmaxima, temperatura_calibracion, humedad_calibracion, incertidumbre, incertidumbre_expandida)
        agregar_imagenes_pdf3(img_fondo_path3, os.path.join(output_directory3, certficado + ".pdf"), errores_promedio[errorpos], desviaciones[errorpos], primerpatrones_list[errorpos], segundopatron_list[errorpos], primeras_list[errorpos], segundas_list[errorpos])
        img_superior_path1 = os.path.join(inferior_directory, certficado + ".png")
        img_superior_path2 = os.path.join(superior_directory, certficado + ".png")
        agregar_imagenes_pdf4(img_fondo_path4, img_superior_path1, img_superior_path2, os.path.join(output_directory4, certficado + ".pdf"),
                             yinferior=80, ysuperior=420, error_promedio=errores_promedio[errorpos], desviacion=desviaciones[errorpos])
        agregar_imagenes_pdf5(img_fondo_path5, os.path.join(output_directory5, certficado + ".pdf"), notas[errorpos] )
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
    c.setFont("ArialI", 12)
    c.drawString(305, 668, nombrecertificado)
    c.setFont("Arial", 15)
    c.drawString(320, 265, fecha)
    c.drawString(320, 235, fecha)
    c.setFont("Arial", 10)
    c.drawString(320, 210, nombreEse)
    c.setFont("Arial", 15)
    c.drawString(320, 180, metrologo)
    c.save()
def agregar_imagenes_pdf2(fondo_path, output_pdf_path, temperaturaminima, temperaturamaxima, presionbarometrica, humedadminima, humedadmaxima, temperatura_calibracion, humedad_calibracion, incertidumbre, incertidumbre_expandida):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(325, 650, str(temperaturaminima))
    c.drawString(438, 650, str(temperaturamaxima))
    c.drawString(370, 625, str(presionbarometrica))
    c.drawString(325, 600, str(humedadminima))
    c.drawString(438, 600, str(humedadmaxima))
    c.drawString(400, 520, str(temperatura_calibracion))
    c.drawString(400, 490, str(humedad_calibracion) + "%")
    c.drawString(400, 300, str(incertidumbre_expandida))
    c.drawString(400, 275, str(incertidumbre))
    c.save()

def agregar_imagenes_pdf3(fondo_path, output_pdf_path, error_promedio, desviacion, primerpatron, segundopatron, primera, segunda):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    c.setFont("Arial", 11)
    c.drawString(400, 290, str(error_promedio))
    c.drawString(400, 250, str(desviacion))
    for i in range(8):
        c.drawString(205 + (i * 42), 570, str(primerpatron[i]))
    for i in range(8):
        c.drawString(205 + (i * 42), 545, str(primera[i]))
    for i in range(8):
        c.drawString(205 + (i * 42), 470, str(segundopatron[i]))
    for i in range(8):
        c.drawString(205 + (i * 42), 445, str(segunda[i]))
    c.save()
def agregar_imagenes_pdf4(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior, error_promedio, desviacion):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior = Image.open(img_superior_path1).convert("RGBA")
    image_inferior = Image.open(img_superior_path2).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    image_superior = image_superior.resize((int(image_superior.width / 5), int(image_superior.height / 5)), Image.LANCZOS)
    image_inferior = image_inferior.resize((int(image_inferior.width / 5), int(image_inferior.height / 5)), Image.LANCZOS)
    # Calcular la posición de las imágenes
    x_fondo = 0  # Colocar el fondo en 0 para ocupar toda la página
    xsuperior = (carta_ancho - image_inferior.width) // 2 - 30
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, xsuperior, yinferior, width=image_superior.width, height=image_superior.height)
    c.drawImage(img_superior_path2, xsuperior, ysuperior, width=image_inferior.width, height=image_inferior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)

    c.save()


def agregar_imagenes_pdf5(fondo_path, output_pdf_path , nota):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    c.setFont("ArialBold", 14)
    c.drawString(250, 605, "OBSERVACIONES")
    c.setFont("Arial", 12)
    max_length = 75  # Maximum characters per line
    if len(nota) > max_length:
        nota_line1 = nota[:max_length]
        nota_line2 = nota[max_length:]
        c.drawString(63, 605, nota_line1)
        c.drawString(63, 585, nota_line2)
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