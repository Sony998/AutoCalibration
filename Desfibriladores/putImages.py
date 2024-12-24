from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
import argparse
sheetname = "DESFIBRILADOR"
desviaciones = []
certificados = []
errores_list = []
errores_promedio = []
incertidumbres = []
patrones = []
patroncarga_list = []
segundos_list = []
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
inferior_directory = "OUTPUT/Graficos/Error"
superior_directory = "OUTPUT/Graficos/Desviacion"
def crear_paginas(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    print(df)
    fila_inicial = 0
    errorpos = 0
    while True:
        if fila_inicial >= len(df):
            print("Se han creado todos los reportes.")
            break
        nombreEse = dfdatos.iat[3,1]
        fecha = dfdatos.iat[4, 1]
        direccion = dfdatos.iat[6, 1]
        metrologo = dfdatos.iat[7,1]
        temperaturaminima =dfdatos.iat[10,1]
        temperaturamaxima = dfdatos.iat[10,2]
        humedadminima = dfdatos.iat[11,1]
        humedadmaxima = dfdatos.iat[11,2]
        presionbarometrica = dfdatos.iat[12,1]
        certificado = str(df.iat[fila_inicial + 2, 5])
        if certificado == "nan":
            break
        patron = df.iloc[fila_inicial + 13, 1:7].astype(int).tolist()
        patroncarga = df.iloc[fila_inicial + 25, 1:7].fillna(0).astype(int).tolist()
        print(patroncarga)
        segundos = df.iloc[fila_inicial + 26, 1:7].astype(float).tolist()
        print(segundos)
        tiempopromedio = df.iat[fila_inicial + 27, 1]
        segundos_list.append(segundos)
        patrones.append(patron)
        patroncarga_list.append(patroncarga)
        tiempopromedio_list.append(tiempopromedio)
        print(patron)
        error_promedio = df.iat[fila_inicial + 20, 1]
        nota = df.iat[fila_inicial + 1, 5]
        if pd.isna(nota):
            nota = "No se realizan observaciones"
        else:
            nota = str(nota)
        desviacionestandar = df.iat[fila_inicial + 21, 1]
        incertidumbre =  df.iat[fila_inicial + 22, 1]
        errores_promedio.append(error_promedio)
        errores = df.iloc[fila_inicial + 19, 1:7].astype(float).tolist()
        if len(errores) < 5:
            errores.extend([0.0] * (5 - len(errores)))
        incertidumbres.append(incertidumbre)
        certificados.append(certificado)
        desviaciones.append(desviacionestandar)
        errores_list.append(errores)
        notas.append(nota)
        fila_inicial += 28
    for certficado in certificados:
        agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha, metrologo, nombreEse)
        img_superior_path1 = os.path.join(inferior_directory, certficado + ".png")
        img_superior_path2 = os.path.join(superior_directory, certficado + ".png")
        output_pdf_path = os.path.join(output_directory3, certficado + ".pdf")
        agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), certficado, patrones[errorpos], errores_list[errorpos], temperaturaminima, temperaturamaxima, presionbarometrica, humedadminima, humedadmaxima)
        agregar_imagenes_pdf3(img_fondo_path3, output_pdf_path, errores_promedio[errorpos], desviaciones[errorpos], incertidumbres[errorpos], patroncarga_list[errorpos], segundos_list[errorpos], tiempopromedio_list[errorpos])
        agregar_imagenes_pdf4(img_fondo_path4, img_superior_path1, img_superior_path2, os.path.join(output_directory4, certficado + ".pdf"), 400, 100)
        print("Se ha creado el reporte completo para el desfibrilador con el certificado: ", certficado) 
        agregar_imagenes_pdf5(img_fondo_path5, os.path.join(output_directory5, certficado + ".pdf"), notas[errorpos])
        errorpos += 1
        if errorpos >= len(certificados):
            print("Se han creado todos los reportes.")
            #os.system("python3 UnirPartes.py")
            break

def agregar_imagenes_pdf1(fondo_path, output_pdf_path, nombrecertificado, fecha, metrologo, ubicacion):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    pdfmetrics.registerFont(TTFont('Baskerville', 'Formatos/Fuentes/Baskerville-Bold.ttf'))
    c.setFont("Baskerville", 12)
    c.drawString(330, 665, nombrecertificado)
    c.setFont("Arial", 15)
    c.drawString(285, 210, str(fecha))  # desplazado 20 unidades hacia arriba
    c.drawString(285, 180, str(fecha))  # desplazado 20 unidades hacia arriba
    c.setFont("Arial", 12)
    c.drawString(285, 150, str(ubicacion))  # desplazado 20 unidades hacia arriba
    c.setFont("Arial", 15)
    c.drawString(285, 120, str(metrologo))  # desplazado 20 unidades hacia arriba
    c.save()
def agregar_imagenes_pdf2(fondo_path, output_pdf_path, nombrecertificado,  patron, errores, temperaturaminima, temperaturamaxima, presionbarometrica, humedadminima, humedadmaxima):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(345, 640, str(temperaturaminima))
    c.drawString(438, 640, str(temperaturamaxima))
    c.drawString(390, 615, str(presionbarometrica))
    c.drawString(345, 595, str(humedadminima))
    c.drawString(438, 595, str(humedadmaxima))
    c.setFont("ArialI", 10)
    for i in range(6):
        c.drawString(260 + i * 48, 130 , "{:.2f}".format(float(f"{patron[i]:.2f}")))
    for i in range(6):
        c.drawString(260 + i * 48,110 , "{:.2f}".format(float(f"{errores[i]:.2f}")))
    c.save()

def agregar_imagenes_pdf3(img_fondo_path, output_pdf_path, error_promedio, desviacion, incertidumbre, patroncarga, segundos, tiempo_promedio):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(img_fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 18)
    c.drawString(400, 285, "{:.2f}".format(float(f"{error_promedio:.2f}")))
    c.drawString(400, 250, "{:.2f}".format(float(f"{desviacion:.2f}")))
    c.drawString(400, 210, "{:.2f}".format(float(f"{incertidumbre:.2f}")))
    c.setFont("Arial", 12)
    for i in range(9):
        c.drawString(140 + i * 45, 630 , "0.00")
    c.drawString(140, 600, "0.00")
    for i in range(6):
        c.drawString(170 + i * 65, 530 , "{:.2f}".format(float(f"{patroncarga[i]:.2f}")))   
        c.drawString(170 + i * 65, 500 , "{:.2f}".format(float(f"{segundos[i]:.2f}")))
    c.drawString(170, 470, "{:.2f}".format(float(f"{tiempo_promedio:.2f}")))
    c.save()


def agregar_imagenes_pdf4(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior):
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