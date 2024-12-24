from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
sheetname = "TERMOMETRO INFRARROJO"


archivo_excel = '/home/raven/SANTANA.xlsx'
df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
fila_inicial = 0
desviaciones = []
certificados = []
errores_list = []
errores_promedio = []
primeras = []
segundas = []   
incertidumbres_expandidas = []
incertidumbres = []
notas = []
fecha = str(df.iat[1, 12])
img_fondo_path1 = "Formatos/partesReporte/Pagina1.png"
img_fondo_path2 = "Formatos/partesReporte/Pagina2.png"
img_fondo_path3 = "Formatos/partesReporte/Pagina3.png"
img_fondo_path4 = "Formatos/partesReporte/Pagina4.png"
output_directory1 = "OUTPUT/Reportes/1"
output_directory2 = "OUTPUT/Reportes/2"
output_directory3 = "OUTPUT/Reportes/3"
output_directory4 = "OUTPUT/Reportes/4"
inferior_directory = "OUTPUT/Graficos/Error"
superior_directory = "OUTPUT/Graficos/Desviacion"
errorpos = 0
while True:
    if fila_inicial >= len(df):
        break
    nocertificado = df.iat[fila_inicial + 2 , 5]
    error_promedio = df.iat[fila_inicial + 9, 1]
    nota = df.iat[fila_inicial + 1, 5]
    if pd.isna(nota):
        nota = "No se realizan observaciones"
    else:
        nota = str(nota)
    fecha = df.iat[5, 15]
    desviacionestandar = df.iat[fila_inicial + 10, 1]
    print(nocertificado, nota)
    primera = df.iloc[fila_inicial + 6, 1:12].astype(float).tolist()
    segunda = df.iloc[fila_inicial + 7, 1:12].astype(float).tolist()
    incertidumbre =  df.iat[fila_inicial + 11, 1]
    incertidumbres_expandida =  df.iat[fila_inicial + 12, 1]
    errores_promedio.append(error_promedio)
    errores = df.iloc[fila_inicial + 8, 1:12].astype(float).tolist()
    primeras.append(primera)
    segundas.append(segunda)
    incertidumbres.append(incertidumbre)
    incertidumbres_expandidas.append(incertidumbres_expandida)
    certificados.append(nocertificado)
    desviaciones.append(desviacionestandar)
    errores_list.append(errores)
    notas.append(nota)
    fila_inicial += 13

def agregar_imagenes_pdf1(fondo_path, output_pdf_path, nombrecertificado, fecha):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(315, 665, nombrecertificado)
    c.setFont("Arial", 15)
    c.drawString(270, 210, fecha)
    c.drawString(270, 176, fecha)
    c.drawString(270, 148, "Santana, Boyaca")
    c.drawString(270, 115, "Ingeniera Luz Alejandra Vargas")
    c.save()
def agregar_imagenes_pdf2(fondo_path, output_pdf_path, nombrecertificado, incertidumbre, incertidumbre_expandida, primera, segunda, errores_list):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(330, 650, "22.5")
    c.drawString(405, 650, "26.8")
    c.drawString(355, 610, "1012")
    c.drawString(330, 570, "59")
    c.drawString(410, 570, "68")
    c.drawString(330, 410, "{:.2f}".format(float(f"{incertidumbre_expandida:.2f}")))
    c.drawString(330, 392, "{:.2f}".format(float(f"{incertidumbre:.2f}")))
    c.setFont("ArialI", 10)
    margen_inicial = 80  # Distancia desde el borde izquierdo
    ancho_columna = 46  # Espaciado horizontal uniforme entre columnas

    # Dibujar la primera fila
    for i in range(11):
        c.drawString(margen_inicial + i * ancho_columna + i , 140, "{:.2f}".format(primera[i]))

    # Dibujar la segunda fila
    for i in range(11):
        c.drawString(margen_inicial + i * ancho_columna + i, 115, "{:.2f}".format(segunda[i]))

    # Dibujar la tercera fila con errores
    c.setFont("ArialI", 11)
    for i in range(11):
        c.drawString(margen_inicial + i * ancho_columna + i, 90, "{:.2f}".format(errores_list[i]))

    c.save()


def agregar_imagenes_pdf3(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior, error_promedio, desviacion):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior = Image.open(img_superior_path1).convert("RGBA")
    image_inferior = Image.open(img_superior_path2).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    image_superior = image_superior.resize((int(image_superior.width / 6), int(image_superior.height / 6)), Image.LANCZOS)
    image_inferior = image_inferior.resize((int(image_inferior.width / 6), int(image_inferior.height / 6)), Image.LANCZOS)
    # Calcular la posición de las imágenes
    x_fondo = 0  # Colocar el fondo en 0 para ocupar toda la página
    xsuperior = (carta_ancho - image_superior.width) // 2 - 30
    xinferior = (carta_ancho - image_inferior.width) // 2 - 30
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path2, xinferior, yinferior, width=image_inferior.width, height=image_inferior.height)
    c.drawImage(img_superior_path1, xsuperior, ysuperior, width=image_superior.width, height=image_superior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    c.drawString(350, 690, "{:.2f}".format(float(f"{error_promedio:.2f}")))
    c.drawString(350, 672, "{:.2f}".format(float(f"{desviacion:.2f}")))
    c.save()

def agregar_imagenes_pdf4(fondo_path, output_pdf_path , nota):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    c.setFont("ArialBold", 14)
    c.drawString(220, 660, "OBSERVACIONES")
    c.setFont("Arial", 12)
    max_length = 75  # Maximum characters per line
    if len(nota) > max_length:
        nota_line1 = nota[:max_length]
        nota_line2 = nota[max_length:]
        c.drawString(63, 630, nota_line1)
        c.drawString(63, 610, nota_line2)
    else:
        c.drawString(63, 630, nota)
    c.save()

print(errores_list)
for certficado, error_promedio, desviacion in zip(certificados, errores_list, desviaciones):
    agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha)
    img_superior_path1 = os.path.join(inferior_directory, certficado + ".png")
    img_superior_path2 = os.path.join(superior_directory, certficado + ".png")
    output_pdf_path = os.path.join(output_directory3, certficado + ".pdf")
    agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), certficado, incertidumbres[errorpos], incertidumbres_expandidas[errorpos], primeras[errorpos], segundas[errorpos], errores_list[errorpos])
    agregar_imagenes_pdf3(img_fondo_path3, img_superior_path1, img_superior_path2, output_pdf_path, 
                         yinferior=140, ysuperior=405, error_promedio=errores_promedio[errorpos], desviacion=desviaciones[errorpos])
    agregar_imagenes_pdf4(img_fondo_path4, os.path.join(output_directory4, certficado + ".pdf"), notas[errorpos] )

    errorpos += 1
    print(f"Se ha creado el archivo {output_pdf_path} con error promedio de {error_promedio}")