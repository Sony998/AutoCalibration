from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd


archivo_excel = '/home/raven/ramiriqui.xlsx'   
df = pd.read_excel(archivo_excel, sheet_name='BASCULA PESA BEBE', header=None)
dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
fila_inicial = 0
desviaciones = []
nocertificados = []
errores_list = []
errores_promedio = []
primeras = []
segundas = []   
repetibilidades = []
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
inferior_directory = "OUTPUT/Graficos/Error"
superior_directory = "OUTPUT/Graficos/Desviacion"
errorpos = 0
while True:
    if fila_inicial >= len(df):
        break
    nombreEse = dfdatos.iat[3,1]
    fecha = dfdatos.iat[4, 1]
    metrologo = dfdatos.iat[7,1]
    temperaturaminima =dfdatos.iat[10,1]
    temperaturamaxima = dfdatos.iat[10,2]
    humedadminima = dfdatos.iat[11,1]
    humedadmaxima = dfdatos.iat[11,2]
    presionbarometrica = dfdatos.iat[12,1]
    nocertificado = df.iat[fila_inicial + 2 , 5]
    error_promedio = df.iat[fila_inicial + 9, 1]
    nota = df.iat[fila_inicial + 1, 5]
    if pd.isna(nota):
        nota = "No se realizan observaciones"
    else:
        nota = str(nota)
    desviacion = df.iat[fila_inicial + 10, 1]
    primera = df.iloc[fila_inicial + 6, 1:5].astype(float).tolist()
    segunda = df.iloc[fila_inicial + 7, 1:5].astype(float).tolist()
    repetibilidad = df.iloc[fila_inicial + 5:fila_inicial + 12, 11].astype(float).tolist()
    repetibilidades.append(repetibilidad)
    print(repetibilidad)
    incertidumbre = float(df.iat[fila_inicial + 11, 1])
    incertidumbres_expandida = float(df.iat[fila_inicial + 12, 1])
    errores_promedio.append(error_promedio)
    errores = df.iloc[fila_inicial + 8, 1:5].astype(float).tolist()
    primeras.append(primera)
    segundas.append(segunda)
    incertidumbres.append(incertidumbre)
    incertidumbres_expandidas.append(incertidumbres_expandida)
    nocertificados.append(nocertificado)
    desviaciones.append(desviacion)
    errores_list.append(errores)
    notas.append(nota)
    print(f"Certificado: {nocertificado}, Error promedio: {error_promedio}, desviacion {desviacion} nota {nota}")
    fila_inicial += 13
def agregar_imagenes_pdf1(fondo_path, output_pdf_path, nombrecertificado, fecha, nombreEse, metrologo):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(312, 664, nombrecertificado)
    c.setFont("Arial", 15)
    c.drawString(210, 240, fecha)
    c.drawString(210, 205, fecha)
    c.setFont("Arial", 12)
    c.drawString(210, 175, nombreEse)
    c.setFont("Arial", 15)
    c.drawString(210, 145, metrologo)
    c.save()
def agregar_imagenes_pdf2(img_fondo_path, output_pdf_path, incertidumbre, incertidumbre_expandida,temperaturaminima, temperaturamaxima, humedadminima, humedadmaxima, presionbarometrica):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    x_fondo = 0  # Colocar el fondo en 0 para ocupar toda la página
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("Arial", 11)
    c.setFont("ArialI", 14)
    c.drawString(330, 670, str(temperaturaminima))
    c.drawString(398, 670, str(temperaturamaxima))
    c.drawString(360, 642, str(presionbarometrica))
    c.drawString(330, 610, str(humedadminima))
    c.drawString(398, 610, str(humedadmaxima))
    c.drawString(330, 432, "{:.2f}".format(float(f"{incertidumbre_expandida:.2f}")))
    c.drawString(330, 418, "{:.2f}".format(float(f"{incertidumbre:.2f}")))
    c.save()

def agregar_imagenes_pdf4(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior, error_promedio, desviacion):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior = Image.open(img_superior_path1).convert("RGBA")
    image_inferior = Image.open(img_superior_path2).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    image_superior = image_superior.resize((int(image_superior.width / 6.8), int(image_superior.height / 6.8)), Image.LANCZOS)
    image_inferior = image_inferior.resize((int(image_inferior.width / 6), int(image_inferior.height / 6)), Image.LANCZOS)
    # Calcular la posición de las imágenes
    x_fondo = 0  # Colocar el fondo en 0 para ocupar toda la página
    xsuperior = (carta_ancho - image_inferior.width) // 2 - 30
    xinferior = (carta_ancho - image_superior.width) // 2 - 20
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, xinferior, yinferior, width=image_superior.width, height=image_superior.height)
    c.drawImage(img_superior_path2, xsuperior, ysuperior, width=image_inferior.width, height=image_inferior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    c.drawString(350, 675, "{:.2f}".format(float(f"{error_promedio:.2f}")))
    c.drawString(350, 660, "{:.2f}".format(float(f"{desviacion:.2f}")))
    c.save()


def agregar_imagenes_pdf3(fondo_path, output_pdf_path,repetibilidad, primera, segunda, errores_list):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(313, 375, "5")
    for i in range(7):
        c.drawString(313, 320 - i * 17, "{:.1f}".format(float(f"{repetibilidad[i]:.1f}")))
    c.setFont("ArialI", 10)
    for i in range(4):
        c.drawString(262 + i * 40, 88 , "{:.2f}".format(float(f"{primera[i]:.2f}")))
    for i in range(4):
        c.drawString(262 + i * 40, 63 , "{:.2f}".format(float(f"{segunda[i]:.2f}")))
    for i in range(4):
        c.drawString(262 + i * 40, 35 , "{:.2f}".format(float(f"{errores_list[i]:.2f}")))
    c.save()

def agregar_imagenes_pdf5(fondo_path, output_pdf_path , nota):
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

for certficado, error_promedio, desviacion in zip(nocertificados, errores_list, desviaciones):
    agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha, nombreEse, metrologo)
    agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), incertidumbres[errorpos], incertidumbres_expandidas[errorpos], temperaturaminima, temperaturamaxima, humedadminima, humedadmaxima, presionbarometrica)
    agregar_imagenes_pdf3(img_fondo_path3, os.path.join(output_directory3, certficado + ".pdf"), repetibilidades[errorpos], primeras[errorpos], segundas[errorpos], errores_list[errorpos])
    agregar_imagenes_pdf5(img_fondo_path5, os.path.join(output_directory5, certficado + ".pdf"), notas[errorpos] )
    agregar_imagenes_pdf4(img_fondo_path4, os.path.join(superior_directory, certficado + ".png"), os.path.join(inferior_directory, certficado + ".png"), os.path.join(output_directory4, certficado + ".pdf"), 150, 400, errores_promedio[errorpos], desviaciones[errorpos])
    errorpos += 1
