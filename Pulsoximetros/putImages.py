from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
sheetname = "PULSO OXIMETRO"


archivo_excel = '/home/raven/Quipama.xlsx'
df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
fila_inicial = 0
desviacion_pulso_list = []
desvaicion_saturacion_list = []
certificados = []
errores_promedio_saturacion_list = []
errores_promedio_pulso_list = []
errores_pulso_list = []
errores_saturacion_list = []
primeras_saturacion_list = []
primeras_pulso_list = []
incertidumbres_expandidas = []
incertidumbres = []
notas = []
nombreEse = dfdatos.iat[3,1]
fecha = dfdatos.iat[4, 1]
metrologo = dfdatos.iat[7,1]
temperaturaminima =dfdatos.iat[10,1]
temperaturamaxima = dfdatos.iat[10,2]
humedadminima = dfdatos.iat[11,1]
humedadmaxima = dfdatos.iat[11,2]
presionbarometrica = dfdatos.iat[12,1]
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

errorpulso_directory = "OUTPUT/Graficos/Pulso/Error"
desviacionpulso_directory = "OUTPUT/Graficos/Pulso/Desviacion"
errorsaturacion_directory = "OUTPUT/Graficos/Saturacion/Error"
desviacionsaturacion_directory = "OUTPUT/Graficos/Saturacion/Desviacion"
errorpos = 0
while True:
    if fila_inicial >= len(df):
        break
    nocertificado = df.iat[fila_inicial + 2 , 5]
    error_promedio_saturacion = df.iat[fila_inicial + 9, 1]
    desviacion_saturacion = df.iat[fila_inicial + 10, 1]
    desviacion_pulso = df.iat[fila_inicial + 18, 1]
    errores_promedio_pulso = df.iat[fila_inicial + 17, 1]
    nota = df.iat[fila_inicial + 1, 5]
    if pd.isna(nota):
        nota = "No se realizan observaciones"
    else:
        nota = str(nota)
    desviacionestandar = df.iat[fila_inicial + 10, 1]
    print(nocertificado, nota)
    saturacion = df.iloc[fila_inicial + 7, 1:6].astype(float).tolist()
    pulso = df.iloc[fila_inicial + 15, 1:12].astype(float).tolist()
    incertidumbre =  df.iat[fila_inicial + 11, 1]
    incertidumbres_expandida =  df.iat[fila_inicial + 12, 1]
    errores_promedio_saturacion_list.append(error_promedio_saturacion)
    errores_promedio_pulso_list.append(errores_promedio_pulso)
    errores_saturacion = df.iloc[fila_inicial + 8, 1:6].astype(float).tolist()
    errores_pulso = df.iloc[fila_inicial + 16, 1:6].astype(float).tolist()
    errores_saturacion_list.append(errores_saturacion)
    errores_pulso_list.append(errores_pulso)
    primeras_saturacion_list.append(saturacion)
    primeras_pulso_list.append(pulso)
    incertidumbres.append(incertidumbre)
    incertidumbres_expandidas.append(incertidumbres_expandida)
    certificados.append(nocertificado)
    desviacion_pulso_list.append(desviacion_pulso)
    desvaicion_saturacion_list.append(desviacion_saturacion)
    errores_pulso_list.append(errores_pulso)
    errores_saturacion_list.append(errores_saturacion)
    notas.append(nota)
    fila_inicial += 21

def agregar_imagenes_pdf1(fondo_path, output_pdf_path, nombrecertificado, fecha):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(355, 685, nombrecertificado)
    c.setFont("Arial", 12)
    c.drawString(345, 292, fecha)
    c.drawString(345, 270, fecha)
    c.drawString(345, 246, nombreEse)
    c.drawString(345, 226, metrologo)
    c.setFont("ArialI", 14)
    c.drawString(345, 135, str(temperaturaminima))
    c.drawString(460, 135, str(temperaturamaxima))
    c.drawString(390, 110, str(presionbarometrica))
    c.drawString(345, 85, str(humedadminima))
    c.drawString(460, 85, str(humedadmaxima))
    c.save()
def agregar_imagenes_pdf2(fondo_path, output_pdf_path, nombrecertificado, incertidumbre, incertidumbre_expandida, primeras_saturacion, errores_saturacion, primeras_pulso, errores_pulso, desviacion_saturacion, desviacion_pulso, error_promedio_saturacion, error_promedio_pulso):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(285, 85, "1.8")
    c.drawString(440, 85, "2.2")
    c.drawString(285, 135, "{:.2f}".format(float(f"{error_promedio_saturacion:.2f}")))
    c.drawString(285, 110, "{:.2f}".format(float(f"{desviacion_saturacion:.2f}")))
    c.drawString(440, 135, "{:.2f}".format(float(f"{error_promedio_pulso:.2f}")))
    c.drawString(440, 110, "{:.2f}".format(float(f"{desviacion_pulso:.2f}")))
    c.setFont("ArialI", 10)
    for i in range(5):
        c.drawString(240 + i * 45, 345, "{:.2f}".format(float(f"{primeras_saturacion[i]:.2f}")))
    for i in range(5):
        c.drawString(240 + i * 45, 320 , "{:.2f}".format(float(f"{errores_saturacion[i]:.2f}")))
    for i in range(5):
        c.drawString(240 + i * 45, 260 , "{:.2f}".format(float(f"{primeras_pulso[i]:.2f}")))
    for i in range(5):
        c.drawString(240 + i * 45, 230 , "{:.2f}".format(float(f"{errores_pulso[i]:.2f}")))
    c.setFont("Arial", 11)
    c.save()
def agregar_imagenes_pdf3(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior):
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
    xsuperior = (carta_ancho - image_superior.width) // 2 - 30
    xinferior = (carta_ancho - image_inferior.width) // 2 - 30
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, xsuperior, ysuperior, width=image_superior.width, height=image_superior.height)
    c.drawImage(img_superior_path2, xinferior, yinferior, width=image_inferior.width, height=image_inferior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
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
    x_fondo = 0  # Colocar el fondo en 0 para ocupar toda la página
    xsuperior = (carta_ancho - image_superior.width) // 2 - 30
    xinferior = (carta_ancho - image_inferior.width) // 2 - 30
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, xsuperior, ysuperior, width=image_superior.width, height=image_superior.height)
    c.drawImage(img_superior_path2, xinferior, yinferior, width=image_inferior.width, height=image_inferior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    c.save()

def agregar_imagenes_pdf5(fondo_path, output_pdf_path , nota):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    c.setFont("Arial", 12)
    max_length = 75  # Maximum characters per line
    if len(nota) > max_length:
        nota_line1 = nota[:max_length]
        nota_line2 = nota[max_length:]
        c.drawString(63, 630, nota_line1)
        c.drawString(63, 610, nota_line2)
    else:
        c.drawString(88, 620, nota)
    c.save()

for certficado, error_promedio_saturacion, desviacion_saturacion, desviacion_pulso, error_promedio_pulso in zip(certificados, errores_promedio_saturacion_list, desvaicion_saturacion_list, desviacion_pulso_list, errores_promedio_pulso_list):
    agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha)
    imagen_error_saturacion = os.path.join(errorsaturacion_directory, certficado + ".png")
    imagen_desviacion_saturacion = os.path.join(desviacionsaturacion_directory, certficado + ".png")
    imagen_error_pulso = os.path.join(errorpulso_directory, certficado + ".png")
    imagen_desviacion_pulso = os.path.join(desviacionpulso_directory, certficado + ".png")
    output_pdf_path = os.path.join(output_directory3, certficado + ".pdf")
    agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), certficado, incertidumbres[errorpos], incertidumbres_expandidas[errorpos], primeras_saturacion_list[errorpos], errores_saturacion_list[errorpos], primeras_pulso_list[errorpos], errores_pulso_list[errorpos], desvaicion_saturacion_list[errorpos], desviacion_pulso_list[errorpos], errores_promedio_saturacion_list[errorpos], errores_promedio_pulso_list[errorpos])
    agregar_imagenes_pdf3(img_fondo_path3, imagen_error_saturacion, imagen_desviacion_saturacion, output_pdf_path, 
                         yinferior=50, ysuperior=400)
    output_pdf_path = os.path.join(output_directory4, certficado + ".pdf")
    agregar_imagenes_pdf4(img_fondo_path4,imagen_error_pulso, imagen_desviacion_pulso, output_pdf_path, 
                         yinferior=50, ysuperior=400)
    agregar_imagenes_pdf5(img_fondo_path5, os.path.join(output_directory5, certficado + ".pdf"), notas[errorpos] )
    errorpos += 1
