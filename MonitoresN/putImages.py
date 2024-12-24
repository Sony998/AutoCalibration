from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd


archivo_excel = '/home/raven/Quipama.xlsx'   
df = pd.read_excel(archivo_excel, sheet_name='MONITORES MULTIPARAMETROS', header=None)
dfdatos = pd.read_excel(archivo_excel, sheet_name='DATOS SOLICITANTE', header=None)

fila_actual = 0
certificados= []
desviaciones = []
errores_list = []
errores_promedio = []
ecg_list = []
fdepulso_list = []
diastolica_list = []
sistolica_list = []
saturacion_list = []
fp_list = []
respiracion_list = []
incertidumbres_list = []
notas = []
errorsistolica_directory = "OUTPUT/Graficos/Error/Sistolica"
errordiastolica_directory = "OUTPUT/Graficos/Error/Diastolica"
errorfrecuencia_directory = "OUTPUT/Graficos/Error/Frecuencia"
errorsaturacion_directory = "OUTPUT/Graficos/Error/Saturacion"
errorpulso_directory = "OUTPUT/Graficos/Error/Pulso"


desviacionsistolica_directory = "OUTPUT/Graficos/Desviacion/Sistolica"
desviaciondiastolica_directory = "OUTPUT/Graficos/Desviacion/Diastolica"
desviacionfrecuencia_directory = "OUTPUT/Graficos/Desviacion/Frecuencia"
desviacionsaturacion_directory = "OUTPUT/Graficos/Desviacion/Saturacion"
desviacionpulso_directory = "OUTPUT/Graficos/Desviacion/Pulso"

img_fondo_path1 = "Formatos/partesReporte/Pagina1.png"
img_fondo_path2 = "Formatos/partesReporte/Pagina2.png"
img_fondo_path3 = "Formatos/partesReporte/Pagina3.png"
img_fondo_path4 = "Formatos/partesReporte/Pagina4.png"
img_fondo_path5 = "Formatos/partesReporte/Pagina5.png"
img_fondo_path6 = "Formatos/partesReporte/Pagina6.png"
img_fondo_path7 = "Formatos/partesReporte/Pagina7.png"
output_directory1 = "OUTPUT/Reportes/1/"
output_directory2 = "OUTPUT/Reportes/2/"
output_directory3 = "OUTPUT/Reportes/3/"
output_directory4 = "OUTPUT/Reportes/4/"
output_directory5 = "OUTPUT/Reportes/5/"
output_directory6 = "OUTPUT/Reportes/6/"
output_directory7 = "OUTPUT/Reportes/7/"
output_directory8 = "OUTPUT/Reportes/8/"
inferior_directory = "OUTPUT/Graficos/Error"
superior_directory = "OUTPUT/Graficos/Desviacion"
errorpos = 0
while True:
    if fila_actual >= len(df):
        break
    certficado= df.iat[fila_actual + 2, 5]
    certificados.append(certficado)
    nombreEse = dfdatos.iat[3,1]
    fecha = dfdatos.iat[4, 1]
    metrologo = dfdatos.iat[7,1]
    temperaturaminima =dfdatos.iat[10,1]
    temperaturamaxima = dfdatos.iat[10,2]
    humedadminima = dfdatos.iat[11,1]
    humedadmaxima = dfdatos.iat[11,2]
    presionbarometrica = dfdatos.iat[12,1]
    print(certficado)
    nota = df.iat[fila_actual + 1,5]
    if pd.isna(nota):
        nota = "No se realizan observaciones"
    else:
        nota = str(nota)
    print(nota)
    notas.append(nota)
    ecg = df.iloc[fila_actual+39:fila_actual + 42,1].values.astype(float)
    print(ecg)
    ecg_list.append(ecg)
    fdepulso = df.iloc[fila_actual + 31:fila_actual+ 34, 1].values.astype(float)
    fdepulso_list.append(fdepulso)
    diastolica = df.iloc[fila_actual + 21:fila_actual + 24, 1].values.astype(float)
    diastolica_list.append(diastolica)
    sistolica = df.iloc[fila_actual + 11: fila_actual+ 14, 1].values.astype(float)
    sistolica_list.append(sistolica)
    saturacion = df.iloc[fila_actual + 57: fila_actual + 60, 1].values.astype(float)
    saturacion_list.append(saturacion)
    fp = df.iloc[fila_actual + 66:fila_actual + 69, 1].values.astype(float)
    fp_list.append(fp)
    respiracion = df.iloc[fila_actual + 47: fila_actual +50, 1].values.astype(float)
    respiracion_list.append(respiracion)
    fila_actual += 70


def agregar_imagenes_pdf1(fondo_path, output_pdf_path, nombrecertificado, metrologo, fecha, ubicacion,
                          
                           ):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(305, 690, nombrecertificado)
    c.setFont("Arial", 12)
    c.drawString(325, 322, fecha)
    c.drawString(325, 302, fecha)
    c.setFont("Arial", 9)
    c.drawString(325, 282, ubicacion)
    c.setFont("Arial", 12)
    c.drawString(325, 262, metrologo)
    c.setFont("Arial", 13)
    c.drawString(335, 160, str(temperaturaminima))
    c.drawString(425, 160, str(temperaturamaxima))
    c.drawString(375, 145, str(presionbarometrica))
    c.drawString(340, 125, str(humedadminima))
    c.drawString(430, 125, str(humedadmaxima))
    c.save()
    
def agregar_imagenes_pdf2(fondo_path, output_pdf_path, ecg_list, fdepulso_list, diastolica_list, sistolica_list, saturacion_list, fp_list, respiracion_list): 
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    for i in range(3):
        c.drawString(195 + i * 128, 302 , "{:.2f}".format(float(ecg_list[i])))
    for i in range(3):
        c.drawString(195 + i * 128, 282 , "{:.2f}".format(float(fdepulso_list[i])))
    for i in range(3):
        c.drawString(195 + i * 128, 265 , "{:.2f}".format(float(diastolica_list[i])))
    for i in range(3):
        c.drawString(195 + i * 128, 246 , "{:.2f}".format(float(sistolica_list[i])))
    for i in range(3):
        c.drawString(195 + i * 128, 226 , "{:.2f}".format(float(saturacion_list[i])))
    for i in range(3):
        c.drawString(195 + i * 128, 207 , "{:.2f}".format(float(fp_list[i])))
    for i in range(3):
        c.drawString(195 + i * 128, 190 , "{:.2f}".format(float(respiracion_list[i])))
    c.showPage()
    c.save()
    print("Tabla agregada")
    
def agregar_imagenes_pdf3(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior = Image.open(img_superior_path1).convert("RGBA")
    image_inferior = Image.open(img_superior_path2).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")
    # Increase the size of the images by reducing the division factor
    image_superior = image_superior.resize((int(image_superior.width / 5), int(image_superior.height / 5)), Image.LANCZOS)
    image_inferior = image_inferior.resize((int(image_inferior.width / 5), int(image_inferior.height / 5)), Image.LANCZOS)
    x_fondo = 0
    xsuperior = (carta_ancho - image_superior.width) // 2 - 30
    xinferior = (carta_ancho - image_inferior.width) // 2 - 30
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, xsuperior, ysuperior, width=image_superior.width, height=image_superior.height)
    c.drawImage(img_superior_path2, xinferior, yinferior, width=image_inferior.width, height=image_inferior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    c.save()
def agregar_imagenes_pdf5(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior = Image.open(img_superior_path1).convert("RGBA")
    image_inferior = Image.open(img_superior_path2).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")
    image_superior = image_superior.resize((int(image_superior.width / 5), int(image_superior.height / 5)), Image.LANCZOS)
    image_inferior = image_inferior.resize((int(image_inferior.width / 5), int(image_inferior.height / 5)), Image.LANCZOS)
    x_fondo = 0
    xsuperior = (carta_ancho - image_superior.width) // 2 - 30
    xinferior = (carta_ancho - image_inferior.width) // 2 - 30
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, xsuperior, ysuperior, width=image_superior.width, height=image_superior.height)
    c.drawImage(img_superior_path2, xinferior, yinferior, width=image_inferior.width, height=image_inferior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    c.save()
def agregar_imagenes_pdf6(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior = Image.open(img_superior_path1).convert("RGBA")
    image_inferior = Image.open(img_superior_path2).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")
    image_superior = image_superior.resize((int(image_superior.width / 5), int(image_superior.height / 5)), Image.LANCZOS)
    image_inferior = image_inferior.resize((int(image_inferior.width / 5), int(image_inferior.height / 5)), Image.LANCZOS)
    x_fondo = 0
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
    x_fondo = 0
    xsuperior = (carta_ancho - image_superior.width) // 2 - 30
    xinferior = (carta_ancho - image_inferior.width) // 2 - 30
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, xsuperior, ysuperior, width=image_superior.width, height=image_superior.height)
    c.drawImage(img_superior_path2, xinferior, yinferior, width=image_inferior.width, height=image_inferior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    c.save()
    
def agregar_imagenes_pdf7(fondo_path, output_pdf_path , nota):
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
        c.drawString(70, 620, nota)
    c.save()

""" 


def agregar_imagenes_pdf4(fondo_path, output_pdf_path , nota):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'ArialBold.ttf'))
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
    c.save() """




for certficado in certificados:
    agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, metrologo, fecha, nombreEse)
    imagen_error_sistolica = os.path.join(errorsistolica_directory, certficado + ".png")
    imagen_desviacion_sistolica = os.path.join(desviacionsistolica_directory, certficado + ".png")
    agregar_imagenes_pdf3(img_fondo_path3, imagen_error_sistolica, imagen_desviacion_sistolica, output_directory3 + certficado +".pdf", yinferior=60, ysuperior=410)     #  agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado)
    imagen_error_diastolica = os.path.join(errordiastolica_directory, certficado + ".png")
    imagen_desviacion_diastolica = os.path.join(desviaciondiastolica_directory, certficado + ".png")
    agregar_imagenes_pdf4(img_fondo_path4, imagen_error_diastolica, imagen_desviacion_diastolica, output_directory4 + certficado +".pdf", yinferior=60, ysuperior=410)
    imagen_error_frecuencia = os.path.join(errorfrecuencia_directory, certficado + ".png") 
    imagen_desviacion_frecuencia = os.path.join(desviacionfrecuencia_directory, certficado + ".png")
    agregar_imagenes_pdf5(img_fondo_path5, imagen_error_frecuencia, imagen_desviacion_frecuencia, output_directory5 + certficado +".pdf", yinferior=60, ysuperior=410)
    imagen_error_saturacion = os.path.join(errorsaturacion_directory, certficado + ".png")
    imagen_desviacion_saturacion = os.path.join(desviacionsaturacion_directory, certficado + ".png")
    agregar_imagenes_pdf6(img_fondo_path6, imagen_error_saturacion, imagen_desviacion_saturacion, output_directory6 + certficado +".pdf", yinferior=60, ysuperior=410)
    agregar_imagenes_pdf7(img_fondo_path7, os.path.join(output_directory7, certficado + ".pdf"), notas[errorpos] )
    agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), ecg_list[errorpos], fdepulso_list[errorpos], diastolica_list[errorpos], sistolica_list[errorpos], saturacion_list[errorpos], fp_list[errorpos], respiracion_list[errorpos])
   # agregar_imagenes_pdf4(img_fondo_path4, os.path.join(output_directory4, certficado + ".pdf"), notas[errorpos] )
  ##  img_superior_path1 = os.path.join(inferior_directory, certficado + ".png")
    #img_superior_path2 = os.path.join(superior_directory, certficado + ".png")
    #output_pdf_path = os.path.join(output_directory3, certficado + ".pdf")
   # agregar_imagenes_pdf3(img_fondo_path3, img_superior_path1, img_superior_path2, output_pdf_path, 
    #                     yinferior=125, ysuperior=378, error_promedio=errores_promedio[errorpos], desviacion=desviacion)
    errorpos += 1

