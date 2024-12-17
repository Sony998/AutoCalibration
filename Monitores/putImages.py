import textwrap
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
archivo_excel = '/home/raven/TensiometrosSantana.xlsx'   
df = pd.read_excel(archivo_excel, sheet_name='TENSIOMETRO DIGITAL', header=None)
fila_actual = 0
certificados= []
desviaciones = []
errores_list = []
frecuencia_list = []
datos_frecuencia_list = []
diastolica_list = []
datos_diastolica_list = []
sistolica_list = []
tabla_sistolica_list = []
tabla_diastolica_list = []
tabla_frecuencia_list = []
datos_sistolica_list = []
incertidumbres_list = []
notas = []
fecha = df.iat[4, 10]
ciudad = df.iat[3,10]
metrologo = df.iat[10,10]
errorsistolica_directory = "OUTPUT/Graficos/Error/Sistolica"
errordiastolica_directory = "OUTPUT/Graficos/Error/Diastolica"
errorfrecuencia_directory = "OUTPUT/Graficos/Error/Frecuencia"
desviacionsistolica_directory = "OUTPUT/Graficos/Desviacion/Sistolica"
desviaciondiastolica_directory = "OUTPUT/Graficos/Desviacion/Diastolica"
desviacionfrecuencia_directory = "OUTPUT/Graficos/Desviacion/Frecuencia"
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
errorpos = 0
while True:
    if fila_actual >= len(df):
        break
    certficado = df.iat[fila_actual + 2, 5]
    nota = df.iat[fila_actual + 1,5]
    certificados.append(certficado)
    if pd.isna(nota):
        nota = "No se realizan observaciones"
    else:
        nota = str(nota)
    notas.append(nota)
    datos_sistolica = df.loc[[fila_actual + 11, fila_actual + 12, fila_actual+14],1].astype(float)
    datos_sistolica_list.append(datos_sistolica)
    datos_diastolica = df.loc[[fila_actual + 21, fila_actual + 22, fila_actual+24],1].astype(float)
    datos_diastolica_list.append(datos_diastolica)
    datos_frecuencia = df.loc[[fila_actual + 31, fila_actual + 32, fila_actual+34],1].astype(float) 
    datos_frecuencia_list.append(datos_frecuencia)
    sistolicas_tabla = df.iloc[fila_actual+7:fila_actual+11, 1:7]
    tabla_sistolica_list.append(sistolicas_tabla)
    diastolica_tabla = df.iloc[fila_actual+17:fila_actual+21, 1:7]
    tabla_diastolica_list.append(diastolica_tabla)
    frecuencia_tabla = df.iloc[fila_actual+27:fila_actual+31, 1:5]
    tabla_frecuencia_list.append(frecuencia_tabla)
    fila_actual += 35

"""     datos_diastolica = df.loc[fila_actual + 37, [2, 4, 6, 8,10 ]].astype(float)
    datos_frecuencia = df.loc[fila_actual + 47, [2, 4, 6, 8,10 ]].astype(float) 
    sistolica =  df.iloc[fila_actual + 10, 1:7].astype(float) 
    sistolica_list.append(sistolica)
    diastolica = df.iloc[fila_actual + 20, 1:7].astype(float) 
    diastolica_list.append(diastolica)
    fdepulso = df.iloc[fila_actual + 30, 1:5].astype(float)
    fdepulso_list.append(fdepulso) """

def agregar_imagenes_pdf1(fondo_path, output_pdf_path, nombrecertificado):
    print("Generando paginas para",nombrecertificado)
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    pdfmetrics.registerFont(TTFont('Baskerville', 'Formatos/Fuentes/Baskerville-Bold.ttf'))
    c.setFont("Baskerville", 14)
    c.drawString(302, 660, nombrecertificado)
    c.setFont("Arial", 16)
    c.drawString(310, 222, fecha)
    c.drawString(310, 192, fecha)
    c.drawString(310, 162, ciudad)
    c.drawString(310, 130, metrologo)
    c.save()

"""     """

def agregar_imagenes_pdf2(fondo_path, output_pdf_path, sistolica, diastolica, frecuencia ):
    print("Generando pagina 2" )
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.setFont("Arial", 13)
    c.drawString(360, 625, "24")
    c.drawString(455, 625, "26")
    c.drawString(385, 605, "1011")
    c.drawString(360, 585, "45")
    c.drawString(455, 585, "56") 
    for i in range(3):
        c.drawString(180 + i * 115, 124 , "{:.2f}".format(float(sistolica.iloc[i])))
        c.drawString(180 + i * 115, 104 , "{:.2f}".format(float(diastolica.iloc[i])))
        c.drawString(180 + i * 115, 82 , "{:.2f}".format(float(frecuencia.iloc[i])))
    c.showPage()
    c.save()

def generar_pagina_3(fondo_path, output_pdf_path, tabla_sistolica, tabla_diastolica, tabla_frecuencia):
    print("Generando pagina 3" )
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 14)
    data = tabla_sistolica.values.tolist() if hasattr(tabla_sistolica, 'values') else tabla_sistolica
    for fila_idx, fila in enumerate(data):
        for col_idx, valor in enumerate(fila):
            c.drawString(
                190 + col_idx * 58,  # Eliminado el `+ col_idx`
                500 - fila_idx * 32,
                "{:.1f}".format(float(valor))  # Formateo con 1 decimal
            )
    data = tabla_diastolica.values.tolist() if hasattr(tabla_diastolica, 'values') else tabla_diastolica
    for fila_idx, fila in enumerate(data):
        for col_idx, valor in enumerate(fila):
            c.drawString(
                200 + col_idx * 60,  # Eliminado el `+ col_idx`
                290 - fila_idx * 28,
                "{:.1f}".format(float(valor))  # Formateo con 1 decimal
            )    
    c.setFontSize(13)
    data = tabla_frecuencia.values.tolist() if hasattr(tabla_frecuencia, 'values') else tabla_frecuencia
    for fila_idx, fila in enumerate(data):
        for col_idx, valor in enumerate(fila):
            c.drawString(
                250 + col_idx * 80,  # Eliminado el `+ col_idx`
                115 - fila_idx * 20,
                "{:.1f}".format(float(valor))  # Formateo con 1 decimal
            )    
    c.showPage()
    c.save()


""" def agregar_imagenes_pdf3(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior):
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
    c.save() """

def agregar_imagenes_pdf4(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior):
    print("Generando pagina 4" )
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

def agregar_imagenes_pdf5(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior):
    print("Generando pagina 5" )
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
    print("Generando pagina 6" )
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
    
def agregar_imagenes_pdf7(fondo_path, output_pdf_path, nota):
    print("Generando pagina 7" )
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    c.setFont("Arial", 12)
    max_width = 80  # Longitud máxima en caracteres aproximados
    line_height = 15  # Espaciado entre líneas
    wrapped_lines = textwrap.wrap(nota, width=max_width)
    x = 90
    y = 560
    for line in wrapped_lines:
        c.drawString(x, y, line)
        y -= line_height 
    c.save()

for certficado in certificados:
    agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado)
    agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), datos_sistolica_list[errorpos], datos_diastolica_list[errorpos], datos_frecuencia_list[errorpos])
    generar_pagina_3(img_fondo_path3, os.path.join(output_directory3, certficado + ".pdf"), tabla_sistolica_list[errorpos], tabla_diastolica_list[errorpos], tabla_frecuencia_list[errorpos])
    imagen_error_sistolica = os.path.join(errorsistolica_directory, certficado + ".png")
    imagen_desviacion_sistolica = os.path.join(desviacionsistolica_directory, certficado + ".png")
    agregar_imagenes_pdf4(img_fondo_path4, imagen_error_sistolica, imagen_desviacion_sistolica, output_directory4 + certficado +".pdf", yinferior=50, ysuperior=390)
    imagen_error_diastolica = os.path.join(errordiastolica_directory, certficado + ".png")
    imagen_desviacion_diastolica = os.path.join(desviaciondiastolica_directory, certficado + ".png")
    agregar_imagenes_pdf5(img_fondo_path5, imagen_error_diastolica, imagen_desviacion_diastolica, output_directory5 + certficado +".pdf", yinferior=50, ysuperior=390)
    imagen_error_frecuencia = os.path.join(errorfrecuencia_directory, certficado + ".png")
    imagen_desviacion_frecuencia = os.path.join(desviacionfrecuencia_directory, certficado + ".png")
    agregar_imagenes_pdf6(img_fondo_path6, imagen_error_frecuencia, imagen_desviacion_frecuencia, output_directory6 + certficado +".pdf", yinferior=35, ysuperior=375)
    agregar_imagenes_pdf7(img_fondo_path7, os.path.join(output_directory7, certficado + ".pdf"), notas[errorpos])
    errorpos += 1
