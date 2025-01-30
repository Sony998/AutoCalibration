from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
import argparse
sheetname = "ELECTROESTIMULADOR DE 4 CANALES"
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
tipo_corriente_1_list = []
tipo_corriente_2_list = []
tipo_corriente_3_list = []
tipo_corriente_4_list = []
voltaje_1_list = []
voltaje_2_list = []
voltaje_3_list = []
voltaje_4_list = []
corriente_1_list = []
corriente_2_list = []
corriente_3_list = []
corriente_4_list = []
frecuencia_1_1_list = []
frecuencia_2_1_list = []
frecuencia_1_2_list = []
frecuencia_2_2_list = []
frecuencia_1_3_list = []
frecuencia_2_3_list = []
frecuencia_1_4_list = []
frecuencia_2_4_list = []
tipo_corriente_normal_1_list = []
tipo_corriente_normal_2_list = []
tipo_corriente_normal_3_list = []
tipo_corriente_normal_4_list = []
voltaje_normal_1_list = []
voltaje_normal_2_list = []
voltaje_normal_3_list = []
voltaje_normal_4_list = []
corriente_normal_1_list = []
corriente_normal_2_list = []
corriente_normal_3_list = []
corriente_normal_4_list = []
frecuencia_1_normal_1_list = []
frecuencia_2_normal_1_list = []
frecuencia_1_normal_2_list = []
frecuencia_2_normal_2_list = []
frecuencia_1_normal_3_list = []
frecuencia_2_normal_3_list = []
frecuencia_1_normal_4_list = []
frecuencia_2_normal_4_list = []

notas = []
img_fondo_path1 = "Formatos/partesReporte/Pagina1.png"
img_fondo_path2 = "Formatos/partesReporte/Pagina2.png"
img_fondo_path3 = "Formatos/partesReporte/Pagina3.png"
img_fondo_path4 = "Formatos/partesReporte/Pagina4.png"
img_fondo_path5 = "Formatos/partesReporte/Pagina5.png"
img_fondo_path6 = "Formatos/partesReporte/Pagina6.png"
img_fondo_path7 = "Formatos/partesReporte/Pagina7.png"
img_fondo_path8 = "Formatos/partesReporte/Pagina8.png"
img_fondo_path9 = "Formatos/partesReporte/Pagina9.png"
output_directory1 = "OUTPUT/Reportes/1"
output_directory2 = "OUTPUT/Reportes/2"
output_directory3 = "OUTPUT/Reportes/3"
output_directory4 = "OUTPUT/Reportes/4"
output_directory5 = "OUTPUT/Reportes/5"
output_directory6 = "OUTPUT/Reportes/6"
output_directory7 = "OUTPUT/Reportes/7"
output_directory8 = "OUTPUT/Reportes/8"
output_directory9 = "OUTPUT/Reportes/9"

voltaje_1_directory = "OUTPUT/Graficos/Canal 1/Voltaje"
corriente_1_directory = "OUTPUT/Graficos/Canal 1/Corriente"
voltaje_2_directory = "OUTPUT/Graficos/Canal 2/Voltaje"
corriente_2_directory = "OUTPUT/Graficos/Canal 2/Corriente"
voltaje_3_directory = "OUTPUT/Graficos/Canal 3/Voltaje"
corriente_3_directory = "OUTPUT/Graficos/Canal 3/Corriente"
voltaje_4_directory = "OUTPUT/Graficos/Canal 4/Voltaje"
corriente_4_directory = "OUTPUT/Graficos/Canal 4/Corriente"


def crear_paginas(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    fila_inicial = 0
    errorpos = 0
    while True:
        if fila_inicial >= len(df):
            break
        nocertificado = str(df.iat[fila_inicial + 2 , 5])
        error_promedio = df.iat[fila_inicial + 7, 1]
        nota = df.iat[fila_inicial + 1, 5]
        if pd.isna(nota):
            nota = "No se realizan observaciones"
        else:
            nota = str(nota)
        nombreEse = dfdatos.iat[3,1]
        fecha = dfdatos.iat[4, 1]
        metrologo = dfdatos.iat[7,1]
        primera_magnitud = df.iloc[fila_inicial + 6, 1]
        temperatura_calibracion = dfdatos.iat[13,1]
        humedad_calibracion = dfdatos.iat[14,1]
        temperaturaminima =dfdatos.iat[10,1]
        temperaturamaxima = dfdatos.iat[10,2]
        humedadminima = dfdatos.iat[11,1]
        humedadmaxima = dfdatos.iat[11,2]
        presionbarometrica = dfdatos.iat[12,1]
        desviacionestandar = df.iat[fila_inicial + 8, 1]
        incertidumbre = df.iat[fila_inicial + 9, 1]
        incertidumbre_expandida = df.iat[fila_inicial + 10, 1]
        #### Tablas complejas
        ## Canal 1 
        tipo_corriente_1 = df.iat[fila_inicial + 5, 1]
        voltaje_1 = df.iloc[fila_inicial + 8, 1:11].astype(float).tolist()
        corriente_1 = df.iloc[fila_inicial + 9, 1:11].astype(float).tolist()
        frecuencia_1_1 = df.iloc[fila_inicial + 12, 1:3].astype(str).tolist()
        frecuencia_2_1 = df.iloc[fila_inicial + 13, 1:3].astype(str).tolist()
        tipo_corriente_1_list.append(tipo_corriente_1)
        voltaje_1_list.append(voltaje_1)
        corriente_1_list.append(corriente_1)
        frecuencia_1_1_list.append(frecuencia_1_1)
        frecuencia_2_1_list.append(frecuencia_2_1)
        ## Normal
        tipo_corriente_normal_1 = df.iat[fila_inicial + 17, 1]
        voltaje_normal_1 = df.iloc[fila_inicial + 20, 1:11].astype(float).tolist()
        corriente_normal_1 = df.iloc[fila_inicial + 21, 1:11].astype(float).tolist()
        frecuencia_1_normal_1 = df.iloc[fila_inicial + 24, 1:3].astype(str).tolist()
        frecuencia_2_normal_1 = df.iloc[fila_inicial + 25, 1:3].astype(str).tolist()
        tipo_corriente_normal_1_list.append(tipo_corriente_normal_1)
        voltaje_normal_1_list.append(voltaje_normal_1)
        corriente_normal_1_list.append(corriente_normal_1)
        frecuencia_1_normal_1_list.append(frecuencia_1_normal_1)
        frecuencia_2_normal_1_list.append(frecuencia_2_normal_1)
        ## Canal 2
        tipo_corriente_2 = df.iat[fila_inicial + 5, 13]
        voltaje_2 = df.iloc[fila_inicial + 8, 13:23].astype(float).tolist()
        corriente_2 = df.iloc[fila_inicial + 9, 13:23].astype(float).tolist()
        frecuencia_1_2 = df.iloc[fila_inicial + 12, 13:15].astype(str).tolist()
        frecuencia_2_2 = df.iloc[fila_inicial + 13, 13:15].astype(str).tolist()
        tipo_corriente_2_list.append(tipo_corriente_2)
        voltaje_2_list.append(voltaje_2)
        corriente_2_list.append(corriente_2)
        frecuencia_1_2_list.append(frecuencia_1_2)
        frecuencia_2_2_list.append(frecuencia_2_2)
        ## Normal
        tipo_corriente_normal_2 = df.iat[fila_inicial + 17, 13]
        voltaje_normal_2 = df.iloc[fila_inicial + 20, 13:23].astype(float).tolist()
        corriente_normal_2 = df.iloc[fila_inicial + 21, 13:23].astype(float).tolist()
        frecuencia_1_normal_2 = df.iloc[fila_inicial + 24, 13:15].astype(str).tolist()
        frecuencia_2_normal_2 = df.iloc[fila_inicial + 25, 13:15].astype(str).tolist()
        tipo_corriente_normal_2_list.append(tipo_corriente_normal_2)
        voltaje_normal_2_list.append(voltaje_normal_2)
        corriente_normal_2_list.append(corriente_normal_2)
        frecuencia_1_normal_2_list.append(frecuencia_1_normal_2)
        frecuencia_2_normal_2_list.append(frecuencia_2_normal_2)

        ## Canal 3
        tipo_corriente_3 = df.iat[fila_inicial + 5, 25]
        voltaje_3 = df.iloc[fila_inicial + 8, 25:35].astype(float).tolist()
        corriente_3 = df.iloc[fila_inicial + 9, 25:35].astype(float).tolist()
        frecuencia_1_3 = df.iloc[fila_inicial + 12, 25:27].astype(str).tolist()
        frecuencia_2_3 = df.iloc[fila_inicial + 13, 25:27].astype(str).tolist()
        tipo_corriente_3_list.append(tipo_corriente_3)
        voltaje_3_list.append(voltaje_3)
        corriente_3_list.append(corriente_3)
        frecuencia_1_3_list.append(frecuencia_1_3)
        frecuencia_2_3_list.append(frecuencia_2_3)
        ## Normal 
        tipo_corriente_normal_3 = df.iat[fila_inicial + 17, 25]
        voltaje_normal_3 = df.iloc[fila_inicial + 20, 25:35].astype(float).tolist()
        corriente_normal_3 = df.iloc[fila_inicial + 21, 25:35].astype(float).tolist()
        frecuencia_1_normal_3 = df.iloc[fila_inicial + 24, 25:27].astype(str).tolist()
        frecuencia_2_normal_3 = df.iloc[fila_inicial + 25, 25:27].astype(str).tolist()
        tipo_corriente_normal_3_list.append(tipo_corriente_normal_3)
        voltaje_normal_3_list.append(voltaje_normal_3)
        corriente_normal_3_list.append(corriente_normal_3)
        frecuencia_1_normal_3_list.append(frecuencia_1_normal_3)
        frecuencia_2_normal_3_list.append(frecuencia_2_normal_3)
        
        ## Canal 4
        tipo_corriente_4 = df.iat[fila_inicial + 5, 37]
        voltaje_4 = df.iloc[fila_inicial + 8, 37:47].astype(float).tolist()
        corriente_4 = df.iloc[fila_inicial + 9, 37:47].astype(float).tolist()
        frecuencia_1_4 = df.iloc[fila_inicial + 12, 37:39].astype(str).tolist()
        frecuencia_2_4 = df.iloc[fila_inicial + 13, 37:39].astype(str).tolist()
        tipo_corriente_4_list.append(tipo_corriente_4)
        voltaje_4_list.append(voltaje_4)
        corriente_4_list.append(corriente_4)
        frecuencia_1_4_list.append(frecuencia_1_4)
        frecuencia_2_4_list.append(frecuencia_2_4)
        ## Normal
        tipo_corriente_normal_4 = df.iat[fila_inicial + 17, 37]
        voltaje_normal_4 = df.iloc[fila_inicial + 20, 37:47].astype(float).tolist()
        corriente_normal_4 = df.iloc[fila_inicial + 21, 37:47].astype(float).tolist()
        frecuencia_1_normal_4 = df.iloc[fila_inicial + 24, 37:39].astype(str).tolist()
        frecuencia_2_normal_4 = df.iloc[fila_inicial + 25, 37:39].astype(str).tolist()
        tipo_corriente_normal_4_list.append(tipo_corriente_normal_4)
        voltaje_normal_4_list.append(voltaje_normal_4)
        corriente_normal_4_list.append(corriente_normal_4)
        frecuencia_1_normal_4_list.append(frecuencia_1_normal_4)
        frecuencia_2_normal_4_list.append(frecuencia_2_normal_4)
        #### Tablas complejas

        if desviacionestandar == "N.R":
            desviacionestandar = str(desviacionestandar)
        if incertidumbre == "N.R":
            incertidumbre = str(incertidumbre)
        if incertidumbre_expandida == "N.R":
            incertidumbre_expandida = str(incertidumbre_expandida)
        errores_promedio.append(error_promedio)
        incertidumbres.append(incertidumbre)
        incertidumbres_expandidas.append(incertidumbre_expandida)
        certificados.append(nocertificado)
        desviaciones.append(desviacionestandar)
        notas.append(nota)
        fila_inicial += 26
    for certficado in certificados:
        agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha, metrologo, nombreEse)
        agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), tipo_corriente_1_list[errorpos], temperaturaminima, temperaturamaxima, presionbarometrica, humedadminima, humedadmaxima, tipo_corriente_1_list[errorpos], voltaje_1_list[errorpos], corriente_1_list[errorpos], frecuencia_1_1_list[errorpos], frecuencia_2_1_list[errorpos])
        agregar_imagenes_pdf3(img_fondo_path3, os.path.join(output_directory3, certficado + ".pdf"), tipo_corriente_normal_1_list[errorpos], voltaje_normal_1_list[errorpos], corriente_normal_1_list[errorpos], frecuencia_1_normal_1_list[errorpos], frecuencia_2_normal_1_list[errorpos], tipo_corriente_2_list[errorpos], voltaje_2_list[errorpos], corriente_2_list[errorpos], frecuencia_1_2_list[errorpos], frecuencia_2_2_list[errorpos])
        agregar_imagenes_pdf4(img_fondo_path4, os.path.join(output_directory4, certficado + ".pdf"), tipo_corriente_normal_2_list[errorpos], voltaje_normal_2_list[errorpos], corriente_normal_2_list[errorpos], frecuencia_1_normal_2_list[errorpos], frecuencia_2_normal_2_list[errorpos], tipo_corriente_3_list[errorpos], voltaje_3_list[errorpos], corriente_3_list[errorpos])
        agregar_imagenes_pdf5(img_fondo_path5, os.path.join(output_directory5, certficado + ".pdf"), frecuencia_1_3_list[errorpos], frecuencia_2_3_list[errorpos], tipo_corriente_normal_3_list[errorpos], voltaje_normal_3_list[errorpos], corriente_normal_3_list[errorpos], frecuencia_1_normal_3_list[errorpos], frecuencia_2_normal_3_list[errorpos], tipo_corriente_4_list[errorpos], voltaje_4_list[errorpos], corriente_4_list[errorpos], frecuencia_1_4_list[errorpos])
        agregar_imagenes_pdf6(img_fondo_path6, os.path.join(output_directory6, certficado + ".pdf"), frecuencia_1_4[errorpos], frecuencia_2_4_list[errorpos], tipo_corriente_normal_4_list[errorpos], voltaje_normal_4_list[errorpos], corriente_normal_4_list[errorpos], frecuencia_1_normal_4_list[errorpos], frecuencia_2_normal_4_list[errorpos])
        img_corriente_1 = os.path.join(corriente_1_directory, certficado + ".png")
        img_voltaje_1 = os.path.join(voltaje_1_directory, certficado + ".png")
        img_voltaje_2 = os.path.join(voltaje_2_directory, certficado + ".png")
        img_corriente_2 = os.path.join(corriente_2_directory, certficado + ".png")
        agregar_imagenes_pdf7(img_fondo_path7, img_voltaje_1, img_corriente_1, img_voltaje_2, img_corriente_2, os.path.join(output_directory7, certficado + ".pdf"), yinferior=120, ysuperior=420)
        img_voltaje_3 = os.path.join(voltaje_3_directory, certficado + ".png")
        img_corriente_3 = os.path.join(corriente_3_directory, certficado + ".png")
        img_voltaje_4 = os.path.join(voltaje_4_directory, certficado + ".png")
        img_corriente_4 = os.path.join(corriente_4_directory, certficado + ".png")
        agregar_imagenes_pdf8(img_fondo_path8, img_voltaje_3, img_corriente_3, img_voltaje_4, img_corriente_4, os.path.join(output_directory8, certficado + ".pdf"), yinferior=120, ysuperior=470)
        agregar_imagenes_pdf9(img_fondo_path9, os.path.join(output_directory9, certficado + ".pdf"), notas[errorpos])
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
    c.drawString(330, 665, nombrecertificado)
    c.setFont("Arial", 15)
    c.drawString(270, 265, fecha)
    c.drawString(270, 235, fecha)
    c.setFont("Arial", 11)
    c.drawString(270, 210, nombreEse)
    c.setFont("Arial", 15)
    c.drawString(270, 175, metrologo)
    c.save()
def agregar_imagenes_pdf2(fondo_path, output_pdf_path, tipo_corriente_normal_1 ,temperaturaminima, temperaturamaxima, presionbarometrica, humedadminima, humedadmaxima, tipo_corriente_1, voltaje_1, corriente_1, frecuencia_1_1, frecuencia_2_1):
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
    c.setFont("Arial", 11)
    c.drawString(290, 446, str(tipo_corriente_1))
    c.drawString(290, 85, str(tipo_corriente_normal_1))
    for i in range(10):
        #285
        c.drawString(295, 387 - (i * 20), str(voltaje_1[i]))
        c.drawString(410, 387 - (i * 20), str(corriente_1[i]))
    for i in range(2):
        c.drawString(295 + (i * 100), 158, str(frecuencia_1_1[i]))
        c.drawString(295 + (i * 100), 142, str(frecuencia_2_1[i]))
    c.save()

def agregar_imagenes_pdf3(fondo_path, output_pdf_path,tipo_corriente_normal_1 ,voltaje_normal_1, corriente_normal_1, frecuencia_1_normal_1, frecuencia_2_normal_1, tipo_corriente_2, voltaje_2, corriente_2, frecuencia_1_2, frecuencia_2_2,):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    c.setFont("Arial", 11)
    c.drawString(320, 80, str(tipo_corriente_normal_1))
    for i in range(10):
        c.drawString(295, 688 - (i * 20), str(voltaje_normal_1[i]))
        c.drawString(410, 688 - (i * 20), str(corriente_normal_1[i]))
    for i in range(2):
        c.drawString(295 + (i * 100), 458, str(frecuencia_1_normal_1[i]))
        c.drawString(295 + (i * 100), 442, str(frecuencia_2_normal_1[i]))
    c.drawString(320, 394, str(tipo_corriente_2))
    for i in range(10):
        c.drawString(295, 340 - (i * 20), str(voltaje_2[i]))
        c.drawString(410, 340 - (i * 20), str(corriente_2[i]))
    for i in range(2):
        c.drawString(295 + (i * 100), 123, str(frecuencia_1_2[i]))
        c.drawString(295 + (i * 100), 110, str(frecuencia_2_2[i]))
    c.save()

def agregar_imagenes_pdf4(img_fondo_path, output_pdf_path, tipo_corriente_normal_2, voltaje_normal_2, corriente_normal_2, frecuencia_1_normal_2, frecuencia_2_normal_2, tipo_corriente_3, voltaje_3, corriente_3):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    c.drawString(320, 688, str(tipo_corriente_normal_2))
    for i in range(10):
        c.drawString(295, 615 - (i * 20), str(voltaje_normal_2[i]))
        c.drawString(410, 615 - (i * 20), str(corriente_normal_2[i]))
    for i in range(2):
        c.drawString(295 + (i * 100), 388, str(frecuencia_1_normal_2[i]))
        c.drawString(295 + (i * 100), 372, str(frecuencia_2_normal_2[i]))
    c.drawString(320, 305, str(tipo_corriente_3))
    for i in range(10):
        c.drawString(295, 252 - (i * 20), str(voltaje_3[i]))
        c.drawString(410, 252 - (i * 20), str(corriente_3[i]))
    c.save()
def agregar_imagenes_pdf5(img_fondo_path, output_pdf_path , frecuencia_1_3, frecuencia_2_3, tipo_corriente_normal_3, voltaje_normal_3, corriente_normal_3, frecuencia_normal_1_3, frecuencia_normal_2_3, tipo_corriente_4, voltaje_4, corriente_4, frecuencia_1_4):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    for i in range(2):
        c.drawString(295 + (i * 100), 692, str(frecuencia_1_3[i]))
        c.drawString(295 + (i * 100), 677, str(frecuencia_2_3[i]))
    c.drawString(320, 622, str(tipo_corriente_normal_3))
    for i in range(10):
        c.drawString(295, 575 - (i * 20), str(voltaje_normal_3[i]))
        c.drawString(410, 575 - (i * 20), str(corriente_normal_3[i]))
    for i in range(2):
        c.drawString(295 + (i * 100), 350, str(frecuencia_normal_1_3[i]))
        c.drawString(295 + (i * 100), 335, str(frecuencia_normal_2_3[i]))
    c.drawString(320, 279, str(tipo_corriente_4))
    for i in range(10):
        c.drawString(295, 220 - (i * 20), str(voltaje_4[i]))
        c.drawString(410, 220 - (i * 20), str(corriente_4[i]))
    c.save()
def agregar_imagenes_pdf6(img_fondo_path, output_pdf_path, frecuencia_1_4, frecuencia_2_4, tipo_corriente_normal_4, voltaje_normal_4, corriente_normal_4, frecuencia_normal_1_4, frecuencia_normal_2_4):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    for i in range(2):
        c.drawString(295 + (i * 100), 680, str(frecuencia_1_4[i]))
        c.drawString(295 + (i * 100), 665, str(frecuencia_2_4[i]))
    c.drawString(320, 597, str(tipo_corriente_normal_4))
    for i in range(10):
        c.drawString(295, 542 - (i * 20), str(voltaje_normal_4[i]))
        c.drawString(410, 542 - (i * 20), str(corriente_normal_4[i]))
    for i in range(2):
        c.drawString(295 + (i * 100), 295, str(frecuencia_normal_1_4[i]))
        c.drawString(295 + (i * 100), 280, str(frecuencia_normal_2_4[i]))
    c.save()

def agregar_imagenes_pdf7(img_fondo_path, img_superior_path1, img_superior_path2, img_inferior_path1, img_inferior_path2, output_pdf_path, yinferior, ysuperior):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior1 = Image.open(img_superior_path1).convert("RGBA")
    image_superior2 = Image.open(img_superior_path2).convert("RGBA")
    image_inferior1 = Image.open(img_inferior_path1).convert("RGBA")
    image_inferior2 = Image.open(img_inferior_path2).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    
    # Resize images to fit the page and have the same size
    new_width = int(carta_ancho / 2.0)
    new_height = int(image_superior1.height * new_width / image_superior1.width)
    image_superior1 = image_superior1.resize((new_width, new_height), Image.LANCZOS)
    image_superior2 = image_superior2.resize((new_width, new_height), Image.LANCZOS)
    image_inferior1 = image_inferior1.resize((new_width, new_height), Image.LANCZOS)
    image_inferior2 = image_inferior2.resize((new_width, new_height), Image.LANCZOS)
    
    # Calculate positions
    x_fondo = 0
    x_superior1 = (carta_ancho / 4) - (new_width / 2)
    x_superior2 = (3 * carta_ancho / 4) - (new_width / 2)
    x_inferior1 = (carta_ancho / 4) - (new_width / 2)
    x_inferior2 = (3 * carta_ancho / 4) - (new_width / 2)
    
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, x_superior1, ysuperior, width=new_width, height=new_height)
    c.drawImage(img_superior_path2, x_superior2, ysuperior, width=new_width, height=new_height)
    c.drawImage(img_inferior_path1, x_inferior1, yinferior, width=new_width, height=new_height)
    c.drawImage(img_inferior_path2, x_inferior2, yinferior, width=new_width, height=new_height)
    
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    c.save()
def agregar_imagenes_pdf8(img_fondo_path, img_superior_path1, img_superior_path2, img_inferior_path1, img_inferior_path2, output_pdf_path, yinferior, ysuperior):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior1 = Image.open(img_superior_path1).convert("RGBA")
    image_superior2 = Image.open(img_superior_path2).convert("RGBA")
    image_inferior1 = Image.open(img_inferior_path1).convert("RGBA")
    image_inferior2 = Image.open(img_inferior_path2).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    
    # Resize images to fit the page and have the same size
    new_width = int(carta_ancho / 2.0)
    new_height = int(image_superior1.height * new_width / image_superior1.width)
    image_superior1 = image_superior1.resize((new_width, new_height), Image.LANCZOS)
    image_superior2 = image_superior2.resize((new_width, new_height), Image.LANCZOS)
    image_inferior1 = image_inferior1.resize((new_width, new_height), Image.LANCZOS)
    image_inferior2 = image_inferior2.resize((new_width, new_height), Image.LANCZOS)
    
    # Calculate positions
    x_fondo = 0
    x_superior1 = (carta_ancho / 4) - (new_width / 2)
    x_superior2 = (3 * carta_ancho / 4) - (new_width / 2)
    x_inferior1 = (carta_ancho / 4) - (new_width / 2)
    x_inferior2 = (3 * carta_ancho / 4) - (new_width / 2)
    
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, x_superior1, ysuperior, width=new_width, height=new_height)
    c.drawImage(img_superior_path2, x_superior2, ysuperior, width=new_width, height=new_height)
    c.drawImage(img_inferior_path1, x_inferior1, yinferior, width=new_width, height=new_height)
    c.drawImage(img_inferior_path2, x_inferior2, yinferior, width=new_width, height=new_height)
    
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    c.save()


def agregar_imagenes_pdf9(fondo_path, output_pdf_path , nota):
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
        c.drawString(85, 605, nota_line1)
        c.drawString(85, 585, nota_line2)
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