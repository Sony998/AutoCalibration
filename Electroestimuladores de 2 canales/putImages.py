from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
import argparse
sheetname = "ELECTROESTIMULADOR"
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
voltaje_1_list = []
corriente_1_list = []
voltaje_2_list = []
corriente_2_list = []
tipo_corriente_1_list = []
tipo_corriente_2_list = []
frecuencia_1_1_list = []
frecuencia_2_1_list = []
frecuencia_1_2_list = []
frecuencia_2_2_list = []
tipo_corriente_normal_1_list = []
tipo_corriente_normal_2_list = []
voltaje_normal_1_list = []
voltaje_normal_2_list = []
corriente_normal_1_list = []
corriente_normal_2_list = []
frecuencia_1_1_normal_list = []
frecuencia_2_1_normal_list = []
frecuencia_1_2_normal_list = []
frecuencia_2_2_normal_list = []
tipo_corriente_modulacion_1_list = []
tipo_corriente_modulacion_2_list = []
voltaje_modulacion_1_list = []
voltaje_modulacion_2_list = []
corriente_modulacion_1_list = []
corriente_modulacion_2_list = []
frecuencia_1_1_modulacion_list = []
frecuencia_2_1_modulacion_list = []
frecuencia_1_2_modulacion_list = []
frecuencia_2_2_modulacion_list = []
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
voltaje_1_directory = "OUTPUT/Graficos/Canal 1/Voltaje"
corriente_1_directory = "OUTPUT/Graficos/Canal 1/Corriente"
voltaje_2_directory = "OUTPUT/Graficos/Canal 2/Voltaje"
corriente_2_directory = "OUTPUT/Graficos/Canal 2/Corriente"



def crear_paginas(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    fila_inicial = 0
    errorpos = 0
    while True:
        if fila_inicial >= len(df):
            break
        nocertificado = str(df.iat[fila_inicial + 2 , 5])
        if pd.isna(nocertificado):
            break
        if nocertificado == "nan":
            break
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
        voltaje_1 = df.iloc[fila_inicial + 8, 1:5].astype(float).tolist()
        corriente_1 = df.iloc[fila_inicial + 9, 1:5].astype(float).tolist()
        frecuencia_1_1 = df.iloc[fila_inicial + 12, 1:3].astype(str).tolist()
        frecuencia_2_1 = df.iloc[fila_inicial + 13, 1:3].astype(str).tolist()


        tipo_corriente_1_list.append(tipo_corriente_1)
        voltaje_1_list.append(voltaje_1)
        corriente_1_list.append(corriente_1)
        frecuencia_1_1_list.append(frecuencia_1_1)
        frecuencia_2_1_list.append(frecuencia_2_1)

        ## Normal 
        tipo_corriente_normal_1 = df.iat[fila_inicial + 17, 1]
        voltaje_normal_1 = df.iloc[fila_inicial + 20, 1:5].astype(float).tolist()
        corriente_normal_1 = df.iloc[fila_inicial + 21, 1:5].astype(float).tolist()
        frecuencia_1_1_normal = df.iloc[fila_inicial + 24, 1:3].astype(str).tolist()
        frecuencia_2_1_normal = df.iloc[fila_inicial + 25, 1:3].astype(str).tolist()

        tipo_corriente_normal_1_list.append(tipo_corriente_normal_1)
        voltaje_normal_1_list.append(voltaje_normal_1)
        corriente_normal_1_list.append(corriente_normal_1)
        frecuencia_1_1_normal_list.append(frecuencia_1_1_normal)
        frecuencia_2_1_normal_list.append(frecuencia_2_1_normal)
        ## Modulacion 
        tipo_corriente_modulacion_1 = df.iat[fila_inicial + 29, 1]
        voltaje_modulacion_1 = df.iloc[fila_inicial + 32, 1:5].astype(float).tolist()
        corriente_modulacion_1 = df.iloc[fila_inicial + 33, 1:5].astype(float).tolist()
        frecuencia_1_1_modulacion = df.iloc[fila_inicial + 36, 1:3].astype(str).tolist()
        frecuencia_2_1_modulacion = df.iloc[fila_inicial + 37, 1:3].astype(str).tolist()

        tipo_corriente_modulacion_1_list.append(tipo_corriente_modulacion_1)
        voltaje_modulacion_1_list.append(voltaje_modulacion_1)
        corriente_modulacion_1_list.append(corriente_modulacion_1)
        frecuencia_1_1_modulacion_list.append(frecuencia_1_1_modulacion)
        frecuencia_2_1_modulacion_list.append(frecuencia_2_1_modulacion)


        ## Canal 2
        tipo_corriente_2 = df.iat[fila_inicial+5,7]
        print(tipo_corriente_2)
        voltaje_2 = df.iloc[fila_inicial + 8,7:11].astype(float).tolist()
        corriente_2 = df.iloc[fila_inicial + 9,7:11].astype(float).tolist()
        frecuencia_1_2 = df.iloc[fila_inicial + 12, 7:9].astype(str).tolist()
        frecuencia_2_2 = df.iloc[fila_inicial + 13, 7:9].astype(str).tolist()
        tipo_corriente_2_list.append(tipo_corriente_2)
        voltaje_2_list.append(voltaje_2)
        corriente_2_list.append(corriente_2)
        frecuencia_1_2_list.append(frecuencia_1_2)
        frecuencia_2_2_list.append(frecuencia_2_2)
        ## Normal 
        tipo_corriente_normal_2 = df.iat[fila_inicial + 17, 7]
        voltaje_normal_2 = df.iloc[fila_inicial + 20, 7:11].astype(float).tolist()
        corriente_normal_2 = df.iloc[fila_inicial + 21, 7:11].astype(float).tolist()
        frecuencia_1_2_normal = df.iloc[fila_inicial + 24, 7:9].astype(str).tolist()
        frecuencia_2_2_normal = df.iloc[fila_inicial + 25, 7:9].astype(str).tolist()

        tipo_corriente_normal_2_list.append(tipo_corriente_normal_2)
        voltaje_normal_2_list.append(voltaje_normal_2)
        corriente_normal_2_list.append(corriente_normal_2)
        frecuencia_1_2_normal_list.append(frecuencia_1_2_normal)
        frecuencia_2_2_normal_list.append(frecuencia_2_2_normal)
        ## Modulacion
        tipo_corriente_modulacion_2 = df.iat[fila_inicial + 29, 7]
        voltaje_modulacion_2 = df.iloc[fila_inicial + 32, 7:11].astype(float).tolist()
        corriente_modulacion_2 = df.iloc[fila_inicial + 33, 7:11].astype(float).tolist()
        frecuencia_1_2_modulacion = df.iloc[fila_inicial + 36, 7:9].astype(str).tolist()
        frecuencia_2_2_modulacion = df.iloc[fila_inicial + 37, 7:9].astype(str).tolist()

        tipo_corriente_modulacion_2_list.append(tipo_corriente_modulacion_2)
        voltaje_modulacion_2_list.append(voltaje_modulacion_2)
        corriente_modulacion_2_list.append(corriente_modulacion_2)
        frecuencia_1_2_modulacion_list.append(frecuencia_1_2_modulacion)
        frecuencia_2_2_modulacion_list.append(frecuencia_2_2_modulacion)

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
        fila_inicial += 38
    for certficado in certificados:
        agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha, metrologo, nombreEse, temperaturaminima, temperaturamaxima, presionbarometrica, humedadminima, humedadmaxima)
        agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), tipo_corriente_1_list[errorpos], voltaje_1_list[errorpos], corriente_1_list[errorpos], frecuencia_1_1_list[errorpos], frecuencia_2_1_list[errorpos], tipo_corriente_normal_1_list[errorpos], voltaje_normal_1_list[errorpos], corriente_normal_1_list[errorpos], frecuencia_1_1_normal_list[errorpos], frecuencia_2_1_normal_list[errorpos], tipo_corriente_modulacion_1_list[errorpos], voltaje_modulacion_1_list[errorpos], corriente_modulacion_1_list[errorpos], frecuencia_1_1_modulacion_list[errorpos], frecuencia_2_1_modulacion_list[errorpos])
        agregar_imagenes_pdf3(img_fondo_path3, os.path.join(output_directory3, certficado + ".pdf"), tipo_corriente_2_list[errorpos], voltaje_2_list[errorpos], corriente_2_list[errorpos], frecuencia_1_2_list[errorpos], frecuencia_2_2_list[errorpos], tipo_corriente_normal_2_list[errorpos], voltaje_normal_2_list[errorpos], corriente_normal_2_list[errorpos], frecuencia_1_2_normal_list[errorpos], frecuencia_2_2_normal_list[errorpos], tipo_corriente_modulacion_2_list[errorpos], voltaje_modulacion_2_list[errorpos], corriente_modulacion_2_list[errorpos], frecuencia_1_2_modulacion_list[errorpos], frecuencia_2_2_modulacion_list[errorpos])
        img_voltaje_1 = os.path.join(voltaje_1_directory, certficado + ".png")
        img_corriente_1 = os.path.join(corriente_1_directory, certficado + ".png")
        img_voltaje_2 = os.path.join(voltaje_2_directory, certficado + ".png")
        img_corriente_2 = os.path.join(corriente_2_directory, certficado + ".png")
        agregar_imagenes_pdf4(img_fondo_path4, img_voltaje_1, img_corriente_1, img_voltaje_2, img_corriente_2, os.path.join(output_directory4, certficado + ".pdf"), 200, 400)
        agregar_imagenes_pdf5(img_fondo_path5, os.path.join(output_directory5, certficado + ".pdf"), notas[errorpos] )
        print("Se ha creado el reporte completo para la lampara de fotocurado con el certificado: ", certficado) 
        errorpos += 1
        if errorpos >= len(certificados):
            break
    print("Se han creado todos los reportes completos para las lamparas de fotocurado.")
    os.system("python3 UnirPartes.py")

def agregar_imagenes_pdf1(fondo_path, output_pdf_path, nombrecertificado, fecha, metrologo, nombreEse, temperaturaminima, temperaturamaxima, presionbarometrica, humedadminima, humedadmaxima):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 12)
    c.drawString(335, 688, nombrecertificado)
    c.setFont("Arial", 15)
    c.drawString(270, 315, fecha)
    c.drawString(270, 285, fecha)
    c.setFont("Arial", 11)
    c.drawString(270, 260, nombreEse)
    c.setFont("Arial", 15)
    c.drawString(270, 225, metrologo)
    c.setFont("ArialI", 14)
    c.drawString(325, 135, str(temperaturaminima))
    c.drawString(438, 135, str(temperaturamaxima))
    c.drawString(370, 110, str(presionbarometrica))
    c.drawString(325, 85, str(humedadminima))
    c.drawString(438, 85, str(humedadmaxima))
    c.save()
def agregar_imagenes_pdf2(fondo_path, output_pdf_path, tipo_corriente_1, voltaje_1, corriente_1, frecuencia_1_1, frecuencia_2_1, tipo_corriente_normal_1, voltaje_normal_1, corriente_normal_1, frecuencia_1_1_normal, frecuencia_2_1_normal, tipo_corriente_modulacion_1, voltaje_modulacion_1, corriente_modulacion_1, frecuencia_1_1_modulacion, frecuencia_2_1_modulacion):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("Arial", 11)
    c.drawString(240, 595, str(tipo_corriente_1))
    for i in range(4):
        c.drawString(220 + (i * 85), 555, str(voltaje_1[i]))
        c.drawString(220 + (i * 85), 540, str(corriente_1[i]))
    for i in range(2):
        c.drawString(320 + (i * 110), 492, str(frecuencia_1_1[i]))
        c.drawString(320 + (i * 110), 477, str(frecuencia_2_1[i]))
    c.drawString(260, 433, str(tipo_corriente_normal_1))
    for i in range(4):
        c.drawString(220 + (i * 85), 394, str(voltaje_normal_1[i]))
        c.drawString(220 + (i * 85), 378, str(corriente_normal_1[i]))
    for i in range(2):
        c.drawString(320 + (i * 110), 328, str(frecuencia_1_1_normal[i]))
        c.drawString(320 + (i * 110), 312, str(frecuencia_2_1_normal[i]))
    c.drawString(260, 272, str(tipo_corriente_modulacion_1))
    for i in range(4):
        c.drawString(220 + (i * 85), 232, str(voltaje_modulacion_1[i]))
        c.drawString(220 + (i * 85), 219, str(corriente_modulacion_1[i]))
    for i in range(2):
        c.drawString(320 + (i * 110), 167, str(frecuencia_1_1_modulacion[i]))
        c.drawString(320 + (i * 110), 150, str(frecuencia_2_1_modulacion[i]))
    c.save()

def agregar_imagenes_pdf3(fondo_path, output_pdf_path, tipo_corriente_2, voltaje_2, corriente_2, frecuencia_1_2, frecuencia_2_2, tipo_corriente_normal_2, voltaje_normal_2, corriente_normal_2, frecuencia_1_2_normal, frecuencia_2_2_normal, tipo_corriente_modulacion_2, voltaje_modulacion_2, corriente_modulacion_2, frecuencia_1_2_modulacion, frecuencia_2_2_modulacion):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    c.setFont("Arial", 11)
    c.drawString(240, 638, str(tipo_corriente_2))
    for i in range(4):
        c.drawString(220 + (i * 85), 598, str(voltaje_2[i]))
        c.drawString(220 + (i * 85), 583, str(corriente_2[i]))
    for i in range(2):
        c.drawString(320 + (i * 110), 535, str(frecuencia_1_2[i]))
        c.drawString(320 + (i * 110), 520, str(frecuencia_2_2[i]))
    c.drawString(260, 475, str(tipo_corriente_normal_2))
    for i in range(4):
        c.drawString(220 + (i * 85), 436, str(voltaje_normal_2[i]))
        c.drawString(220 + (i * 85), 420, str(corriente_normal_2[i]))
    for i in range(2):
        c.drawString(320 + (i * 110), 370, str(frecuencia_1_2_normal[i]))
        c.drawString(320 + (i * 110), 355, str(frecuencia_2_2_normal[i]))
    c.drawString(260, 315, str(tipo_corriente_modulacion_2))
    for i in range(4):
        c.drawString(220 + (i * 85), 275, str(voltaje_modulacion_2[i]))
        c.drawString(220 + (i * 85), 260, str(corriente_modulacion_2[i]))
    for i in range(2):
        c.drawString(320 + (i * 110), 210, str(frecuencia_1_2_modulacion[i]))
        c.drawString(320 + (i * 110), 195, str(frecuencia_2_2_modulacion[i]))

    
    c.save()
def agregar_imagenes_pdf4(img_fondo_path, img_superior_path1, img_superior_path2, img_inferior_path1, img_inferior_path2, output_pdf_path, yinferior, ysuperior):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior1 = Image.open(img_superior_path1).convert("RGBA")
    image_superior2 = Image.open(img_superior_path2).convert("RGBA")
    image_inferior1 = Image.open(img_inferior_path1).convert("RGBA")
    image_inferior2 = Image.open(img_inferior_path2).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    
    # Resize images to fit the page
    new_width = int(carta_ancho / 2.0)
    image_superior1 = image_superior1.resize((new_width, int(image_superior1.height * new_width / image_superior1.width)), Image.LANCZOS)
    image_superior2 = image_superior2.resize((new_width, int(image_superior2.height * new_width / image_superior2.width)), Image.LANCZOS)
    image_inferior1 = image_inferior1.resize((new_width, int(image_inferior1.height * new_width / image_inferior1.width)), Image.LANCZOS)
    image_inferior2 = image_inferior2.resize((new_width, int(image_inferior2.height * new_width / image_inferior2.width)), Image.LANCZOS)
    
    # Calculate positions
    x_fondo = 0
    x_superior1 = (carta_ancho / 4) - (image_superior1.width / 2)
    x_superior2 = (3 * carta_ancho / 4) - (image_superior2.width / 2)
    x_inferior1 = (carta_ancho / 4) - (image_inferior1.width / 2)
    x_inferior2 = (3 * carta_ancho / 4) - (image_inferior2.width / 2)
    
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, x_superior1, ysuperior + 50, width=image_superior1.width, height=image_superior1.height)
    c.drawImage(img_superior_path2, x_superior2, ysuperior + 50, width=image_superior2.width, height=image_superior2.height)
    c.drawImage(img_inferior_path1, x_inferior1, yinferior - 50, width=image_inferior1.width, height=image_inferior1.height)
    c.drawImage(img_inferior_path2, x_inferior2, yinferior - 50, width=image_inferior2.width, height=image_inferior2.height)
    
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
        c.drawString(80, 605, nota_line1)
        c.drawString(80, 585, nota_line2)
    else:
        c.drawString(80, 580, nota)
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