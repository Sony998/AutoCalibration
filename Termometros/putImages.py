import argparse
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd

def main(excel_file):
    df = pd.read_excel(excel_file, sheet_name='TERMOMETRO', header=None)
    dfdatos = pd.read_excel(excel_file, sheet_name='DATOS SOLICITANTE', header=None)
    fila_actual = 0
    desviaciones = []
    nocertificados = []
    errores_list = []
    errores_promedio = []
    primeras = []
    segundas = []   
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
    output_directory1 = "OUTPUT/Reportes/1"
    output_directory2 = "OUTPUT/Reportes/2"
    output_directory3 = "OUTPUT/Reportes/3"
    output_directory4 = "OUTPUT/Reportes/4"
    inferior_directory = "OUTPUT/Graficos/Error"
    superior_directory = "OUTPUT/Graficos/Desviacion"
    errorpos = 0
    while True:
        if fila_actual >= len(df):
            break
        nocertificado = df.iat[fila_actual + 2 , 5]
        error_promedio = df.iat[fila_actual + 7, 1]
        if 7 < len(df.columns):
            nota = df.iat[fila_actual + 4, 7]
            if pd.isna(nota):
                nota = "No se realizan observaciones"
            else:
                nota = str(nota)
        else:
            nota = "No se realizan observaciones"
        desviacion = df.iat[fila_actual + 8, 1]
        primera = df.iloc[fila_actual + 5, 1:7].astype(float).tolist()
        incertidumbre = float(df.iat[fila_actual + 10, 1])
        incertidumbres_expandida = float(df.iat[fila_actual + 11, 1])
        errores_promedio.append(error_promedio)
        errores = df.iloc[fila_actual + 6, 1:7].astype(float).tolist()
        primeras.append(primera)
        incertidumbres.append(incertidumbre)
        incertidumbres_expandidas.append(incertidumbres_expandida)
        nocertificados.append(nocertificado)
        desviaciones.append(desviacion)
        errores_list.append(errores)
        notas.append(nota)
        print(f"nocertificado: {nocertificado}, Error promedio: {error_promedio}, desviacion {desviacion} nota {nota}")
        fila_actual += 13

    def agregar_imagenes_pdf3(img_fondo_path, img_superior_path1, img_superior_path2, output_pdf_path, yinferior, ysuperior, error_promedio, desviacion):
        img_fondo = Image.open(img_fondo_path).convert("RGBA")
        image_superior = Image.open(img_superior_path1).convert("RGBA")
        image_inferior = Image.open(img_superior_path2).convert("RGBA")
        carta_ancho, carta_alto = letter
        fondo_path = "temp_fondo.png"
        img_fondo.save(fondo_path, format="PNG")  
        image_superior = image_superior.resize((int(image_superior.width / 7), int(image_superior.height / 7)), Image.LANCZOS)
        image_inferior = image_inferior.resize((int(image_inferior.width / 7), int(image_inferior.height / 7)), Image.LANCZOS)
        yinferior += 50  # Move the inferior image up by 50 units
        ysuperior += 50  # Move the superior image up by 50 units
        x_fondo = 0  # Colocar el fondo en 0 para ocupar toda la página
        xsuperior = (carta_ancho - image_inferior.width) // 2 - 30
        c = canvas.Canvas(output_pdf_path, pagesize=letter)
        c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
        c.drawImage(img_superior_path1, xsuperior, yinferior, width=image_superior.width, height=image_superior.height)
        c.drawImage(img_superior_path2, xsuperior, ysuperior, width=image_inferior.width, height=image_inferior.height)
        pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
        c.setFont("Arial", 11)
        c.drawString(350, 690, "{:.2f}".format(float(f"{error_promedio:.2f}")))
        c.drawString(350, 675, "{:.2f}".format(float(f"{desviacion:.2f}")))
        c.save()

    def agregar_imagenes_pdf1(fondo_path, output_pdf_path, nombrenocertificado, fecha, nombreEse, metrologo):
        carta_ancho, carta_alto = letter
        c = canvas.Canvas(output_pdf_path, pagesize=letter)
        c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
        pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
        pdfmetrics.registerFont(TTFont('ArialBold', 'ArialBold.ttf'))
        pdfmetrics.registerFont(TTFont('ArialI', 'ArialI.ttf'))
        c.setFont("ArialI", 14)
        c.drawString(312, 664, nombrenocertificado)
        c.setFont("Arial", 15)
        c.drawString(210, 235, fecha)
        c.drawString(210, 205, fecha)
        c.setFont("Arial", 13)
        c.drawString(210, 177, nombreEse)
        c.setFont("Arial", 15)
        c.drawString(210, 150, metrologo)
        c.save()

    def agregar_imagenes_pdf2(fondo_path, output_pdf_path, nombrenocertificado, incertidumbre, incertidumbre_expandida, primera, errores_list):
        carta_ancho, carta_alto = letter
        c = canvas.Canvas(output_pdf_path, pagesize=letter)
        c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
        pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
        pdfmetrics.registerFont(TTFont('ArialBold', 'ArialBold.ttf'))
        pdfmetrics.registerFont(TTFont('ArialI', 'ArialI.ttf'))
        c.setFont("ArialI", 14)
        c.drawString(330, 680, str(temperaturaminima))
        c.drawString(398, 680, str(temperaturamaxima))
        c.drawString(360, 650, str(presionbarometrica))
        c.drawString(330, 620, str(humedadminima))
        c.drawString(398, 620, str(humedadmaxima))
        c.drawString(330, 410, "{:.2f}".format(float(f"{incertidumbre_expandida:.2f}")))
        c.drawString(330, 390, "{:.2f}".format(float(f"{incertidumbre:.2f}")))
        c.setFont("ArialI", 12)
        for i in range(6):
            c.drawString(175 + i * 52, 147, "{:.2f}".format(float(f"{primera[i]:.2f}")))
        for i in range(6):
            c.drawString(175 + i * 52, 125, "{:.2f}".format(float(f"{errores_list[i]:.2f}")))
        c.save()

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
        c.save()

    for certficado, error_promedio, desviacion in zip(nocertificados, errores_list, desviaciones):
        agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha, nombreEse, metrologo)
        agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), certficado, incertidumbres[errorpos], incertidumbres_expandidas[errorpos], primeras[errorpos], errores_list[errorpos])
        agregar_imagenes_pdf4(img_fondo_path4, os.path.join(output_directory4, certficado + ".pdf"), notas[errorpos] )
        img_superior_path1 = os.path.join(inferior_directory, certficado + ".png")
        img_superior_path2 = os.path.join(superior_directory, certficado + ".png")
        output_pdf_path = os.path.join(output_directory3, certficado + ".pdf")
        agregar_imagenes_pdf3(img_fondo_path3, img_superior_path1, img_superior_path2, output_pdf_path, 
                             yinferior=378, ysuperior=125, error_promedio=errores_promedio[errorpos], desviacion=desviacion)
        errorpos += 1
        print(f"Se ha creado el archivo {output_pdf_path} con error promedio de {error_promedio}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument('--f', type=str, required=True, help='Nombre del archivo de Excel')
    parser.add_argument('--c', type=str, help='Carpeta')
    args = parser.parse_args()
    main(args.f)
