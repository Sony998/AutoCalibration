from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
import argparse
sheetname = "TERMOHIGROMETRO"
certificados = []
desviaciones_temperatura_list = []
desviaciones_humedad_list = []
primeras_temperatura_list = []
segundas_temperatura_list = []
errores_temperatura_list = []
primeras_humedad_list = []
errores_humedad_list = []
error_promediohumedad_list = []
error_promediotemperatura_list = []
notas = []
img_fondo_path1 = "Formatos/partesReporte/Pagina1.png"
img_fondo_path2 = "Formatos/partesReporte/Pagina2.png"
img_fondo_path3 = "Formatos/partesReporte/Pagina3.png"
img_fondo_path4 = "Formatos/partesReporte/Pagina4.png"
output_directory1 = "OUTPUT/Reportes/1"
output_directory2 = "OUTPUT/Reportes/2"
output_directory3 = "OUTPUT/Reportes/3"
output_directory4 = "OUTPUT/Reportes/4"
inferior_directory = "OUTPUT/Graficos/Error"
def crear_paginas(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    fila_inicial = 0
    errorpos = 0
    while True:
        if fila_inicial >= len(df):
            break
        nocertificado = df.iat[fila_inicial + 2 , 5]
        nota = df.iat[fila_inicial + 1, 5]
        if pd.isna(nota):
            nota = "No se realizan observaciones"
        else:
            nota = str(nota)
        nombreEse = dfdatos.iat[3,1]
        fecha = dfdatos.iat[4, 1]
        metrologo = dfdatos.iat[7,1]
        error_promedio_temperatura = df.iat[fila_inicial + 8, 1]
        desviacionestandar_temperatura = df.iat[fila_inicial + 9, 1]
        errorpromedio_humedad = df.iat[fila_inicial + 14, 1]
        desviacionestandard_humedad = df.iat[fila_inicial + 15, 1]
        print(errorpromedio_humedad)
        primeratemperatura = df.iloc[fila_inicial + 5, 1:7].apply(lambda x: str(x) if x == "N.R" else float(x)).tolist()
        segundatemperatura = df.iloc[fila_inicial + 6, 1:7].apply(lambda x: str(x) if x == "N.R" else float(x)).tolist()
        errortemperatura = df.iloc[fila_inicial + 7, 1:7].apply(lambda x: str(x) if x == "N.R" else float(x)).tolist()
        primerahumedad = df.iloc[fila_inicial + 12, 1:5].apply(lambda x: str(x) if x == "N.R" else float(x)).tolist()
        errorhumedad = df.iloc[fila_inicial + 13, 1:5].apply(lambda x: str(x) if x == "N.R" else float(x)).tolist()
        certificados.append(nocertificado)
        desviaciones_temperatura_list.append(desviacionestandar_temperatura)
        desviaciones_humedad_list.append(desviacionestandard_humedad)
        primeras_temperatura_list.append(primeratemperatura)
        segundas_temperatura_list.append(segundatemperatura)
        errores_temperatura_list.append(errortemperatura)
        primeras_humedad_list.append(primerahumedad)
        errores_humedad_list.append(errorhumedad)
        error_promediohumedad_list.append(errorpromedio_humedad)
        error_promediotemperatura_list.append(error_promedio_temperatura)
        notas.append(nota)
        fila_inicial += 18
    for certficado in certificados:
        agregar_imagenes_pdf1(img_fondo_path1, os.path.join(output_directory1, certficado + ".pdf"), certficado, fecha, metrologo, nombreEse)
        img_superior_path1 = os.path.join(inferior_directory, certficado + ".png")
        output_pdf_path = os.path.join(output_directory3, certficado + ".pdf")
        agregar_imagenes_pdf2(
            img_fondo_path2, 
            os.path.join(output_directory2, certficado + ".pdf"), 
            primeras_temperatura_list[errorpos], 
            segundas_temperatura_list[errorpos], 
            errores_temperatura_list[errorpos], 
            primeras_humedad_list[errorpos], 
            errores_humedad_list[errorpos], 
            error_promediohumedad_list[errorpos], 
            desviaciones_humedad_list[errorpos], 
            error_promediotemperatura_list[errorpos], 
            desviaciones_temperatura_list[errorpos]
        )
        agregar_imagenes_pdf3(
            img_fondo_path3, 
            img_superior_path1, 
            output_pdf_path, 
            yinferior=270, 
        )
        agregar_imagenes_pdf4(
            img_fondo_path4, 
            os.path.join(output_directory4, certficado + ".pdf"), 
            notas[errorpos] 
        )
        errorpos += 1

"""         agregar_imagenes_pdf2(img_fondo_path2, os.path.join(output_directory2, certficado + ".pdf"), certficado, incertidumbres[errorpos], incertidumbres_expandidas[errorpos], primeras[errorpos], segundas[errorpos], errores_list[errorpos])
        agregar_imagenes_pdf3(img_fondo_path3, img_superior_path1, img_superior_path2, output_pdf_path, 
                            yinferior=125, ysuperior=378, error_promedio=errores_promedio[errorpos], desviacion=desviaciones[errorpos])
        agregar_imagenes_pdf4(img_fondo_path4, os.path.join(output_directory4, certficado + ".pdf"), notas[errorpos] )
        print("Se ha creado el reporte completo para el tensiometro con el certificado: ", certficado) 
        if errorpos >= len(certificados):
            print("Se han creado todos los reportes.")
            os.system("python3 UnirPartes.py")
            break  """

def agregar_imagenes_pdf1(fondo_path, output_pdf_path, nombrecertificado, fecha, metrologo, nombreEse):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    c.drawString(300, 145, "24")
    c.drawString(445, 145, "27")
    c.drawString(360, 115, "1011")
    c.drawString(300, 85, "52")
    c.drawString(445, 85, "60")
    c.setFont("ArialI", 12)
    c.drawString(330, 668, nombrecertificado)
    c.setFont("Arial", 12)
    c.drawString(320, 300, fecha)
    c.drawString(320, 276, fecha)
    c.setFont("Arial", 12)
    c.drawString(320, 255, nombreEse )
    c.setFont("Arial", 12)
    c.drawString(320, 235, metrologo)
    c.save()
def agregar_imagenes_pdf2(fondo_path, output_pdf_path, primera_temperatura, segunda_temperatura, errores_temperatura, primera_humedad, errores_humedad, error_promedio_humedad, desviacion_humedad ,error_promedio_temperatura, desviacion_temperatura):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    pdfmetrics.registerFont(TTFont('ArialI', 'Formatos/Fuentes/ArialI.ttf'))
    c.setFont("ArialI", 14)
    for i in range(6):
        c.drawString(160 + 70 * i, 575, "{:.2f}".format(float(f"{primera_temperatura[i]:.2f}")))
        c.drawString(160 + 70 * i, 545, "{:.2f}".format(float(f"{segunda_temperatura[i]:.2f}")))
        c.drawString(160 + 70 * i, 515, "{:.2f}".format(float(f"{errores_temperatura[i]:.2f}")))
    for i in range(4):
        c.drawString(240 + 72 * i, 440, "{:.2f}".format(float(f"{primera_humedad[i]:.2f}")))
        c.drawString(240 + 72 * i, 400, "{:.2f}".format(float(f"{errores_humedad[i]:.2f}")))
    c.setFont("Arial", 12)
    c.drawString(410, 250, "{:.2f}".format(float(f"{error_promedio_humedad:.2f}")))
    c.drawString(410, 220, "{:.2f}".format(float(f"{desviacion_humedad:.2f}")))

    c.drawString(410, 130, "{:.2f}".format(float(f"{error_promedio_temperatura:.2f}")))
    c.drawString(410, 100, "{:.2f}".format(float(f"{desviacion_temperatura:.2f}")))
    c.save()
def agregar_imagenes_pdf3(img_fondo_path, img_superior_path1, output_pdf_path, yinferior):
    img_fondo = Image.open(img_fondo_path).convert("RGBA")
    image_superior = Image.open(img_superior_path1).convert("RGBA")
    carta_ancho, carta_alto = letter
    fondo_path = "temp_fondo.png"
    img_fondo.save(fondo_path, format="PNG")  
    image_superior = image_superior.resize((int(image_superior.width / 4), int(image_superior.height / 4)), Image.LANCZOS)
    # Calcular la posición de la imagen
    x_fondo = 0  # Colocar el fondo en 0 para ocupar toda la página
    xsuperior = (carta_ancho - image_superior.width) // 2 - 30
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, x_fondo, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    c.drawImage(img_superior_path1, xsuperior, yinferior, width=image_superior.width, height=image_superior.height)
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Arial", 11)
    c.save()


def agregar_imagenes_pdf4(fondo_path, output_pdf_path , nota):
    carta_ancho, carta_alto = letter
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    c.drawImage(fondo_path, 0, 0, width=carta_ancho, height=carta_alto, preserveAspectRatio=True, mask='auto')
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArialBold', 'Formatos/Fuentes/ArialBold.ttf'))
    c.setFont("ArialBold", 14)
    c.drawString(80, 620, "OBSERVACIONES")
    c.setFont("Arial", 12)
    max_length = 75  # Maximum characters per line
    if len(nota) > max_length:
        nota_line1 = nota[:max_length]
        nota_line2 = nota[max_length:]
        c.drawString(80, 590, nota_line1)
        c.drawString(80, 570, nota_line2)
    else:
        c.drawString(80, 590, nota)
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