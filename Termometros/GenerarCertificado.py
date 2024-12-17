from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
archivo_excel = 'pulido.xlsx'
df = pd.read_excel(archivo_excel, sheet_name='TERMOMETRO', header=None)
fila_inicial = 5
def create_pdf(output_path, background_image_path, text_data, nocertificado):
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    background = ImageReader(background_image_path)
    c.drawImage(background, 0, 0, width=width, height=height)
    pdfmetrics.registerFont(TTFont('Fraunces', 'Fraunces.ttf'))
    pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
    c.setFont("Fraunces", 18)
    c.setFillColor("#d7a534")
    c.drawString(265, 647, nocertificado)
    for text, position in text_data.items():
        x, y = position
        c.setFont("Arial", 10)
        c.setFillColor("black")
        c.drawString(x, y, text)
    c.save()
while True:
    if fila_inicial >= len(df):
        print("Fin del archivo")
        break
    else:
        nocertificado = df.iat[fila_inicial, 7]
        print(nocertificado)
        output_path = "Certificados/" + nocertificado + ".pdf"
        background_image_path = "backCertificado.png"
        tipo = str(df.iat[fila_inicial, 0])
        marca = str(df.iat[fila_inicial+1, 2])
        modelo = str(df.iat[fila_inicial + 2, 2])
        serie = str(df.iat[fila_inicial + 1, 7])
        resolucion = str(df.iat[fila_inicial + 12, 1])   
        print(serie)
        if str(df.iat[fila_inicial + 3, 2]) == "nan":
            inventario = "N.R"
        else:
            inventario = str(df.iat[fila_inicial + 3, 2])
            print(inventario)
        nombreEse = str(df.iat[1, 2])
        direccion = str(df.iat[3, 2])
        ubicacion = str(df.iat[fila_inicial + 2, 7])
        fecha = "7 de noviembre de 2024"
        text_data = {
            "Grados Celsius": (315, 595),
            tipo: (315, 575),
            marca: (315, 555),
            modelo: (315, 535),
            serie: (315, 515),
            inventario: (315, 495),
            "Grados": (315, 475),
            resolucion: (315, 455),
            "16.8 - 37.21": (315, 435),
            nombreEse: (315, 390),
            direccion: (315, 370),
            ubicacion: (315, 350),
            fecha: (315, 330),
            "5": (315, 310)
        }
        create_pdf(output_path, background_image_path, text_data, nocertificado)
        fila_inicial += 19