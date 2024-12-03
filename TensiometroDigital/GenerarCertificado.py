from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd

# Ruta al archivo Excel
archivo_excel = '/home/raven/Tensiometros.xlsx'
sheetname = "TENSIOMETRO DIGITAL"
df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
fila_actual = 0

def create_pdf(output_path, background_image_path, text_data, certificado):
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    background = ImageReader(background_image_path)
    c.drawImage(background, 0, 0, width=width, height=height)
    
    # Registrar fuentes
    pdfmetrics.registerFont(TTFont('Fraunces', 'Formatos/Fuentes/Fraunces.ttf'))
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Fraunces", 18)
    c.setFillColor("#d7a534")
    c.drawString(265, 647, certificado)
    
    # Dibujar el resto de los datos
    for position, texts in text_data.items():
        x, y = position
        c.setFont("Arial", 10)
        c.setFillColor("black")
        for i, text in enumerate(texts):
            c.drawString(x, y - i * 15, str(text))  # Ajusta la separación vertical
    
    c.save()

while True:
    if fila_actual >= len(df):
        print("Fin del archivo")
        break
    else:
        certificado = df.iat[fila_actual + 2, 5]
        print("Creando certificado: " + certificado)  
        output_path = "OUTPUT/Certificados/" + certificado + ".pdf"
        background_image_path = "Formatos/Imagenes/backCertificado.png"
        tipo = sheetname
        if pd.isna(df.iat[fila_actual + 3, 2]):
            inventario = "N.R"
        else:
            inventario = df.iat[fila_actual + 3, 2]
        marca = df.iat[fila_actual + 1, 1] if pd.notna(df.iat[fila_actual + 1, 1]) else "N.R"
        modelo = df.iat[fila_actual + 2, 1] if pd.notna(df.iat[fila_actual + 2, 1]) else "N.R"
        serie = df.iat[fila_actual + 1, 3] if pd.notna(df.iat[fila_actual + 1, 3]) else "N.R"
        ubicacion = df.iat[fila_actual + 2, 3] if pd.notna(df.iat[fila_actual + 2, 3]) else "N.R"
        nombreEse = df.iat[3, 10]
        fecha = df.iat[5, 10]
        direccion = df.iat[7, 10]
        text_data = {
            (315, 595): ["NIBP"],
            (315, 575): [tipo],
            (315, 555): [marca],
            (315, 535): [modelo],
            (315, 515): [serie],
            (315, 495): [inventario],
            (315, 475): ["1mmHg , 1LPM"],
            (315, 455): ["2mmHg"],
            (315, 435): ["SISTOLICA 60-220, DIASTOLICA 30-140, FP 60-90"],
            (315, 390): [nombreEse],
            (315, 370): [direccion],
            (315, 350): [ubicacion],
            (315, 330): [fecha],
            (315, 310): ["8"]
        }
        create_pdf(output_path, background_image_path, text_data, certificado)
        print(certificado + "ha sido creado")
        fila_actual += 35
