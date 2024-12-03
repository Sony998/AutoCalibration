from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd

# Ruta al archivo Excel
archivo_excel = '/home/raven/Tensiometros.xlsx'
sheetname = "TERMOMETRO INFRARROJO"

# Cargar la hoja del archivo Excel
df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)

# Fila inicial
fila_inicial = 0

def create_pdf(output_path, background_image_path, text_data, nocertificado):
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    background = ImageReader(background_image_path)
    c.drawImage(background, 0, 0, width=width, height=height)
    
    # Registrar fuentes
    pdfmetrics.registerFont(TTFont('Fraunces', 'Formatos/Fuentes/Fraunces.ttf'))
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Fraunces", 18)
    c.setFillColor("#d7a534")
    c.drawString(265, 647, nocertificado)
    
    # Dibujar el resto de los datos
    for position, texts in text_data.items():
        x, y = position
        c.setFont("Arial", 10)
        c.setFillColor("black")
        for i, text in enumerate(texts):
            c.drawString(x, y - i * 15, str(text))  # Ajusta la separación vertical
    
    c.save()

while True:
    if fila_inicial >= len(df):
        print("Fin del archivo")
        break
    else:
        # Obtener los valores desde el archivo Excel
        nocertificado = df.iat[fila_inicial + 2, 5]
        output_path = "OUTPUT/Certificados/" + nocertificado + ".pdf"
        background_image_path = "Formatos/Imagenes/backCertificado.png"
        tipo = sheetname
        if pd.isna(df.iat[fila_inicial + 3, 2]):
            inventario = "N.R"
        else:
            inventario = df.iat[fila_inicial + 3, 2]
        marca = df.iat[fila_inicial + 1, 1] if pd.notna(df.iat[fila_inicial + 1, 1]) else "N.R"
        modelo = df.iat[fila_inicial + 2, 1] if pd.notna(df.iat[fila_inicial + 2, 1]) else "N.R"
        serie = df.iat[fila_inicial + 1, 3] if pd.notna(df.iat[fila_inicial + 1, 3]) else "N.R"
        ubicacion = df.iat[fila_inicial + 2, 3] if pd.notna(df.iat[fila_inicial + 2, 3]) else "N.R"
        nombreEse = df.iat[3, 15]
        fecha = df.iat[5, 15]
        direccion = df.iat[7, 15]

        # Depuración: Imprimir cada valor
        # Construir el diccionario text_data
        text_data = {
            (315, 595): ["TEMPERATURA"],
            (315, 575): [tipo],
            (315, 555): [marca],
            (315, 535): [modelo],
            (315, 515): [serie],
            (315, 495): [inventario],
            (315, 475): ["°C"],
            (315, 455): ["0.1° C"],
            (315, 435): ["32.55 °C - 42.56 °C"],
            (315, 390): [nombreEse],
            (315, 370): [direccion],
            (315, 350): [ubicacion],
            (315, 330): [fecha],
            (315, 310): ["5"]
        }
        create_pdf(output_path, background_image_path, text_data, nocertificado)
        print(nocertificado + "ha sido creado")
        fila_inicial += 13
