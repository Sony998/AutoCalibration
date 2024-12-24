from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
sheetname = "MONITOR DE SIGNOS VITALES"   
archivo_excel = '/home/raven/Quipama.xlsx'
df = pd.read_excel(archivo_excel, sheetname, header=None)
dfdatos = pd.read_excel(archivo_excel, sheet_name='DATOS SOLICITANTE', header=None)

fila_actual = 0
def create_pdf(output_path, background_image_path, text_data, certificado):
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    background = ImageReader(background_image_path)
    c.drawImage(background, 0, 0, width=width, height=height)
    pdfmetrics.registerFont(TTFont('Fraunces', 'Formatos/Fuentes/Fraunces.ttf'))
    pdfmetrics.registerFont(TTFont('Arial', 'Formatos/Fuentes/Arial.ttf'))
    c.setFont("Fraunces", 18)
    c.setFillColor("#d7a534")
    c.drawString(265, 647, certificado)
    for text, position in text_data.items():
        x, y = position
        c.setFont("Arial", 10)
        c.setFillColor("black")
        c.drawString(x, y, str(text))
    c.save()
while True:
    if fila_actual >= len(df):
        print("Fin del archivo")
        break
    else:
        if fila_actual + 2 >= len(df) or fila_actual + 1 >= len(df):
            print("Fin del archivo")
            break
        certificado = df.iat[fila_actual + 2, 5]
        print(certificado)
        output_path = "OUTPUT/Certificados/" + certificado + ".pdf"
        background_image_path = "Formatos/Imagenes/backCertificado.png"
        tipo = sheetname
        marca = df.iat[fila_actual + 1, 1] if pd.notna(df.iat[fila_actual + 1, 1]) else "N.R"
        modelo = df.iat[fila_actual + 2, 1] if pd.notna(df.iat[fila_actual + 2, 1]) else "N.R"
        serie = df.iat[fila_actual + 1, 3] if pd.notna(df.iat[fila_actual + 1, 3]) else "N.R"
        ubicacion = df.iat[fila_actual + 2, 3] if pd.notna(df.iat[fila_actual + 2, 3]) else "N.R"
        if 7 < len(df.columns):
            inventario = df.iat[fila_actual + 1, 7]
            if pd.isna(inventario):
                inventario = "N.R"
            else:
                inventario = str(inventario)
        else:
            inventario = "N.R"
        nombreEse = dfdatos.iat[3,1]
        fecha = dfdatos.iat[4, 1]
        direccion = dfdatos.iat[6, 1]
        text_data = {
            "NIBP, SPO2": (315, 595),
            tipo: (315, 575),
            marca: (315, 555),
            modelo: (315, 535),
            serie: (315, 515),
            "N.R": (315, 495),
            "mmHg, % , RPM, LPM": (315, 475),
            "2mmHg, 1 %, 1rpm, 1lpm": (315, 455),
            "60-220, 85%-100% ": (315, 435),
            nombreEse: (315, 390),
            direccion: (315, 370),
            ubicacion: (315, 350),
            fecha: (315, 330),
            "8": (315, 310)
        }
        create_pdf(output_path, background_image_path, text_data, certificado)
        fila_actual += 70