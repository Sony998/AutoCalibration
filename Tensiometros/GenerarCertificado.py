from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
import argparse

sheetname = "ESFIGMOMANOMETRO"
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
    
    # Dibujar el resto de los datos
    for position, texts in text_data.items():
        x, y = position
        c.setFont("Arial", 10)
        c.setFillColor("black")
        for i, text in enumerate(texts):
            c.drawString(x, y - i * 15, str(text))  # Ajusta la separación vertical
    
    c.save()
def generar_certificado(archivo_excel):
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    fila_inicial = 0
    while True:
        if fila_inicial >= len(df):
            print("Fin del archivo")
            break
        else:
            certificado = df.iat[fila_inicial + 2, 5]
            output_path = "OUTPUT/Certificados/" + certificado + ".pdf"
            background_image_path = "Formatos/Imagenes/backCertificado.png"
            tipo = sheetname
            if pd.isna(df.iat[fila_inicial + 1, 6]):
                inventario = "N.R"
            else:
                inventario = df.iat[fila_inicial + 1, 6]
            marca = df.iat[fila_inicial + 1, 1] if pd.notna(df.iat[fila_inicial + 1, 1]) else "N.R"
            modelo = df.iat[fila_inicial + 2, 1] if pd.notna(df.iat[fila_inicial + 2, 1]) else "N.R"
            serie = df.iat[fila_inicial + 1, 3] if pd.notna(df.iat[fila_inicial + 1, 3]) else "N.R"
            ubicacion = df.iat[fila_inicial + 2, 3] if pd.notna(df.iat[fila_inicial + 2, 3]) else "N.R"
            nombreEse = df.iat[3, 15]
            fecha = df.iat[5, 15]
            direccion = df.iat[7, 15]
            text_data = {
                (315, 595): ["PRESION"],
                (315, 575): [tipo],
                (315, 555): [marca],
                (315, 535): [modelo],
                (315, 515): [serie],
                (315, 495): [inventario],
                (315, 475): ["mmHg"],
                (315, 455): ["2mmHg"],
                (315, 435): ["40 - 280"],
                (315, 390): [nombreEse],
                (315, 370): [direccion],
                (315, 350): [ubicacion],
                (315, 330): [fecha],
                (315, 310): ["5"]
            }
            create_pdf(output_path, background_image_path, text_data, certificado)
            print("El certificado" + certificado + " ha sido creado")
            fila_inicial += 13

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
        generar_certificado(args.f)

if __name__ == "__main__":
    main()