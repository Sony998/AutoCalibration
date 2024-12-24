import qrcode
from PIL import Image, ImageDraw, ImageFont
import json
import pandas as pd
import os
sheetname = "PULSO OXIMETRO"

# Cargar archivo Excel
archivo_excel = '/home/raven/Quipama.xlsx'
df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
# Inicialización
fila_actual = 0
output_directory = "OUTPUT/QRS"
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Cargar los datos de los certificados desde el archivo JSON
with open("file_urls.json", "r") as file:
    data = json.load(file)

# Lista para almacenar certificados y sus series
certificados_series = []

# Recorrer el Excel y guardar certificados y sus series
while fila_actual < len(df):
    certificado = df.iat[fila_actual + 2 , 5]
    serie = str(df.iat[fila_actual + 1, 3])
    certificados_series.append((certificado, serie))
    fila_actual += 21  # Avanzar al siguiente bloque

print(f"Certificados y series encontrados: {certificados_series}")

# Crear imágenes QR solo para los certificados que están en el JSON
background_template = Image.open("Formatos/Imagenes/backqr.png")
bg_width, bg_height = background_template.size
fecha = dfdatos.iat[4, 1]
font = ImageFont.truetype("Formatos/Fuentes/Arial.ttf", 80)
font2 = ImageFont.truetype("Formatos/Fuentes/Arial.ttf", 50)

for certificado, serie in certificados_series:
    if certificado in data:
        link = data[certificado]
        print(f"Generando QR para certificado: {certificado}, Serie: {serie}, Link: {link}")
        
        # Crear el código QR
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=18,
            border=1,
        )
        qr.add_data(link)
        qr.make(fit=True)
        qr_image = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
        
        # Crear la imagen de fondo
        background = background_template.copy()
        draw = ImageDraw.Draw(background)

        # Añadir el texto del certificado, fecha y serie
        draw.text((130, 580), certificado, font=font, fill="white")
        draw.text((210, 722), fecha, font=font2, fill="white")
        draw.text((200, 815), serie, font=font2, fill="white")

        # Añadir el código QR
        qr_width, qr_height = qr_image.size
        pos = (int((bg_width / 2) + 200), int((bg_height / 2) - 40))
        background.paste(qr_image, pos, qr_image)

        # Guardar la imagen generada
        output_path = f"{output_directory}/{certificado}.png"
        background.save(output_path)
        print(f"Imagen guardada en: {output_path}")
    else:
        print(f"Certificado {certificado} no encontrado en el JSON.")
