import argparse
import qrcode
from PIL import Image, ImageDraw, ImageFont
import json
import pandas as pd
import os
def main(archivo_excel):
    sheetname = "LAMPARA DE FOTOCURADO"
    output_directory = "OUTPUT/QRS"
    df = pd.read_excel(archivo_excel, sheet_name=sheetname, header=None)
    dfdatos = pd.read_excel(archivo_excel, sheet_name="DATOS SOLICITANTE", header=None)
    fila_actual = 0

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    with open("file_urls.json", "r") as file:
        data = json.load(file)
    certificados_series = []
    while fila_actual < len(df):
        certificado = df.iat[fila_actual + 2 , 5]
        serie = str(df.iat[fila_actual + 1, 3])
        certificados_series.append((certificado, serie))
        fila_actual += 18
    print(f"Certificados y series encontrados: {certificados_series}")
    background_template = Image.open("Formatos/Imagenes/backqr.png")
    bg_width, bg_height = background_template.size
    fecha = dfdatos.iat[4, 1]
    font = ImageFont.truetype("Formatos/Fuentes/Arial.ttf", 95)
    font2 = ImageFont.truetype("Formatos/Fuentes/Arial.ttf", 70)

    for certificado, serie in certificados_series:
        if certificado in data:
            link = data[certificado]
            print(f"Generando QR para certificado: {certificado}, Serie: {serie}, Link: {link}")
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=18,
                border=1,
            )
            qr.add_data(link)
            qr.make(fit=True)
            qr_image = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
            background = background_template.copy()
            draw = ImageDraw.Draw(background)

            draw.text((120, 515), certificado, font=font, fill="white")
            draw.text((250, 660), fecha, font=font2, fill="white")
            draw.text((240, 780), serie, font=font2, fill="white")
            qr_width, qr_height = qr_image.size
            pos = (int((bg_width / 2) + 200), int((bg_height / 2) - 40))
            background.paste(qr_image, pos, qr_image)

            # Guardar la imagen generada
            output_path = f"{output_directory}/{certificado}.png"
            background.save(output_path)
            print(f"Imagen guardada en: {output_path}")
        else:
            print(f"Certificado {certificado} no encontrado en el JSON.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Ejecutar scripts en orden con un archivo específico.")
    parser.add_argument(
        "--f", 
        required=True, 
        help="Especifica el archivo que deben usar los scripts, por ejemplo: Tensiometros.xlsx"
    )
    parser.add_argument(
        "--c", 
        nargs="+", 
        help="Especifica el -nombre de la nueva carpeta de drive"
    )
    args = parser.parse_args()
    main(args.f)
