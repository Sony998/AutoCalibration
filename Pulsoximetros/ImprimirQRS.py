import os
from PIL import Image
from fpdf import FPDF

# Directorios a buscar
directories = ["OUTPUT/QRS","/home/raven/AutoCalibration/Infrarrojos/OUTPUT/QRS"]#/home/raven/AutoCalibration/Tensiometros/OUTPUT/QRS","/home/raven/AutoCalibration/Infrarrojos/OUTPUT/QRS"]#/home/raven/CODE/PULIR/Tensiometros/Completos","/home/raven/CODE/PULIR/Termometros/Completos"]
output_pdf = 'OUTPUT/Imprimir/QRSPulsoximetros.pdf'

images = []
for directory in directories:
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.png'):
                images.append(os.path.join(root, file))

# Crear un documento PDF tamaño carta
pdf = FPDF('P', 'mm', 'Letter')
pdf.set_auto_page_break(auto=True, margin=10)

# Configuración de la cuadrícula y dimensiones de imagen
images_per_row = 4
page_width, page_height = 215.9, 279.4  # Tamaño carta en mm
margin = 10  # Margen en mm
spacing = 5  # Espacio entre imágenes en mm
usable_width = page_width - 2 * margin - (images_per_row - 1) * spacing
image_width = usable_width / images_per_row
aspect_ratio = 1772 / 1181
image_height = image_width / aspect_ratio

# Añadir imágenes al PDF
for i, image_path in enumerate(images):
    if i % images_per_row == 0 and (i % (images_per_row * (page_height // (image_height + spacing))) == 0):
        pdf.add_page()
    row = (i // images_per_row) % int(page_height // (image_height + spacing))
    col = i % images_per_row
    x = margin + col * (image_width + spacing)
    y = margin + row * (image_height + spacing)
    pdf.image(image_path, x=x, y=y, w=image_width, h=image_height)

# Guardar el PDF
pdf.output(output_pdf)

print(f"PDF generado en {output_pdf}")