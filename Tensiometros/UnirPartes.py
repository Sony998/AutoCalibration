import argparse
import os
import PyPDF2

folders = ["OUTPUT/Certificados","OUTPUT/Reportes/1", "OUTPUT/Reportes/2", "OUTPUT/Reportes/3", "OUTPUT/Reportes/4"]
pdf_files = {}

def obtener_archivos_pdf():
    for folder in folders:
        for filename in os.listdir(folder):
            if filename.endswith(".pdf"):
                if filename not in pdf_files:
                    pdf_files[filename] = []
                pdf_files[filename].append(os.path.join(folder, filename))

def unir_archivos_pdf():
    merger = PyPDF2.PdfMerger()
    for filename, paths in pdf_files.items():
        for path in paths:
            merger.append(path)
        for filename, paths in pdf_files.items():
            merger = PyPDF2.PdfMerger()
            for path in paths:
                merger.append(path)
            output_path = os.path.join("OUTPUT/Completos", filename)
            print(f"El reporte {filename} ha sido unido con el certificado")
            with open(output_path, "wb") as f_out:
                merger.write(f_out)

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
        help="Especifica el nombre de la nueva carpeta de drive"
    )
    obtener_archivos_pdf()
    unir_archivos_pdf()
