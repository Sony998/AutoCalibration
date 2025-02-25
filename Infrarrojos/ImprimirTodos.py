import argparse
import os
import PyPDF2
folders = ["OUTPUT/Completos"]#"/home/raven/CODE/PULIR/Tensiometros/Completos","/home/raven/CODE/PULIR/Termometros/Completos"]
merger = PyPDF2.PdfMerger()
def obtener_archivos_pdf():
    for folder in folders:
        for filename in os.listdir(folder):
            if filename.endswith(".pdf"):
                merger.append(os.path.join(folder, filename))
    output_path = os.path.join("../Para Imprimir/Certificados", "Infrarrojos.pdf")
    with open(output_path, "wb") as f_out:
        merger.write(f_out)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Ejecutar scripts en orden con un archivo específico.")
    parser.add_argument(
        "--f", 
        help="Especifica el archivo que deben usar los scripts, por ejemplo: Tensiometros.xlsx"
    )
    parser.add_argument(
        "--c", 
        nargs="+", 
        help="Especifica el nombre de la nueva carpeta de drive"
    )
    args = parser.parse_args()
    obtener_archivos_pdf()

