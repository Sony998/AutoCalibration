import os
import PyPDF2
folders = ["OUTPUT/Completos"]
merger = PyPDF2.PdfMerger()
for folder in folders:
    for filename in os.listdir(folder):
        if filename.endswith(".pdf"):
            merger.append(os.path.join(folder, filename))

output_path = os.path.join("../Para Imprimir/Certificados/", "ElectroEstimuladoresDobleCanal.pdf")
with open(output_path, "wb") as f_out:
    merger.write(f_out)