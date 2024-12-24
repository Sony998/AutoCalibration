import os
file = "file_urls.json"
directories_for_clean = [
                        "OUTPUT/Certificados",
                        "OUTPUT/Completos",
                        "OUTPUT/Graficos/Error/Diastolica",
                        "OUTPUT/Graficos/Error/Sistolica",
                        "OUTPUT/Graficos/Error/Frecuencia",
                        "OUTPUT/Graficos/Error/Saturacion",
                        "OUTPUT/Graficos/Error/Pulso",
                        "OUTPUT/Graficos/Desviacion/Diastolica",
                        "OUTPUT/Graficos/Desviacion/Sistolica",
                        "OUTPUT/Graficos/Desviacion/Frecuencia",
                        "OUTPUT/Graficos/Desviacion/Saturacion",
                        "OUTPUT/Graficos/Desviacion/Pulso",
                        "OUTPUT/Certificados",
                        "OUTPUT/Reportes/1",
                        "OUTPUT/Reportes/2",
                        "OUTPUT/Reportes/3",
                        "OUTPUT/Reportes/4",
                        "OUTPUT/Reportes/5",
                        "OUTPUT/Reportes/6",
                        "OUTPUT/Reportes/7",    
                        "OUTPUT/QRS",
                        "OUTPUT/Imprimir"
                        ]
for directory in directories_for_clean:
    print("Limpiando", directory)
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)

if os.path.isfile(file):
    os.remove(file) 
print("Archivo files eliminado")