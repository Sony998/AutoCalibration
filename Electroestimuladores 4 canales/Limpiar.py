import os
file = "file_urls.json"
directories_for_clean = [
                        "OUTPUT/Certificados",
                        "OUTPUT/Completos",
                        "OUTPUT/Graficos/Canal 1/Voltaje",
                        "OUTPUT/Graficos/Canal 1/Corriente",
                        "OUTPUT/Graficos/Canal 2/Voltaje",
                        "OUTPUT/Graficos/Canal 2/Corriente",
                        "OUTPUT/Graficos/Canal 3/Voltaje",
                        "OUTPUT/Graficos/Canal 3/Corriente",
                        "OUTPUT/Graficos/Canal 4/Voltaje",
                        "OUTPUT/Graficos/Canal 4/Corriente",
                        "OUTPUT/Certificados",
                        "OUTPUT/Reportes/1",
                        "OUTPUT/Reportes/2",
                        "OUTPUT/Reportes/3",
                        "OUTPUT/Reportes/4",
                        "OUTPUT/Reportes/5",
                        "OUTPUT/Reportes/6",
                        "OUTPUT/Reportes/7",
                        "OUTPUT/Reportes/8",
                        "OUTPUT/Reportes/9",
                        "OUTPUT/Reportes/10",
                        "OUTPUT/Reportes/11",
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