import os 
import json

nombretipo = input("Ingrese el nombre del tipo de equipo:")
directorios = ["Cetificados", "Completos", "Graficos", "partesReporte"  "QRS", "Rerportes/1","Reportes/2","Reportes/3","Reportes/4"]
archivo_json = 'enlaces.json'
links = []

if not os.path.exists(nombretipo):
    for directorio in directorios:
        os.makedirs(nombretipo + "/" + directorio)
        print("Directorio " + directorio + " creado")

def obtener_urls():
    with open(archivo_json) as json_file:
        data = json.load(json_file)
        for p in data['links']:
            link = {'url': p['url'], 'name': p['name']}
            links.append(link)
        return data

def download_files():
    for link in links:
        os.system( "wget " + link['url'])
        os.system("mv " + link['name'] + " " + nombretipo + "/")
        print("Archivo descargado")

obtener_urls()
download_files()
