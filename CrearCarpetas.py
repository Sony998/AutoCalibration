import argparse
import os
import json
from googleapiclient.http import MediaFileUpload
from Google import Create_Service

CLIENT_SECRET_FILE = "client_secrets.json"
API_NAME = "drive"
API_VERSION = "v3"
SCOPES = ["https://www.googleapis.com/auth/drive"]

directorios_found_list = []
service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
directorio_base = os.getcwd()

def search_in_directories(directorio_base):
    for root, dirs, files in os.walk(directorio_base):
        for dir_name in dirs:
            if dir_name == 'Completos':
                output_dir = os.path.join(root, dir_name)
                parent_dir = os.path.basename(os.path.dirname(root))
                if os.listdir(output_dir):
                    print(f"Se han encontrado reportes en la carpeta {output_dir}")
                    directorios_found_list.append(parent_dir)

def crear_subcarpeta(nombre_carpeta, parent_id):
    file_metadata = {
        'name': nombre_carpeta,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    file = service.files().create(
        body=file_metadata,
        fields='id'
    ).execute()
    print(f"Carpeta '{nombre_carpeta}' creada con ID: {file['id']}")
    with open("folder_id.json", "r") as f:
        data = json.load(f)
    data[nombre_carpeta] = file['id']
    with open("folder_id.json", "w") as f:
        json.dump(data, f, indent=4)
    return file['id']

def main():
    search_in_directories(directorio_base)
    
    with open("folder_id.json", "r") as f:
        data = json.load(f)
        id_principal = data.get("id_principal")
    
    if id_principal:
        for directorio in directorios_found_list:
            crear_subcarpeta(directorio, id_principal)
    else:
        print("No se encontró el ID principal en el archivo folder_id.json")

if __name__ == "__main__":
    main()
