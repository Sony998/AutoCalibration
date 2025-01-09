import argparse
import os
import json
import re
from googleapiclient.http import MediaFileUpload
from Google import Create_Service

CLIENT_SECRET_FILE = "client_secrets.json"
API_NAME = "drive"
API_VERSION = "v3"
SCOPES = ["https://www.googleapis.com/auth/drive"]

directorios_found_list = []
# directory = "OUTPUT/Completos"
# directories = os.listdir(directory)
service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
directorio_base = os.getcwd()
def search_in_directories(directorio_base):
    for root, dirs, files in os.walk(directorio_base):
        for dir_name in dirs:
            if dir_name == 'Completos':
                output_dir = os.path.join(root, dir_name) ## Directorio completo
                parent_dir = os.path.basename(os.path.dirname(root)) ### Nombre corto
                #print("Procesando directorio:", output_dir)
                #print("Directorio padre:", parent_dir)
                # for item in os.listdir(output_dir):
                #     item_path = os.path.join(output_dir, item)
                #     if os.path.isfile(item_path):
                #         print("Archivo encontrado:", item_path)
                # vaciar_contenido(output_dir)
                if not os.listdir(output_dir):
                    nothing = True
                else:
                    print(f"Se han encontrado reportes en la carpeta {output_dir}")
                    directorios_found_list.append(parent_dir)
def vaciar_contenido(directorio):
    for item in os.listdir(directorio):
        item_path = os.path.join(directorio, item)
        if os.path.isfile(item_path):
            print("Eliminando archivo:", item_path)
            # os.remove(item_path)  
        elif os.path.isdir(item_path):
            print("Accediendo al subdirectorio:", item_path)
            vaciar_contenido(item_path)


def crear_carpeta(nombre_carpeta):
    """Crear carpeta padre"""
    file_metadata = {
        'name': nombre_carpeta,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    file = service.files().create(
        body=file_metadata,
        fields='id'
    ).execute()
    print(f"Carpeta '{nombre_carpeta}' creada con ID: {file['id']}")
    with open("folder_id.json", "w") as f:
        json.dump({"id_principal": file['id']}, f, indent=4)
    for directorio in directorios_found_list:
        crear_subcarpeta(directorio, file['id'])
    return file['id']

# def search_in_folders(folder_name):
#     for folder in directories:
#         if re.search(folder_name, folder):
#             return folder
def crear_subcarpeta(nombre_carpeta, parent_id=None):
    """Crear una carpeta en Google Drive dentro de otra carpeta existente"""
    file_metadata = {
        'name': nombre_carpeta,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    if parent_id is not None:
        file_metadata['parents'] = [parent_id]
    else:
        print("No se especificó una carpeta padre, error.")
        return
    file = service.files().create(
        body=file_metadata,
        fields='id'
    ).execute()
    print(f"Carpeta '{nombre_carpeta}' creada con ID: {file['id']}")
    data = {}
    if os.path.exists("folder_id.json"):
        with open("folder_id.json", "r") as f:
            data = json.load(f)
    data[nombre_carpeta] = file['id']
    with open("folder_id.json", "w") as f:
        json.dump(data, f, indent=4)
    return file['id']


# def upload_files_to_drive(folder_id, directory_files):
#     """Subir los archivos a la carpeta en Google Drive"""
#     file_urls = []
#     file_names = os.listdir(directory_files)
#     print(f"Archivos a subir: {file_names}")
#     mime_types = ["application/pdf"]
#     file_data = {}

#     for file_name in file_names:
#         file_name_without_extension = os.path.splitext(file_name)[0]
#         mime_type = "application/pdf"  # Asumiendo que todos los archivos son PDF
#         file_metadata = {
#             "name": file_name_without_extension,  # Guardamos el nombre sin '.pdf'
#             "parents": [folder_id]
#         }
#         media = MediaFileUpload(
#             os.path.join(directory_files, file_name),
#             mimetype=mime_type
#         )
#         file = service.files().create(
#             body=file_metadata,
#             media_body=media,
#             fields="id"
#         ).execute()
#         file_url = f"https://drive.google.com/file/d/{file.get('id')}/view"
#         file_urls.append(file_url)
#         print(f"Archivo subido: {file_name}, ID: {file.get('id')}, URL: {file_url}")
#         file_data[file_name_without_extension] = file_url  # Guardamos el nombre sin '.pdf'
#     json_file_path = os.path.join("file_urls.json")
#     with open(json_file_path, mode='w') as file:
#         json.dump(file_data, file, indent=4)
#     print(f"JSON generado en {json_file_path}")

def main():
    parser = argparse.ArgumentParser(description="Subir archivos a Google Drive en una carpeta específica.")
    parser.add_argument(
        "--c", 
        required=True, 
        help="Especifica el nombre de la carpeta en Google Drive donde se subirán los archivos."
    )
    args = parser.parse_args()
    search_in_directories(directorio_base)
    crear_carpeta(args.c)
if __name__ == "__main__":
    main()
