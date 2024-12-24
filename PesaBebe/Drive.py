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

directory = "OUTPUT/Completos"
service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
def crear_carpeta(nombre_carpeta):
    """Crear una carpeta en Google Drive"""
    file_metadata = {
        'name': nombre_carpeta,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    file = service.files().create(
        body=file_metadata,
        fields='id'
    ).execute()
    print(f"Carpeta '{nombre_carpeta}' creada con ID: {file['id']}")
    return file['id']

def upload_files_to_drive(folder_id, directory_files):
    """Subir los archivos a la carpeta en Google Drive"""
    file_urls = []
    file_names = os.listdir(directory_files)
    print(f"Archivos a subir: {file_names}")
    mime_types = ["application/pdf"]
    file_data = {}

    for file_name in file_names:
        file_name_without_extension = os.path.splitext(file_name)[0]
        mime_type = "application/pdf"  # Asumiendo que todos los archivos son PDF
        file_metadata = {
            "name": file_name_without_extension,  # Guardamos el nombre sin '.pdf'
            "parents": [folder_id]
        }
        media = MediaFileUpload(
            os.path.join(directory_files, file_name),
            mimetype=mime_type
        )
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id"
        ).execute()
        file_url = f"https://drive.google.com/file/d/{file.get('id')}/view"
        file_urls.append(file_url)
        print(f"Archivo subido: {file_name}, ID: {file.get('id')}, URL: {file_url}")
        file_data[file_name_without_extension] = file_url  # Guardamos el nombre sin '.pdf'
    json_file_path = os.path.join("file_urls.json")
    with open(json_file_path, mode='w') as file:
        json.dump(file_data, file, indent=4)
    print(f"JSON generado en {json_file_path}")

def main():
    print(os.listdir(directory))
    parser = argparse.ArgumentParser(description="Subir archivos a Google Drive en una carpeta específica.")
    parser.add_argument(
        "--c", 
        required=True, 
        help="Especifica el nombre de la carpeta en Google Drive donde se subirán los archivos."
    )
    parser.add_argument(
        "--f", 
        required=True, 
        help="Especifica el archivo que deben usar los scripts, por ejemplo: Tensiometros.xlsx"
    )
    args = parser.parse_args()
    folder_id = crear_carpeta(args.c)
    upload_files_to_drive(folder_id, directory)
if __name__ == "__main__":
    main()
