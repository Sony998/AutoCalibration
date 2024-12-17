from googleapiclient.http import MediaFileUpload
from Google import Create_Service
import os
import json

CLIENT_SECRET_FILE = "client_secrets.json"
API_NAME = "drive"
API_VERSION = "v3"
SCOPES = ["https://www.googleapis.com/auth/drive"]

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
folder_id = "1NRa_gm6YdP0Rn6qLr4I4HDa5ZbYlcxBK"
directory_files = 'Completos'
specific_file_name = "T13081124133.pdf"
file_path = os.path.join(directory_files, specific_file_name)

if os.path.exists(file_path):
    mime_type = "application/pdf"
    file_metadata = {
        "name": specific_file_name,
        "parents": [folder_id]
    }
    media = MediaFileUpload(file_path, mimetype=mime_type)
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()
    file_url = f"https://drive.google.com/file/d/{file.get('id')}/view"
    print(f"Archivo subido: {specific_file_name}, ID: {file.get('id')}, URL: {file_url}")

    json_file_path = os.path.join("file_urls.json")
    file_data = {specific_file_name.replace(".pdf", ""): file_url}
    with open(json_file_path, mode='w') as file:
        json.dump(file_data, file, indent=4)
    print(f"JSON generado en {json_file_path}")
else:
    print(f"El archivo {specific_file_name} no se encuentra en la carpeta {directory_files}")