from googleapiclient.http import MediaFileUpload
from Google import Create_Service
import os
import json
CLIENT_SECRET_FILE = "client_secrets.json"
API_NAME = "drive"
API_VERSION = "v3"
SCOPES = ["https://www.googleapis.com/auth/drive"]

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
file_urls = []
folder_id = "1NRa_gm6YdP0Rn6qLr4I4HDa5ZbYlcxBK"
directory_files = 'Completos'
file_names = os.listdir(directory_files)
print(file_names)
mime_types = ["application/pdf"]
for file_name in file_names:
    mime_type = "application/pdf"  # Assuming all files are PDFs, adjust if necessary
    file_metadata = {
        "name": file_name,
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
    json_file_path = os.path.join("file_urls.json")
    file_data = {file_name.replace(".pdf", ""): file_url for file_name, file_url in zip(file_names, file_urls)}
    with open(json_file_path, mode='w') as file:
        json.dump(file_data, file, indent=4)
    print(f"JSON generado en {json_file_path}")