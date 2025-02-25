import tkinter as tk
from tkinter import filedialog
import os
import subprocess
import threading
import json
from Google import Create_Service

CLIENT_SECRET_FILE = "client_secrets.json"
API_NAME = "drive"
API_VERSION = "v3"
SCOPES = ["https://www.googleapis.com/auth/drive"]
service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

selected_file_path = ""
directorios_found_list = []
directorio_base = os.getcwd()

def browse_file(entry):
    global selected_file_path
    selected_file_path = filedialog.askopenfilename()
    if selected_file_path:
        entry.delete(0, tk.END)
        entry.insert(0, selected_file_path)
        print(selected_file_path)

def search_in_directories_on_start(directorio_base):
    for root, dirs, files in os.walk(directorio_base):
        for dir_name in dirs:
            if dir_name == 'Completos':
                output_dir = os.path.join(root, dir_name)  # Directorio completo
                parent_dir = os.path.basename(os.path.dirname(root))  # Nombre corto
                if not os.listdir(output_dir):
                    nothing = True
                else:
                    directorios_found_list.append(parent_dir)
    print(f"Directorios con equipos encontrados: {directorios_found_list}\n")

def upload_files(directorio):
    with open("folder_id.json", "r") as f:
        folder_ids = json.load(f)
    
    if directorio in folder_ids:
        folder_id = folder_ids[directorio]
        os.chdir(directorio)
        print(directorio, folder_id)
        command = ["python3", "Drive2.py", "--i", folder_id, "--f", selected_file_path]
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        for line in iter(process.stdout.readline, ''):
            output_text.insert(tk.END, line)
            output_text.see(tk.END)
        process.stdout.close()
        process.wait()
        os.chdir(directorio_base)
        if process.returncode != 0:
            errors = process.stderr.read()
            output_text.insert(tk.END, f"\nErrores:\n{errors}")
            output_text.see(tk.END)
    else:
        output_text.insert(tk.END, f"Error: No se encontró el ID para el directorio '{directorio}' en folder_id.json\n")
        output_text.see(tk.END)

def search_in_directories(directorio_base):
    output_text.delete(1.0, tk.END)  # Limpia el área de salida antes de iniciar
    for root, dirs, files in os.walk(directorio_base):
        for dir_name in dirs:
            if dir_name == 'Completos':
                output_dir = os.path.join(root, dir_name)  # Directorio completo
                parent_dir = os.path.basename(os.path.dirname(root))  # Nombre corto
                if not os.listdir(output_dir):
                    nothing = True
                else:
                    if parent_dir not in directorios_found_list:
                        directorios_found_list.append(parent_dir)
    output_text.insert(tk.END, f"Directorios con equipos encontrados: {directorios_found_list}\n")
    output_text.see(tk.END)
    create_folder_padre(entry_drive_folder.get())

def upload_all_files():
    output_text.delete(1.0, tk.END)  # Limpia el área de salida antes de iniciar
    for directorio in directorios_found_list:
        upload_files(directorio)
    output_text.insert(tk.END, "Todos los equipos han sido subidos a su carpeta correspondiente\n")
    output_text.see(tk.END)

def create_folder_padre(nombre_carpeta):
    """Crear carpeta padre"""
    file_metadata = {
        'name': nombre_carpeta,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    file = service.files().create(
        body=file_metadata,
        fields='id'
    ).execute()
    output_text.insert(tk.END, f"Carpeta '{nombre_carpeta}' creada con ID: {file['id']}\n")
    output_text.see(tk.END)
    with open("folder_id.json", "w") as f:
        json.dump({"id_principal": file['id']}, f, indent=4)
    for directorio in directorios_found_list:
        crear_subcarpeta(directorio, file['id'])
    return file['id']

def crear_subcarpeta(nombre_carpeta, parent_id=None):
    """Crear una carpeta en Google Drive dentro de otra carpeta existente"""
    file_metadata = {
        'name': nombre_carpeta,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    if parent_id is not None:
        file_metadata['parents'] = [parent_id]
    else:
        output_text.insert(tk.END, "No se especificó una carpeta padre, error.\n")
        output_text.see(tk.END)
        return
    file = service.files().create(
        body=file_metadata,
        fields='id'
    ).execute()
    output_text.insert(tk.END, f"Carpeta '{nombre_carpeta}' creada con ID: {file['id']}\n")
    output_text.see(tk.END)
    data = {}
    if os.path.exists("folder_id.json"):
        with open("folder_id.json", "r") as f:
            data = json.load(f)
    data[nombre_carpeta] = file['id']
    with open("folder_id.json", "w") as f:
        json.dump(data, f, indent=4)
    return file['id']

def start_qr_pdf_generation():
    def run_command():
        command = ["python3", "ImprimirTodosQRS.py"]
        output_text.delete(1.0, tk.END)  # Limpia el área de salida antes de iniciar
        output_text.insert(tk.END, "Empezando la generacion de QRS...\n")
        output_text.see(tk.END)
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        for line in iter(process.stdout.readline, ''):
            output_text.insert(tk.END, line)  # Añade la línea al área de texto
            output_text.see(tk.END)  # Auto-scroll al final
        process.stdout.close()
        process.wait()  # Esperar a que el proceso termine
        if process.returncode != 0:
            errors = process.stderr.read()
            output_text.insert(tk.END, f"\nErrores:\n{errors}")
            output_text.see(tk.END)
    
    threading.Thread(target=run_command).start()

def on_option_select(*args):
    selected_option = selected_var.get()
    print(selected_option)

def on_metrologo_select(*args):
    selected_metrologo = selected_var_metrologo.get()
    print(selected_metrologo)

def get_drive_folder_text():
    return entry_drive_folder.get()

def eliminar_equipos_generados():
    def run_command():
        command = ["python3", "LimpiarTodo.py"]
        output_text.delete(1.0, tk.END)
        output_text.insert(tk.END, "Empezando a eliminar los equipos generados...\n")
        output_text.see(tk.END)
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        for line in iter(process.stdout.readline, ''):
            output_text.insert(tk.END, line)
            output_text.see(tk.END)
        process.stdout.close()
        process.wait()
        if process.returncode != 0:
            errors = process.stderr.read()
            output_text.insert(tk.END, f"\nErrores:\n{errors}")
            output_text.see(tk.END)
    threading.Thread(target=run_command).start()

def create_folders():
    def run_command():
        command = ["python3", "CrearCarpetas.py"]
        output_text.delete(1.0, tk.END)  # Limpia el área de salida antes de iniciar
        output_text.insert(tk.END, "Empezando a crear las carpetas...\n")
        output_text.see(tk.END)
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        for line in iter(process.stdout.readline, ''):
            output_text.insert(tk.END, line)  # Añade la línea al área de texto
            output_text.see(tk.END)  # Auto-scroll al final
        process.stdout.close()
        process.wait()  # Esperar a que el proceso termine
        if process.returncode != 0:
            errors = process.stderr.read()
            output_text.insert(tk.END, f"\nErrores:\n{errors}")
            output_text.see(tk.END)
    threading.Thread(target=run_command).start()

def start_generation():
    def run_command():
        output_text.delete(1.0, tk.END)  # Limpia el área de salida antes de iniciar
        print("Empezando el proceso...")
        try:
            ruta_excel = selected_file_path
            drive_name = entry_drive_folder.get()
            tipo_equipo = selected_var.get()
            metrologo = selected_var_metrologo.get()
            
            if not ruta_excel:
                raise ValueError("No se seleccionó un archivo Excel")
            if not metrologo:
                raise ValueError("No se seleccionó un metrologo")
            if not tipo_equipo:
                raise ValueError("No se seleccionó un tipo de equipo")
        except (NameError, ValueError) as e:
            output_text.insert(tk.END, f"Error: {str(e)}\n")
            output_text.see(tk.END)
            return
        os.chdir(tipo_equipo)
        print(os.getcwd())
        print(ruta_excel, drive_name, tipo_equipo, metrologo)
        if metrologo == "Ingeniero Ruben Ospina":
            metrologo = "Ruben"
        elif metrologo == "Ingeniera Luz Alejandra Vargas":
            metrologo = "Luz"
        command = ["python3", "main.py", "--f", ruta_excel,"--m", metrologo]

        # Ejecutar el comando y capturar la salida en tiempo real
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        for line in iter(process.stdout.readline, ''):
            output_text.insert(tk.END, line)  # Añade la línea al área de texto
            output_text.see(tk.END)  # Auto-scroll al final
        process.stdout.close()
        process.wait()  # Esperar a que el proceso termine
        output_text.insert(tk.END, "Reportes generados para el grupo de", selected_var.get(), "por favor verificar que los documentos se han generado correctamente antes de subirlos, los documentos se encuentran en la carpeta Para Imprimir/Certificados/", selected_var.get())
        os.chdir(directorio_base)
        if process.returncode != 0:
            errors = process.stderr.read()
            output_text.insert(tk.END, f"\nErrores:\n{errors}")
            output_text.see(tk.END)

    # Ejecutar el comando en un hilo para no bloquear la interfaz gráfica
    threading.Thread(target=run_command).start()

def main():
    global selected_var, selected_var_metrologo, entry_drive_folder, options, output_text
    search_in_directories_on_start(directorio_base)
    root = tk.Tk()
    root.title("Generacion de reportes de calibracion Ruben Ospina")
    root.geometry("700x500")

    # Crear un canvas y un frame para el contenido
    canvas = tk.Canvas(root)
    scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollbar_horizontal = tk.Scrollbar(root, orient="horizontal", command=canvas.xview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set, xscrollcommand=scrollbar_horizontal.set)

    # Colocar el canvas y los scrollbars en la ventana principal
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    scrollbar_horizontal.pack(side="bottom", fill="x")

    # Añadir widgets al frame scrollable
    label_file = tk.Label(scrollable_frame, text="Selecciona el archivo excel:")
    label_file.pack(pady=5, anchor="center")
    entry = tk.Entry(scrollable_frame, width=60)
    entry.pack(pady=5, anchor="center")
    browse_button = tk.Button(scrollable_frame, text="Abrir archivo", command=lambda: browse_file(entry))
    browse_button.pack(pady=10, anchor="center")

    label_picker = tk.Label(scrollable_frame, text="Selecciona un tipo de equipo para generar los certificados:")
    label_picker.pack(pady=10, anchor="center")
    options = ["Presiona aqui", "Pulsoximetros","Infrarrojos", "Basculas", "PesaBebe" ,"Tensiometros", "Monitores de signos", "Tensiometro digital",
               "Lamparas Fotocurado"]
    selected_var = tk.StringVar(scrollable_frame)
    selected_var.set(options[0]) 
    option_menu = tk.OptionMenu(scrollable_frame, selected_var, *options, command=on_option_select)
    option_menu.pack(pady=5, anchor="center")

    label_picker_metrologo = tk.Label(scrollable_frame, text="Selecciona el metrologo", font=("Arial", 12))
    label_picker_metrologo.pack(pady=5, anchor="center")
    options_metrologo = ["Presiona aqui", "Ingeniero Ruben Ospina", "Ingeniera Luz Alejandra Vargas"]
    selected_var_metrologo = tk.StringVar(scrollable_frame)
    selected_var_metrologo.set(options_metrologo[0])
    option_menu_metrologo = tk.OptionMenu(scrollable_frame, selected_var_metrologo, *options_metrologo, command=on_metrologo_select)
    option_menu_metrologo.pack(pady=5, anchor="center")

    start_button = tk.Button(scrollable_frame, text="Empezar a generar los certificados", command=lambda: start_generation())
    start_button.pack(pady=10, anchor="center")

    label_drive_folder = tk.Label(scrollable_frame, text="Escribe la carpeta para drive Ej. Puerto Boyaca 2024:")
    label_drive_folder.pack(pady=5, anchor="center")
    entry_drive_folder = tk.Entry(scrollable_frame, width=90)
    entry_drive_folder.pack(pady=5, anchor="center")

    create_folders_button = tk.Button(scrollable_frame, text="Crear carpetas de cada equipo encontrado", command=lambda: create_folders())
    create_folders_button.pack(pady=10, anchor="center")

    save_button = tk.Button(scrollable_frame, text="Crear carpeta de drive con las carpetas de cada equipo", command=lambda: search_in_directories(directorio_base))
    save_button.pack(pady=10, anchor="center")

    upload_button = tk.Button(scrollable_frame, text="Subir todos los equipos en su carpeta", command=lambda: upload_all_files())
    upload_button.pack(pady=10, anchor="center")

    genqr_button = tk.Button(scrollable_frame, text="Generar pdf con todos los QR", command=lambda: start_qr_pdf_generation())
    genqr_button.pack(pady=10, anchor="center")

    delete_all_button = tk.Button(scrollable_frame, text="Eliminar equipos generados", command=lambda: eliminar_equipos_generados())
    delete_all_button.pack(pady=10, anchor="center")

    output_text = tk.Text(scrollable_frame, wrap=tk.WORD, height=15, width=90)
    output_text.pack(pady=10, anchor="center")

    root.mainloop()

if __name__ == "__main__":
    main()
