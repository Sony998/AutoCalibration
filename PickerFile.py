import tkinter as tk
from tkinter import filedialog
import os
import subprocess
import threading

def browse_file(entry):
    global selected_file_path
    selected_file_path = filedialog.askopenfilename()
    if selected_file_path:
        entry.delete(0, tk.END)
        entry.insert(0, selected_file_path)
        print(selected_file_path)

def on_option_select(*args):
    selected_option = selected_var.get()
    print(selected_option)

def get_drive_folder_text():
    return entry_drive_folder.get()

def start_generation():
    def run_command():
        output_text.delete(1.0, tk.END)  # Limpia el área de salida antes de iniciar
        print("Empezando el proceso...")
        ruta_excel = selected_file_path
        drive_name = entry_drive_folder.get()
        tipo_equipo = selected_var.get()

        os.chdir(tipo_equipo)
        print(os.getcwd())
        print(ruta_excel, drive_name, tipo_equipo)
        command = f"python3 main.py --f {ruta_excel} --c {drive_name}"
        process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        for line in process.stdout:
            output_text.insert(tk.END, line)
            output_text.see(tk.END)
        process.stdout.close()
        process.wait()
    # Ejecutar el comando en un hilo para no bloquear la interfaz gráfica
    threading.Thread(target=run_command).start()

def main():
    global selected_var, entry_drive_folder, options, output_text
    root = tk.Tk()
    root.title("Generacion de reportes de calibracion Ruben Ospina")
    root.geometry("700x500")

    # Entrada para seleccionar archivo
    label_file = tk.Label(root, text="Selecciona el archivo excel:")
    label_file.pack(pady=5)
    entry = tk.Entry(root, width=60)
    entry.pack(pady=5)
    browse_button = tk.Button(root, text="Abrir archivo", command=lambda: browse_file(entry))
    browse_button.pack(pady=10)

    # Menú de opciones
    label_picker = tk.Label(root, text="Selecciona un tipo de equipo para generar los certificados:")
    label_picker.pack(pady=10)
    options = ["Presiona aqui", "Pulsoximetros", "Tensiometros", "Monitores de signos", "Tensiometro digital"]
    selected_var = tk.StringVar(root)
    selected_var.set(options[0]) 

    option_menu = tk.OptionMenu(root, selected_var, *options, command=on_option_select)
    option_menu.pack(pady=5)

    # Entrada para carpeta de Drive
    label_drive_folder = tk.Label(root, text="Escribe la carpeta para drive Ej. Puerto Boyaca 2024:")
    label_drive_folder.pack(pady=5)
    entry_drive_folder = tk.Entry(root, width=90)
    entry_drive_folder.pack(pady=5)
    save_button = tk.Button(root, text="Guardar nombre de la carpeta de drive", command=lambda: print(get_drive_folder_text()))
    save_button.pack(pady=10)

    # Botón para iniciar el proceso
    start_button = tk.Button(root, text="Empezar a generar los certificados", command=lambda: start_generation())
    start_button.pack(pady=10)

    # Área de texto para la salida
    output_text = tk.Text(root, wrap=tk.WORD, height=15, width=90)
    output_text.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
