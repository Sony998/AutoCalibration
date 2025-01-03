import subprocess
import argparse


scripts_generacion = [
    "GenerarCertificado.py",
    "GenerarGraficaError.py",
    "GenerarGraficaDesviacion.py",
    "putImages.py",
    "UnirPartes.py",
    "Drive.py",
    "GenQR.py",
    "ImprimirTodos.py",
    "ImprimirQRS.py"
]
def main(file_name, folder_name):
    print(f"Iniciando ejecución de scripts con el archivo: {file_name}")
    for script in scripts_generacion:
        try:
            print(f"Ejecutando {script} con el archivo: {file_name}")
            result = subprocess.run(
                ["python3", script, "--f", file_name, "--c", folder_name],
                check=True,
                text=True
            )
            print(f"{script} ejecutado con éxito.")
        except subprocess.CalledProcessError as e:
            print(f"Error al ejecutar {script}: {e}")
            break
    print("Ejecución completada en todo el proceso puede generar los reportes para otro tipo de equipo!!!.") 
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Ejecutar scripts en orden con un archivo específico.")
    parser.add_argument(
        "--f", 
        required=True, 
        help="Especifica el archivo que deben usar los scripts, por ejemplo: Tensiometros.xlsx"
    )
    parser.add_argument(
        "--c", 
        nargs="+", 
        help="Especifica el nombre de la nueva carpeta de drive"
    )
    args = parser.parse_args()
    args.c = " ".join(args.c)
    main(args.f, args.c)
