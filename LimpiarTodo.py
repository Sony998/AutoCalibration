import os
import shutil
directorio_base = os.getcwd()
para_imprimir_certificados = os.path.join(directorio_base, "Para Imprimir/Certificados")
para_imprimir_qrs = os.path.join(directorio_base, "Para Imprimir/QRS")
def limpiar_directorio(directorio_base):
    for root, dirs, files in os.walk(directorio_base):
        for dir_name in dirs:
            if dir_name == 'OUTPUT':
                output_dir = os.path.join(root, dir_name)
                print("Procesando carpeta OUTPUT:", output_dir)
                vaciar_contenido(output_dir)

def vaciar_contenido(directorio):
    for item in os.listdir(directorio):
        item_path = os.path.join(directorio, item)
        if os.path.isfile(item_path):
            print("Eliminando archivo:", item_path)
            os.remove(item_path)  
        elif os.path.isdir(item_path):
            print("Accediendo al subdirectorio:", item_path)
            vaciar_contenido(item_path)
                            
if __name__ == "__main__":
    print("Limpiando directorio base", directorio_base)
    limpiar_directorio(directorio_base)
    for file in os.listdir(para_imprimir_certificados):
        file_path = os.path.join(para_imprimir_certificados, file)
        try:
            if os.path.isfile(file_path):
                print("Eliminando archivo:", file_path)
                os.remove(file_path)
            elif os.path.isdir(file_path):
                print("Eliminando directorio:", file_path)
                shutil.rmtree(file_path)
        except Exception as e:
            print("Error al eliminar:", e)
    print("Proceso de limpieza de carpetas completado.")
    for file in os.listdir(para_imprimir_qrs):
        file_path = os.path.join(para_imprimir_qrs, file)
        try:
            if os.path.isfile(file_path):
                print("Eliminando archivo:", file_path)
                os.remove(file_path)
            elif os.path.isdir(file_path):
                print("Eliminando directorio:", file_path)
                shutil.rmtree(file_path)
        except Exception as e:
            print("Error al eliminar:", e)
    print("Proceso de limpieza de carpeta para imprimir completado.")

