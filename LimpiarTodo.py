import os
import shutil
directorio_base = os.getcwd()

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
