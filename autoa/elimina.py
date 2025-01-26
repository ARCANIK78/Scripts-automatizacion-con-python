import os
import re

# Ruta de la carpeta donde están los archivos
carpeta = r'C:\Users\DELL\Documents\Tareas de trabajos y cursos\sedeges\autoa\Nueva carpeta'

# Expresión regular para encontrar y eliminar las líneas con "max capacity"
patron_max_capacity = re.compile(r'.*max capacity.*', re.IGNORECASE)

# Iterar sobre los archivos .txt en la carpeta
for archivo in os.listdir(carpeta):
    if archivo.endswith('.txt'):
        ruta_archivo = os.path.join(carpeta, archivo)
        
        # Abrir el archivo en modo lectura
        with open(ruta_archivo, 'r', encoding='latin-1') as file:
            contenido = file.readlines()

        # Filtrar las líneas que no contienen "max capacity"
        contenido_limpio = [linea for linea in contenido if not patron_max_capacity.match(linea)]

        # Escribir el contenido limpio de nuevo en el archivo
        with open(ruta_archivo, 'w', encoding='latin-1') as file:
            file.writelines(contenido_limpio)

        print(f"Se ha eliminado la línea con 'max capacity' en {archivo}.")
