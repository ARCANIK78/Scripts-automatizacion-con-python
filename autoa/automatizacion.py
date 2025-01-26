import os
import re
import pandas as pd

# Ruta de la carpeta donde están los archivos
carpeta = r'C:\Users\DELL\Documents\Tareas de trabajos y cursos\sedeges\autoa\Nueva carpeta'

# Expresiones regulares para extraer información clave
patrones = {
    'procesador': re.compile(r'Specification\s*[:=]*\s*(.*)', re.IGNORECASE),
    'tarjeta_madre': re.compile(r'Mainboard Model\s*[:=]*\s*(.*)', re.IGNORECASE),
    'chipset': re.compile(r'Northbridge\s*[:=]*\s*(Intel.*?rev\.\s+\d+|AMD.*?rev\.\s+\d+)', re.IGNORECASE),
    'DDR': re.compile(r'Memory Type\s*[:=]*\s*(.*)', re.IGNORECASE),
    'ram': re.compile(r'Memory Size\s*[:=]*\s*(\d+\s*GBytes)', re.IGNORECASE),
    # Solo extraemos el Capacity y el Bus Type
    'disco_duro': re.compile(r'Capacity\s*[:=]*\s*([\d.]+\s*GB)', re.IGNORECASE),
    'bus_type': re.compile(r'Bus\s*Type\s*[:=]*\s*(\S+\s*\(\d+\)|\S+)', re.IGNORECASE),  # Captura cualquier tipo de Bus Type
    'tarjeta_grafica': re.compile(r'Display\s*adapter[\s\S]*?Name\s*[:=]*\s*(.*)', re.IGNORECASE),
    'sistema_operativo': re.compile(r'Windows Version\s*[:=]*\s*(Microsoft\sWindows\s.*)', re.IGNORECASE)
}

# Lista para almacenar los datos
data = []

# Iterar sobre los archivos .txt en la carpeta
for archivo in os.listdir(carpeta):
    if archivo.endswith('.txt'):
        ruta_archivo = os.path.join(carpeta, archivo)
        
        with open(ruta_archivo, 'r', encoding='latin-1') as file:
            contenido = file.read()

            # Buscar el valor de Capacity y Bus Type
            capacidad = patrones['disco_duro'].search(contenido)
            bus_type = patrones['bus_type'].search(contenido)

            # Extraer la información
            info_pc = {
                'Codigo BP': archivo,  # Agregar el nombre del archivo
                'Procesador': patrones['procesador'].search(contenido).group(1) if patrones['procesador'].search(contenido) else 'N/A',
                'Modelo': patrones['tarjeta_madre'].search(contenido).group(1) if patrones['tarjeta_madre'].search(contenido) else 'N/A',
                'Marca': patrones['chipset'].search(contenido).group(1) if patrones['chipset'].search(contenido) else 'N/A',
                'DDR': patrones['DDR'].search(contenido).group(1) if patrones['DDR'].search(contenido) else 'N/A',
                'RAM': patrones['ram'].search(contenido).group(1) if patrones['ram'].search(contenido) else 'N/A',
                'Disco Duro': capacidad.group(1) if capacidad else 'N/A',  # Solo el valor de Capacity
                'Bus Type': bus_type.group(1) if bus_type else 'N/A',  # Extraer el valor de Bus Type
                'Tarjeta Gráfica': patrones['tarjeta_grafica'].search(contenido).group(1) if patrones['tarjeta_grafica'].search(contenido) else 'N/A',
                'Sistema Operativo': patrones['sistema_operativo'].search(contenido).group(1) if patrones['sistema_operativo'].search(contenido) else 'N/A'
            }

            # Agregar los resultados al listado
            data.append(info_pc)

# Crear un DataFrame de Pandas y guardarlo en un archivo Excel
df = pd.DataFrame(data)

# Guardar el DataFrame en un archivo Excel
df.to_excel('informe_de_todas_las_pc.xlsx', index=False)
print("Informe generado en 'informe_de_todas_las_pc.xlsx'")
