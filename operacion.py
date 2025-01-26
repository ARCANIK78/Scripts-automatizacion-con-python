import platform
import psutil
import openpyxl
from openpyxl import Workbook
import subprocess
import re
import winreg 

def get_system_inf():
    try:
        if platform.system()=="Windows":
            brand = subprocess.check_output("wmic computersystem get manufacturer", shell=True).decode().split('\n')[1].strip()
        else:
            brand = "Desconoido"
    except:
        brand = "No detectado"
    
    try:
        if platform.system() == "Windows":
            procesador = subprocess.check_output("wmic cpu get Name", shell=True).decode().split('\n')[1].strip()
        else:
            procesador = platform.processor()  # Alternativa para otros sistemas
    except:
        procesador = "No detectado"
    
    try:
        if platform.system() == "Windows":
            motherboard = subprocess.check_output("wmic baseboard get product,Manufacturer", shell=True).decode().split('\n')[1].strip()
        else:
            motherboard = "Desconocido"
    except:
        motherboard= "No encontrado"
    
    ram = f"{round(psutil.virtual_memory().total / (1024**3), 2)} GB"
    
    try: #Tarejta de Video
        if platform.system() == "Windows":
            gpu = subprocess.check_output("wmic path win32_videocontroller get name", shell=True).decode().split('\n')[1].strip()
        else:
            gpu = "Desconocido"
    except:
        gpu = "No Detecto"
    
    # Precisa detección del sistema operativo para Windows
    sistema_Operativo_version = get_precise_windows_version()
    
    office_version = get_office_version()
            
    return{
        "Marca": brand,
        "Procesador": procesador,
        "Tajeta Madre": motherboard,
        "Memoria Ram": ram,
        "Tarjeta de video": gpu,
        "Sistema Operativo": sistema_Operativo_version,
        "Office": office_version
    }

def get_office_version():
    # Rutas en el registro para diferentes versiones de Office
    office_reg_paths = [
        r"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",  # Office 2016 y 2019
        r"SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot",  # Office 2013
        r"SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot",  # Office 2010
        r"SOFTWARE\Microsoft\Office\12.0\Common\InstallRoot",  # Office 2007
        r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",  # Office 365 Click-to-Run
    ]

    office_version = "No detectado"
    version_mapping = {
        "16": "Office 2016/2019",
        "15": "Office 2013",
        "14": "Office 2010",
        "12": "Office 2007"
    }

    for path in office_reg_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                version_number = winreg.QueryValueEx(key, "VersionToReport")[0]
                major_version = version_number.split('.')[0]  # Obtener el número principal (16, 15, etc.)
                office_version = version_mapping.get(major_version, f"Office versión desconocida ({version_number})")
                break
        except FileNotFoundError:
            continue

    return office_version
def get_precise_windows_version():
    if platform.system() == "Windows":
        try:
            # Obtiene la versión detallada de Windows con `wmic`
            version_output = subprocess.check_output("wmic os get Caption,Version", shell=True).decode().split('\n')
            caption = version_output[1].strip()  # Nombre del sistema operativo (ej. Windows 10 Pro)
            version = version_output[2].strip()  # Número de versión completo (ej. 10.0.22000)

            # Combina la información para obtener una descripción completa
            sistema_Operativo_version = f"{caption} (Versión {version})"
        except Exception as e:
            sistema_Operativo_version = f"Windows (versión no detectada): {str(e)}"
    else:
        sistema_Operativo_version = f"{platform.system()} {platform.version()}"

    return sistema_Operativo_version
  

info = get_system_inf()

