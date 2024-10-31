import pyautogui
import time
import shutil
import os

# Ruta del acceso directo que abre la aplicación
ruta_acceso_directo = r'C:\Users\Administrador\Desktop\Reporte - copia\fts.jar - Acceso directo.lnk'

# Función para abrir la aplicación
def abrir_aplicacion():
    os.startfile(ruta_acceso_directo)

COORDENADAS= {
    "jatmmmon_usuario": (972,518), #Uusario de jatmmon
    "jatmmon_contraseña": (955,555),
    "Boton de Jatmmon": (1259,360),
    "Export to excel": (1127,315),
    "Close": (1895,13),
    "Confirmacion_1": (893,552),
    "Confirmacion_2": (902,546),
    "Desktop": (1206,320),
    "Nombre_archivo": (836, 630)
}

def mover_mouse_y_clic(x, y, delay=0.2):
    """Función para mover el mouse a una posición y hacer clic."""
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.click()

def escribir_texto(texto, delay=0.001):
    """Función para escribir texto con un retraso entre cada carácter."""
    pyautogui.write(texto, interval=delay)

def automatizar_jatmmon ():
    abrir_aplicacion()

    time.sleep(5)

    mover_mouse_y_clic(*COORDENADAS["jatmmmon_usuario"])
    escribir_texto("jatoche")

    mover_mouse_y_clic(*COORDENADAS["jatmmon_contraseña"])
    escribir_texto("Valeria99,")

    pyautogui.press("enter")

    mover_mouse_y_clic(*COORDENADAS["Boton de Jatmmon"])

    mover_mouse_y_clic(*COORDENADAS["Export to excel"])
    
    mover_mouse_y_clic(*COORDENADAS["Desktop"])

    mover_mouse_y_clic(*COORDENADAS["Nombre_archivo"])
    escribir_texto("jatmmon_indicadores.xls")
    pyautogui.press("enter")
    pyautogui.press("enter")

    time.sleep(3) 

    mover_mouse_y_clic(*COORDENADAS["Close"])
    mover_mouse_y_clic(*COORDENADAS["Confirmacion_1"])

    mover_mouse_y_clic(*COORDENADAS["Close"])
    mover_mouse_y_clic(*COORDENADAS["Confirmacion_2"])

def mover_archivo_jatmmon(nombre_archivo_jatmmon, ruta_origen_jatmmon, ruta_destino_jatmmon):
    """Función que busca un archivo por su nombre en la ruta de origen y lo mueve a la ruta de destino."""
    archivo_origen = os.path.join(ruta_origen_jatmmon, nombre_archivo_jatmmon)
    archivo_destino = os.path.join(ruta_destino_jatmmon, nombre_archivo_jatmmon)

    # Verificar si el archivo existe en la ruta de origen
    if os.path.exists(archivo_origen):
        try:
            # Mover el archivo a la ruta de destino
            shutil.move(archivo_origen, archivo_destino)
            print(f"El archivo '{nombre_archivo_jatmmon}' ha sido movido exitosamente a '{ruta_destino_jatmmon}'.")
        except Exception as e:
            print(f"Ocurrió un error al mover el archivo: {e}")
    else:
        print(f"El archivo '{nombre_archivo_jatmmon}' no fue encontrado en '{ruta_origen_jatmmon}'.")


nombre_archivo_jatmmon = "jatmmon_indicadores.xls"
ruta_origen_jatmmon = r"C:\Users\Administrador\Desktop"
ruta_destino_jatmmon = r"C:\Users\Administrador\Desktop\Reporte - copia\DATA"


automatizar_jatmmon()
mover_archivo_jatmmon(nombre_archivo_jatmmon, ruta_origen_jatmmon, ruta_destino_jatmmon)
    
