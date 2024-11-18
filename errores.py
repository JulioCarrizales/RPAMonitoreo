import os
import pyautogui
import time
import win32com.client as win32
import pythoncom
from datetime import datetime
import shutil

# Ruta del acceso directo que abre la aplicación
ruta_acceso_directo = r'C:\Users\Administrador\Desktop\Reporte - copia\ErroresRemesas_MonitoreoCajeros.exe - Acceso directo.lnk'

# Función para abrir la aplicación
def abrir_aplicacion():
    os.startfile(ruta_acceso_directo)


COORDENADAS= {
     "Usuario": (1057, 528),
     "Password": (1035, 579),
     "Reporte_estadistico": (1867,659),
     "Fecha": (418,162),
     "Mes": (767,204),
     "Estado": (357,361),
     "Calcular":(1541,182),
     "Excel": (1541,352),
     "Salir": (1530,517),
     "Cerrar_1":(1893,21),
     "Cerrar_2":(1247,324),
     "Cerrar_3" : (1845,305),
     "Extra": (1363,1043),
     "Bug": (1498,310)
}

def mover_mouse_y_clic(x, y, delay=0.2):
    """Función para mover el mouse a una posición y hacer clic."""
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.click()

def escribir_texto(texto, delay=0.05):
    """Función para escribir texto con un retraso entre cada carácter."""
    pyautogui.write(texto, interval=delay)

def obtener_mes_anterior():
    """Función que obtiene el número del mes anterior en formato de dos dígitos."""
    mes_actual = datetime.now().month
    mes_anterior = mes_actual - 1 if mes_actual > 1 else 12
    mes_anterior_str = f"{mes_anterior:02d}"
    return mes_anterior_str

def seleccionar_mes_anterior():
    """Función que selecciona el mes anterior en el sistema."""
    mes_anterior = obtener_mes_anterior()
    pyautogui.write(mes_anterior)


def automatizar_errores():
    time.sleep(10)

    abrir_aplicacion()

    time.sleep(5)

    mover_mouse_y_clic(*COORDENADAS["Usuario"])
    escribir_texto("JPALOMINO")

    mover_mouse_y_clic(*COORDENADAS["Password"])
    escribir_texto("JPALOMINO1")

    pyautogui.press("enter")

    time.sleep(15)

    mover_mouse_y_clic(*COORDENADAS["Reporte_estadistico"])
    time.sleep(10)

    mover_mouse_y_clic(*COORDENADAS["Fecha"])
    time.sleep(3)

    mover_mouse_y_clic(*COORDENADAS["Mes"])
    seleccionar_mes_anterior()

    mover_mouse_y_clic(*COORDENADAS["Estado"])

    mover_mouse_y_clic(*COORDENADAS["Calcular"])
    time.sleep(3)

    mover_mouse_y_clic(*COORDENADAS["Excel"])
    time.sleep(75)  # Espera adicional para asegurarse de que Excel esté abierto

    
    mover_mouse_y_clic(*COORDENADAS["Extra"])
    time.sleep(5)
    pyautogui.press("enter")
    mover_mouse_y_clic(*COORDENADAS["Cerrar_3"])
    time.sleep(3)
    pyautogui.press("enter")
    time.sleep(3)

    escribir_texto("rep_TicketsPendientes.xlsx")

    pyautogui.press("enter")
    time.sleep(3)
    pyautogui.press("tab")
    pyautogui.press("enter")
    time.sleep(3)
    mover_mouse_y_clic(*COORDENADAS["Salir"])
    mover_mouse_y_clic(*COORDENADAS["Cerrar_1"])
    mover_mouse_y_clic(*COORDENADAS["Cerrar_2"])
    time.sleep(10)


def mover_archivo_errores(nombre_archivo_errores, ruta_origen_errores, ruta_destino_errores):
    """Función que busca un archivo por su nombre en la ruta de origen y lo mueve a la ruta de destino."""
    archivo_origen = os.path.join(ruta_origen_errores, nombre_archivo_errores)
    archivo_destino = os.path.join(ruta_destino_errores, nombre_archivo_errores)

    # Verificar si el archivo existe en la ruta de origen
    if os.path.exists(archivo_origen):
        try:
            # Mover el archivo a la ruta de destino
            shutil.move(archivo_origen, archivo_destino)
            print(f"El archivo '{nombre_archivo_errores}' ha sido movido exitosamente a '{ruta_destino_errores}'.")
        except Exception as e:
            print(f"Ocurrió un error al mover el archivo: {e}")
    else:
        print(f"El archivo '{nombre_archivo_errores}' no fue encontrado en '{ruta_origen_errores}'.")

nombre_archivo_errores = "rep_TicketsPendientes.xlsx"
ruta_origen_errores = r"C:\Users\Administrador\Documents"
ruta_destino_errores = r"C:\Users\Administrador\Desktop\Reporte - copia\DATA"

automatizar_errores()
mover_archivo_errores(nombre_archivo_errores,ruta_origen_errores,ruta_destino_errores)