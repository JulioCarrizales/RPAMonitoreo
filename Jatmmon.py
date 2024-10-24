import pyautogui
import time
import os

# Ruta del acceso directo que abre la aplicación
ruta_acceso_directo = r'C:\Users\Administrador\Desktop\Reporte - copia\fts.jar - Acceso directo.lnk'

# Función para abrir la aplicación
def abrir_aplicacion():
    os.startfile(ruta_acceso_directo)
