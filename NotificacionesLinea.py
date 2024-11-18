import pyautogui
import time
import os

# Ruta del acceso directo que abre la aplicación
ruta_acceso_directo = r'C:\Users\Administrador\Desktop\Reporte - copia\Notificaiones en linea.lnk'

# Función para abrir la aplicación
def abrir_aplicacion():
    os.startfile(ruta_acceso_directo)

# Coordenadas a ajustar según tu pantallar
COORDENADAS = {
    "campo_usuario": (236,447),  # Coordenadas para escribir el usuario
    "campo_password": (217,518),  # Coordenadas para escribir la contraseña
    "boton_login": (113,600),  # Coordenadas para hacer clic en el botón de login
    "tablero_cajeros": (714,95),  # Coordenadas para hacer clic en el menú de tablero de cajeros
    "boton_excel": (136,154),  # Coordenadas para hacer clic en el icono de descargar excel
    "guardar_ruta": (883,228),  # Coordenadas para el campo de la ruta en el explorador
    "nombre_archivo": (858,636),  # Coordenadas para el campo de nombre del archivo
    "reemplazar_archivo" : (1055,535), #Coordenadas para la confirmación del reemplazo de archivo
    "boton_guardar": (1124,814),  # Coordenadas para hacer clic en el botón de guardar
    "cerrar": (1890,18) #Cerrar ventana
}

def mover_mouse_y_clic(x, y, delay=0.2):
    """Función para mover el mouse a una posición y hacer clic."""
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.click()

def escribir_texto(texto, delay=0.001):
    """Función para escribir texto con un retraso entre cada carácter."""
    pyautogui.write(texto, interval=delay)

# Proceso de automatización
def automatizar_notificaciones():
    abrir_aplicacion()
    
    time.sleep(5)  # Esperar a que la aplicación cargue
    
    # Rellenar usuario y contraseña
    mover_mouse_y_clic(*COORDENADAS["campo_usuario"])
    escribir_texto("egonzalesz")
    
    mover_mouse_y_clic(*COORDENADAS["campo_password"])
    escribir_texto("Banco12.")
    
    # Clic en login
    mover_mouse_y_clic(*COORDENADAS["boton_login"])
    
    time.sleep(5)  # Esperar a que cargue el sistema después de login
    
    # Clic en el menú Tablero de cajeros
    mover_mouse_y_clic(*COORDENADAS["tablero_cajeros"])
    
    time.sleep(3)  # Esperar a que cargue el tablero
    
    # Clic en el botón para descargar excel
    mover_mouse_y_clic(*COORDENADAS["boton_excel"])
    
    time.sleep(2)  # Esperar a que aparezca la ventana de guardar archivo
    
    # Escribir la ruta donde se guardará el archivo
    mover_mouse_y_clic(*COORDENADAS["guardar_ruta"])
    escribir_texto(r"C:\Users\Administrador\Desktop\Reporte - copia\DATA")
    pyautogui.press('enter')
    
    # Escribir el nombre del archivo
    mover_mouse_y_clic(*COORDENADAS["nombre_archivo"])
    escribir_texto("snl_saldos.xls")
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')



    mover_mouse_y_clic(*COORDENADAS["reemplazar_archivo"])

    mover_mouse_y_clic(*COORDENADAS["cerrar"])

# Ejecutar la automatización
automatizar_notificaciones()
