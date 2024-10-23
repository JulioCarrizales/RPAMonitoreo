import os
import pyautogui
import time
import logging
import webbrowser

# Configuración del logging
logging.basicConfig(level=logging.INFO)

# URL a abrir
url = "https://tsps-vip.bn.com.pe:8043/#/monitoring/events"

# Abre la URL en el navegador predeterminado
webbrowser.open(url)

# Coordenadas proporcionadas
COORDENADAS = {
    "primer_archivo": (81, 817),
    "ultimo_archivo": (83, 954),
    "boton_filtro": (355, 224),
    "opcion_filtro": (435, 363),
    "boton_aplicar_filtro": (1490, 368),
    "boton_cargar": (300, 400),
    "boton_aceptar_exito": (400, 500),
    "boton_exportar": (968, 682),
    "opcion_csv": (541, 402),
    "boton_exportar_final": (646, 496),
}

def obtener_numero_archivos():
    """Función simulada para obtener el número de archivos disponibles en la página."""
    return 22  # Cambia este valor según sea necesario

def mover_mouse_y_clic(x, y, delay=0.5):
    """Función para mover el mouse a una posición y hacer clic."""
    logging.info(f"Moviendo el mouse a las coordenadas ({x}, {y}) y haciendo clic.")
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.click()

def aplicar_filtro():
    """Función para aplicar el filtro antes de seleccionar archivos."""
    logging.info("Aplicando el filtro...")
    
    x_filtro, y_filtro = COORDENADAS["boton_filtro"]
    mover_mouse_y_clic(x_filtro, y_filtro)
    time.sleep(1)
    
    x_opcion, y_opcion = COORDENADAS["opcion_filtro"]
    mover_mouse_y_clic(x_opcion, y_opcion)
    time.sleep(1)

    x_aplicar, y_aplicar = COORDENADAS["boton_aplicar_filtro"]
    mover_mouse_y_clic(x_aplicar, y_aplicar)
    time.sleep(2)
    logging.info("Filtro aplicado exitosamente.")

def seleccionar_archivos():
    """Función para seleccionar archivos dependiendo de la cantidad disponible."""
    numero_archivos = obtener_numero_archivos()
    logging.info(f"Número de archivos disponibles: {numero_archivos}")

    archivos_a_seleccionar = min(numero_archivos, 100)

    x1, y1 = COORDENADAS["primer_archivo"]
    mover_mouse_y_clic(x1, y1)
    time.sleep(1)

    if archivos_a_seleccionar > 1:
        pyautogui.scroll(-1000)
        time.sleep(1)

        pyautogui.keyDown('shift')
        for i in range(1, archivos_a_seleccionar):
            x2, y2 = COORDENADAS["ultimo_archivo"]
            mover_mouse_y_clic(x2, y2)
            time.sleep(0.5)
        pyautogui.keyUp('shift')

    logging.info(f"Se seleccionaron {archivos_a_seleccionar} archivos exitosamente.")

def cargar_y_aceptar():
    """Función para hacer clic en el botón de 'Cargar' y luego aceptar el mensaje."""
    x, y = COORDENADAS["boton_cargar"]
    mover_mouse_y_clic(x, y)
    time.sleep(2)
    esperar_y_aceptar_exito()

def esperar_y_aceptar_exito():
    """Esperar a que aparezca el mensaje de 'Datos cargados exitosamente' y dar clic en 'Aceptar'."""
    logging.info("Esperando a que aparezca el mensaje de éxito...")
    time.sleep(3)
    x, y = COORDENADAS["boton_aceptar_exito"]
    mover_mouse_y_clic(x, y)
    logging.info("Mensaje de éxito aceptado.")

def exportar_archivos():
    """Función para exportar archivos seleccionados."""
    logging.info("Exportando archivos...")
    
    x_exportar, y_exportar = COORDENADAS["boton_exportar"]
    mover_mouse_y_clic(x_exportar)
    time.sleep(1)
    
    x_opcion_csv, y_opcion_csv = COORDENADAS["opcion_csv"]
    mover_mouse_y_clic(x_opcion_csv)
    time.sleep(1)

    x_exportar_final, y_exportar_final = COORDENADAS["boton_exportar_final"]
    mover_mouse_y_clic(x_exportar_final)
    time.sleep(2)
    logging.info("Exportación completada.")

def automatizar_proceso():
    """Automatiza el proceso de selección de archivos, carga y exportación."""
    logging.info("Iniciando proceso de selección, carga y exportación de archivos...")
    webbrowser.open(url)
    time.sleep(10)  # Esperar a que la página cargue completamente

    aplicar_filtro()
    seleccionar_archivos()
    cargar_y_aceptar()
    exportar_archivos()

# Ejecutar el proceso de automatización
automatizar_proceso()
