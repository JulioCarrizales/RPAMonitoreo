import pyautogui
import time
import os
import webbrowser
import pyperclip
import math

# Asegúrate de que pyautogui no falle por fail-safe
pyautogui.FAILSAFE = True

# Coordenadas de elementos clave en la pantalla (ajusta según tu pantalla)
COORDENADAS = {
    "usuario": (958, 486),
    "contrasena": (933, 533),
    "primer_filtro": (308, 229),
    "segundo_filtro": (124, 284),
    "tercer_filtro": (351, 296),
    "cuarto_filtro": (95,586),
    "quinto_filtro": (1461,639),
    "total_events": (1671, 522),  # Coordenadas del número total de eventos
    "expandir_tabla": (1729, 208),
    "export_boton": (965, 687),
    "export_csv": (549, 402),
    "export_complete": (608, 501),
    "primera_fila_tabla": (83, 443),  # Coordenadas de la primera fila
    "cerrar": (1890,18)
}

# Credenciales de inicio de sesión
usuario = "ATM01"
contrasena = "Teraware"

# Rutas y nombres de archivos
ruta_descarga = "C:/Users/Administrador/Downloads/mcevent.csv"
ruta_destino = "C:/Users/Administrador/Desktop/Reporte - copia/DATA"
nombre_nuevo = "truesight_noretiro.csv"

def abrir_pagina():
    """Abre el enlace en el navegador predeterminado."""
    webbrowser.open("https://tsps-vip.bn.com.pe:8043")
    time.sleep(5)

def mover_mouse_y_clic(x, y, delay=0.2):
    """Mueve el mouse a una posición y hace clic."""
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.click()

def escribir_texto(texto, delay=0.1):
    """Escribe texto con un retraso entre caracteres."""
    pyautogui.write(texto, interval=delay)

def iniciar_sesion():
    """Inicia sesión en la aplicación."""
    mover_mouse_y_clic(*COORDENADAS["usuario"])
    escribir_texto(usuario)
    mover_mouse_y_clic(*COORDENADAS["contrasena"])
    escribir_texto(contrasena)
    pyautogui.press("enter")
    time.sleep(5)

def aplicar_filtros():
    """Aplica filtros antes de la exportación."""
    mover_mouse_y_clic(*COORDENADAS["primer_filtro"])
    time.sleep(1)
    mover_mouse_y_clic(*COORDENADAS["segundo_filtro"])
    time.sleep(1)
    mover_mouse_y_clic(*COORDENADAS["tercer_filtro"])
    time.sleep(1)
    mover_mouse_y_clic(*COORDENADAS["cuarto_filtro"])
    time.sleep(1)
    mover_mouse_y_clic(*COORDENADAS["quinto_filtro"])
    time.sleep(1)

def obtener_total_eventos():
    """Obtiene el número total de eventos en pantalla."""
    mover_mouse_y_clic(*COORDENADAS["total_events"])
    time.sleep(0.5)
    pyautogui.doubleClick()
    pyautogui.hotkey("ctrl", "c")  # Copia al portapapeles
    time.sleep(0.5)
    try:
        total_events = int(pyperclip.paste())  # Pega el valor copiado
    except ValueError:
        total_events = 0
    return total_events

def seleccionar_y_descargar_lote(lote_numero, eventos_restantes):
    """Selecciona y descarga un lote de hasta 100 eventos usando Shift + flecha hacia abajo."""
    # Expande la vista de la tabla
    mover_mouse_y_clic(*COORDENADAS["expandir_tabla"])
    time.sleep(2)

    # Mueve el cursor a la primera fila y hace clic para seleccionarla
    mover_mouse_y_clic(*COORDENADAS["primera_fila_tabla"])
    time.sleep(0.5)

    # Calcula cuántas veces presionar la flecha hacia abajo
    filas_a_seleccionar = min(eventos_restantes - 1, 99)  # Ya seleccionamos una fila

    # Mantiene presionada la tecla Shift
    pyautogui.keyDown('shift')

    # Presiona la flecha hacia abajo para extender la selección
    for _ in range(filas_a_seleccionar):
        pyautogui.press('down')
        time.sleep(0.1)  # Pequeña pausa entre cada pulsación

    # Suelta la tecla Shift
    pyautogui.keyUp('shift')
    time.sleep(0.5)

    # Contrae la vista de la tabla
    mover_mouse_y_clic(*COORDENADAS["expandir_tabla"])
    time.sleep(1)

    # Descargar el archivo
    descargar_archivo()

    # Esperar a que la interfaz procese la descarga
    time.sleep(2)

def descargar_archivo():
    """Descarga el archivo en formato CSV y lo mueve a la ruta destino con el nuevo nombre."""
    mover_mouse_y_clic(*COORDENADAS["export_boton"])
    time.sleep(1)
    mover_mouse_y_clic(*COORDENADAS["export_csv"])
    time.sleep(1)
    mover_mouse_y_clic(*COORDENADAS["export_complete"])
    time.sleep(3)

    # Espera hasta que el archivo esté disponible y no esté en uso
    max_wait_time = 60  # Tiempo máximo de espera en segundos
    start_time = time.time()
    while True:
        if os.path.exists(ruta_descarga):
            try:
                # Intenta abrir el archivo para verificar si está en uso
                with open(ruta_descarga, 'r'):
                    pass
                break  # Si se puede abrir, salir del bucle
            except PermissionError:
                pass  # Si el archivo está en uso, continúa esperando
        if time.time() - start_time > max_wait_time:
            print("Error: Tiempo de espera excedido para la descarga del archivo.")
            return
        time.sleep(1)

    ruta_nueva = os.path.join(ruta_destino, nombre_nuevo)
    
    try:
        os.replace(ruta_descarga, ruta_nueva)
        print(f"Archivo descargado y movido a {ruta_nueva}")
    except PermissionError as e:
        print(f"Error al mover el archivo: {e}")


def proceso_completo():
    """Ejecuta todo el proceso de inicio de sesión, aplicación de filtros y descargas por lotes."""
    abrir_pagina()
    time.sleep(5)
    iniciar_sesion()
    aplicar_filtros()

    total_events = obtener_total_eventos()
    print(f"Total Events: {total_events}")

    if total_events > 0:
        lotes = math.ceil(total_events / 100)
        eventos_restantes = total_events

        for lote_numero in range(1, lotes + 1):
            print(f"Procesando lote {lote_numero} de {lotes}")
            seleccionar_y_descargar_lote(lote_numero, eventos_restantes)
            eventos_restantes -= 100

            if eventos_restantes > 0:
                # Refrescar la página y reaplicar los filtros
                pyautogui.hotkey('ctrl', 'r')
                time.sleep(5)
                aplicar_filtros()
                time.sleep(2)
            else:
                break
    else:
        print("No hay eventos para descargar.")

# Ejecutar el proceso completo
proceso_completo()
mover_mouse_y_clic(*COORDENADAS["cerrar"])
