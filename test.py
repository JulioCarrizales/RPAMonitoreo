import pyautogui
import time
import os
import webbrowser
import pyperclip

# Asegúrate de que pyautogui no falle por fail-safe
pyautogui.FAILSAFE = True

# Coordenadas de elementos clave en la pantalla (ajusta según tu pantalla)
COORDENADAS = {
    "usuario": (958, 486),
    "contrasena": (933, 533),
    "primer_filtro": (357, 220),
    "segundo_filtro": (395, 354),
    "tercer_filtro": (195, 372),
    "total_events": (1679, 682),  # Coordenadas del número total de eventos
    "expandir_tabla": (1729, 208),
    "export_boton": (965, 687),
    "export_csv": (549, 402),
    "export_complete": (608, 501),
    "primera_fila_tabla": (83, 443),  # Coordenadas de la primera fila
}

# Credenciales de inicio de sesión
usuario = "ATM01"
contrasena = "Teraware"

# Ruta para mover y renombrar el archivo descargado
ruta_descarga = "C:/Users/Administrador/Downloads/mcevent.csv"
ruta_destino = "C:/Users/Administrador/Desktop/Reporte - copia/DATA"

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

def seleccionar_todas_las_filas():
    """Selecciona todas las filas desde la primera hasta la última sin mantener Shift durante el scroll."""
    # Expande la vista de la tabla
    mover_mouse_y_clic(*COORDENADAS["expandir_tabla"])
    time.sleep(2)

    x = COORDENADAS["primera_fila_tabla"][0]
    y_inicio = COORDENADAS["primera_fila_tabla"][1]

    # Paso 1: Hacer clic en la primera fila sin mantener Shift
    pyautogui.moveTo(x, y_inicio)
    pyautogui.click()
    time.sleep(0.5)

    # Paso 2: Realizar scroll hasta el final de la tabla sin mantener Shift
    # Mover el cursor dentro del área de la tabla antes de scrollear
    pyautogui.moveTo(x, y_inicio + 100)
    time.sleep(0.2)

    # Realizar scroll hacia abajo hasta el final
    scroll_times = 10  # Ajusta este número según sea necesario
    for _ in range(scroll_times):
        pyautogui.scroll(-500)  # Scroll hacia abajo
        time.sleep(0.5)

    # Paso 3: Mover el cursor a la última fila visible
    # Usamos una posición relativa en la pantalla
    screen_width, screen_height = pyautogui.size()
    x_fin = x  # Mantenemos la misma coordenada X
    y_fin = screen_height - 150  # Un poco por encima del borde inferior

    pyautogui.moveTo(x_fin, y_fin)
    time.sleep(0.5)

    # Paso 4: Mantener Shift y hacer clic en la última fila
    pyautogui.keyDown('shift')
    pyautogui.click()
    pyautogui.keyUp('shift')
    time.sleep(0.5)

    # Contrae la vista de la tabla
    mover_mouse_y_clic(*COORDENADAS["expandir_tabla"])
    time.sleep(1)

def descargar_archivo():
    """Descarga el archivo en formato CSV."""
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

    # Renombra y mueve el archivo descargado
    nuevo_nombre = os.path.join(ruta_destino, f"truesight_noretiro_{int(time.time())}.csv")
    try:
        os.rename(ruta_descarga, nuevo_nombre)
        print(f"Archivo descargado y movido a {nuevo_nombre}")
    except PermissionError as e:
        print(f"Error al mover el archivo: {e}")

def proceso_completo():
    """Ejecuta todo el proceso de inicio de sesión, aplicación de filtros y descargas."""
    abrir_pagina()
    time.sleep(5)
    iniciar_sesion()
    aplicar_filtros()

    total_events = obtener_total_eventos()
    print(f"Total Events: {total_events}")
    if total_events > 0:
        seleccionar_todas_las_filas()
        descargar_archivo()
    else:
        print("No hay eventos para descargar.")

# Ejecutar el proceso completo
proceso_completo()