import pyautogui
import time
import os
import webbrowser
import pyperclip
import csv

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
    "export_boton": (969, 315),
    "export_csv": (549, 402),
    "export_complete": (608, 501),
    "primera_fila_tabla": (83, 443),  # Coordenadas de la primera fila
    "cerrar": (1890,18),
    "abrir":(1624,222),
    "bloque": (81, 953),
}

# Credenciales de inicio de sesión
usuario = "ATM01"
contrasena = "Teraware"

# Rutas y nombres de archivos
ruta_descarga = "C:/Users/Administrador/Downloads/mcevent.csv"
ruta_destino = "C:/Users/Administrador/Desktop/Reporte - copia/DATA"
nombre_nuevo_base = "truesight_critical"
nombre_archivo_final = "truesight_critical.csv"

def abrir_pagina():
    """Abre el enlace en el navegador predeterminado."""
    webbrowser.open("https://tsps-vip.bn.com.pe:8043")
    mover_mouse_y_clic(*COORDENADAS["abrir"])
    time.sleep(5)
    mover_mouse_y_clic(*COORDENADAS["usuario"])
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

def seleccionar_y_descargar_lote(lote_numero, eventos_restantes, primera_iteracion=False):
    """Selecciona y descarga un lote de hasta 100 eventos."""
    if primera_iteracion:
        # Expande la vista de la tabla solo una vez al inicio
        mover_mouse_y_clic(*COORDENADAS["expandir_tabla"])
        time.sleep(2)
        # Mueve el cursor a la primera fila y hace clic para seleccionarla
        mover_mouse_y_clic(*COORDENADAS["primera_fila_tabla"])
        time.sleep(0.5)
    else:
        # Mueve al siguiente evento sin deseleccionar
        pyautogui.press('down')  # Mueve al siguiente evento
        time.sleep(0.2)
        # Hace clic en 'bloque'
        mover_mouse_y_clic(*COORDENADAS["bloque"])
        time.sleep(0.2)

    # Calcula cuántas veces presionar la flecha hacia abajo
    filas_a_seleccionar = min(eventos_restantes - 1, 99)  # Ya seleccionamos una fila

    # Mantiene presionada la tecla Shift
    pyautogui.keyDown('shift')

    # Presiona la flecha hacia abajo para extender la selección
    for _ in range(filas_a_seleccionar):
        pyautogui.press('down')
        time.sleep(0.05)  # Pequeña pausa entre cada pulsación

    # Suelta la tecla Shift
    pyautogui.keyUp('shift')
    time.sleep(0.5)

    # Descargar el archivo
    descargar_archivo(lote_numero)

    # Esperar a que la interfaz procese la descarga (15-20 segundos)
    time.sleep(20)

def descargar_archivo(lote_numero):
    """Descarga el archivo en formato CSV y lo mueve a la ruta destino con un nombre único."""
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
                with open(ruta_descarga, 'r', encoding='utf-8'):
                    pass
                break  # Si se puede abrir, salir del bucle
            except PermissionError:
                pass  # Si el archivo está en uso, continúa esperando
        if time.time() - start_time > max_wait_time:
            print("Error: Tiempo de espera excedido para la descarga del archivo.")
            return
        time.sleep(1)

    # Genera un nombre único para cada lote
    nombre_nuevo = f"{nombre_nuevo_base}_lote_{lote_numero}.csv"
    ruta_nueva = os.path.join(ruta_destino, nombre_nuevo)

    try:
        os.replace(ruta_descarga, ruta_nueva)
        print(f"Archivo del lote {lote_numero} descargado y movido a {ruta_nueva}")
    except PermissionError as e:
        print(f"Error al mover el archivo: {e}")

def combinar_archivos():
    """Combina los archivos descargados en uno solo, ignorando la cabecera de los archivos subsecuentes."""
    archivos = []
    # Obtiene la lista de archivos descargados, excluyendo el archivo combinado
    for archivo in os.listdir(ruta_destino):
        if archivo.startswith(nombre_nuevo_base) and archivo.endswith('.csv') and archivo != nombre_archivo_final:
            archivos.append(os.path.join(ruta_destino, archivo))

    archivos.sort()  # Ordena los archivos para mantener el orden de los lotes

    ruta_archivo_combinado = os.path.join(ruta_destino, nombre_archivo_final)

    with open(ruta_archivo_combinado, 'w', newline='', encoding='utf-8') as archivo_salida:
        escritor = None
        for idx, archivo in enumerate(archivos):
            with open(archivo, 'r', encoding='utf-8') as f:
                lector = csv.reader(f)
                if idx == 0:
                    # Escribe la cabecera y los datos
                    escritor = csv.writer(archivo_salida)
                    for fila in lector:
                        escritor.writerow(fila)
                else:
                    # Ignora la primera fila (cabecera)
                    next(lector, None)
                    for fila in lector:
                        escritor.writerow(fila)
    print(f"Archivos combinados y guardados en {ruta_archivo_combinado}")

    # Opcional: eliminar los archivos individuales después de combinarlos, excepto el combinado
    for archivo in archivos:
        os.remove(archivo)
        print(f"Archivo temporal eliminado: {archivo}")

def proceso_completo_noretiros():
    """Ejecuta todo el proceso de inicio de sesión, aplicación de filtros y descargas por lotes."""
    abrir_pagina()
    time.sleep(5)
    iniciar_sesion()
    aplicar_filtros()

    total_events = obtener_total_eventos()
    print(f"Total de eventos: {total_events}")

    if total_events > 0:
        eventos_restantes = total_events
        primera_iteracion = True
        lote_numero = 1

        while eventos_restantes > 0:
            print(f"Procesando lote {lote_numero}")
            seleccionar_y_descargar_lote(lote_numero, eventos_restantes, primera_iteracion)
            eventos_procesados = min(eventos_restantes, 100)
            eventos_restantes -= eventos_procesados

            # Actualiza variables para la siguiente iteración
            primera_iteracion = False
            lote_numero += 1

        print("Todos los lotes han sido procesados.")

        # Combinar los archivos descargados
        combinar_archivos()

    else:
        print("No hay eventos para descargar.")

    # Cerrar la aplicación (si es necesario)
    mover_mouse_y_clic(*COORDENADAS["cerrar"])
    
# Ejecutar el proceso completo
proceso_completo_noretiros()
