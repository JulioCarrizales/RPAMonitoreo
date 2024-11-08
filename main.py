import os
import pyautogui
import time
import logging
import win32com.client as win32
import pythoncom
import webbrowser
import subprocess


# Configuración del log para guardar mensajes en un archivo
logging.basicConfig(filename='carga_datos.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Ruta de la carpeta donde está el archivo de Excel
ruta_aplicativo = r"C:\Users\Administrador\Desktop\Reporte - copia\Aplicativo"
ruta_excel = os.path.join(ruta_aplicativo, "ggrpt_monitoreo_atms.xlsm")

# Ruta de la carpeta donde están los archivos de datos
ruta_data = r"C:\Users\Administrador\Desktop\Reporte - copia\DATA"
ruta_reporte = os.path.join(ruta_aplicativo, "rpt_monitoreo_atms.xlsm")

# Relación de hojas y archivos correspondientes
archivos_necesarios = {
    "scaj_registro": "rpt_SuperDetallado.xlsx",
    "eeyrr_ticketsPendiente": "rep_TicketsPendientes.xlsx",
    "truesight_noRetiro": "truesight_noretiro.csv",
    "snl_saldos": "snl_saldos.xls",
    "truesight_alertasCritical": "truesight_critical.csv",
    "jatmmon_indicadores": "jatmmon_indicadores.xls"
}

# Coordenadas proporcionadas
COORDENADAS = {
    "boton_cargar_datos": (730, 632),
    "boton_adjuntar": (1293, 494),
    "menu_desplegable_hoja": (1319, 415),
    "boton_cargar": (1106, 621),
    "boton_salir": (832, 618),
    "boton_generar_reporte": (1057, 624),
    "boton_aceptar_exito": (1020, 619), # Coordenada ajustada del botón "Aceptar"
    "boton_aceptar_exito": (1020, 619),  # Coordenada ajustada del botón "Aceptar"
    "whatsapp_adjuntar": (989, 952),  # Coordenada del icono "adjuntar" en WhatsApp Web
    "whatsapp_adjuntar_archivo": (1078, 575),  # Coordenada del icono "adjuntar archivo"
    "whatsapp_enviar": (1858, 951),  # Coordenada del botón "Enviar"
    "cerrar": (1890,18),
    "abrir_mai": (1740,303),
    "porsiacaso": (1325,1040)
}

def abrir_excel():
    """Función para abrir el archivo de Excel automáticamente."""
    pythoncom.CoInitialize()
    try:
        excel = win32.Dispatch('Excel.Application')  # Crear una nueva instancia de Excel
        excel.Visible = True  # Mostrar Excel
        workbook = excel.Workbooks.Open(ruta_excel)  # Abrir el archivo
        mover_mouse_y_clic(*COORDENADAS["porsiacaso"])
        pyautogui.press("enter")
        mover_mouse_y_clic(*COORDENADAS["abrir_mai"])
        logging.info(f"Archivo {ruta_excel} abierto correctamente.")
        return excel, workbook
    except Exception as e:
        logging.error(f"Error al abrir el archivo de Excel: {e}")
        return None, None

def mover_mouse_y_clic(x, y, delay=0.2):
    """Función para mover el mouse a una posición y hacer clic."""
    logging.info(f"Moviendo el mouse a las coordenadas ({x}, {y}) y haciendo clic.")
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.click()

def escribir_texto(texto, delay=0.001):
    """Función para escribir texto con un retraso entre cada carácter."""
    pyautogui.write(texto, interval=delay)

def seleccionar_hoja(nombre_hoja):
    """Función para seleccionar la hoja en el menú desplegable."""
    x, y = COORDENADAS["menu_desplegable_hoja"]
    mover_mouse_y_clic(x, y)  # Menú desplegable del nombre de la hoja
    time.sleep(1)
    escribir_texto(nombre_hoja)
    pyautogui.press('enter')
    time.sleep(1)

def seleccionar_archivo_y_adjuntar(archivo):
    """Función para seleccionar un archivo desde el explorador de Windows."""
    x, y = COORDENADAS["boton_adjuntar"]
    mover_mouse_y_clic(x, y)  # Botón de "Adjuntar"
    time.sleep(2)
    ruta_archivo = os.path.join(ruta_data, archivo)
    escribir_texto(ruta_archivo)
    pyautogui.press('enter')
    time.sleep(2)

def esperar_y_aceptar_exito():
    """Esperar a que aparezca el mensaje de 'Datos cargados exitosamente' y dar clic en 'Aceptar'."""
    logging.info("Esperando a que aparezca el mensaje de éxito...")
    time.sleep(3)  # Esperamos un momento a que se procese la carga
    x, y = COORDENADAS["boton_aceptar_exito"]
    mover_mouse_y_clic(x, y)  # Clic en el botón "Aceptar" del mensaje de éxito
    logging.info("Mensaje de éxito aceptado.")

def cargar_y_aceptar():
    """Función para hacer clic en el botón de 'Cargar' y luego aceptar el mensaje."""
    x, y = COORDENADAS["boton_cargar"]
    mover_mouse_y_clic(x, y)  # Botón de "Cargar"
    time.sleep(2)

    # Esperar el mensaje de éxito y aceptarlo
    esperar_y_aceptar_exito()

def conectar_sbdir():
    try:
        print("Intentando conectar a la red SBDIR...")
        # Comando para conectarse a la red WiFi SBDIR
        comando = ['netsh', 'wlan', 'connect', 'name=SBDIR']
        print(f"Ejecutando comando: {' '.join(comando)}")
        resultado = subprocess.run(comando, capture_output=True, text=True)
        print(f"Resultado del comando: {resultado.stdout}")

        # Verificar si la conexión fue exitosa
        if "completed successfully" in resultado.stdout.lower():
            print("Conectado a la red SBDIR exitosamente.")
        else:
            print("No se pudo conectar a la red SBDIR. Verifica que la red esté guardada y al alcance.")
            print(resultado.stdout)
    except Exception as e:
        print(f"Ocurrió un error: {e}")

def enviar_reporte_por_whatsapp():
    conectar_sbdir()
    """Función para enviar el archivo de Excel por WhatsApp usando pyautogui."""
    try:
        # Abrir WhatsApp Web
        webbrowser.open("https://web.whatsapp.com/")
        logging.info("Abriendo WhatsApp Web...")
        time.sleep(60)  # Esperar a que se cargue completamente
        # Buscar el contacto o número de teléfono
        pyautogui.click(412,235)  # Ajusta la coordenada para el cuadro de búsqueda de WhatsApp
        time.sleep(2)
        escribir_texto("Reportes ATM")  # Escribir el número de contacto
        time.sleep(2)
        pyautogui.press('enter')  # Abrir el chat del contacto
        time.sleep(2)
        # Adjuntar archivo
        x, y = COORDENADAS["whatsapp_adjuntar"]
        mover_mouse_y_clic(x, y)  # Clic en el icono de adjuntar (clip)
        time.sleep(2)
        x, y = COORDENADAS["whatsapp_adjuntar_archivo"]
        mover_mouse_y_clic(x, y)  # Clic en "adjuntar archivo"
        time.sleep(2)
        # Escribir la ruta del archivo de Excel y enviarlo
        escribir_texto(ruta_reporte)
        time.sleep(2)
        pyautogui.press('enter')  # Seleccionar el archivo
        time.sleep(2)
        # Clic en "Enviar"
        x, y = COORDENADAS["whatsapp_enviar"]
        mover_mouse_y_clic(x, y)  # Botón de "Enviar"
        logging.info("Archivo enviado por WhatsApp exitosamente.")
    except Exception as e:
        logging.error(f"Error al enviar el archivo por WhatsApp: {e}")

def generar_reporte_y_salir():
    """Función para salir y luego generar el reporte."""
    x, y = COORDENADAS["boton_salir"]
    mover_mouse_y_clic(x, y)  # Botón de "Salir"
    time.sleep(2)

    x, y = COORDENADAS["boton_generar_reporte"]
    mover_mouse_y_clic(x, y)  # Botón de "Generar Reporte"
    time.sleep(70)

    # Después de generar el reporte, esperar y enviar el archivo por WhatsApp
    enviar_reporte_por_whatsapp()

def conectar_bncorp():
    try:
        print("Intentando conectar a la red BNCORP...")
        # Comando para conectarse a la red WiFi BNCORP
        comando = ['netsh', 'wlan', 'connect', 'name=BNCORP']
        print(f"Ejecutando comando: {' '.join(comando)}")
        resultado = subprocess.run(comando, capture_output=True, text=True)
        print(f"Resultado del comando: {resultado.stdout}")

        # Verificar si la conexión fue exitosa
        if "completed successfully" in resultado.stdout.lower():
            print("Conectado a la red BNCORP exitosamente.")
        else:
            print("No se pudo conectar a la red BNCORP. Verifica que la red esté guardada y al alcance.")
            print(resultado.stdout)
    except Exception as e:
        print(f"Ocurrió un error: {e}")

def toggle_wifi(interface_name="Wi-Fi", delay=5):
    """
    Apaga y enciende el adaptador Wi-Fi especificado.

    Args:
        interface_name (str): Nombre del adaptador de red Wi-Fi. Por defecto es "Wi-Fi".
        delay (int): Tiempo en segundos para esperar entre apagar y encender. Por defecto es 5 segundos.
    """
    try:
        # Apagar el adaptador de Wi-Fi
        print(f"Apagando el adaptador de red: {interface_name}...")
        subprocess.check_call(['netsh', 'interface', 'set', 'interface', interface_name, 'admin=disable'], shell=True)
        print(f"Adaptador {interface_name} apagado. Esperando {delay} segundos...")
        time.sleep(delay)

        # Encender el adaptador de Wi-Fi
        print(f"Encendiendo el adaptador de red: {interface_name}...")
        subprocess.check_call(['netsh', 'interface', 'set', 'interface', interface_name, 'admin=enable'], shell=True)
        print(f"Adaptador {interface_name} encendido exitosamente.")

    except subprocess.CalledProcessError as e:
        print(f"Error al cambiar el estado del adaptador {interface_name}: {e}")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")


def cerrar_todo_excel():
    try:
        excel_app = win32.Dispatch("Excel.application")
        for workbook in excel_app.Workbooks:
            workbook.Close(SaveChanges = False)

        excel_app.Quit()
    except Exception as e:
        print ("Error: ", e)


def automatizar_proceso():
    toggle_wifi()
    time.sleep(5)
    conectar_sbdir()
    time.sleep(5)
    conectar_sbdir()
    time.sleep(10)
    """Automatiza todo el proceso de carga de archivos y generación de reportes."""
    logging.info("Iniciando proceso de carga de datos...")

    # Abrir Excel automáticamente
    excel, workbook = abrir_excel()
    if excel is None or workbook is None:
        logging.error("No se pudo abrir el archivo de Excel. Proceso abortado.")
        return

    # Esperar 10 segundos antes de comenzar la interacción con Excel
    logging.info("Esperando 10 segundos para que Excel se cargue completamente...")
    time.sleep(10)

    # Hacer clic en el botón de "Cargar Datos"
    x, y = COORDENADAS["boton_cargar_datos"]
    mover_mouse_y_clic(x, y)
    time.sleep(2)

    # Repetir el proceso para cada hoja y archivo
    for hoja, archivo in archivos_necesarios.items():
        logging.info(f"Seleccionando hoja {hoja} y cargando archivo {archivo}")

        # Seleccionar la hoja desde el menú desplegable
        seleccionar_hoja(hoja)

        # Adjuntar el archivo correspondiente
        seleccionar_archivo_y_adjuntar(archivo)

        # Cargar y aceptar el mensaje de éxito
        cargar_y_aceptar()

    # Después de cargar "jatmmon_indicadores", salir y generar el reporte
    logging.info("Cargando última hoja 'jatmmon_indicadores'. Saliendo y generando reporte.")
    generar_reporte_y_salir()
    time.sleep(20)
    conectar_bncorp()
    mover_mouse_y_clic(*COORDENADAS["cerrar"])
    pyautogui.press("enter")
    cerrar_todo_excel()

# Ejecutar el proceso de automatización
automatizar_proceso()