import os
import pyautogui
import win32com.client as win32
import time
import logging
import pythoncom

# Configuración del log para guardar mensajes en un archivo
logging.basicConfig(filename='carga_datos.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Ruta de la carpeta donde están los archivos de datos
ruta_data = r"C:\Users\Administrador\Desktop\Reporte - copia\DATA"

# Nombre del archivo Excel
archivo_excel = "ggrpt_monitoreo_atms.xlsm"

# Relación de hojas y archivos correspondientes
archivos_necesarios = {
    "scaj_registro": "rpt_SuperDetallado.xlsx",
    "eeyrr_ticketsPendiente": "rep_TicketsPendientes.xlsx",
    "truesight_noRetiro": "rpt_TSnoRetiro.csv",
    "snl_saldos": "TableAllATMs.xls",
    "truesight_alertasCritical": "rpt_AlertasCritical.csv",
    "jatmmon_indicadores": "rep_Indicadores.xls"
}

def verificar_excel_abierto():
    """Verifica si Excel está abierto y si el archivo correcto está activo."""
    pythoncom.CoInitialize()
    try:
        excel = win32.GetObject(None, "Excel.Application")  # Obtener Excel si está abierto
        for workbook in excel.Workbooks:
            if workbook.Name == archivo_excel:
                logging.info(f"Archivo {archivo_excel} encontrado y activo.")
                return excel, workbook
        logging.error(f"El archivo {archivo_excel} no está abierto.")
        return None, None
    except Exception as e:
        logging.error("Excel no está abierto o no se puede conectar: " + str(e))
        return None, None

def mover_mouse_y_clic(x, y, delay=0.5):
    """Función para mover el mouse a una posición y hacer clic."""
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.click()

def escribir_texto(texto, delay=0.2):
    """Función para escribir texto con un retraso entre cada carácter."""
    pyautogui.write(texto, interval=delay)

def seleccionar_hoja(nombre_hoja):
    """Función para seleccionar la hoja en el menú desplegable."""
    mover_mouse_y_clic(400, 350)  # Ajusta las coordenadas a la posición del menú desplegable
    time.sleep(1)
    escribir_texto(nombre_hoja)
    pyautogui.press('enter')
    time.sleep(1)

def seleccionar_archivo_y_adjuntar(archivo):
    """Función para seleccionar un archivo desde el explorador de Windows."""
    mover_mouse_y_clic(500, 400)  # Coordenadas del botón "Adjuntar"
    time.sleep(2)
    ruta_archivo = os.path.join(ruta_data, archivo)
    escribir_texto(ruta_archivo)
    pyautogui.press('enter')
    time.sleep(2)

def cargar_y_salir():
    """Función para hacer clic en el botón de 'Cargar' y luego en 'Salir'."""
    mover_mouse_y_clic(500, 500)  # Coordenadas del botón "Cargar" (ajusta según tu pantalla)
    time.sleep(2)
    mover_mouse_y_clic(500, 550)  # Coordenadas del botón "Salir"
    time.sleep(2)

def generar_reporte():
    """Función para hacer clic en el botón de 'Generar Reporte'."""
    mover_mouse_y_clic(550, 600)  # Coordenadas del botón "Generar Reporte"
    time.sleep(5)

def automatizar_proceso():
    """Automatiza todo el proceso de carga de archivos y generación de reportes."""
    logging.info("Iniciando proceso de carga de datos...")
    
    # Verificar si el archivo de Excel está abierto
    excel, workbook = verificar_excel_abierto()
    if excel is None or workbook is None:
        logging.error("No se pudo encontrar el archivo Excel. Proceso abortado.")
        return
    
    # Repetir el proceso para cada hoja y archivo
    for hoja, archivo in archivos_necesarios.items():
        logging.info(f"Seleccionando hoja {hoja} y cargando archivo {archivo}")
        
        # Seleccionar la hoja desde el menú desplegable
        seleccionar_hoja(hoja)
        
        # Adjuntar el archivo correspondiente
        seleccionar_archivo_y_adjuntar(archivo)
    
    # Cargar y salir
    logging.info("Haciendo clic en 'Cargar' y luego en 'Salir'")
    cargar_y_salir()
    
    # Generar el reporte
    logging.info("Haciendo clic en 'Generar Reporte'")
    generar_reporte()

# Ejecutar el proceso de automatización
automatizar_proceso()
