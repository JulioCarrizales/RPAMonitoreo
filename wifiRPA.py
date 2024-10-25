import subprocess
import time

# Función para conectar a la red BNCORP
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

# Función para conectar a la red SBDIR
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

