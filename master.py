import time
import schedule

from errores import (
    automatizar_errores,
    mover_archivo_errores,
    nombre_archivo_errores,
    ruta_origen_errores,
    ruta_destino_errores
)
from Jatmmon import (
    automatizar_jatmmon,
    mover_archivo_jatmmon,
    nombre_archivo_jatmmon,
    ruta_origen_jatmmon,
    ruta_destino_jatmmon
)

from truesight_noretiro import proceso_completo_noretiros
from truesight_alertascriticas import proceso_completo_criticos

from NotificacionesLinea import automatizar_notificaciones

from main import automatizar_proceso  # Importamos solo la función necesaria

def proceso_completo():
    automatizar_errores()
    mover_archivo_errores(
        nombre_archivo_errores,
        ruta_origen_errores,
        ruta_destino_errores
    )
    time.sleep(5)
    automatizar_jatmmon()
    mover_archivo_jatmmon(
        nombre_archivo_jatmmon,
        ruta_origen_jatmmon,
        ruta_destino_jatmmon
    )
    time.sleep(5)
    proceso_completo_noretiros()
    time.sleep(5)
    proceso_completo_criticos()
    time.sleep(5)
    automatizar_notificaciones()
    automatizar_proceso()
    pass

schedule.every().hour.do(proceso_completo)

while True:
    schedule.run_pending()
    time.sleep(1)