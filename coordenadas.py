import pyautogui
import time

while True:
    x, y = pyautogui.position()
    print(f"Posición actual del mouse: {x}, {y}")
    time.sleep(1)  # Esperar un segundo antes de actualizar