import pyautogui
import time
#LINK DE PRUEBAS: http://10.7.10.196:9080/MonitorMultipagos/usuarios/sessionTerminada.action
#LINK ORIGINAL: http://10.7.10.164:9080/MonitorMultipagos/usuarios/sessionTerminada.action
while True:
    x, y = pyautogui.position()
    print(f"Posición actual del mouse: {x}, {y}")
    time.sleep(1)  # Esperar un segundo antes de actualizar           