import pandas as pd
import os
import pyautogui
import time
import logging
import webbrowser
import subprocess
import pyperclip

COORDENADAS = {
    "boton_cargar_datos": (468, 435),
    "boton_adjuntar": (1037, 421),
    "menu_desplegable_hoja": (1039, 358),
    "boton_cargar": (893,509),
    "boton_salir": (716,509),
    "boton_generar_reporte": (704,442),
    "boton_aceptar_exito": (847,501), # Coordenada ajustada del botón "Aceptar"
    "boton_aceptar_exito": (847,501),  # Coordenada ajustada del botón "Aceptar"
    "whatsapp_adjuntar": (745,807),  # Coordenada del icono "adjuntar" en WhatsApp Web
    "whatsapp_adjuntar_archivo": (645,408),  # Coordenada del icono "adjuntar archivo"
    "whatsapp_enviar": (1531,784),  # Coordenada del botón "Enviar"
    "cerrar": (1577,17),
    "abrir_mai": (1272,161),
    "porsiacaso": (712,877),
}


def escribir_texto(texto, delay=0.001):
    """Función para escribir texto con un retraso entre cada carácter."""
    pyautogui.write(texto, interval=delay)


def mover_mouse_y_clic(x, y, delay=0.2):
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.click()

def triple_click(x, y, delay=0.2):
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.tripleClick()

def double_click(x, y, delay=0.2):
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.doubleClick()

def anti_click(x, y, delay=0.2):
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.rightClick()

def escribir_texto(texto, delay=0.001):
    pyautogui.write(texto, interval=delay)

def whatsapp ():
    try:
        
        # Abrir WhatsApp Web
        webbrowser.open("https://web.whatsapp.com/")
        time.sleep(30)  # Esperar a que se cargue completamente
        # Buscar el contacto o número de teléfono
        pyautogui.click(179,178)  # Ajusta la coordenada para el cuadro de búsqueda de WhatsApp
        time.sleep(2)
        escribir_texto("Reportes Canales")  # Escribir el número de contacto
        time.sleep(2)
        pyautogui.press('enter')  # Abrir el chat del contacto
        time.sleep(2)
        pyautogui.click(702,802)
        # Escribir la ruta del archivo de Excel y enviarlo
        escribir_texto("RPA_PAGALO - TICKETS PLOMOS (1/2) OPERATIVO")
        pyautogui.press('enter')  # Seleccionar el archivo
        time.sleep(10)
        logging.info("Archivo enviado por WhatsApp exitosamente.")
        pyautogui.click(1578,16)
        time.sleep(2)
        pyautogui.press('enter')  # Seleccionar el archivo
        time.sleep(2)
    except Exception as e:
        logging.error(f"Error al enviar el archivo por WhatsApp: {e}")


if __name__ == "__main__":
       # Llamar a la función whatsapp
    whatsapp()