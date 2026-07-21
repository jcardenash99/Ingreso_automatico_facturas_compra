# -*- coding: utf-8 -*-
"""
Utilidad para calibrar coordenadas de pantalla.

Ejecuta este script, luego mueve el mouse sobre cada botón/casilla
de Siigo (Crear, Factura de compra, casilla Tipo, casilla Fecha, etc.)
y anota las coordenadas (x, y) que se imprimen en la consola.

Presiona Ctrl+C en la consola para detener.
"""

import ctypes
import pyautogui
import time

# Debe coincidir con el fix del script principal, o las coordenadas
# que leas aquí no van a coincidir con las que usa siigo_automatizacion.py
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    ctypes.windll.user32.SetProcessDPIAware()

print("Mueve el mouse sobre los elementos que quieras calibrar.")
print("Coordenadas actuales cada 1 segundo (Ctrl+C para detener):\n")

try:
    while True:
        x, y = pyautogui.position()
        print(f"x={x:5d}  y={y:5d}", end="\r")
        time.sleep(1)
except KeyboardInterrupt:
    print("\nCalibración detenida.")
