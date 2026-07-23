# -*- coding: utf-8 -*-
"""
Semi-automatización de ingreso de facturas de compra en Siigo Nube
desde un archivo de Excel.

FASE 1: Encabezado de la factura (fecha, prefijo, no. factura, proveedor)
La iteración de la tabla de ítems se agrega en una fase posterior.

Requisitos:
    pip install pywin32 pyautogui
"""

import os
import time
import ctypes
import datetime
import win32com.client
import win32gui
import win32con
import win32clipboard
import win32process
import win32api
import pyautogui

# ---------------------------------------------------------
# IMPORTANTE: sin esto, en pantallas con escalado (125%, 150%, etc.)
# las coordenadas que calibraste NO van a coincidir con donde realmente
# hace clic pyautogui. Esta línea debe ir ANTES de cualquier otra cosa.
# ---------------------------------------------------------
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_AWARE
except Exception:
    ctypes.windll.user32.SetProcessDPIAware()  # fallback para versiones viejas de Windows

# =========================================================
# CONFIGURACIÓN - AJUSTA ESTOS VALORES ANTES DE EJECUTAR
# =========================================================

RUTA_EXCEL = r"C:\ruta\a\tu_archivo.xlsx"
NOMBRE_HOJA = 1  # o el nombre de la hoja, ej. "Facturas"

# Nota: ya no identificamos las ventanas por título (Chrome cambia el
# título con cada pestaña), sino por clase de ventana de Windows, que es
# estable. Ver funciones ir_a_siigo() / ir_a_excel() más abajo.

# Coordenadas de pantalla (x, y) de cada elemento en Siigo.
# Para obtenerlas: ejecuta el script "calibrar_coordenadas.py" (ver abajo)
# y pasa el mouse sobre cada botón/casilla para leer su posición.
# Nota: prefijo, no. factura y proveedor NO necesitan coordenada propia,
# porque una vez estás en la casilla fecha, el "tab" navega solo entre
# esos campos.
COORD = {
    "crear": (1372, 143),
    "factura_compra": (1183, 257),
    "casilla_tipo": (595, 383),
    "casilla_fecha": (635, 418),
}

# Fila donde están los datos de encabezado en tu Excel.
# (Fila 1 = encabezado mayor, Fila 2 = encabezado de fecha/prefijo/factura/proveedor,
#  Fila 3 = primera y única fila de datos de encabezado)
FILA_ENCABEZADO = 3
COL_FECHA = 1
COL_PREFIJO = 2
COL_NO_FACTURA = 3
COL_PROVEEDOR = 4

# --- Tabla de ítems ---
# Fila 5 = encabezado de la tabla de ítems (código, bodega, cantidad, precio, descuento, iva)
# Fila 6 = primera fila de datos de ítems (de aquí en adelante, itera hacia abajo)
FILA_ITEMS_INICIO = 6
COL_CODIGO = 1
COL_BODEGA = 2
COL_CANTIDAD = 3
COL_PRECIO = 4
COL_DESCUENTO = 5
COL_IVA = 6

# Pausa entre movimientos de mouse (segundos) - ajusta según qué tan rápido
# responde tu equipo/Siigo
PAUSA_CORTA = 0.4
pyautogui.PAUSE = 0.15  # pausa por defecto entre acciones de pyautogui


# =========================================================
# FUNCIONES DE VENTANAS
# =========================================================
#
# En vez de forzar el cambio de ventana con SetForegroundWindow (que
# Windows bloquea con frecuencia, dando 'Acceso denegado', y que además
# obligaba a usar ShowWindow/SW_RESTORE — encogiendo ventanas maximizadas),
# usamos Alt+Tab simulado: es una acción de teclado genuina que Windows
# siempre permite, y no toca el tamaño ni la posición de ninguna ventana.
#
# Identificamos la ventana de destino por su PROCESO (chrome.exe,
# EXCEL.EXE) y no por la "clase" de ventana, porque varias apps modernas
# (VS Code, Slack, Discord, Teams, etc.) están hechas con Electron y
# comparten la misma clase de ventana que Chrome ("Chrome_WidgetWin_1").
# Identificar por proceso evita que el script confunda VS Code con Chrome.

PROCESO_EXCEL = "excel.exe"
PROCESO_CHROME = "chrome.exe"


def nombre_proceso_de_ventana(hwnd):
    """Devuelve el nombre del .exe dueño de una ventana (en minúsculas),
    o '' si no se puede determinar."""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        handle = win32api.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        ruta = win32process.GetModuleFileNameEx(handle, 0)
        win32api.CloseHandle(handle)
        return os.path.basename(ruta).lower()
    except Exception:
        return ""


def proceso_ventana_activa():
    hwnd = win32gui.GetForegroundWindow()
    return nombre_proceso_de_ventana(hwnd)


def alt_tab():
    pyautogui.keyDown("alt")
    time.sleep(0.1)
    pyautogui.press("tab")
    time.sleep(0.1)
    pyautogui.keyUp("alt")
    time.sleep(0.35)  # dar tiempo a Windows a completar la animación de cambio


def cambiar_a_proceso(nombre_proceso_esperado, max_intentos=8):
    """Presiona Alt+Tab hasta que la ventana activa pertenezca al proceso
    deseado (o se agoten los intentos). max_intentos es más alto que
    antes porque ahora puede haber más ventanas de por medio (ej. VS Code)."""
    for _ in range(max_intentos):
        if proceso_ventana_activa() == nombre_proceso_esperado:
            return True
        alt_tab()

    if proceso_ventana_activa() == nombre_proceso_esperado:
        return True

    print(f"⚠️  No se pudo llegar al proceso '{nombre_proceso_esperado}'. "
          f"Proceso activo actual: '{proceso_ventana_activa()}'")
    return False


def ir_a_siigo():
    cambiar_a_proceso(PROCESO_CHROME)


def ir_a_excel():
    cambiar_a_proceso(PROCESO_EXCEL)


# =========================================================
# FUNCIONES DE MOUSE / TECLADO
# =========================================================

def mover_y_click(coord):
    """Mueve el mouse a una coordenada y hace clic izquierdo."""
    x, y = coord
    pyautogui.moveTo(x, y, duration=0.3)
    pyautogui.click()
    time.sleep(PAUSA_CORTA)


def mover_mouse(coord):
    """Solo mueve el mouse, sin clic."""
    x, y = coord
    pyautogui.moveTo(x, y, duration=0.3)
    time.sleep(PAUSA_CORTA)


def triple_click(coord, pausa_entre_clics=0.3):
    """Mueve el mouse a una coordenada y hace 3 clics izquierdos,
    con una pausa entre cada uno."""
    x, y = coord
    pyautogui.moveTo(x, y, duration=0.3)
    for _ in range(3):
        pyautogui.click()
        time.sleep(pausa_entre_clics)


def pegar():
    pyautogui.hotkey("ctrl", "v")
    time.sleep(PAUSA_CORTA)


def enter():
    pyautogui.press("enter")
    time.sleep(PAUSA_CORTA)


def tab(n=1):
    for _ in range(n):
        pyautogui.press("tab")
        time.sleep(0.15)

def shift_tab(n=1):
    for _ in range(n):
        pyautogui.hotkey('shift', 'tab')
        time.sleep(0.15)


# =========================================================
# FUNCIONES DE EXCEL (lectura de celda + portapapeles)
# =========================================================

def formatear_valor(valor):
    """Convierte el valor crudo de una celda de Excel a texto, evitando
    los problemas típicos de win32com:
    - Números enteros que llegan como float (5 -> '5.0') se limpian a '5'.
    - Los que sí tienen decimales conservan solo los necesarios (sin ceros
      de sobra) y sin notación científica.
    - Las fechas (que llegan como objeto datetime) se formatean dd/mm/aaaa.
    """
    if valor is None:
        return ""

    if isinstance(valor, datetime.datetime):
        return valor.strftime("%d/%m/%Y")

    if isinstance(valor, float):
        if valor.is_integer():
            return str(int(valor))
        texto = f"{valor:.10f}".rstrip("0").rstrip(".")
        return texto

    return str(valor)


def poner_texto_en_portapapeles(texto):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(str(texto), win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()


def copiar_celda(hoja, fila, col):
    """Lee el valor de una celda de Excel, lo formatea correctamente y lo
    pone en el portapapeles (equivalente a 'copiar' esa celda)."""
    valor_crudo = hoja.Cells(fila, col).Value
    texto = formatear_valor(valor_crudo)
    poner_texto_en_portapapeles(texto)
    return texto


def copiar_desde_excel_y_pegar_en_siigo(hoja, fila, col):
    """Flujo repetido: ir a Excel, copiar celda, ir a Siigo, pegar.
    No mueve el mouse: se asume que el foco ya quedó en la casilla
    correcta gracias al 'tab' anterior."""
    ir_a_excel()
    copiar_celda(hoja, fila, col)
    ir_a_siigo()
    pegar()


# =========================================================
# CONEXIÓN A EXCEL (libro ya abierto)
# =========================================================

def obtener_hoja_excel():
    """Se conecta al Excel que YA está abierto (no abre uno nuevo)."""
    excel = win32com.client.GetObject(Class="Excel.Application")
    libro = excel.Workbooks(1)  # ajusta el índice/nombre si tienes varios libros abiertos
    hoja = libro.Sheets(NOMBRE_HOJA)
    return hoja


# =========================================================
# FLUJO PRINCIPAL - FASE 1 (encabezado de factura)
# =========================================================

def ingresar_encabezado_factura(hoja):
    # 1. Ir a Siigo y crear factura de compra
    ir_a_siigo()
    #mover_y_click(COORD["crear"])
    #mover_y_click(COORD["factura_compra"])

    print("Esperando 5 seg a que cargue la página...")
    time.sleep(5)

    # 2. Casilla tipo
    mover_y_click(COORD["casilla_tipo"])
    #pyautogui.press("down")
    time.sleep(0.2)
    enter()

    # 3. Fecha (requiere 3 clics, con pausa entre cada uno, antes de pegar)
    mover_y_click(COORD["casilla_fecha"])
    #copiar_desde_excel_y_pegar_en_siigo(hoja, FILA_ENCABEZADO, COL_FECHA)
    tab()

    # 4. Prefijo factura
    copiar_desde_excel_y_pegar_en_siigo(hoja, FILA_ENCABEZADO, COL_PREFIJO)
    tab()

    # 5. No. factura
    copiar_desde_excel_y_pegar_en_siigo(hoja, FILA_ENCABEZADO, COL_NO_FACTURA)
    tab()

    # 6. Proveedor
    copiar_desde_excel_y_pegar_en_siigo(hoja, FILA_ENCABEZADO, COL_PROVEEDOR)
    time.sleep(3)
    enter()
    tab(8)

    print("✅ Encabezado de factura ingresado. Listo para continuar con la tabla de ítems.")


# =========================================================
# FLUJO PRINCIPAL - FASE 2 (tabla de ítems)
# =========================================================

def celda_vacia(hoja, fila, col):
    """True si la celda no tiene ningún valor (o solo espacios)."""
    valor = hoja.Cells(fila, col).Value
    return valor is None or str(valor).strip() == ""


def manejar_iva(hoja, fila):
    """Lee el valor de IVA de Excel y decide si oprime flecha abajo antes
    de los 3 tabs finales. Se asume que Chrome/Siigo ya tiene el foco."""
    valor_iva = hoja.Cells(fila, COL_IVA).Value
    if valor_iva and str(valor_iva).strip().upper() == "SI":
        pyautogui.press("down")
        time.sleep(0.2)
    tab(3)


def ingresar_primer_item(hoja, fila):
    """Ingresa el primer renglón de la tabla de ítems (incluye bodega)."""
    # Código
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, COL_CODIGO)
    time.sleep(2)
    enter()
    tab(4)

    # Bodega
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, COL_BODEGA)
    time.sleep(1)
    enter()
    tab(2)

    # Cantidad
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, COL_CANTIDAD)
    tab(1)

    # Precio
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, COL_PRECIO)
    tab(1)

    # Descuento
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, COL_DESCUENTO)
    tab(1)

    # IVA
    manejar_iva(hoja, fila)


def ingresar_item_siguiente(hoja, fila):
    """Ingresa los renglones 2 en adelante (sin bodega)."""
    # Código
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, COL_CODIGO)
    time.sleep(1)
    enter()

    # Cantidad
    shift_tab(1)
    tab(1)
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, COL_CANTIDAD)  
    tab(1)

    # Precio
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, COL_PRECIO)
    tab(1)

    # Descuento
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, COL_DESCUENTO)
    tab(1)

    # IVA
    manejar_iva(hoja, fila)


def ingresar_items(hoja):
    """Itera fila por fila desde FILA_ITEMS_INICIO hasta encontrar
    la primera celda de 'código' vacía."""
    fila = FILA_ITEMS_INICIO
    primero = True

    while not celda_vacia(hoja, fila, COL_CODIGO):
        print(f"Ingresando ítem de la fila {fila}...")

        ingresar_item_siguiente(hoja, fila)

        # Al terminar el renglón, Enter deja el cursor en (columna 1, fila+1)
        ir_a_siigo()
        enter()

        fila += 1

    print(f"✅ Ítems ingresados hasta la fila {fila - 1}.")


if __name__ == "__main__":
    print("Tienes 5 segundos para asegurarte de que Excel y Chrome/Siigo estén abiertos...")
    time.sleep(5)

    hoja_excel = obtener_hoja_excel()
    ingresar_encabezado_factura(hoja_excel)
    ingresar_items(hoja_excel)
