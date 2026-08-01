# -*- coding: utf-8 -*-
"""
Semi-automatización de ingreso de facturas de compra en Siigo Nube.

Soporta 2 usuarios de Siigo, cada uno con su propio flujo de ingreso
(coordenadas y algoritmo de campos), y 2 modos de arranque:
    - "nueva": crea la factura desde cero (Crear -> Factura de compra)
    - "retomar": continúa un ingreso interrumpido (la factura ya está
      abierta en Siigo, no se debe volver a crear)

Al arrancar, el programa pregunta qué usuario y qué modo usar.

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
# DPI awareness: sin esto, en pantallas con escalado (125%, 150%, etc.)
# las coordenadas calibradas no coinciden con donde hace clic pyautogui.
# ---------------------------------------------------------
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_AWARE
except Exception:
    ctypes.windll.user32.SetProcessDPIAware()

NOMBRE_HOJA = 1  # o el nombre de la hoja, ej. "Facturas"
PAUSA_CORTA = 0.4
pyautogui.PAUSE = 0.15


# =========================================================
# FUNCIONES DE VENTANAS (comunes a ambos usuarios)
# =========================================================
# Identificamos la ventana de destino por su PROCESO (chrome.exe,
# EXCEL.EXE), no por título ni por clase, porque apps Electron (VS Code,
# Slack, etc.) comparten la clase de ventana con Chrome. Cambiamos de
# ventana con Alt+Tab simulado, que Windows siempre permite.

PROCESO_EXCEL = "excel.exe"
PROCESO_CHROME = "chrome.exe"


def nombre_proceso_de_ventana(hwnd):
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
    time.sleep(0.35)


def cambiar_a_proceso(nombre_proceso_esperado, max_intentos=8):
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
# FUNCIONES DE MOUSE / TECLADO (comunes)
# =========================================================

def mover_y_click(coord):
    x, y = coord
    pyautogui.moveTo(x, y, duration=0.3)
    pyautogui.click()
    time.sleep(PAUSA_CORTA)


def mover_mouse(coord):
    x, y = coord
    pyautogui.moveTo(x, y, duration=0.3)
    time.sleep(PAUSA_CORTA)


def triple_click(coord, pausa_entre_clics=0.3):
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
        pyautogui.hotkey("shift", "tab")
        time.sleep(0.15)


# =========================================================
# FUNCIONES DE EXCEL (comunes)
# =========================================================

def formatear_valor(valor):
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
    valor_crudo = hoja.Cells(fila, col).Value
    texto = formatear_valor(valor_crudo)
    poner_texto_en_portapapeles(texto)
    return texto


def copiar_desde_excel_y_pegar_en_siigo(hoja, fila, col):
    ir_a_excel()
    copiar_celda(hoja, fila, col)
    ir_a_siigo()
    pegar()


def obtener_hoja_excel():
    excel = win32com.client.GetObject(Class="Excel.Application")
    libro = excel.Workbooks(1)  # ajusta el índice/nombre si tienes varios libros abiertos
    hoja = libro.Sheets(NOMBRE_HOJA)
    return hoja


def celda_vacia(hoja, fila, col):
    valor = hoja.Cells(fila, col).Value
    return valor is None or str(valor).strip() == ""


def manejar_iva(hoja, fila, perfil):
    """Común a ambos usuarios: lee IVA de Excel, oprime flecha abajo si
    es 'SI', y hace los 3 tabs finales del renglón."""
    valor_iva = hoja.Cells(fila, perfil["col_iva"]).Value
    if valor_iva and str(valor_iva).strip().upper() == "SI":
        pyautogui.press("down")
        time.sleep(0.2)
    tab(3)


# =========================================================
# USUARIO 1 - Juan Pablo (SAIN)
# =========================================================
# Flujo: crear -> factura de compra -> casilla tipo (flecha abajo si es
# nueva) -> fecha (triple clic + pegar) -> prefijo -> no. factura ->
# proveedor -> ítems (el primer ítem sí incluye bodega, los siguientes no).

COORD_USUARIO_1 = {
    "crear": (1623, 143),
    "factura_compra": (1183, 257),
    "casilla_tipo": (595, 383),
    "casilla_fecha": (635, 418),
}


def encabezado_usuario1(hoja, perfil, modo):
    coord = perfil["coord"]
    ir_a_siigo()

    if modo == "nueva":
        mover_y_click(coord["crear"])
        mover_y_click(coord["factura_compra"])

    print("Esperando 5 seg a que cargue la página...")
    time.sleep(5)

    mover_y_click(coord["casilla_tipo"])
    if modo == "nueva":
        pyautogui.press("down")
    time.sleep(0.2)
    enter()

    if modo == "nueva":
        triple_click(coord["casilla_fecha"])
        copiar_desde_excel_y_pegar_en_siigo(hoja, perfil["fila_encabezado"], perfil["col_fecha"])
    else:
        # En modo retomar no se vuelve a pegar la fecha: se asume que ya
        # quedó correcta antes de la interrupción, solo se avanza el foco.
        mover_y_click(coord["casilla_fecha"])
    tab()

    copiar_desde_excel_y_pegar_en_siigo(hoja, perfil["fila_encabezado"], perfil["col_prefijo"])
    tab()

    copiar_desde_excel_y_pegar_en_siigo(hoja, perfil["fila_encabezado"], perfil["col_no_factura"])
    tab()

    copiar_desde_excel_y_pegar_en_siigo(hoja, perfil["fila_encabezado"], perfil["col_proveedor"])
    time.sleep(3)
    enter()
    if modo == "nueva":
        tab(6)
        enter()
        tab(2)
    else:
        tab(8)

    print("✅ Encabezado de factura ingresado (Usuario 1). Listo para los ítems.")


def item_primero_usuario1(hoja, fila, perfil, modo):
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_codigo"])
    time.sleep(2)
    enter()
    tab(4)

    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_bodega"])
    time.sleep(1)
    if modo == "nueva":
        tab(3)
    else:
        enter()
        tab(2)

    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_cantidad"])
    tab(1)

    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_precio"])
    tab(1)

    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_descuento"])
    tab(1)

    manejar_iva(hoja, fila, perfil)


def item_siguiente_usuario1(hoja, fila, perfil):
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_codigo"])
    time.sleep(1)
    enter()

    shift_tab(1)
    tab(1)
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_cantidad"])
    tab(1)

    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_precio"])
    tab(1)

    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_descuento"])
    tab(1)

    manejar_iva(hoja, fila, perfil)


def items_usuario1(hoja, perfil, modo):
    fila = perfil["fila_items_inicio"]
    primero = True

    while not celda_vacia(hoja, fila, perfil["col_codigo"]):
        print(f"Ingresando ítem de la fila {fila}...")

        if primero:
            item_primero_usuario1(hoja, fila, perfil, modo)
            primero = False
        else:
            item_siguiente_usuario1(hoja, fila, perfil)

        ir_a_siigo()
        enter()
        fila += 1

    print(f"✅ Ítems ingresados hasta la fila {fila - 1} (Usuario 1).")


# =========================================================
# USUARIO 2
# =========================================================
# Flujo: crear -> factura de compra -> casilla fecha (clic + tab, SIN
# pegar fecha desde Excel) -> prefijo -> no. factura -> proveedor ->
# ítems (nunca usa bodega; mismo patrón en todos los renglones).

COORD_USUARIO_2 = {
    "crear": (1623, 143),
    "factura_compra": (1166, 258),
    "casilla_fecha": (587, 386),
}


def encabezado_usuario2(hoja, perfil, modo):
    coord = perfil["coord"]
    ir_a_siigo()

    if modo == "nueva":
        mover_y_click(coord["crear"])
        mover_y_click(coord["factura_compra"])

    print("Esperando 5 seg a que cargue la página...")
    time.sleep(5)

    mover_y_click(coord["casilla_fecha"])
    tab()

    copiar_desde_excel_y_pegar_en_siigo(hoja, perfil["fila_encabezado"], perfil["col_prefijo"])
    tab()

    copiar_desde_excel_y_pegar_en_siigo(hoja, perfil["fila_encabezado"], perfil["col_no_factura"])
    tab()

    copiar_desde_excel_y_pegar_en_siigo(hoja, perfil["fila_encabezado"], perfil["col_proveedor"])
    time.sleep(1)
    enter()
    tab(6)
    enter()
    tab(2)

    print("✅ Encabezado de factura ingresado (Usuario 2). Listo para los ítems.")


def item_siguiente_usuario2(hoja, fila, perfil):
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_codigo"])
    time.sleep(1)
    enter()

    tab(4)
    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_cantidad"])
    tab(1)

    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_precio"])
    tab(1)

    copiar_desde_excel_y_pegar_en_siigo(hoja, fila, perfil["col_descuento"])
    tab(1)

    manejar_iva(hoja, fila, perfil)


def items_usuario2(hoja, perfil, modo=None):
    # modo no afecta el flujo de ítems de Usuario 2 (siempre es el mismo
    # patrón); se acepta el parámetro solo para tener la misma firma que
    # items_usuario1 y poder llamarlas de forma genérica desde main.
    fila = perfil["fila_items_inicio"]

    while not celda_vacia(hoja, fila, perfil["col_codigo"]):
        print(f"Ingresando ítem de la fila {fila}...")

        item_siguiente_usuario2(hoja, fila, perfil)

        ir_a_siigo()
        enter()
        fila += 1

    print(f"✅ Ítems ingresados hasta la fila {fila - 1} (Usuario 2).")


# =========================================================
# PERFILES DE USUARIO
# =========================================================
# La estructura de filas/columnas en Excel es igual para ambos usuarios;
# lo que cambia es el flujo de clics/tabs (fn_encabezado / fn_items).

PERFILES = {
    "1": {
        "nombre": "Juan Pablo (SAIN)",
        "coord": COORD_USUARIO_1,
        "fila_encabezado": 3,
        "col_fecha": 1,
        "col_prefijo": 2,
        "col_no_factura": 3,
        "col_proveedor": 4,
        "fila_items_inicio": 6,
        "col_codigo": 1,
        "col_bodega": 2,
        "col_cantidad": 3,
        "col_precio": 4,
        "col_descuento": 5,
        "col_iva": 6,
        "fn_encabezado": encabezado_usuario1,
        "fn_items": items_usuario1,
    },
    "2": {
        "nombre": "Usuario 2",
        "coord": COORD_USUARIO_2,
        "fila_encabezado": 3,
        "col_fecha": 1,
        "col_prefijo": 2,
        "col_no_factura": 3,
        "col_proveedor": 4,
        "fila_items_inicio": 6,
        "col_codigo": 1,
        "col_bodega": 2,
        "col_cantidad": 3,
        "col_precio": 4,
        "col_descuento": 5,
        "col_iva": 6,
        "fn_encabezado": encabezado_usuario2,
        "fn_items": items_usuario2,
    },
}


# =========================================================
# MENÚ DE INICIO
# =========================================================

def elegir_perfil():
    print("¿Qué usuario va a ingresar la factura?")
    for clave, datos in PERFILES.items():
        print(f"  [{clave}] {datos['nombre']}")
    while True:
        opcion = input("Escribe el número y presiona Enter: ").strip()
        if opcion in PERFILES:
            return PERFILES[opcion]
        print("Opción no válida, intenta de nuevo.")


def elegir_modo():
    print("\n¿Vas a iniciar una factura nueva o retomar un ingreso interrumpido?")
    print("  [1] Nueva factura (desde cero: Crear -> Factura de compra)")
    print("  [2] Retomar ingreso interrumpido (la factura ya está abierta en Siigo)")
    while True:
        opcion = input("Escribe 1 o 2 y presiona Enter: ").strip()
        if opcion == "1":
            return "nueva"
        if opcion == "2":
            return "retomar"
        print("Opción no válida, intenta de nuevo.")


if __name__ == "__main__":
    print("=== Ingreso automático de facturas de compra - Siigo ===\n")

    try:
        perfil = elegir_perfil()
        modo = elegir_modo()

        print(f"\nUsuario: {perfil['nombre']}  |  Modo: {modo}")
        print("Tienes 5 segundos para asegurarte de que Excel y Chrome/Siigo estén abiertos...")
        time.sleep(5)

        hoja_excel = obtener_hoja_excel()
        perfil["fn_encabezado"](hoja_excel, perfil, modo)
        perfil["fn_items"](hoja_excel, perfil, modo)

        print("\n✅ Proceso terminado sin errores.")
    except Exception:
        import traceback
        print("\n❌ Ocurrió un error y el proceso se detuvo:\n")
        traceback.print_exc()

    input("\nPresiona Enter para cerrar esta ventana...")
