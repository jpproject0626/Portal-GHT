# Portal GHT · Creado por Juan Pablo Salcedo Torres · Smurfit Westrock
"""
============================================================
 GENERADOR DE DATOS DIARIOS - Portal GHT (Grupo Chia)
============================================================
Este script lee el backlog diario, la programacion semanal y las notas
de despacho, aplica toda la logica de negocio (filtro GHT, Estatus OMP
Final, OW Final, cruce de despachos con cantidades) y genera un archivo
"datos.json" listo para subir al portal web.

COMO USARLO:
1. Cada dia, reemplaza (copia y pega encima) estos archivos DENTRO DE
   ESTA MISMA CARPETA (C:\\Portal GHT):
     - BACKLOG.xlsm
     - PROGRAMACION.xlsx
     - El archivo de Notas de Despacho (.xls) mas reciente - el nombre
       puede variar cada dia, el script agarra solo el mas nuevo.
2. Corre este script (o el .bat "actualizar_portal.bat", que ya lo
   hace por ti junto con la subida a GitHub/Vercel):
       python generar_datos.py

MODO AUTOMATICO (para correrlo sin intervencion, p. ej. con el Programador de
tareas de Windows en los cortes de 7:30, 12:00 y 16:00 hora Colombia):
       python generar_datos.py --auto
  Usa el .xls de despachos mas reciente de la carpeta. Si falta un archivo de
  entrada, la clave interna, o el .xls esta bloqueado (reintenta una vez), NO
  escribe nada y termina con un codigo de salida distinto de 0:
       0 ok | 1 error inesperado | 2 falta un archivo de entrada
       3 falta la clave interna (o la libreria cryptography)
       4 archivo de entrada bloqueado/incompleto | 5 config.json invalido
  Las rutas y los cortes se leen de config.json (copia config.ejemplo.json como
  config.json; config.json NO se sube al repo). Cada corrida deja:
  logs/corridas.log, salida_interna/resumen_corrida.json (para el correo de
  aviso, SOLO interno) y salida/GHT_consolidado_AAAA-MM-DD_HHMM.xlsx (el Excel
  para el cliente, con lo mismo que ve GHT en el portal).
  Solo las corridas con --auto:
    - actualizan salida_interna/estado_despachos_reportados.json (que despachos
      ya se reportaron como nuevos): una corrida manual NO consume despachos;
    - escriben SIEMPRE salida_interna/ultima_corrida.json (tambien si fallan),
      sin tocar datos.json, excluidos_avance.json ni resumen_corrida.json;
    - escriben SIEMPRE el correo del corte: salida_interna/correo_asunto.txt y
      salida_interna/correo_cuerpo.html (exito o error). ultima_corrida.json trae
      "adjunto": el Excel consolidado de esa corrida (null si fallo).
    - la corrida --auto que CIERRA el ciclo (pasa de abierto a cerrado) ademas deja
      salida_interna/excluidos_ciclo_AAAA-MM-DD.xlsx (con el mismo contenido que el
      boton de excluidos del portal), lo indica en "adjunto_excluidos" de
      ultima_corrida.json y lo anuncia en el correo. Ese Excel es INTERNO: queda en
      salida_interna/ (en .gitignore) y nunca se sube.

Requiere: pip install openpyxl xlrd cryptography

CLAVE INTERNA (para cifrar excluidos_avance.json): se lee del archivo local
"clave_interna.txt" (esta en .gitignore, NO se sube al repo) o de la variable
de entorno CLAVE_INTERNA_PORTAL. Es la misma clave que se usa para entrar al
portal como usuario interno de Smurfit.
============================================================
"""

import argparse
import base64
import hashlib
import html as html_lib
import json
import os
import struct
import sys
import time
import traceback
from collections import Counter
from datetime import datetime, timedelta, timezone
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

# ------------------------------------------------------------------
# 1. RUTAS DE LOS ARCHIVOS
# ------------------------------------------------------------------
# Los 3 archivos viven en la MISMA carpeta que este script (C:\Portal GHT).
# Cada dia se reemplazan aqui mismo, encima de los de ayer.
RUTA_BACKLOG = "BACKLOG.xlsm"
RUTA_PROGRAMACION = "PROGRAMACION.xlsx"
CARPETA_DESPACHOS = "."  # el script busca aqui el .xls MAS RECIENTE

ARCHIVO_SALIDA = "datos.json"
ARCHIVO_HISTORICO_DESPACHOS = "despachos_historico.json"
# Detalle de los pedidos que NO cuentan en el "Avance general" (uso interno:
# lo lee el boton de exportar que solo aparece con la clave interna). Se
# publica CIFRADO (AES-256-GCM, llave derivada con PBKDF2-SHA256): sin la
# clave interna el archivo no deja leer ningun dato.
ARCHIVO_EXCLUIDOS = "excluidos_avance.json"
ARCHIVO_CLAVE_INTERNA = "clave_interna.txt"  # local, en .gitignore
VARIABLE_CLAVE_INTERNA = "CLAVE_INTERNA_PORTAL"
PBKDF2_ITERACIONES = 600_000  # el navegador lee este valor del propio archivo

# Un pedido ya despachado se sigue mostrando en el portal durante este
# numero de dias despues de su fecha de despacho, AUNQUE el backlog ya
# lo haya quitado de su lista. Pasado ese tiempo, deja de aparecer.
DIAS_VISIBLE_DESPACHADO = 7


# ------------------------------------------------------------------
# 1b. CONFIGURACION (config.json), CODIGOS DE SALIDA Y MODO AUTOMATICO
# ------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_CONFIG = os.path.join(BASE_DIR, "config.json")
BOGOTA = timezone(timedelta(hours=-5), "America/Bogota")  # Colombia no tiene horario de verano

SALIDA_OK = 0
SALIDA_ERROR = 1            # error inesperado
SALIDA_FALTA_ARCHIVO = 2    # falta un archivo de entrada
SALIDA_FALTA_CLAVE = 3      # falta la clave interna o la libreria cryptography
SALIDA_BLOQUEADO = 4        # archivo de entrada bloqueado / incompleto (tras reintentar)
SALIDA_CONFIG = 5           # config.json invalido

# Todas las rutas son relativas a la carpeta de config.json (la del proyecto).
CONFIG_POR_DEFECTO = {
    "backlog": RUTA_BACKLOG,
    "programacion": RUTA_PROGRAMACION,
    "carpeta_despachos": CARPETA_DESPACHOS,   # .xls de despachos y BACKLOG CONSOLIDADO
    "datos_json": ARCHIVO_SALIDA,
    "historico_despachos": ARCHIVO_HISTORICO_DESPACHOS,
    "excluidos_avance": ARCHIVO_EXCLUIDOS,
    "clave_interna": ARCHIVO_CLAVE_INTERNA,
    "carpeta_salida": "salida",                  # Excel para el cliente
    "carpeta_salida_interna": "salida_interna",  # resumen para el correo (NO se sube)
    "carpeta_logs": "logs",
    "espera_reintento_segundos": 5,              # pausa antes del unico reintento
    "cortes": ["07:30", "12:00", "16:00"],       # hora Colombia
    "tolerancia_corte_minutos": 120,
}
CLAVES_DE_RUTA = ("backlog", "programacion", "carpeta_despachos", "datos_json", "historico_despachos",
                  "excluidos_avance", "clave_interna", "carpeta_salida", "carpeta_salida_interna", "carpeta_logs")


class ErrorCorrida(Exception):
    """Error esperado de una corrida: lleva el codigo de salida y un mensaje claro."""
    def __init__(self, codigo, mensaje, corto=None):
        super().__init__(mensaje)
        self.codigo = codigo
        self.mensaje = mensaje
        self.corto = corto or mensaje  # version sin rutas largas, para ultima_corrida.json


def cargar_config(ruta=None):
    """Lee config.json (si no existe usa CONFIG_POR_DEFECTO) y devuelve el
    diccionario con las rutas ya convertidas a absolutas."""
    ruta = os.path.abspath(ruta or ARCHIVO_CONFIG)
    cfg = dict(CONFIG_POR_DEFECTO)
    if os.path.exists(ruta):
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                leido = json.load(f)
        except (OSError, ValueError) as e:
            raise ErrorCorrida(SALIDA_CONFIG, f"No se pudo leer {ruta}: {e}", f"No se pudo leer config.json: {e}")
        if not isinstance(leido, dict):
            raise ErrorCorrida(SALIDA_CONFIG, f"{ruta} debe contener un objeto JSON.", "config.json debe contener un objeto JSON.")
        desconocidas = sorted(set(leido) - set(cfg))
        if desconocidas:
            raise ErrorCorrida(SALIDA_CONFIG, f"{ruta}: claves desconocidas: {', '.join(desconocidas)}",
                               f"config.json: claves desconocidas: {', '.join(desconocidas)}")
        cfg.update(leido)
    for clave in CLAVES_DE_RUTA:
        if not isinstance(cfg[clave], str) or not cfg[clave].strip():
            raise ErrorCorrida(SALIDA_CONFIG, f"config.json: '{clave}' debe ser una ruta (texto).")
    if not isinstance(cfg["espera_reintento_segundos"], (int, float)) or cfg["espera_reintento_segundos"] < 0:
        raise ErrorCorrida(SALIDA_CONFIG, "config.json: 'espera_reintento_segundos' debe ser un numero >= 0.")
    if not isinstance(cfg["tolerancia_corte_minutos"], (int, float)) or cfg["tolerancia_corte_minutos"] < 0:
        raise ErrorCorrida(SALIDA_CONFIG, "config.json: 'tolerancia_corte_minutos' debe ser un numero >= 0.")
    try:
        cfg["cortes"] = [datetime.strptime(c, "%H:%M").strftime("%H:%M") for c in cfg["cortes"]]
    except (TypeError, ValueError):
        raise ErrorCorrida(SALIDA_CONFIG, "config.json: 'cortes' debe ser una lista de horas 'HH:MM'.")
    return resolver_rutas(cfg, os.path.dirname(ruta))


def resolver_rutas(cfg, base):
    """Convierte las rutas de la configuracion (relativas a 'base') en absolutas."""
    cfg = dict(cfg)
    for clave in CLAVES_DE_RUTA:
        cfg[clave] = os.path.normpath(os.path.join(base, cfg[clave]))
    return cfg


def config_por_defecto():
    """Configuracion de respaldo (si config.json no sirve): sirve para dejar
    constancia del error en ultima_corrida.json y en los logs."""
    return resolver_rutas(CONFIG_POR_DEFECTO, BASE_DIR)


def aplicar_config(cfg):
    """Apunta las rutas globales del script a las de la configuracion."""
    global RUTA_BACKLOG, RUTA_PROGRAMACION, CARPETA_DESPACHOS, ARCHIVO_SALIDA
    global ARCHIVO_HISTORICO_DESPACHOS, ARCHIVO_EXCLUIDOS, ARCHIVO_CLAVE_INTERNA
    RUTA_BACKLOG = cfg["backlog"]
    RUTA_PROGRAMACION = cfg["programacion"]
    CARPETA_DESPACHOS = cfg["carpeta_despachos"]
    ARCHIVO_SALIDA = cfg["datos_json"]
    ARCHIVO_HISTORICO_DESPACHOS = cfg["historico_despachos"]
    ARCHIVO_EXCLUIDOS = cfg["excluidos_avance"]
    ARCHIVO_CLAVE_INTERNA = cfg["clave_interna"]


def verificar_lectura(ruta, espera):
    """Abre el archivo en lectura. Si esta bloqueado por otro programa espera y
    reintenta UNA vez; si sigue bloqueado, termina con SALIDA_BLOQUEADO."""
    for intento in (1, 2):
        try:
            with open(ruta, "rb") as f:
                f.read(1)
            return
        except FileNotFoundError:
            raise ErrorCorrida(SALIDA_FALTA_ARCHIVO, f"Falta el archivo de entrada: {ruta}",
                               f"Falta el archivo de entrada: {os.path.basename(ruta)}")
        except PermissionError:
            if intento == 1:
                print(f"  '{os.path.basename(ruta)}' esta bloqueado; reintento en {espera} s...")
                time.sleep(espera)
    raise ErrorCorrida(SALIDA_BLOQUEADO, f"El archivo esta bloqueado (lo tiene abierto otro programa): {ruta}",
                       f"Archivo bloqueado: {os.path.basename(ruta)}")


def verificar_entradas_auto(espera):
    """Comprobaciones previas del modo --auto, ANTES de leer o escribir nada:
    archivos de entrada, clave interna y archivos no bloqueados."""
    for ruta in (RUTA_BACKLOG, RUTA_PROGRAMACION):
        if not os.path.isfile(ruta):
            raise ErrorCorrida(SALIDA_FALTA_ARCHIVO, f"Falta el archivo de entrada: {ruta}",
                               f"Falta el archivo de entrada: {os.path.basename(ruta)}")
    try:
        ruta_xls = encontrar_archivo_mas_reciente(CARPETA_DESPACHOS, ".xls")
    except (FileNotFoundError, NotADirectoryError):
        raise ErrorCorrida(SALIDA_FALTA_ARCHIVO,
                           f"No hay ningun .xls de despachos en la carpeta: {CARPETA_DESPACHOS}",
                           "No hay ningun .xls de despachos")
    ruta_consolidado = encontrar_backlog_consolidado(CARPETA_DESPACHOS)
    if ruta_consolidado is None:
        raise ErrorCorrida(SALIDA_FALTA_ARCHIVO,
                           f"No hay ningun archivo 'consolidado' (.xlsm/.xlsx) en la carpeta: {CARPETA_DESPACHOS}",
                           "Falta el BACKLOG CONSOLIDADO")
    if leer_clave_interna() is None:
        raise ErrorCorrida(SALIDA_FALTA_CLAVE,
                           f"No hay clave interna ('{ARCHIVO_CLAVE_INTERNA}' o variable {VARIABLE_CLAVE_INTERNA}).",
                           "Falta la clave interna")
    try:
        import cryptography  # noqa: F401
    except ImportError:
        raise ErrorCorrida(SALIDA_FALTA_CLAVE, "Falta la libreria 'cryptography' (pip install cryptography).",
                           "Falta la libreria cryptography")
    for ruta in (RUTA_BACKLOG, RUTA_PROGRAMACION, ruta_xls, ruta_consolidado):
        verificar_lectura(ruta, espera)


def cargar_despachos_con_reintento(ruta, espera):
    """cargar_despachos, con UN reintento si el .xls esta bloqueado o todavia
    se esta escribiendo (incompleto)."""
    import xlrd
    motivo = ""
    for intento in (1, 2):
        try:
            return cargar_despachos(ruta)
        except PermissionError:
            motivo = "bloqueado (lo tiene abierto otro programa)"
        except (xlrd.XLRDError, IndexError, EOFError, struct.error) as e:
            # un .xls cortado (todavia se esta escribiendo) puede fallar de varias formas
            motivo = f"ilegible o incompleto ({type(e).__name__}: {e})"
        if intento == 1:
            print(f"  El .xls de despachos esta {motivo}; reintento en {espera} s...")
            time.sleep(espera)
    raise ErrorCorrida(SALIDA_BLOQUEADO, f"El archivo de despachos esta {motivo}: {ruta}",
                       "Archivo de despachos " + ("bloqueado" if motivo.startswith("bloqueado") else "incompleto o ilegible"))


def escribir_todo_o_nada(destinos):
    """destinos: lista de (ruta_final, funcion_escritora(ruta_temporal)). Escribe
    primero todo en archivos temporales y solo si TODO salio bien los mueve a su
    lugar; si algo falla no se sobrescribe ningun archivo."""
    pendientes = []
    try:
        for ruta, escribir in destinos:
            os.makedirs(os.path.dirname(ruta), exist_ok=True)
            tmp = ruta + ".tmp"
            escribir(tmp)
            pendientes.append((tmp, ruta))
        for tmp, ruta in pendientes:
            os.replace(tmp, ruta)
    except BaseException:
        for tmp, _ in pendientes:
            try:
                os.remove(tmp)
            except OSError:
                pass
        raise


def escribir_json(contenido, **opciones):
    def escritor(ruta_tmp):
        with open(ruta_tmp, "w", encoding="utf-8") as f:
            json.dump(contenido, f, **opciones)
    return escritor


def etiqueta_de_corte(ahora, cortes, tolerancia_min):
    """'07:30' / '12:00' / '16:00' segun el corte programado mas cercano a la
    hora de la corrida (dentro de la tolerancia); 'manual' si ninguno aplica."""
    minutos_ahora = ahora.hour * 60 + ahora.minute
    mejor = None
    for corte in cortes:
        h, m = map(int, corte.split(":"))
        distancia = abs(minutos_ahora - (h * 60 + m))
        if mejor is None or distancia < mejor[0]:
            mejor = (distancia, corte)
    return mejor[1] if mejor and mejor[0] <= tolerancia_min else "manual"


def registrar_log(cfg, ahora, modo, texto):
    """Agrega una linea a logs/corridas.log. Un fallo al escribir el log nunca
    debe tumbar la corrida."""
    try:
        os.makedirs(cfg["carpeta_logs"], exist_ok=True)
        with open(os.path.join(cfg["carpeta_logs"], "corridas.log"), "a", encoding="utf-8") as f:
            f.write(f"{ahora:%Y-%m-%d %H:%M:%S} (Colombia) | {modo} | {texto}\n")
    except OSError:
        pass

# El "porcentaje_global" (avance general) se calcula contra un archivo
# APARTE, fijo: el "BACKLOG_CONSOLIDADO" que manda logistica una vez por
# semana con la TOTALIDAD de los pedidos de esa semana. Ese archivo se
# queda quieto en la carpeta toda la semana (no se toca) y define el 100%
# real. El BACKLOG.xlsm de todos los dias sigue igual que siempre - solo
# se usa para el estado/seguimiento de cada pedido, ya no define el total.
# Los adicionales (pedidos que no estaban en el consolidado) SI se siguen
# viendo normal en el portal, pero NO se suman a este % - asi el numero
# nunca baja por una razon que no es culpa de nadie.
#
# Ademas, dentro del consolidado hay pedidos que NO cuentan en el % aunque
# esten en la base (siguen apareciendo normal en el portal con su estado
# real; ver calcular_avance_general):
#   - Regla 1: fecha_entrega POSTERIOR al cierre del ciclo (pedidos de
#     semanas futuras que nunca se esperaba tener listos todavia).
#   - Regla 2: pedidos dentro del ciclo con CERO unidades despachadas, solo
#     una vez que el ciclo ya cerro (ver ciclo_esta_cerrado: el ciclo cierra el
#     MISMO dia de la moda, a las 16:00 hora Colombia).
# El cierre del ciclo es la moda (fecha mas repetida) de fecha_entrega en
# el consolidado vigente; si cambia el consolidado, el ciclo se recalcula
# solo.


def encontrar_backlog_consolidado(carpeta="."):
    """Busca en la carpeta un archivo que tenga 'consolidado' en el nombre
    (mayusculas o minusculas, no importa) y termine en .xlsm o .xlsx - el
    archivo semanal fijo de logistica. Si no hay ninguno todavia, devuelve
    None (ese dia simplemente no se calcula el % de la semana)."""
    for nombre_archivo in os.listdir(carpeta):
        nombre_min = nombre_archivo.lower()
        if nombre_archivo.startswith("~$"):
            continue
        if "consolidado" in nombre_min and (nombre_min.endswith(".xlsm") or nombre_min.endswith(".xlsx")):
            return os.path.join(carpeta, nombre_archivo)
    return None


def cargar_base_semana_ght(ruta_consolidado):
    """Lee el BACKLOG_CONSOLIDADO (mismo formato que el backlog diario,
    hoja 'Formato') y devuelve un diccionario
    {OV|Elemento: {"cant_sol", "fecha_entrega" (date | None), "finca",
    "orden_compra", "ov", "elemento"}} SOLO con las filas de clientes GHT.
    Este es el total FIJO de la semana.
    "fecha_entrega" sale de la columna A del consolidado (la misma que se
    muestra como "Entrega estimada" en el portal)."""
    wb = load_workbook(ruta_consolidado, read_only=True, data_only=True)
    ws = wb["Formato"]
    filas = list(ws.iter_rows(values_only=True))
    encabezado = [limpiar_texto(c) for c in filas[0]]
    datos = filas[1:]

    idx_fecha_entrega = encabezado.index("Fecha entrega")
    idx_id_cliente = encabezado.index("ID Cliente")
    idx_elemento = encabezado.index("Elemento")
    idx_ord_compra = encabezado.index("Ord. de Compra")
    idx_ord_venta = encabezado.index("Ord. de Venta")
    idx_cant_sol = encabezado.index("Cant Sol")

    base = {}
    for fila in datos:
        id_cliente = limpiar_texto(fila[idx_id_cliente])
        if not id_cliente.startswith("GHT"):
            continue
        elemento = limpiar_texto(fila[idx_elemento])
        orden_compra = limpiar_texto(fila[idx_ord_compra])
        ord_venta = limpiar_texto(fila[idx_ord_venta])
        try:
            cant_sol = float(fila[idx_cant_sol]) if fila[idx_cant_sol] not in (None, "") else 0.0
        except (ValueError, TypeError):
            cant_sol = 0.0
        valor_fecha = fila[idx_fecha_entrega]
        fecha_entrega = valor_fecha.date() if isinstance(valor_fecha, datetime) else None
        clave = f"{ord_venta}|{elemento}"
        # Por si el consolidado trae mas de una fila para la misma OV+Elemento:
        # se suman las cantidades y se conserva la fecha de entrega mas temprana.
        previa = base.get(clave)
        if previa is None:
            base[clave] = {
                "cant_sol": cant_sol, "fecha_entrega": fecha_entrega,
                "finca": id_cliente, "orden_compra": orden_compra, "ov": ord_venta,
                "elemento": elemento,
            }
        else:
            previa["cant_sol"] += cant_sol
            fechas = [f for f in (previa["fecha_entrega"], fecha_entrega) if f]
            previa["fecha_entrega"] = min(fechas) if fechas else None
    return base


def calcular_fecha_cierre_ciclo(base_semana):
    """Cierre del ciclo = moda (fecha mas repetida) de fecha_entrega entre
    los pedidos del consolidado. Si dos fechas empatan, gana la mas tardia.
    Devuelve None si el consolidado no trae ninguna fecha de entrega."""
    conteo = Counter(d["fecha_entrega"] for d in base_semana.values() if d["fecha_entrega"])
    if not conteo:
        return None
    return max(conteo.items(), key=lambda item: (item[1], item[0]))[0]


def leer_clave_interna():
    """Devuelve la clave interna (variable de entorno o archivo local) o
    None si no esta configurada. NUNCA debe quedar escrita en el codigo ni
    en ningun archivo que se suba al repo."""
    clave = os.environ.get(VARIABLE_CLAVE_INTERNA, "").strip()
    if clave:
        return clave
    try:
        with open(ARCHIVO_CLAVE_INTERNA, "r", encoding="utf-8") as f:
            return f.readline().strip() or None
    except FileNotFoundError:
        return None


def cifrar_contenido(contenido, clave):
    """Cifra un diccionario con AES-256-GCM. La llave sale de la clave con
    PBKDF2-HMAC-SHA256 y un salt aleatorio nuevo en cada corrida; el IV
    (nonce) tambien es aleatorio. Devuelve el "sobre" que se publica: solo
    parametros del cifrado + texto cifrado en base64, sin ningun dato
    legible. Compatible con Web Crypto (el texto cifrado lleva la etiqueta
    de autenticacion GCM al final)."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM

    salt = os.urandom(16)
    iv = os.urandom(12)
    llave = hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), salt, PBKDF2_ITERACIONES, dklen=32)
    claro = json.dumps(contenido, ensure_ascii=False).encode("utf-8")
    cifrado = AESGCM(llave).encrypt(iv, claro, None)
    b64 = lambda datos: base64.b64encode(datos).decode("ascii")
    return {
        "v": 1,
        "alg": "AES-256-GCM",
        "kdf": "PBKDF2-SHA256",
        "iter": PBKDF2_ITERACIONES,
        "salt": b64(salt),
        "iv": b64(iv),
        "ct": b64(cifrado),
    }


RAZON_FECHA_FUTURA = "Fecha de entrega futura"
RAZON_SIN_DESPACHO = "Sin despacho al cierre del ciclo"


def armar_detalle_excluidos(base_semana, acumulado_por_clave, avance):
    """Arma el contenido de excluidos_avance.json: el detalle de cada pedido
    que quedo fuera del % (finca, elemento, OC, cantidades, fecha de entrega
    y razon), mas un resumen del calculo. La cantidad despachada es la REAL
    acumulada (sin el tope de lo solicitado que usa el calculo del %)."""
    def fila(clave, razon):
        d = base_semana[clave]
        info = acumulado_por_clave.get(clave)
        return {
            "finca": d["finca"], "ov": d["ov"], "elemento": d["elemento"],
            "orden_compra": d["orden_compra"],
            "cantidad_solicitada": d["cant_sol"],
            "cantidad_despachada": info["cantidad"] if info else 0.0,
            "fecha_entrega": d["fecha_entrega"].strftime("%d/%m/%Y") if d["fecha_entrega"] else "",
            "razon": razon,
        }

    def orden(clave):
        d = base_semana[clave]
        return (d["fecha_entrega"].isoformat() if d["fecha_entrega"] else "", d["finca"], d["ov"])

    # Primero los de fecha futura, luego los sin despacho; dentro de cada
    # grupo, por fecha de entrega, finca y OV.
    excluidos = [fila(c, RAZON_FECHA_FUTURA) for c in sorted(avance["futuros"], key=orden)]
    excluidos += [fila(c, RAZON_SIN_DESPACHO) for c in sorted(avance["sin_despacho"], key=orden)]
    cierre = avance["cierre_ciclo"]
    return {
        "cierre_ciclo": cierre.strftime("%d/%m/%Y") if cierre else "",
        "ciclo_cerrado": avance["ciclo_cerrado"],
        "porcentaje_global": avance["porcentaje"],
        "total_solicitado_en_porcentaje": avance["total_sol"],
        "total_despachado_en_porcentaje": avance["total_desp"],
        "pedidos_en_porcentaje": len(avance["incluidos"]),
        "excluidos": excluidos,
    }


HORA_CIERRE_CICLO = (16, 0)  # hora Colombia: desde aqui el DIA de la moda ya cuenta como ciclo cerrado


def ciclo_esta_cerrado(cierre, hoy, hora=None):
    """El ciclo cierra el MISMO dia de la moda (cierre). Esta cerrado si:
      - la fecha de la corrida es POSTERIOR a la moda, o
      - la fecha de la corrida es la moda Y la hora de la corrida es >= 16:00.
    Antes de las 16:00 del dia de la moda no se excluye nada (los despachos
    todavia pueden entrar). Una vez cerrado, las corridas siguientes lo mantienen
    mientras el consolidado (y por tanto la moda) sea el mismo. Sin 'hora' solo
    cuenta lo de "posterior a la moda" (el comportamiento anterior)."""
    if hoy > cierre:
        return True
    return hora is not None and hoy == cierre and (hora.hour, hora.minute) >= HORA_CIERRE_CICLO


def calcular_avance_general(base_semana, acumulado_por_clave, hoy, hora=None):
    """Calcula el "Avance general" sobre la base fija del consolidado,
    sacando de la cuenta (sin quitarlos del portal) dos grupos de pedidos:
      - Regla 1 "futuros": fecha_entrega POSTERIOR al cierre del ciclo.
      - Regla 2 "sin despacho": fecha_entrega dentro del ciclo (o sin fecha),
        CERO unidades despachadas, y el ciclo ya cerro (ver ciclo_esta_cerrado:
        el dia de la moda desde las 16:00, o cualquier dia posterior). Si el
        ciclo sigue abierto, estos pedidos cuentan normal.
    Devuelve un diccionario con el %, los totales y las llaves excluidas."""
    cierre = calcular_fecha_cierre_ciclo(base_semana)
    ciclo_cerrado = cierre is not None and ciclo_esta_cerrado(cierre, hoy, hora)

    futuros, sin_despacho, incluidos = [], [], []
    for clave, datos in base_semana.items():
        fecha = datos["fecha_entrega"]
        if cierre and fecha and fecha > cierre:
            futuros.append(clave)
            continue
        info_desp = acumulado_por_clave.get(clave)
        cantidad_despachada = info_desp["cantidad"] if info_desp else 0.0
        if ciclo_cerrado and cantidad_despachada <= 0:
            sin_despacho.append(clave)
            continue
        incluidos.append(clave)

    total_sol = 0.0
    total_desp = 0.0
    for clave in incluidos:
        cant_sol = base_semana[clave]["cant_sol"]
        info_desp = acumulado_por_clave.get(clave)
        total_sol += cant_sol
        total_desp += min(info_desp["cantidad"], cant_sol) if info_desp else 0.0

    porcentaje = None
    if total_sol > 0:
        porcentaje = round(min(100.0, (total_desp / total_sol) * 100))

    return {
        "porcentaje": porcentaje,
        "total_sol": total_sol,
        "total_desp": total_desp,
        "cierre_ciclo": cierre,
        "ciclo_cerrado": ciclo_cerrado,
        "incluidos": incluidos,
        "futuros": futuros,
        "sin_despacho": sin_despacho,
    }


def limpiar_texto(valor):
    """Convierte a texto y quita espacios, maneja valores vacios/None."""
    if valor is None:
        return ""
    return str(valor).strip()


def clasificar_estado_ght(estatus_omp_final):
    """
    Replica exactamente la formula DAX 'Estado GHT' que ya usa Power BI.
    Traduce el texto crudo de OMP a uno de los 3 estados simples de GHT.
    """
    texto = estatus_omp_final.lower()

    if "programaci" in texto:
        return "Pedido recibido"
    if "plano mestre" in texto:
        return "Pedido recibido"
    if "totalmente" in texto:
        return "Producido"
    if "parcial" in texto:
        return "En producción"
    if texto == "terminado":
        return "Producido"
    if "product" in texto:
        return "En producción"

    return "Sin clasificar"


def cargar_order_capacity(ruta_backlog):
    """Lee la hoja 'Order Capacity' del backlog y arma un mapa ORDER ID -> PRODUCTIONSTATUS."""
    wb = load_workbook(ruta_backlog, read_only=True, data_only=True)
    ws = wb["Order Capacity"]
    filas = list(ws.iter_rows(values_only=True))
    encabezado = filas[0]
    datos = filas[1:]

    idx_order_id = encabezado.index("ORDER ID")
    idx_status = encabezado.index("PRODUCTIONSTATUS")

    mapa = {}
    for fila in datos:
        order_id = limpiar_texto(fila[idx_order_id])
        if order_id:
            mapa[order_id] = limpiar_texto(fila[idx_status])
    return mapa


def cargar_programacion(ruta_programacion):
    """Lee la hoja ORDENTRABAJO y arma un mapa TARJETA (Elemento) -> OW."""
    wb = load_workbook(ruta_programacion, read_only=True, data_only=True)
    ws = wb["ORDENTRABAJO"]
    filas = list(ws.iter_rows(values_only=True))
    encabezado = [limpiar_texto(c) for c in filas[0]]
    datos = filas[1:]

    idx_tarjeta = encabezado.index("TARJETA")
    idx_ow = encabezado.index("OW")

    mapa = {}
    for fila in datos:
        tarjeta = limpiar_texto(fila[idx_tarjeta])
        if tarjeta:
            mapa[tarjeta] = limpiar_texto(fila[idx_ow])
    return mapa


def encontrar_archivo_mas_reciente(carpeta, extension=".xls"):
    """
    Busca dentro de 'carpeta' todos los archivos que terminen en 'extension'
    y devuelve la ruta completa del que se modifico mas recientemente.
    Asi no importa que el nombre del archivo cambie cada dia (ej. con la
    fecha incluida en el nombre) - siempre agarra el ultimo que se subio.
    """
    import os
    candidatos = [
        os.path.join(carpeta, f) for f in os.listdir(carpeta)
        if f.lower().endswith(extension.lower()) and not f.startswith("~$")
    ]
    if not candidatos:
        raise FileNotFoundError(
            f"No se encontro ningun archivo '{extension}' dentro de la carpeta: {carpeta}"
        )
    mas_reciente = max(candidatos, key=os.path.getmtime)
    return mas_reciente


def cargar_despachos(ruta_despachos):
    """
    Lee el archivo de Notas de Despacho (.xls) y devuelve una lista con
    cada LINEA de despacho individual (una por cada envio real de un
    producto puntual, incluyendo envios parciales). Cada linea trae:
      - clave: identificador UNICO de esa linea de despacho, formado por
        'Nota de despacho' + Orden de Venta + Elemento. Se necesita el
        Elemento en la llave (no solo nota+OV) porque una misma Orden de
        Venta a veces se repite en mas de un Elemento en el backlog -
        sin esto, esos productos distintos compartirian por error la
        misma cantidad/estado de despacho.
      - ov: 'Ord. de Venta' (Order Number + '-' + Order Line)
      - elemento: el Elemento especifico que se envio en esta linea
      - cantidad: cantidad enviada en ESA linea puntual
      - fecha: fecha de ese envio
    """
    import xlrd
    wb = xlrd.open_workbook(ruta_despachos)
    ws = wb.sheet_by_name("Sheet1")
    encabezado = [limpiar_texto(ws.cell_value(0, c)) for c in range(ws.ncols)]

    idx_order_number = encabezado.index("Order Number")
    idx_order_line = encabezado.index("Order Line")
    idx_fecha = encabezado.index("Fecha del Despacho")
    idx_nota = encabezado.index("Nota de despacho")
    idx_elemento = encabezado.index("Elemento")
    idx_cantidad = encabezado.index("Cantidad Enviada")
    idx_cancelada = encabezado.index("Cancelada") if "Cancelada" in encabezado else None
    # "Devolución" (1 = la linea es una devolucion, con cantidad negativa). Solo
    # sirve para etiquetar el tipo en el Excel del cliente; no cambia ningun calculo.
    idx_devolucion = encabezado.index("Devolución") if "Devolución" in encabezado else None

    lineas = []
    for r in range(1, ws.nrows):
        # Si el despacho fue cancelado, no cuenta como cantidad realmente enviada.
        if idx_cancelada is not None:
            valor_cancelada = ws.cell_value(r, idx_cancelada)
            if valor_cancelada not in (0, "", None):
                continue

        try:
            order_number = int(ws.cell_value(r, idx_order_number))
            order_line = int(ws.cell_value(r, idx_order_line))
        except (ValueError, TypeError):
            continue
        ov = f"{order_number}-{order_line}"

        nota = limpiar_texto(ws.cell_value(r, idx_nota))
        if not nota:
            continue  # sin numero de nota no podemos evitar contarlo doble; se descarta

        elemento = limpiar_texto(ws.cell_value(r, idx_elemento))

        try:
            cantidad = float(ws.cell_value(r, idx_cantidad))
        except (ValueError, TypeError):
            cantidad = 0.0

        valor_fecha = ws.cell_value(r, idx_fecha)
        try:
            fecha = xlrd.xldate.xldate_as_datetime(valor_fecha, wb.datemode).strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            fecha = ""

        # La llave unica es NOTA + OV + ELEMENTO: una misma nota de
        # despacho puede incluir varias lineas/productos distintos en un
        # mismo envio (una guia con varios items), y una misma OV a
        # veces se repite en mas de un Elemento.
        clave_linea = f"{nota}|{ov}|{elemento}"
        es_devolucion = False
        if idx_devolucion is not None:
            es_devolucion = limpiar_texto(ws.cell_value(r, idx_devolucion)) not in ("", "0", "0.0")
        lineas.append({
            "clave": clave_linea, "ov": ov, "elemento": elemento,
            "nota": nota, "cantidad": cantidad, "fecha": fecha, "devolucion": es_devolucion,
        })

    return lineas


def formatear_cantidad(valor):
    """Muestra 20000 en vez de 20000.0, pero conserva decimales si los hay de verdad."""
    if valor == int(valor):
        return str(int(valor))
    return str(round(valor, 2))


def formatear_fecha_entrega(valor):
    """Convierte la celda 'Fecha entrega' (columna A del backlog) a texto
    dd/mm/YYYY, igual que el resto de fechas del portal. Si la celda viene
    vacia (sin fecha asignada todavia), devuelve "" - el frontend decide
    como mostrar ese caso."""
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    return ""


# Columnas del Excel para el cliente: EXACTAMENTE las del boton "Exportar
# consolidado a Excel" del portal (index.html, hoja 'Consolidado'). Si se cambian
# alla, hay que cambiarlas aqui. No incluye nada interno (ni OV, ni %, ni fechas
# de corte).
COLUMNAS_CONSOLIDADO = [
    ("Finca", "finca"), ("Descripción", "descripcion"), ("Orden de compra", "orden_compra"),
    ("Elemento", "elemento"), ("Orden de trabajo", "ow"), ("Estado", "estado"),
    ("Cantidad solicitada", "cantidad_solicitada"), ("Cantidad despachada", "cantidad_despachada"),
    ("Fecha del último envío", "fecha_despacho"), ("Notas de despacho", "nota_despacho"),
]
COLUMNAS_DESPACHOS_NUEVOS = [
    ("Finca", "finca"), ("Orden de compra", "orden_compra"), ("Elemento", "elemento"),
    ("Descripción", "descripcion"), ("Nota de despacho", "nota"),
    ("Cantidad enviada", "cantidad"), ("Fecha del envío", "fecha"),
    ("Tipo", "tipo"),   # "Despacho" o "Devolución" (campo Devolución del .xls); sin la razon
]
ARCHIVO_ESTADO_DESPACHOS = "estado_despachos_reportados.json"
ARCHIVO_RESUMEN = "resumen_corrida.json"
ARCHIVO_ULTIMA_CORRIDA = "ultima_corrida.json"


def despachos_hasta_del_xls_mas_reciente(cfg):
    """'dd/mm/aaaa hh:mm' del .xls de despachos mas reciente, o None si no hay
    ninguno legible (se usa para dejar constancia cuando una corrida falla)."""
    try:
        ruta = encontrar_archivo_mas_reciente(cfg["carpeta_despachos"], ".xls")
        return datetime.fromtimestamp(os.path.getmtime(ruta), BOGOTA).strftime("%d/%m/%Y %H:%M")
    except (OSError, ValueError):
        return None


def mensaje_sin_rutas(texto, cfg, maximo=140):
    """Quita de un mensaje de error las rutas de las carpetas del proyecto y lo
    recorta: asi el mensaje de ultima_corrida.json es corto y legible."""
    carpetas = {BASE_DIR} | {cfg[k] for k in CLAVES_DE_RUTA if k.startswith("carpeta_")}
    for carpeta in carpetas:
        for variante in (carpeta + os.sep, carpeta.replace("\\", "\\\\") + "\\\\"):
            texto = texto.replace(variante, "")
    return texto[:maximo]


def escribir_ultima_corrida(cfg, ahora, codigo, estado, mensaje_error, despachos_hasta, adjunto=None,
                            adjunto_excluidos=None):
    """salida_interna/ultima_corrida.json: SOLO este archivo, nunca datos.json,
    excluidos_avance.json ni resumen_corrida.json. Si no se puede escribir se
    deja constancia en el log, pero no se cambia el resultado de la corrida."""
    contenido = {
        "corte": etiqueta_de_corte(ahora, cfg["cortes"], cfg["tolerancia_corte_minutos"]),
        "hora": ahora.isoformat(timespec="seconds"),
        "codigo_salida": codigo,
        "estado": estado,                      # ok / sin_despachos_nuevos / error
        "mensaje_error": mensaje_error,        # None si no hubo error
        "despachos_hasta": despachos_hasta,    # en un error: el del .xls mas reciente (o None)
        "adjunto": adjunto,                    # Excel consolidado de esta corrida (ruta relativa); None si fallo
        # Excel de excluidos: solo en la corrida --auto que cierra el ciclo y con al menos 1 excluido; si no, None
        "adjunto_excluidos": adjunto_excluidos,
    }
    try:
        escribir_todo_o_nada([(os.path.join(cfg["carpeta_salida_interna"], ARCHIVO_ULTIMA_CORRIDA),
                               escribir_json(contenido, ensure_ascii=False, indent=2))])
    except Exception as e:  # noqa: BLE001
        registrar_log(cfg, ahora, "auto", f"No se pudo escribir {ARCHIVO_ULTIMA_CORRIDA}: {type(e).__name__}: {e}")


ARCHIVO_CORREO_ASUNTO = "correo_asunto.txt"
ARCHIVO_CORREO_CUERPO = "correo_cuerpo.html"
CREADO_POR = "Juan Pablo Salcedo Torres"   # segunda linea del pie de todos los correos

# Motivo en lenguaje simple y que revisar, segun el codigo de salida (sin rutas completas).
MOTIVOS_ERROR = {
    SALIDA_FALTA_ARCHIVO: (
        "Falta un archivo de entrada que el proceso necesita.",
        ["Que BACKLOG.xlsm, PROGRAMACION.xlsx, el BACKLOG CONSOLIDADO y el .xls de despachos estén en la carpeta del portal.",
         "Que ningún archivo haya cambiado de nombre o se haya movido."]),
    SALIDA_FALTA_CLAVE: (
        "No se encontró la clave interna del portal.",
        ["Que exista el archivo clave_interna.txt en la carpeta del portal (o la variable CLAVE_INTERNA_PORTAL).",
         "Que esté instalada la librería 'cryptography' en el equipo que corre el proceso."]),
    SALIDA_BLOQUEADO: (
        "El archivo de despachos (.xls) —u otro archivo de entrada— está abierto en otro programa o está incompleto.",
        ["Cerrar el archivo si está abierto en Excel.",
         "Esperar a que termine de guardarse o descargarse y volver a ejecutar el proceso."]),
    SALIDA_CONFIG: (
        "La configuración del proceso (config.json) tiene un error.",
        ["Comparar config.json con config.ejemplo.json y corregir lo que difiera."]),
    SALIDA_ERROR: (
        "Ocurrió un error inesperado.",
        ["Revisar el detalle en logs/corridas.log.",
         "Avisar a quien mantiene el proceso si el problema continúa."]),
}


def hora_corta_correo(despachos_hasta_txt, hoy):
    """'hh:mm' si el corte de los despachos es de hoy; 'dd/mm hh:mm' si es de otro dia."""
    try:
        dt = datetime.strptime(despachos_hasta_txt, "%d/%m/%Y %H:%M")
    except (TypeError, ValueError):
        return despachos_hasta_txt or "—"
    return dt.strftime("%H:%M") if dt.date() == hoy else dt.strftime("%d/%m %H:%M")


def corte_para_correo(ahora, corte):
    """Etiqueta del corte (07:30 / 12:00 / 16:00); si la corrida fue fuera de los cortes, la hora real."""
    return ahora.strftime("%H:%M") if corte == "manual" else corte


def _marco_correo(titulo, color_titulo, contenido):
    """HTML simple y sobrio (estilos en linea, para que lo respete cualquier cliente de correo)."""
    return (
        '<!DOCTYPE html>\n<html lang="es"><head><meta charset="utf-8"><title>' + html_lib.escape(titulo) + '</title></head>\n'
        '<body style="margin:0;padding:0;background:#f4f6f8;">\n'
        '<div style="max-width:640px;margin:0 auto;padding:24px 22px;background:#ffffff;font-family:Arial,Helvetica,sans-serif;'
        'font-size:14px;line-height:1.5;color:#1f2933;">\n'
        '<h2 style="margin:0 0 14px;font-size:18px;color:' + color_titulo + ';">' + html_lib.escape(titulo) + '</h2>\n'
        + contenido +
        '<p style="margin:22px 0 0;font-size:12px;color:#7b8794;">Mensaje automático del portal GHT.</p>\n'
        '<p style="margin:8px 0 0;font-size:12px;color:#7b8794;">Autor: ' + html_lib.escape(CREADO_POR) + '</p>\n'
        '</div>\n</body></html>\n'
    )


def _filas_datos(pares):
    celdas = "".join(
        '<tr><td style="padding:3px 14px 3px 0;color:#52606d;">' + html_lib.escape(k) + '</td>'
        '<td style="padding:3px 0;font-weight:bold;">' + html_lib.escape(v) + '</td></tr>' for k, v in pares)
    return '<table style="border-collapse:collapse;margin:0 0 14px;">' + celdas + '</table>\n'


def construir_correo_ok(ahora, resumen, nuevos, omitidos, mismo_archivo, cierre_ciclo_excluidos=None):
    """(asunto, html) del correo de un corte exitoso. Solo datos de GHT: nada interno
    (sin OV, sin razones de devolucion, sin datos de otros clientes)."""
    corte = corte_para_correo(ahora, resumen["corte"])
    hasta = hora_corta_correo(resumen["despachos_hasta"], ahora.date())
    asunto = f"Portal GHT actualizado – corte {corte} – despachos hasta {hasta}"
    if cierre_ciclo_excluidos is not None:   # corte que cierra el ciclo
        asunto = "CIERRE DE CICLO – " + asunto
    avance = f"{resumen['avance_general']} %" if resumen["avance_general"] is not None else "No disponible"
    partes = [_filas_datos([("Corte", f"{corte} ({ahora:%d/%m/%Y})"), ("Despachos hasta", resumen["despachos_hasta"] or "—"),
                            ("Avance general", avance)])]
    if cierre_ciclo_excluidos is not None:   # solo en el corte que cierra el ciclo
        n = cierre_ciclo_excluidos
        if n > 0:
            plural = "s" if n != 1 else ""
            texto = (f"Cierre de ciclo: {n} pedido{plural} excluido{plural} del avance. "
                     "El Excel de excluidos va adjunto.")
        else:
            texto = "Cierre de ciclo: 0 pedidos excluidos del avance."
        partes.append('<p style="margin:0 0 14px;padding:8px 12px;background:#eef3fb;border-left:3px solid #00205B;">'
                      + html_lib.escape(texto) + '</p>\n')
    if mismo_archivo:
        partes.append('<p style="margin:0 0 14px;padding:8px 12px;background:#fff8e6;border-left:3px solid #e0a100;">'
                      'El archivo de despachos no se ha actualizado desde el corte anterior</p>\n')
    if nuevos:
        partes.append('<p style="margin:0 0 6px;"><b>Despachos nuevos de GHT desde el corte anterior: ' + str(len(nuevos)) + '</b></p>\n')
        encabezado = "".join(
            '<th style="text-align:' + alin + ';padding:6px 8px;border-bottom:2px solid #00205B;color:#00205B;">' + t + '</th>'
            for t, alin in (("Finca", "left"), ("Orden de compra", "left"), ("Elemento", "left"), ("Nota", "left"),
                            ("Cantidad", "right"), ("Tipo", "left")))
        filas = ""
        for d in nuevos:
            celdas = [(d["finca"], "left"), (d["orden_compra"], "left"), (d["elemento"], "left"), (d["nota"], "left"),
                      (f"{d['cantidad']:,}".replace(",", " ") if isinstance(d["cantidad"], int) else str(d["cantidad"]), "right"),
                      (d["tipo"], "left")]
            filas += "<tr>" + "".join(
                '<td style="text-align:' + alin + ';padding:5px 8px;border-bottom:1px solid #e4e7eb;">' + html_lib.escape(str(v)) + '</td>'
                for v, alin in celdas) + "</tr>"
        partes.append('<table style="border-collapse:collapse;width:100%;font-size:13px;"><tr>' + encabezado + '</tr>' + filas + '</table>\n')
    else:
        partes.append('<p style="margin:0 0 6px;">Sin nuevos despachos desde el corte anterior</p>\n')
    nombre_excel = os.path.basename(resumen["excel_cliente"])
    partes.append('<p style="margin:16px 0 0;font-size:13px;color:#52606d;">Se adjunta el Excel consolidado: ' + html_lib.escape(nombre_excel) + '</p>\n')
    if omitidos:
        partes.append('<p style="margin:4px 0 0;font-size:12px;color:#7b8794;">Líneas de despachos de otros clientes omitidas: ' + str(omitidos) + '</p>\n')
    return asunto, _marco_correo("Portal GHT actualizado", "#00205B", "".join(partes))


def construir_correo_error(ahora, corte, codigo, corto):
    """(asunto, html) del correo cuando la corrida fallo. Motivo en lenguaje simple, sin rutas completas."""
    corte = corte_para_correo(ahora, corte)
    motivo, revisar = MOTIVOS_ERROR.get(codigo, MOTIVOS_ERROR[SALIDA_ERROR])
    asunto = f"ATENCIÓN: el portal GHT NO se actualizó – corte {corte}"
    partes = [_filas_datos([("Corte", f"{corte} ({ahora:%d/%m/%Y %H:%M})")]),
              '<p style="margin:0 0 6px;"><b>Motivo</b></p>\n<p style="margin:0 0 14px;">' + html_lib.escape(motivo) + '</p>\n',
              '<p style="margin:0 0 6px;"><b>Qué revisar</b></p>\n<ul style="margin:0 0 14px;padding-left:20px;">'
              + "".join("<li>" + html_lib.escape(r) + "</li>" for r in revisar) + '</ul>\n',
              '<p style="margin:0 0 14px;font-size:13px;color:#52606d;">Detalle técnico: ' + html_lib.escape(corto or "—") + '</p>\n',
              '<p style="margin:0;">El portal conserva la última actualización correcta: este corte no modificó ningún dato.</p>\n']
    return asunto, _marco_correo("ATENCIÓN: el portal GHT NO se actualizó", "#b42318", "".join(partes))


def escribir_correo(cfg, ahora, asunto, cuerpo_html):
    """Escribe salida_interna/correo_asunto.txt y correo_cuerpo.html (UTF-8, juntos o ninguno).
    Un fallo aqui se deja en el log pero no cambia el resultado de la corrida."""
    carpeta = cfg["carpeta_salida_interna"]

    def escribir_texto(texto):
        def escritor(ruta_tmp):
            with open(ruta_tmp, "w", encoding="utf-8", newline="") as f:
                f.write(texto)
        return escritor
    try:
        escribir_todo_o_nada([(os.path.join(carpeta, ARCHIVO_CORREO_ASUNTO), escribir_texto(asunto)),
                              (os.path.join(carpeta, ARCHIVO_CORREO_CUERPO), escribir_texto(cuerpo_html))])
    except Exception as e:  # noqa: BLE001
        registrar_log(cfg, ahora, "auto", f"No se pudo escribir el correo: {type(e).__name__}: {e}")


def leer_estado_reportados(cfg):
    """Estado guardado por la ultima corrida --auto (o None si no existe / no se entiende)."""
    try:
        with open(os.path.join(cfg["carpeta_salida_interna"], ARCHIVO_ESTADO_DESPACHOS), "r", encoding="utf-8") as f:
            estado = json.load(f)
        return estado if isinstance(estado, dict) and "claves" in estado else None
    except (OSError, ValueError):
        return None


def cantidad_para_json(valor):
    """20000.0 -> 20000 (conserva decimales si los hay de verdad)."""
    return int(valor) if valor == int(valor) else round(valor, 2)


def calcular_despachos_nuevos(cfg, notas_historico, pedidos_cache, claves_previas_corrida):
    """Lineas de despacho que no se habian reportado en la corrida anterior.
    - "Corrida anterior" = el estado guardado en salida_interna (las llaves ya
      reportadas). Si todavia no existe, se usa lo que habia en el historico
      justo antes de leer el .xls de hoy.
    - Solo cuentan los despachos de pedidos GHT que el portal conoce (el .xls
      trae despachos de todos los clientes de Smurfit): los demas se omiten y
      solo se cuentan, sin mostrar ningun dato.
    Devuelve (lista_de_despachos, cantidad_omitida_de_otros_clientes)."""
    estado_previo = leer_estado_reportados(cfg)
    previas = set(estado_previo["claves"]) if estado_previo else None
    if previas is None:
        previas = claves_previas_corrida

    nuevos, omitidos = [], 0
    for clave, linea in notas_historico.items():
        if clave in previas:
            continue
        pedido = pedidos_cache.get(f"{linea['ov']}|{linea.get('elemento', '')}")
        if pedido is None:
            omitidos += 1
            continue
        nuevos.append({
            "nota": linea.get("nota", ""), "ov": linea["ov"], "finca": pedido["finca"],
            "elemento": linea.get("elemento", ""), "cantidad": cantidad_para_json(linea["cantidad"]),
            # Solo para el Excel del cliente (el resumen interno no los usa):
            "orden_compra": pedido.get("orden_compra", ""), "descripcion": pedido.get("descripcion", ""),
            "fecha": linea.get("fecha", ""),
            # Lineas guardadas por versiones anteriores no traen el campo: todas las
            # devoluciones del .xls tienen cantidad negativa, asi que se infiere de ahi.
            "tipo": "Devolución" if linea.get("devolucion", linea["cantidad"] < 0) else "Despacho",
        })
    return nuevos, omitidos


def escribir_excel_cliente(registros, despachos_nuevos):
    """Devuelve la funcion que escribe el Excel del cliente: hoja 'Consolidado'
    (los mismos pedidos y columnas que ve GHT en el portal) y hoja 'Despachos
    nuevos' (despachos desde el corte anterior)."""
    def escritor(ruta_tmp):
        libro = Workbook()
        hoja = libro.active
        hoja.title = "Consolidado"
        hoja.append([titulo for titulo, _ in COLUMNAS_CONSOLIDADO])
        for r in registros:
            hoja.append([r.get(campo) or "" for _, campo in COLUMNAS_CONSOLIDADO])
        hoja2 = libro.create_sheet("Despachos nuevos")
        hoja2.append([titulo for titulo, _ in COLUMNAS_DESPACHOS_NUEVOS])
        if despachos_nuevos:
            for d in despachos_nuevos:
                hoja2.append([d.get(campo, "") for _, campo in COLUMNAS_DESPACHOS_NUEVOS])
        else:
            hoja2.append(["Sin despachos nuevos desde el corte anterior."])
        for h, anchos in ((hoja, (14, 48, 16, 12, 16, 22, 20, 20, 22, 26)), (hoja2, (14, 16, 12, 48, 18, 16, 16, 14))):
            for celda in h[1]:
                celda.font = Font(bold=True)
            for i, ancho in enumerate(anchos):
                h.column_dimensions[chr(ord("A") + i)].width = ancho
            h.freeze_panes = "A2"
        libro.save(ruta_tmp)
    return escritor


def escribir_excel_excluidos(detalle):
    """Excel de los pedidos excluidos del avance, con el MISMO contenido que descarga
    el boton "Exportar excluidos del avance" del portal (index.html): hoja 'Excluidos'
    (8 columnas, con la razon de exclusion) y hoja 'Resumen'. 'detalle' es el mismo
    diccionario que se cifra en excluidos_avance.json."""
    def numero(v):
        return int(v) if isinstance(v, float) and v == int(v) else v

    def escritor(ruta_tmp):
        excluidos = detalle["excluidos"]
        libro = Workbook()
        hoja = libro.active
        hoja.title = "Excluidos"
        hoja.append(["Finca", "Orden de venta", "Elemento", "Orden de compra", "Cantidad solicitada",
                     "Cantidad despachada", "Fecha de entrega", "Razón de exclusión"])
        for p in excluidos:
            hoja.append([p["finca"], p["ov"], p["elemento"], p["orden_compra"], numero(p["cantidad_solicitada"]),
                         numero(p["cantidad_despachada"]), p["fecha_entrega"], p["razon"]])
        for celda in hoja[1]:
            celda.font = Font(bold=True)
        for i, ancho in enumerate((14, 14, 12, 20, 19, 19, 16, 34)):
            hoja.column_dimensions[chr(ord("A") + i)].width = ancho
        resumen = libro.create_sheet("Resumen")
        por_razon = lambda razon: sum(1 for p in excluidos if p["razon"] == razon)
        for fila in (
            ["Concepto", "Valor"],
            ["Cierre del ciclo (moda de fecha de entrega)", detalle["cierre_ciclo"]],
            ["Ciclo cerrado", "Sí" if detalle["ciclo_cerrado"] else "No"],
            ["Avance general (%)", detalle["porcentaje_global"]],
            ["Pedidos que cuentan en el %", detalle["pedidos_en_porcentaje"]],
            ["Solicitado en el %", numero(detalle["total_solicitado_en_porcentaje"])],
            ["Despachado en el %", numero(detalle["total_despachado_en_porcentaje"])],
            ["Pedidos excluidos", len(excluidos)],
            ["  " + RAZON_FECHA_FUTURA, por_razon(RAZON_FECHA_FUTURA)],
            ["  " + RAZON_SIN_DESPACHO, por_razon(RAZON_SIN_DESPACHO)],
        ):
            resumen.append(fila)
        for celda in resumen[1]:
            celda.font = Font(bold=True)
        resumen.column_dimensions["A"].width = 44
        resumen.column_dimensions["B"].width = 16
        libro.save(ruta_tmp)
    return escritor


def preparar_salidas(cfg, ahora, resultado, notas_historico, pedidos_cache, claves_previas_corrida,
                     avance, despachos_hasta_txt, ruta_despachos, guardar_estado, detalle_excluidos=None):
    """Arma (sin escribir todavia) el resumen para el correo, el Excel del
    cliente y, SOLO si guardar_estado (corridas --auto), el estado de despachos
    reportados. Una corrida manual no consume despachos: genera el Excel y el
    resumen pero deja el estado como estaba. Devuelve (destinos, resumen, extra);
    'extra' trae lo que necesita el correo (despachos nuevos con OC y Tipo, omitidos,
    corte de despachos de la corrida --auto anterior)."""
    nuevos, omitidos = calcular_despachos_nuevos(cfg, notas_historico, pedidos_cache, claves_previas_corrida)
    # corte de despachos de la corrida --auto anterior (para avisar si el .xls no se ha actualizado)
    estado_previo = leer_estado_reportados(cfg)
    hasta_anterior = estado_previo.get("despachos_hasta") if estado_previo else None
    # Corte que CIERRA el ciclo: ya esta cerrado y la corrida --auto anterior lo dejo abierto
    # (o era de otro ciclo). Si el estado anterior no trae esta informacion no se puede saber,
    # y por prudencia no se avisa.
    cierre_txt = avance["cierre_ciclo"].strftime("%d/%m/%Y") if avance["cierre_ciclo"] else ""
    ciclo_previo = estado_previo.get("ciclo") if estado_previo else None
    cierra_ahora = bool(avance["ciclo_cerrado"] and ciclo_previo
                        and not (ciclo_previo.get("cierre") == cierre_txt and ciclo_previo.get("cerrado")))
    excluidos_ciclo = len(avance["futuros"]) + len(avance["sin_despacho"])
    nombre_excel = f"GHT_consolidado_{ahora:%Y-%m-%d_%H%M}.xlsx"
    ruta_excel = os.path.join(cfg["carpeta_salida"], nombre_excel)
    resumen = {
        "modo": "auto" if guardar_estado else "manual",
        "corte": etiqueta_de_corte(ahora, cfg["cortes"], cfg["tolerancia_corte_minutos"]),
        "hora": ahora.isoformat(timespec="seconds"),
        "hora_texto": ahora.strftime("%d/%m/%Y %H:%M"),
        "estado": "ok" if nuevos else "sin_despachos_nuevos",
        "avance_general": avance["porcentaje"],
        "despachos_hasta": despachos_hasta_txt,
        "archivo_despachos": os.path.basename(ruta_despachos),
        "total_despachos_nuevos": len(nuevos),
        "despachos_nuevos": [
            {"nota": d["nota"], "ov": d["ov"], "finca": d["finca"], "elemento": d["elemento"], "cantidad": d["cantidad"]}
            for d in nuevos
        ],
        "despachos_nuevos_otros_clientes_omitidos": omitidos,
        "excel_cliente": f"{os.path.basename(cfg['carpeta_salida'])}/{nombre_excel}",
    }
    estado = {"actualizado": ahora.isoformat(timespec="seconds"), "despachos_hasta": despachos_hasta_txt,
              "ciclo": {"cierre": cierre_txt, "cerrado": bool(avance["ciclo_cerrado"])},
              "claves": sorted(notas_historico)}
    carpeta_interna = cfg["carpeta_salida_interna"]
    destinos = [
        (ruta_excel, escribir_excel_cliente(resultado, nuevos)),
        (os.path.join(carpeta_interna, ARCHIVO_RESUMEN), escribir_json(resumen, ensure_ascii=False, indent=2)),
    ]
    if guardar_estado:
        destinos.append((os.path.join(carpeta_interna, ARCHIVO_ESTADO_DESPACHOS), escribir_json(estado, ensure_ascii=False)))
    # Solo la corrida --auto que CIERRA el ciclo y tiene al menos 1 excluido deja el Excel de excluidos.
    excel_excluidos = None
    if guardar_estado and cierra_ahora and excluidos_ciclo > 0 and detalle_excluidos is not None:
        nombre_excluidos = f"excluidos_ciclo_{avance['cierre_ciclo']:%Y-%m-%d}.xlsx"
        destinos.append((os.path.join(carpeta_interna, nombre_excluidos), escribir_excel_excluidos(detalle_excluidos)))
        excel_excluidos = f"{os.path.basename(carpeta_interna)}/{nombre_excluidos}"
    extra = {"nuevos": nuevos, "omitidos": omitidos, "hasta_anterior": hasta_anterior,
             "cierre_ciclo_excluidos": excluidos_ciclo if cierra_ahora else None,
             "adjunto_excluidos": excel_excluidos}
    return destinos, resumen, extra


def generar_datos(auto=False, cfg=None, ahora=None):
    if cfg is None:
        cfg = cargar_config()
        aplicar_config(cfg)
    ahora = ahora or datetime.now(BOGOTA)
    espera = cfg["espera_reintento_segundos"]
    if auto:
        print("Modo automatico: verificando archivos de entrada y clave interna...")
        verificar_entradas_auto(espera)

    print("Leyendo Order Capacity...")
    mapa_order_capacity = cargar_order_capacity(RUTA_BACKLOG)

    print("Leyendo Programacion...")
    mapa_programacion = cargar_programacion(RUTA_PROGRAMACION)

    print("Leyendo Notas de Despacho...")
    ruta_despachos = encontrar_archivo_mas_reciente(CARPETA_DESPACHOS, ".xls")
    print(f"  Archivo encontrado: {ruta_despachos}")
    # "Despachos hasta": fecha-hora de modificacion del .xls (hora Colombia).
    despachos_hasta_txt = datetime.fromtimestamp(os.path.getmtime(ruta_despachos), BOGOTA).strftime("%d/%m/%Y %H:%M")
    print(f"  Despachos hasta: {despachos_hasta_txt} (fecha de modificacion del archivo)")
    if auto:
        lineas_nuevas = cargar_despachos_con_reintento(ruta_despachos, espera)
    else:
        lineas_nuevas = cargar_despachos(ruta_despachos)

    # --- Cargar el historial acumulado de corridas anteriores ---
    # "notas": cada envio individual, identificado por su Nota de Despacho
    #          (unica por envio) - evita contar 2 veces el mismo envio
    #          aunque el archivo de un dia se solape en fechas con otro.
    # "pedidos": la ultima info completa vista en el backlog para cada
    #            Orden de Venta (finca, elemento, cantidad solicitada...),
    #            para poder seguir mostrando un pedido despues de que el
    #            backlog ya lo haya quitado de su lista.
    try:
        with open(ARCHIVO_HISTORICO_DESPACHOS, "r", encoding="utf-8") as f:
            historico = json.load(f)
    except FileNotFoundError:
        historico = {}

    if "notas" not in historico or "pedidos" not in historico:
        # Migracion desde el formato viejo (OV -> registro plano). Se
        # conserva como cache de "pedidos" para no perder informacion,
        # pero sin el detalle de cantidades por nota (no exist�a antes).
        pedidos_viejo = {}
        for ov, valor in historico.items():
            if isinstance(valor, dict) and "finca" in valor:
                pedidos_viejo[ov] = valor
        historico = {"notas": {}, "pedidos": pedidos_viejo}

    notas_historico = historico["notas"]
    pedidos_cache = historico["pedidos"]

    # --- Migrar cache de corridas anteriores al esquema con Elemento ---
    # Antes, "notas" se guardaba con llave "nota|ov" (sin Elemento) y
    # "pedidos" con llave "ov" sola. Se recupera el Elemento que falta
    # usando "pedidos" (que si trae el elemento de esa OV) - funciona
    # bien salvo en el caso raro de una OV con mas de un Elemento, donde
    # ese dato puntual del historico se pierde y se vuelve a acumular
    # con los despachos de los proximos dias.
    pedidos_migrados = {}
    for clave, info in pedidos_cache.items():
        nueva_clave = clave if "|" in clave else f"{clave}|{info.get('elemento', '')}"
        pedidos_migrados[nueva_clave] = info
    pedidos_cache = pedidos_migrados

    notas_migradas = {}
    for clave, linea in notas_historico.items():
        if linea.get("elemento"):
            notas_migradas[clave] = linea
            continue
        ov_vieja = linea.get("ov", "")
        elemento_recuperado = ""
        for clave_p, info_p in pedidos_cache.items():
            if clave_p.startswith(f"{ov_vieja}|"):
                elemento_recuperado = info_p.get("elemento", "")
                break
        if not elemento_recuperado:
            continue  # no se pudo recuperar el elemento; se descarta esta entrada vieja
        linea_migrada = dict(linea, elemento=elemento_recuperado)
        notas_migradas[f"{linea.get('nota','')}|{ov_vieja}|{elemento_recuperado}"] = linea_migrada
    notas_historico = notas_migradas

    # Guardamos qué pedidos ya conocíamos ANTES de esta corrida, para poder
    # diagnosticar al final cuáles "desaparecieron" hoy sin haberse
    # despachado (ver diagnóstico al final de esta función). Las llaves
    # son "ov|elemento", no solo "ov".
    ovs_conocidas_antes = set(pedidos_cache.keys())

    # Llaves que ya conociamos ANTES de sumar este .xls (base para "despachos
    # nuevos" cuando todavia no existe el estado de la corrida anterior).
    claves_previas_corrida = set(notas_historico.keys())

    notas_nuevas = 0
    for linea in lineas_nuevas:
        if linea["clave"] not in notas_historico:
            notas_nuevas += 1
        notas_historico[linea["clave"]] = {
            "ov": linea["ov"], "elemento": linea["elemento"], "nota": linea["nota"],
            "cantidad": linea["cantidad"], "fecha": linea["fecha"],
            "devolucion": linea["devolucion"],
        }

    # Acumular cantidad total despachada, fecha del ultimo envio, y la
    # lista de notas de despacho involucradas, por Elemento+OV (NO solo
    # por OV: una misma Orden de Venta a veces se repite en mas de un
    # Elemento en el backlog, y cada uno debe llevar su propio despacho).
    acumulado_por_clave = {}
    for linea_guardada in notas_historico.values():
        clave_acum = f"{linea_guardada['ov']}|{linea_guardada.get('elemento', '')}"
        entrada = acumulado_por_clave.setdefault(clave_acum, {"cantidad": 0.0, "fecha": None, "notas": set()})
        entrada["cantidad"] += linea_guardada["cantidad"]
        if linea_guardada.get("nota"):
            entrada["notas"].add(linea_guardada["nota"])
        try:
            f_linea = datetime.strptime(linea_guardada["fecha"], "%d/%m/%Y").date()
            if entrada["fecha"] is None or f_linea > entrada["fecha"]:
                entrada["fecha"] = f_linea
        except (ValueError, TypeError):
            pass

    print(f"Lineas de despacho nuevas en este archivo: {notas_nuevas}")
    print(f"Total de lineas de despacho acumuladas en el historial: {len(notas_historico)}")

    print("Leyendo backlog (hoja Formato)...")
    wb = load_workbook(RUTA_BACKLOG, read_only=True, data_only=True)
    ws = wb["Formato"]
    filas = list(ws.iter_rows(values_only=True))
    encabezado = [limpiar_texto(c) for c in filas[0]]
    datos = filas[1:]

    idx_fecha_entrega = encabezado.index("Fecha entrega")
    idx_id_cliente = encabezado.index("ID Cliente")
    idx_elemento = encabezado.index("Elemento")
    idx_descripcion = encabezado.index("Descripción")
    idx_ord_compra = encabezado.index("Ord. de Compra")
    idx_ord_venta = encabezado.index("Ord. de Venta")
    idx_ord_trabajo = encabezado.index("Ord. de Trabajo")
    idx_estatus_omp = encabezado.index("Estatus OMP")
    idx_cant_sol = encabezado.index("Cant Sol")

    def construir_registro(id_cliente, elemento, descripcion, ow_final, orden_compra, ov, cant_sol_original, estado_produccion, fecha_entrega=""):
        """
        Decide el estado final de un pedido comparando lo despachado
        (acumulado_por_clave, indexado por Elemento+OV) contra lo
        solicitado ORIGINALMENTE (cant_sol_original, congelado desde la
        primera vez que se vio el pedido en el backlog - ver mas abajo
        donde se arma pedidos_cache):
          - Nada despachado todavia -> se respeta el estado de produccion.
          - Se despacho TODO (o mas, por redondeos) -> "Despachado".
          - Se despacho una parte -> "Parcialmente despachado".

        Tambien calcula "porcentaje_despachado": lo acumulado despachado
        sobre el total ORIGINAL del pedido (no sobre lo que diga el backlog
        de hoy), para que el % avance de forma consistente semana a semana
        y no salte si la cantidad solicitada cambia mas adelante.
        """
        info_despacho = acumulado_por_clave.get(f"{ov}|{elemento}")
        cantidad_despachada = info_despacho["cantidad"] if info_despacho else 0.0
        fecha_ultimo = info_despacho["fecha"] if info_despacho else None
        notas_texto = ""
        if info_despacho and info_despacho["notas"]:
            notas_texto = ", ".join(sorted(info_despacho["notas"]))

        estado = estado_produccion
        fecha_despacho_txt = ""
        if cantidad_despachada > 0:
            if cant_sol_original and cantidad_despachada >= cant_sol_original:
                estado = "Despachado"
            else:
                estado = "Parcialmente despachado"
            if fecha_ultimo:
                fecha_despacho_txt = fecha_ultimo.strftime("%d/%m/%Y")

        porcentaje_despachado = None
        if cant_sol_original:
            porcentaje_despachado = round(min(100.0, (cantidad_despachada / cant_sol_original) * 100))

        return {
            "finca": id_cliente,
            "elemento": elemento,
            "descripcion": descripcion,
            "ow": ow_final,
            "orden_compra": orden_compra,
            "estado": estado,
            "fecha_despacho": fecha_despacho_txt,
            "cantidad_solicitada": formatear_cantidad(cant_sol_original) if cant_sol_original else "",
            "cantidad_despachada": formatear_cantidad(cantidad_despachada) if cantidad_despachada else "",
            "porcentaje_despachado": porcentaje_despachado,
            "nota_despacho": notas_texto,
            # Este campo queda vacio hasta que TI termine de ajustar el bot
            # de despachos para que incluya la fecha/hora ESTIMADA (antes
            # de que el pedido salga).
            "fecha_estimada_despacho": "",
            # Columna A del backlog ("Fecha entrega"): estimado de cuando
            # llegara el pedido. Viene de la MISMA fila que elemento/OV/OC,
            # asi que no necesita cruce aparte - ya esta garantizado que
            # corresponde al pedido correcto.
            "fecha_entrega": fecha_entrega,
        }, fecha_ultimo

    resultado = []
    ov_cubiertas = set()
    total_leidas = 0
    total_ght = 0
    pedidos_sin_estatus_hoy = []  # diagnóstico: ver mensaje al final

    for fila in datos:
        id_cliente = limpiar_texto(fila[idx_id_cliente])
        if not id_cliente:
            continue
        total_leidas += 1

        # Filtro: solo clientes que empiecen con GHT
        if not id_cliente.startswith("GHT"):
            continue
        total_ght += 1

        elemento = limpiar_texto(fila[idx_elemento])
        descripcion = limpiar_texto(fila[idx_descripcion])
        orden_compra = limpiar_texto(fila[idx_ord_compra])
        ord_venta = limpiar_texto(fila[idx_ord_venta])
        fecha_entrega = formatear_fecha_entrega(fila[idx_fecha_entrega])

        try:
            cant_sol = float(fila[idx_cant_sol]) if fila[idx_cant_sol] not in (None, "") else 0.0
        except (ValueError, TypeError):
            cant_sol = 0.0

        # --- OW Final: si el backlog ya trae Ord. de Trabajo propio, se respeta ---
        ord_trabajo_original = limpiar_texto(fila[idx_ord_trabajo])
        ow_desde_programacion = mapa_programacion.get(elemento, "")
        if ord_trabajo_original in ("", "-", "0"):
            ow_final = ow_desde_programacion
        else:
            ow_final = ord_trabajo_original

        # --- Estatus OMP Final: si ya tenia estatus, se respeta; si no, se trae de Order Capacity ---
        estatus_original = limpiar_texto(fila[idx_estatus_omp])
        estatus_nuevo = mapa_order_capacity.get(ow_final, "")
        if estatus_original in ("", "-"):
            estatus_omp_final = estatus_nuevo
        else:
            estatus_omp_final = estatus_original

        estado_produccion = clasificar_estado_ght(estatus_omp_final)

        # --- Cantidad solicitada ORIGINAL: se congela la primera vez que
        # vemos este pedido (Elemento+Orden de Venta) y nunca se vuelve a
        # pisar, aunque el backlog cambie el valor mas adelante. Asi el %
        # de avance siempre se mide contra el total inicial del pedido.
        clave_pedido = f"{ord_venta}|{elemento}"
        cache_previo = pedidos_cache.get(clave_pedido)
        if cache_previo and cache_previo.get("cant_sol_original"):
            cant_sol_original = cache_previo["cant_sol_original"]
        else:
            cant_sol_original = cant_sol

        registro, _ = construir_registro(
            id_cliente, elemento, descripcion, ow_final, orden_compra,
            ord_venta, cant_sol_original, estado_produccion, fecha_entrega,
        )

        # No exponemos los pedidos "Sin clasificar" al cliente (decision ya tomada en el proyecto)
        if registro["estado"] == "Sin clasificar":
            # DIAGNÓSTICO: si este pedido YA lo conocíamos de una corrida
            # anterior (estaba en el portal antes) y hoy no trae estatus
            # (ni en su propia fila del backlog ni en la carga de máquinas
            # de hoy), avisamos aquí para poder revisarlo con datos reales.
            if clave_pedido in ovs_conocidas_antes:
                pedidos_sin_estatus_hoy.append({
                    "finca": id_cliente, "elemento": elemento, "ov": ord_venta,
                })
            continue

        if ord_venta:
            ov_cubiertas.add(clave_pedido)
            # Guardamos SIEMPRE la info mas reciente de este pedido (no solo
            # cuando ya se despacho), para poder recuperarla despues si el
            # backlog lo quita de su lista antes de completarse el despacho.
            # "cant_sol_original" queda congelada (ver arriba); no se
            # sobrescribe con el valor del dia. La llave es Elemento+OV,
            # no solo OV: una misma Orden de Venta a veces se repite en
            # mas de un Elemento en el backlog.
            pedidos_cache[clave_pedido] = {
                "finca": id_cliente, "elemento": elemento, "descripcion": descripcion,
                "ow": ow_final, "orden_compra": orden_compra,
                "cant_sol_original": cant_sol_original,
                "estado_produccion": estado_produccion,
                "fecha_entrega": fecha_entrega,
            }

        resultado.append(registro)

    # --- Agregar pedidos despachados/parciales recientes que el backlog YA no muestra ---
    hoy = datetime.now().date()
    agregados_desde_historico = 0
    ov_recuperadas = set()
    for clave_pedido, info in pedidos_cache.items():
        if clave_pedido in ov_cubiertas:
            continue  # ya viene del backlog de hoy, no lo dupliquemos

        ov_de_clave = clave_pedido.split("|", 1)[0]

        # Compatibilidad: cache viejo (antes de este cambio) solo tenia
        # "cant_sol"; el nuevo tiene "cant_sol_original" congelada.
        cant_sol_original = info.get("cant_sol_original", info.get("cant_sol", 0))

        registro, fecha_ultimo = construir_registro(
            info["finca"], info["elemento"], info["descripcion"], info["ow"],
            info["orden_compra"], ov_de_clave, cant_sol_original, info.get("estado_produccion", "Sin clasificar"),
            info.get("fecha_entrega", ""),
        )
        # Solo tiene sentido recuperarlo si de verdad hubo algun despacho
        # (si nunca se despacho nada, y ya no esta en el backlog, no
        # tenemos nada confiable que mostrar de el).
        if not fecha_ultimo:
            continue
        if (hoy - fecha_ultimo).days <= DIAS_VISIBLE_DESPACHADO:
            resultado.append(registro)
            agregados_desde_historico += 1
            ov_recuperadas.add(clave_pedido)

    # DIAGNÓSTICO: pedidos que ya conocíamos, nunca se habían despachado,
    # y hoy no quedaron en ningún lado del resultado (ni backlog de hoy,
    # ni recuperados del histórico porque nunca tuvieron un despacho que
    # los sostenga). Estos son los que "desaparecen" del portal sin que
    # el cliente vea por qué.
    ov_desaparecidas_del_todo = ovs_conocidas_antes - ov_cubiertas - ov_recuperadas

    # ------------------------------------------------------------
    # PORCENTAJE DE AVANCE DE LA SEMANA, contra el BACKLOG_CONSOLIDADO fijo
    # (la totalidad real de los pedidos de la semana, segun logistica). No
    # se recorre "resultado" para esto - se recorre directamente la base
    # del consolidado, y para cada pedido de ahi se busca cuanto se ha
    # despachado en total (acumulado_por_clave, que viene de TODAS las
    # notas de despacho acumuladas, sin importar si el pedido ya salio del
    # backlog diario o no). Los adicionales que no estan en el consolidado
    # simplemente no entran a esta suma - siguen viendose normal en el
    # portal con su propio estado, pero no afectan este numero.
    # ------------------------------------------------------------
    ruta_consolidado = encontrar_backlog_consolidado(CARPETA_DESPACHOS)
    base_semana = {}
    if ruta_consolidado:
        print(f"Backlog consolidado de la semana encontrado: {ruta_consolidado}")
        base_semana = cargar_base_semana_ght(ruta_consolidado)
    else:
        print("No se encontro ningun archivo 'consolidado' en la carpeta - "
              "el % de avance de la semana no se puede calcular hoy.")

    # La fecha y la hora del cierre del ciclo salen de la hora de la corrida (Colombia).
    avance = calcular_avance_general(base_semana, acumulado_por_clave, ahora.date(), ahora.time())
    total_sol_semana = avance["total_sol"]
    total_desp_semana = avance["total_desp"]
    porcentaje_global = avance["porcentaje"]
    cierre_ciclo = avance["cierre_ciclo"]

    for registro in resultado:
        registro["porcentaje_global"] = porcentaje_global
        registro["despachos_hasta"] = despachos_hasta_txt

    # El detalle de los pedidos excluidos del % NO va en datos.json (que ve el
    # cliente): se guarda aparte, CIFRADO, para el boton interno de exportar
    # a Excel. Nunca se escribe en claro: si falta la clave o la libreria, se
    # avisa y el archivo anterior se deja como estaba.
    destinos = []  # (ruta_final, escritor): se escriben TODOS juntos al final (todo o nada)
    detalle_excluidos = armar_detalle_excluidos(base_semana, acumulado_por_clave, avance)
    clave_interna = leer_clave_interna()
    if clave_interna is None:
        print(f"*** AVISO: no hay clave interna ('{ARCHIVO_CLAVE_INTERNA}' o variable "
              f"{VARIABLE_CLAVE_INTERNA}). NO se actualizo {ARCHIVO_EXCLUIDOS}. ***")
    else:
        try:
            sobre = cifrar_contenido(detalle_excluidos, clave_interna)
        except ImportError:
            print(f"*** AVISO: falta la libreria 'cryptography' (pip install cryptography). "
                  f"NO se actualizo {ARCHIVO_EXCLUIDOS}. ***")
        else:
            destinos.append((ARCHIVO_EXCLUIDOS, escribir_json(sobre, indent=2)))

    historico = {"notas": notas_historico, "pedidos": pedidos_cache}
    destinos.append((ARCHIVO_HISTORICO_DESPACHOS, escribir_json(historico, ensure_ascii=False, indent=2)))
    destinos.append((ARCHIVO_SALIDA, escribir_json(resultado, ensure_ascii=False, indent=2)))

    # Resumen para el correo (interno), Excel del cliente y estado de despachos.
    destinos_extra, resumen, extra = preparar_salidas(
        cfg, ahora, resultado, notas_historico, pedidos_cache, claves_previas_corrida,
        avance, despachos_hasta_txt, ruta_despachos, guardar_estado=auto, detalle_excluidos=detalle_excluidos)
    destinos += destinos_extra
    escribir_todo_o_nada(destinos)

    print()
    print(f"Total filas leidas en el backlog: {total_leidas}")
    print(f"Total filas de GHT (antes de quitar 'Sin clasificar'): {total_ght}")
    print(f"Pedidos recuperados del historico (ya no estan en el backlog): {agregados_desde_historico}")
    print(f"Total filas exportadas a {ARCHIVO_SALIDA}: {len(resultado)}")
    print(f"Pedidos GHT en el backlog consolidado de la semana: {len(base_semana)}")
    print(f"Cierre del ciclo (moda de fecha_entrega del consolidado): "
          f"{cierre_ciclo.strftime('%d/%m/%Y') if cierre_ciclo else '—'} "
          f"({'ya cerro' if avance['ciclo_cerrado'] else 'aun abierto'})")
    print(f"  Excluidos del % por Regla 1 (fecha_entrega posterior al ciclo): {len(avance['futuros'])}")
    for clave in avance["futuros"]:
        d = base_semana[clave]
        print(f"    - {clave} | entrega {d['fecha_entrega'].strftime('%d/%m/%Y')} | solicitado {formatear_cantidad(d['cant_sol'])}")
    print(f"  Excluidos del % por Regla 2 (sin ningun despacho al cerrar el ciclo): {len(avance['sin_despacho'])}")
    for clave in avance["sin_despacho"]:
        d = base_semana[clave]
        print(f"    - {clave} | solicitado {formatear_cantidad(d['cant_sol'])}")
    print(f"  Pedidos que SI cuentan en el %: {len(avance['incluidos'])}")
    print(f"Porcentaje de avance de la semana: "
          f"{porcentaje_global if porcentaje_global is not None else '—'}% "
          f"({formatear_cantidad(total_desp_semana)} / {formatear_cantidad(total_sol_semana)})")
    print()

    # ------------------------------------------------------------
    # DIAGNÓSTICO — pedidos que hoy quedaron "sin estatus" (carga de
    # máquinas no los trajo) y pedidos que desaparecieron del todo.
    # Esto NO afecta el datos.json generado, es solo para revisar contigo
    # si hay pedidos pendientes que se están perdiendo de vista.
    # ------------------------------------------------------------
    print("=" * 60)
    print("DIAGNÓSTICO (no afecta el portal, solo para revisar)")
    print("=" * 60)
    if pedidos_sin_estatus_hoy:
        print(f"\n{len(pedidos_sin_estatus_hoy)} pedido(s) que ya conocíamos de antes "
              f"NO trajeron estatus hoy (ni en su fila del backlog, ni en la carga "
              f"de máquinas) y por eso NO aparecen hoy en el portal:")
        for p in pedidos_sin_estatus_hoy[:20]:
            print(f"  - Finca: {p['finca']} | Elemento: {p['elemento']} | OV: {p['ov']}")
        if len(pedidos_sin_estatus_hoy) > 20:
            print(f"  ... y {len(pedidos_sin_estatus_hoy) - 20} más.")
    else:
        print("\nNo hubo pedidos que 'perdieran' su estatus hoy. Bien.")

    if ov_desaparecidas_del_todo:
        print(f"\n{len(ov_desaparecidas_del_todo)} pedido(s) que ya conocíamos, nunca se "
              f"han despachado, y hoy no quedaron en el datos.json (ni en el backlog de "
              f"hoy ni recuperados del histórico):")
        for clave_pedido in list(ov_desaparecidas_del_todo)[:20]:
            info = pedidos_cache.get(clave_pedido, {})
            ov_mostrar = clave_pedido.split("|", 1)[0]
            print(f"  - Finca: {info.get('finca','?')} | Elemento: {info.get('elemento','?')} | OV: {ov_mostrar}")
    else:
        print("\nNingún pedido pendiente desapareció por completo hoy. Bien.")
    print("=" * 60)
    print()

    print(f"Corte: {resumen['corte']} | estado: {resumen['estado']} | despachos nuevos (GHT): "
          f"{resumen['total_despachos_nuevos']} | despachos hasta: {despachos_hasta_txt}")
    print(f"Excel del cliente: {resumen['excel_cliente']}")
    if not auto:
        print("(Corrida manual: no se actualizo el estado de despachos reportados; "
              "los despachos nuevos de arriba se volveran a reportar en la proxima corrida --auto.)")
    print("Listo. Sube el archivo 'datos.json' al portal web.")
    return resumen, extra


def main(argv=None):
    parser = argparse.ArgumentParser(description="Genera datos.json del portal GHT.")
    parser.add_argument("--auto", action="store_true",
                        help="modo automatico: sin preguntas; si falta algo, no escribe nada y sale con error")
    parser.add_argument("--config", help="ruta de config.json (por defecto, el de la carpeta del script)")
    parser.add_argument("--ahora", help="SOLO PARA PRUEBAS: simula la fecha y hora (Colombia) de la corrida, "
                                        "p. ej. '2026-10-06 16:00'")
    args = parser.parse_args(argv)
    modo = "auto" if args.auto else "manual"
    if args.auto:
        for flujo in (sys.stdout, sys.stderr):
            try:
                flujo.reconfigure(encoding="utf-8", errors="replace")  # tareas programadas: sin consola
            except Exception:
                pass
    if args.ahora:
        for formato in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
            try:
                ahora = datetime.strptime(args.ahora, formato).replace(tzinfo=BOGOTA)
                break
            except ValueError:
                ahora = None
        if ahora is None:
            print("ERROR: --ahora debe tener el formato 'AAAA-MM-DD HH:MM'.", file=sys.stderr)
            return SALIDA_CONFIG
    else:
        ahora = datetime.now(BOGOTA)
    cfg = None
    try:
        cfg = cargar_config(args.config)
        aplicar_config(cfg)
        resumen, extra = generar_datos(auto=args.auto, cfg=cfg, ahora=ahora)
    except ErrorCorrida as e:
        print(f"\nERROR (codigo {e.codigo}): {e.mensaje}", file=sys.stderr)
        print("No se modifico ningun archivo del portal ni de las salidas.", file=sys.stderr)
        cfg_ok = cfg or config_por_defecto()
        registrar_log(cfg_ok, ahora, modo, f"ERROR codigo={e.codigo} | {e.mensaje}")
        if args.auto:
            corte = etiqueta_de_corte(ahora, cfg_ok["cortes"], cfg_ok["tolerancia_corte_minutos"])
            asunto, cuerpo = construir_correo_error(ahora, corte, e.codigo, e.corto)
            escribir_correo(cfg_ok, ahora, asunto, cuerpo)
            escribir_ultima_corrida(cfg_ok, ahora, e.codigo, "error", e.corto, despachos_hasta_del_xls_mas_reciente(cfg_ok), None)
        return e.codigo
    except Exception as e:
        cfg_ok = cfg or config_por_defecto()
        registrar_log(cfg_ok, ahora, modo, f"ERROR codigo={SALIDA_ERROR} | inesperado: {type(e).__name__}: {e}")
        if not args.auto:
            raise
        print(f"\nERROR (codigo {SALIDA_ERROR}): error inesperado: {type(e).__name__}: {e}", file=sys.stderr)
        traceback.print_exc()
        corto = f"Error inesperado: {type(e).__name__}: {mensaje_sin_rutas(str(e), cfg_ok, 110)}"
        corte = etiqueta_de_corte(ahora, cfg_ok["cortes"], cfg_ok["tolerancia_corte_minutos"])
        asunto, cuerpo = construir_correo_error(ahora, corte, SALIDA_ERROR, corto)
        escribir_correo(cfg_ok, ahora, asunto, cuerpo)
        escribir_ultima_corrida(cfg_ok, ahora, SALIDA_ERROR, "error", corto,
                                despachos_hasta_del_xls_mas_reciente(cfg_ok), None)
        return SALIDA_ERROR
    registrar_log(cfg, ahora, modo,
                  f"OK | corte={resumen['corte']} | estado={resumen['estado']} | avance={resumen['avance_general']}% | "
                  f"despachos_nuevos={resumen['total_despachos_nuevos']} | despachos_hasta={resumen['despachos_hasta']} | "
                  f"xls={resumen['archivo_despachos']} | excel={resumen['excel_cliente']}")
    if args.auto:
        mismo_archivo = extra["hasta_anterior"] is not None and extra["hasta_anterior"] == resumen["despachos_hasta"]
        asunto, cuerpo = construir_correo_ok(ahora, resumen, extra["nuevos"], extra["omitidos"], mismo_archivo,
                                             extra["cierre_ciclo_excluidos"])
        escribir_correo(cfg, ahora, asunto, cuerpo)
        escribir_ultima_corrida(cfg, ahora, SALIDA_OK, resumen["estado"], None, resumen["despachos_hasta"],
                                resumen["excel_cliente"], extra["adjunto_excluidos"])
    return SALIDA_OK


if __name__ == "__main__":
    sys.exit(main())
