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

Requiere: pip install openpyxl xlrd cryptography

CLAVE INTERNA (para cifrar excluidos_avance.json): se lee del archivo local
"clave_interna.txt" (esta en .gitignore, NO se sube al repo) o de la variable
de entorno CLAVE_INTERNA_PORTAL. Es la misma clave que se usa para entrar al
portal como usuario interno de Smurfit.
============================================================
"""

import base64
import hashlib
import json
import os
from collections import Counter
from datetime import datetime, timedelta
from openpyxl import load_workbook

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
#     una vez que el ciclo ya cerro (hoy > cierre).
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


def calcular_avance_general(base_semana, acumulado_por_clave, hoy):
    """Calcula el "Avance general" sobre la base fija del consolidado,
    sacando de la cuenta (sin quitarlos del portal) dos grupos de pedidos:
      - Regla 1 "futuros": fecha_entrega POSTERIOR al cierre del ciclo.
      - Regla 2 "sin despacho": fecha_entrega dentro del ciclo (o sin fecha),
        CERO unidades despachadas, y el ciclo ya cerro (hoy > cierre). Si el
        ciclo sigue abierto, estos pedidos cuentan normal.
    Devuelve un diccionario con el %, los totales y las llaves excluidas."""
    cierre = calcular_fecha_cierre_ciclo(base_semana)
    ciclo_cerrado = cierre is not None and hoy > cierre

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
        lineas.append({
            "clave": clave_linea, "ov": ov, "elemento": elemento,
            "nota": nota, "cantidad": cantidad, "fecha": fecha,
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


def generar_datos():
    print("Leyendo Order Capacity...")
    mapa_order_capacity = cargar_order_capacity(RUTA_BACKLOG)

    print("Leyendo Programacion...")
    mapa_programacion = cargar_programacion(RUTA_PROGRAMACION)

    print("Leyendo Notas de Despacho...")
    ruta_despachos = encontrar_archivo_mas_reciente(CARPETA_DESPACHOS, ".xls")
    print(f"  Archivo encontrado: {ruta_despachos}")
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

    notas_nuevas = 0
    for linea in lineas_nuevas:
        if linea["clave"] not in notas_historico:
            notas_nuevas += 1
        notas_historico[linea["clave"]] = {
            "ov": linea["ov"], "elemento": linea["elemento"], "nota": linea["nota"],
            "cantidad": linea["cantidad"], "fecha": linea["fecha"],
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

    avance = calcular_avance_general(base_semana, acumulado_por_clave, hoy)
    total_sol_semana = avance["total_sol"]
    total_desp_semana = avance["total_desp"]
    porcentaje_global = avance["porcentaje"]
    cierre_ciclo = avance["cierre_ciclo"]

    for registro in resultado:
        registro["porcentaje_global"] = porcentaje_global

    # El detalle de los pedidos excluidos del % NO va en datos.json (que ve el
    # cliente): se guarda aparte, CIFRADO, para el boton interno de exportar
    # a Excel. Nunca se escribe en claro: si falta la clave o la libreria, se
    # avisa y el archivo anterior se deja como estaba.
    clave_interna = leer_clave_interna()
    if clave_interna is None:
        print(f"*** AVISO: no hay clave interna ('{ARCHIVO_CLAVE_INTERNA}' o variable "
              f"{VARIABLE_CLAVE_INTERNA}). NO se actualizo {ARCHIVO_EXCLUIDOS}. ***")
    else:
        detalle_excluidos = armar_detalle_excluidos(base_semana, acumulado_por_clave, avance)
        try:
            sobre = cifrar_contenido(detalle_excluidos, clave_interna)
        except ImportError:
            print(f"*** AVISO: falta la libreria 'cryptography' (pip install cryptography). "
                  f"NO se actualizo {ARCHIVO_EXCLUIDOS}. ***")
        else:
            with open(ARCHIVO_EXCLUIDOS, "w", encoding="utf-8") as f:
                json.dump(sobre, f, indent=2)

    historico = {"notas": notas_historico, "pedidos": pedidos_cache}
    with open(ARCHIVO_HISTORICO_DESPACHOS, "w", encoding="utf-8") as f:
        json.dump(historico, f, ensure_ascii=False, indent=2)

    with open(ARCHIVO_SALIDA, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

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

    print("Listo. Sube el archivo 'datos.json' al portal web.")


if __name__ == "__main__":
    generar_datos()
