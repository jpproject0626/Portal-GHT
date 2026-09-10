"""
============================================================
 GENERADOR DE DATOS DIARIOS - Portal GHT (Grupo Chia)
============================================================
Este script lee el backlog diario y la programacion semanal,
aplica la misma logica que ya esta funcionando en Power BI
(filtro GHT, calculo de Estatus OMP Final, OW Final y Estado GHT),
y genera un archivo "datos.json" listo para subir al portal web.

COMO USARLO:
1. Ajusta las 2 rutas de archivo abajo (BACKLOG y PROGRAMACION)
   si cambian de ubicacion.
2. Corre este script cada dia despues de que el backlog se actualice:
       python generar_datos.py
3. Sube el archivo "datos.json" que se genera a la carpeta del
   portal web (o al repositorio, segun como quede desplegado).

Requiere: pip install openpyxl --break-system-packages
============================================================
"""

import json
import os
from datetime import datetime, timedelta
from openpyxl import load_workbook

# ------------------------------------------------------------------
# 1. RUTAS DE LOS ARCHIVOS
# ------------------------------------------------------------------
# Detecta automaticamente la carpeta de OneDrive de Smurfit, sin
# importar en que computador o con que usuario de Windows se corra.
# Windows guarda esa ruta en una de estas variables de entorno.
_ONEDRIVE = os.environ.get("OneDriveCommercial") or os.environ.get("OneDrive")

if _ONEDRIVE:
    RUTA_BACKLOG = os.path.join(_ONEDRIVE, "Backlog", "BACKLOG.xlsm")
    RUTA_PROGRAMACION = os.path.join(_ONEDRIVE, "Backlog", "PROGRAMACION.xlsx")
    CARPETA_DESPACHOS = os.path.join(_ONEDRIVE, "Notas despachos")
else:
    # Si por alguna razon Windows no expone esa variable, se puede
    # escribir la ruta completa a mano aqui como respaldo:
    RUTA_BACKLOG = r"C:\Users\TU_USUARIO\OneDrive - Smurfit Westrock\Backlog\BACKLOG.xlsm"
    RUTA_PROGRAMACION = r"C:\Users\TU_USUARIO\OneDrive - Smurfit Westrock\Backlog\PROGRAMACION.xlsx"
    CARPETA_DESPACHOS = r"C:\Users\TU_USUARIO\OneDrive - Smurfit Westrock\Notas despachos"

ARCHIVO_SALIDA = "datos.json"
ARCHIVO_HISTORICO_DESPACHOS = "despachos_historico.json"

# Un pedido ya despachado se sigue mostrando en el portal durante este
# numero de dias despues de su fecha de despacho, AUNQUE el backlog ya
# lo haya quitado de su lista. Pasado ese tiempo, deja de aparecer.
DIAS_VISIBLE_DESPACHADO = 15


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
        return "Listo"
    if "parcial" in texto:
        return "En producción"
    if texto == "terminado":
        return "Listo"
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
    Lee el archivo de Notas de Despacho (.xls) y arma un mapa
    'Ord. de Venta' -> fecha de despacho real.
    La Ord. de Venta se arma uniendo 'Order Number' + '-' + 'Order Line',
    exactamente igual a como aparece en el backlog (ej. '584722-14').
    """
    import xlrd
    wb = xlrd.open_workbook(ruta_despachos)
    ws = wb.sheet_by_name("Sheet1")
    encabezado = [limpiar_texto(ws.cell_value(0, c)) for c in range(ws.ncols)]

    idx_order_number = encabezado.index("Order Number")
    idx_order_line = encabezado.index("Order Line")
    idx_fecha = encabezado.index("Fecha del Despacho")

    mapa = {}
    for r in range(1, ws.nrows):
        try:
            order_number = int(ws.cell_value(r, idx_order_number))
            order_line = int(ws.cell_value(r, idx_order_line))
        except (ValueError, TypeError):
            continue
        ov = f"{order_number}-{order_line}"

        valor_fecha = ws.cell_value(r, idx_fecha)
        try:
            fecha = xlrd.xldate.xldate_as_datetime(valor_fecha, wb.datemode).strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            fecha = ""

        mapa[ov] = fecha
    return mapa


def generar_datos():
    print("Leyendo Order Capacity...")
    mapa_order_capacity = cargar_order_capacity(RUTA_BACKLOG)

    print("Leyendo Programacion...")
    mapa_programacion = cargar_programacion(RUTA_PROGRAMACION)

    print("Leyendo Notas de Despacho...")
    ruta_despachos = encontrar_archivo_mas_reciente(CARPETA_DESPACHOS, ".xls")
    print(f"  Archivo encontrado: {ruta_despachos}")
    mapa_despachos_nuevo = cargar_despachos(ruta_despachos)

    # --- Cargar el historial acumulado de corridas anteriores ---
    # Este historial guarda el REGISTRO COMPLETO de cada pedido que ya
    # se detecto como despachado (no solo la fecha), para poder seguirlo
    # mostrando en el portal aunque el backlog ya lo haya quitado de su
    # lista (ver DIAS_VISIBLE_DESPACHADO mas arriba).
    try:
        with open(ARCHIVO_HISTORICO_DESPACHOS, "r", encoding="utf-8") as f:
            historico = json.load(f)
    except FileNotFoundError:
        historico = {}

    # Compatibilidad con el formato viejo (OV -> "dd/mm/aaaa" en texto
    # plano, de antes de guardar el registro completo). Se convierte
    # a la nueva forma, aunque sin los demas datos del pedido.
    for ov, valor in list(historico.items()):
        if isinstance(valor, str):
            historico[ov] = {"fecha_despacho": valor}

    # Mapa simple OV -> fecha, combinando lo nuevo con el historial,
    # usado para decidir el estado de cada fila del backlog de hoy.
    mapa_despachos = {ov: v["fecha_despacho"] for ov, v in historico.items()}
    mapa_despachos.update(mapa_despachos_nuevo)

    print(f"Despachos nuevos en este archivo: {len(mapa_despachos_nuevo)}")
    print(f"Total acumulado en el historial: {len(mapa_despachos)}")

    print("Leyendo backlog (hoja Formato)...")
    wb = load_workbook(RUTA_BACKLOG, read_only=True, data_only=True)
    ws = wb["Formato"]
    filas = list(ws.iter_rows(values_only=True))
    encabezado = [limpiar_texto(c) for c in filas[0]]
    datos = filas[1:]

    idx_id_cliente = encabezado.index("ID Cliente")
    idx_elemento = encabezado.index("Elemento")
    idx_descripcion = encabezado.index("Descripción")
    idx_ord_compra = encabezado.index("Ord. de Compra")
    idx_ord_venta = encabezado.index("Ord. de Venta")
    idx_ord_trabajo = encabezado.index("Ord. de Trabajo")
    idx_estatus_omp = encabezado.index("Estatus OMP")

    resultado = []
    ov_cubiertas = set()
    total_leidas = 0
    total_ght = 0

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

        estado_ght = clasificar_estado_ght(estatus_omp_final)

        # --- Si ya existe un despacho real para esta Orden de Venta, ---
        # --- el estado final pasa a ser "Despachado", con su fecha real. ---
        ord_venta = limpiar_texto(fila[idx_ord_venta])
        fecha_despacho = mapa_despachos.get(ord_venta, "")
        if fecha_despacho:
            estado_ght = "Despachado"

        # No exponemos los pedidos "Sin clasificar" al cliente (decision ya tomada en el proyecto)
        if estado_ght == "Sin clasificar":
            continue

        registro = {
            "finca": id_cliente,
            "elemento": elemento,
            "descripcion": descripcion,
            "ow": ow_final,
            "orden_compra": orden_compra,
            "estado": estado_ght,
            "fecha_despacho": fecha_despacho,
            # Este campo queda vacio hasta que TI termine de ajustar el bot
            # de despachos para que incluya la fecha/hora ESTIMADA (antes
            # de que el pedido salga). La fecha_despacho de arriba es la
            # fecha REAL, ya confirmada, que si tenemos.
            "fecha_estimada_despacho": "",
        }

        if ord_venta:
            ov_cubiertas.add(ord_venta)
            if estado_ght == "Despachado":
                # Guardamos el registro completo en el historial, para
                # poder seguir mostrandolo aunque el backlog lo quite
                # de su lista mas adelante.
                historico[ord_venta] = registro

        resultado.append(registro)

    # --- Agregar despachos recientes que el backlog YA no muestra ---
    # Si un pedido se despacho hace poco (dentro de DIAS_VISIBLE_DESPACHADO)
    # pero hoy ya no aparece en el backlog (porque ya se cerro del todo),
    # lo agregamos igual usando el ultimo registro completo que se guardo
    # de el en el historial - asi no desaparece de golpe del portal.
    hoy = datetime.now().date()
    agregados_desde_historico = 0
    for ov, registro in historico.items():
        if ov in ov_cubiertas:
            continue  # ya viene del backlog de hoy, no lo dupliquemos
        if "finca" not in registro:
            continue  # registro viejo (formato antiguo), sin datos suficientes
        try:
            fecha_reg = datetime.strptime(registro["fecha_despacho"], "%d/%m/%Y").date()
        except (ValueError, TypeError):
            continue
        if (hoy - fecha_reg).days <= DIAS_VISIBLE_DESPACHADO:
            resultado.append(registro)
            agregados_desde_historico += 1

    with open(ARCHIVO_HISTORICO_DESPACHOS, "w", encoding="utf-8") as f:
        json.dump(historico, f, ensure_ascii=False, indent=2)

    with open(ARCHIVO_SALIDA, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    print()
    print(f"Total filas leidas en el backlog: {total_leidas}")
    print(f"Total filas de GHT (antes de quitar 'Sin clasificar'): {total_ght}")
    print(f"Pedidos despachados recuperados del historico (ya no estan en el backlog): {agregados_desde_historico}")
    print(f"Total filas exportadas a {ARCHIVO_SALIDA}: {len(resultado)}")
    print()
    print("Listo. Sube el archivo 'datos.json' al portal web.")


if __name__ == "__main__":
    generar_datos()
