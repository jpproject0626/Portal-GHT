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
    Lee el archivo de Notas de Despacho (.xls) y devuelve una lista con
    cada LINEA de despacho individual (una por cada envio real de un
    producto puntual, incluyendo envios parciales). Cada linea trae:
      - clave: identificador UNICO de esa linea de despacho, formado por
        'Nota de despacho' + Orden de Venta. Se necesita esta combinacion
        (no solo la nota) porque una misma nota de despacho puede incluir
        varios productos/lineas distintas en un solo envio.
      - ov: 'Ord. de Venta' (Order Number + '-' + Order Line)
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

        try:
            cantidad = float(ws.cell_value(r, idx_cantidad))
        except (ValueError, TypeError):
            cantidad = 0.0

        valor_fecha = ws.cell_value(r, idx_fecha)
        try:
            fecha = xlrd.xldate.xldate_as_datetime(valor_fecha, wb.datemode).strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            fecha = ""

        # La llave unica es NOTA + OV, no la nota sola: una misma nota de
        # despacho puede incluir varias lineas/productos distintos en un
        # mismo envio (una guia con varios items).
        clave_linea = f"{nota}|{ov}"
        lineas.append({"clave": clave_linea, "ov": ov, "nota": nota, "cantidad": cantidad, "fecha": fecha})

    return lineas


def formatear_cantidad(valor):
    """Muestra 20000 en vez de 20000.0, pero conserva decimales si los hay de verdad."""
    if valor == int(valor):
        return str(int(valor))
    return str(round(valor, 2))


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

    notas_nuevas = 0
    for linea in lineas_nuevas:
        if linea["clave"] not in notas_historico:
            notas_nuevas += 1
        notas_historico[linea["clave"]] = {
            "ov": linea["ov"], "nota": linea["nota"], "cantidad": linea["cantidad"], "fecha": linea["fecha"],
        }

    # Acumular cantidad total despachada, fecha del ultimo envio, y la
    # lista de notas de despacho involucradas, por Orden de Venta.
    acumulado_por_ov = {}
    for linea_guardada in notas_historico.values():
        ov = linea_guardada["ov"]
        entrada = acumulado_por_ov.setdefault(ov, {"cantidad": 0.0, "fecha": None, "notas": set()})
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

    idx_id_cliente = encabezado.index("ID Cliente")
    idx_elemento = encabezado.index("Elemento")
    idx_descripcion = encabezado.index("Descripción")
    idx_ord_compra = encabezado.index("Ord. de Compra")
    idx_ord_venta = encabezado.index("Ord. de Venta")
    idx_ord_trabajo = encabezado.index("Ord. de Trabajo")
    idx_estatus_omp = encabezado.index("Estatus OMP")
    idx_cant_sol = encabezado.index("Cant Sol")

    def construir_registro(id_cliente, elemento, descripcion, ow_final, orden_compra, ov, cant_sol, estado_produccion):
        """
        Decide el estado final de un pedido comparando lo despachado
        (acumulado_por_ov) contra lo solicitado (cant_sol):
          - Nada despachado todavia -> se respeta el estado de produccion.
          - Se despacho TODO (o mas, por redondeos) -> "Despachado".
          - Se despacho una parte -> "Parcialmente despachado".
        """
        info_despacho = acumulado_por_ov.get(ov)
        cantidad_despachada = info_despacho["cantidad"] if info_despacho else 0.0
        fecha_ultimo = info_despacho["fecha"] if info_despacho else None
        notas_texto = ""
        if info_despacho and info_despacho["notas"]:
            notas_texto = ", ".join(sorted(info_despacho["notas"]))

        estado = estado_produccion
        fecha_despacho_txt = ""
        if cantidad_despachada > 0:
            if cant_sol and cantidad_despachada >= cant_sol:
                estado = "Despachado"
            else:
                estado = "Parcialmente despachado"
            if fecha_ultimo:
                fecha_despacho_txt = fecha_ultimo.strftime("%d/%m/%Y")

        return {
            "finca": id_cliente,
            "elemento": elemento,
            "descripcion": descripcion,
            "ow": ow_final,
            "orden_compra": orden_compra,
            "estado": estado,
            "fecha_despacho": fecha_despacho_txt,
            "cantidad_solicitada": formatear_cantidad(cant_sol) if cant_sol else "",
            "cantidad_despachada": formatear_cantidad(cantidad_despachada) if cantidad_despachada else "",
            "nota_despacho": notas_texto,
            # Este campo queda vacio hasta que TI termine de ajustar el bot
            # de despachos para que incluya la fecha/hora ESTIMADA (antes
            # de que el pedido salga).
            "fecha_estimada_despacho": "",
        }, fecha_ultimo

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
        ord_venta = limpiar_texto(fila[idx_ord_venta])

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

        registro, _ = construir_registro(
            id_cliente, elemento, descripcion, ow_final, orden_compra,
            ord_venta, cant_sol, estado_produccion,
        )

        # No exponemos los pedidos "Sin clasificar" al cliente (decision ya tomada en el proyecto)
        if registro["estado"] == "Sin clasificar":
            continue

        if ord_venta:
            ov_cubiertas.add(ord_venta)
            # Guardamos SIEMPRE la info mas reciente de este pedido (no solo
            # cuando ya se despacho), para poder recuperarla despues si el
            # backlog lo quita de su lista antes de completarse el despacho.
            pedidos_cache[ord_venta] = {
                "finca": id_cliente, "elemento": elemento, "descripcion": descripcion,
                "ow": ow_final, "orden_compra": orden_compra, "cant_sol": cant_sol,
                "estado_produccion": estado_produccion,
            }

        resultado.append(registro)

    # --- Agregar pedidos despachados/parciales recientes que el backlog YA no muestra ---
    hoy = datetime.now().date()
    agregados_desde_historico = 0
    for ov, info in pedidos_cache.items():
        if ov in ov_cubiertas:
            continue  # ya viene del backlog de hoy, no lo dupliquemos

        registro, fecha_ultimo = construir_registro(
            info["finca"], info["elemento"], info["descripcion"], info["ow"],
            info["orden_compra"], ov, info.get("cant_sol", 0), info.get("estado_produccion", "Sin clasificar"),
        )
        # Solo tiene sentido recuperarlo si de verdad hubo algun despacho
        # (si nunca se despacho nada, y ya no esta en el backlog, no
        # tenemos nada confiable que mostrar de el).
        if not fecha_ultimo:
            continue
        if (hoy - fecha_ultimo).days <= DIAS_VISIBLE_DESPACHADO:
            resultado.append(registro)
            agregados_desde_historico += 1

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
    print()
    print("Listo. Sube el archivo 'datos.json' al portal web.")


if __name__ == "__main__":
    generar_datos()
