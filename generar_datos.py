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
from openpyxl import load_workbook

# ------------------------------------------------------------------
# 1. RUTAS DE LOS ARCHIVOS (ajustar si cambian)
# ------------------------------------------------------------------
RUTA_BACKLOG = r"C:\Users\salceju\OneDrive - Smurfit Westrock\Backlog\BACKLOG.xlsm"
RUTA_PROGRAMACION = r"C:\Users\salceju\OneDrive - Smurfit Westrock\Backlog\PROGRAMACION.xlsx"
ARCHIVO_SALIDA = "datos.json"


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


def generar_datos():
    print("Leyendo Order Capacity...")
    mapa_order_capacity = cargar_order_capacity(RUTA_BACKLOG)

    print("Leyendo Programacion...")
    mapa_programacion = cargar_programacion(RUTA_PROGRAMACION)

    print("Leyendo backlog (hoja Formato)...")
    wb = load_workbook(RUTA_BACKLOG, read_only=True, data_only=True)
    ws = wb["Formato"]
    filas = list(ws.iter_rows(values_only=True))
    encabezado = [limpiar_texto(c) for c in filas[0]]
    datos = filas[1:]

    idx_id_cliente = encabezado.index("ID Cliente")
    idx_elemento = encabezado.index("Elemento")
    idx_ord_trabajo = encabezado.index("Ord. de Trabajo")
    idx_estatus_omp = encabezado.index("Estatus OMP")

    resultado = []
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

        # No exponemos los pedidos "Sin clasificar" al cliente (decision ya tomada en el proyecto)
        if estado_ght == "Sin clasificar":
            continue

        resultado.append({
            "finca": id_cliente,
            "elemento": elemento,
            "ow": ow_final,
            "estado": estado_ght,
            # Estos 2 campos quedan vacios hasta que TI ajuste el bot de despachos.
            # Cuando lleguen, se agregan aqui mismo sin tocar el resto del script.
            "orden_compra": "",
            "fecha_estimada_despacho": "",
        })

    with open(ARCHIVO_SALIDA, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    print()
    print(f"Total filas leidas en el backlog: {total_leidas}")
    print(f"Total filas de GHT (antes de quitar 'Sin clasificar'): {total_ght}")
    print(f"Total filas exportadas a {ARCHIVO_SALIDA}: {len(resultado)}")
    print()
    print("Listo. Sube el archivo 'datos.json' al portal web.")


if __name__ == "__main__":
    generar_datos()
