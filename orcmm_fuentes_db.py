"""ORCMM — arma un Fuentes leyendo Postgres en vez de un Excel/CSV.

Mismo contrato de salida que leer_fuentes(): una vez armado el Fuentes, el
resto del pipeline (derivar_evidencias, clasificar, escribir_resultado) no
sabe ni le importa de dónde salió — sólo lee de los diccionarios indexados
del objeto. Aquí sólo se arma esa estructura con SELECT en vez de openpyxl.
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import Optional

from psycopg2.extras import RealDictCursor

from orcmm_db import columnas_de, conectar
from orcmm_layout_spec import ORIGEN_CENTRALIZADO
from orcmm_pipeline import (Fuentes, _indexar_eventos, _osa_pct, _revisar_cobertura_de_citas,
                             _revisar_cobertura_de_sima, _texto, aviso_prioridad_3)

# Los 4 eventos (transferencias, pedidos de tienda, pedidos a proveedor,
# citas) se extraen ~un mes antes de la ventana de análisis en el flujo de
# Excel — ver HOJAS["CEDIS_TRANSFERENCIAS"]["ventana"] en
# orcmm_layout_spec.py: "un tránsito de febrero explica faltantes del 1 de
# marzo". Mismo criterio aquí: sin este margen, un pedido u orden generada
# antes de `desde` pero todavía vigente el día D no se encontraría.
LOOKBACK_EVENTOS_DIAS = 30

# Columnas que el análisis realmente consume, por hoja.
#
# Antes se traía SELECT *. Medido sobre la corrida real de Coyoacán: ~1 millón
# de filas, de las cuales 826 mil son de BOPS_OSA. Cada fila se materializa
# como un dict de psycopg2 y cuesta ~1.5 KB, así que leer_fuentes_db sola se
# lleva ~1.9 GB — el 78% de la memoria del proceso (medido por etapas: la
# lectura sube de 59 a 429 MB en una ventana de una semana; clasificar suma 3
# MB). Con eso Coyoacán rebasa los 4 GB de la máquina de Fly y la mata.
#
# Las que no están aquí NO las lee nadie: se verificó una por una contra todo
# lo que consume Fuentes —no sólo el motor, también el scorecard de
# proveedores, los reportes de citas, el expediente y el Excel—. Ojo con los
# nombres que se repiten entre tablas: `tienda_destino` sólo se lee de
# transferencias, `cedis_destino` sólo de pedidos a proveedor, y `estatus` de
# ninguna (lo que se lee es `estatus_cita`).
#
# Se incluyen las fechas que usa _registrar para el rango de cada fuente aunque
# ninguna regla las lea: de ahí sale el panel "Fuentes" del front. CATALOGO se
# queda con SELECT * a propósito: son 26 mil filas (2% del total), es la hoja
# con más consumidores repartidos, y recortarla agregaría riesgo sin mover la
# aguja.
COLUMNAS = {
    "BOPS_OSA": ["sku", "tienda", "fecha", "osa_pct", "venta_perdida_estimada",
                 "alerta_enviada", "alerta_ejecutada"],
    "TABLEAU_INV_TIENDA": ["sku", "tienda", "fecha", "existencia_piezas",
                           "existencia_minima_dia"],
    "TABLEAU_VENTAS": ["sku", "tienda", "fecha", "importe_venta", "unidades_vendidas",
                       "venta_perdida_estimada"],
    "CEDIS_INVENTARIO": ["sku", "cedis", "fecha", "existencia_piezas", "piezas_reservadas"],
    "CEDIS_TRANSFERENCIAS": ["folio", "sku", "tienda_destino", "fecha_generacion",
                             "fecha_salida_cedis", "fecha_recepcion_tienda"],
    # Las cantidades son el insumo de nivel_servicio_tienda, no del motor.
    "SIMA_PEDIDOS_TIENDA": ["folio", "sku", "origen", "fecha_pedido", "fecha_surtido",
                            "cantidad_pedida_piezas", "cantidad_surtida_piezas"],
    "COMPRAS_PEDIDOS_PROV": ["folio", "sku", "proveedor_id", "proveedor_nombre",
                             "cedis_destino", "fecha_pedido", "fecha_cita", "fecha_recibo",
                             "cajas_pedidas", "cajas_entregadas"],
    "CITAS_PROV_CEDIS": ["folio", "folio_cita", "sku", "proveedor_id", "proveedor_nombre",
                         "fecha_pedido", "fecha_cita", "cajas_confirmadas_cita",
                         "cajas_entregadas", "estatus_cita"],
}


def _cols(hoja: str, alias: str = "") -> str:
    """Las columnas de la hoja para el SELECT, con alias opcional de tabla.

    Valida contra el spec: si el layout renombra o quita una columna, esto
    truena aquí y no varias capas más abajo con un None silencioso.
    """
    validas = {c for _, c, _ in columnas_de(hoja)}
    desconocidas = [c for c in COLUMNAS[hoja] if c not in validas]
    if desconocidas:
        raise RuntimeError(
            f"{hoja}: COLUMNAS pide {desconocidas}, que no existen en el layout. "
            f"Si el spec cambió, hay que actualizar orcmm_fuentes_db.COLUMNAS.")
    p = f"{alias}." if alias else ""
    return ", ".join(f"{p}{c}" for c in COLUMNAS[hoja])



def _registrar(fu: Fuentes, hoja: str, filas: list, campos_fecha) -> None:
    """Igual que el `registrar` interno de leer_fuentes: conteo y rango
    salen de lo efectivamente leído, no de los límites de la consulta — así
    el bloque 'Fuentes que no llegaron completas' del front sigue diciendo
    la verdad."""
    if isinstance(campos_fecha, str):
        campos_fecha = [campos_fecha]
    fu.conteo[hoja] = len(filas)
    fechas = [f[c] for f in filas for c in campos_fecha if f.get(c)]
    fu.rango[hoja] = (min(fechas), max(fechas)) if fechas else (None, None)
    if not filas:
        fu.advertencias.append(f"{hoja}: sin filas para esta tienda y periodo.")


def leer_fuentes_db(tienda: str, desde: date, hasta: date,
                     umbral_osa: float = 100.0, dsn: Optional[str] = None,
                     sku: Optional[str] = None) -> Fuentes:
    """Con `sku`, todas las consultas se acotan a ese SKU además de la
    tienda — lo usa orcmm_expediente_db para el detalle diario de un solo
    producto. Con un solo SKU el volumen es mínimo, así que
    TABLEAU_INV_TIENDA/TABLEAU_VENTAS traen TODOS los días del rango (no
    sólo los de faltante): el detalle diario necesita también los días
    sanos, a diferencia del análisis de la tienda completa."""
    fu = Fuentes()
    desde_eventos = desde - timedelta(days=LOOKBACK_EVENTOS_DIAS)

    conn = conectar(dsn)
    try:
        with conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
            # CEDIS que surte a la tienda. cedis_inventario, compras_pedidos_prov
            # y citas_prov_cedis están scoped por cedis, no por tienda.
            sql = "SELECT DISTINCT cedis_surtidor FROM catalogo WHERE tienda = %s"
            params = [tienda]
            if sku:
                sql += " AND sku = %s"
                params.append(sku)
            cur.execute(sql, params)
            cedis_ids = [r["cedis_surtidor"] for r in cur.fetchall() if r["cedis_surtidor"]]

            # 1. CATALOGO — estático, sin fecha.
            sql = "SELECT * FROM catalogo WHERE tienda = %s"
            params = [tienda]
            if sku:
                sql += " AND sku = %s"
                params.append(sku)
            cur.execute(sql, params)
            filas = cur.fetchall()
            for f in filas:
                fu.catalogo[(_texto(f["sku"]), _texto(f["tienda"]))] = f
            fu.conteo["CATALOGO"] = len(filas)
            fu.rango["CATALOGO"] = (None, None)

            # 2. BOPS_OSA — define qué días entran al análisis. Se lee ANTES
            #    de inv_tienda/ventas a propósito: de aquí sale la lista de
            #    (sku, fecha) con faltante real, que es lo único que hace
            #    falta pedirle a esas dos tablas cuando no hay `sku` (ver 3).
            sql = (f"SELECT {_cols('BOPS_OSA')} FROM bops_osa "
                   "WHERE tienda = %s AND fecha BETWEEN %s AND %s")
            params = [tienda, desde, hasta]
            if sku:
                sql += " AND sku = %s"
                params.append(sku)
            cur.execute(sql, params)
            filas = cur.fetchall()
            for f in filas:
                fu.osa[(_texto(f["sku"]), _texto(f["tienda"]), f["fecha"])] = f
            _registrar(fu, "BOPS_OSA", filas, "fecha")

            skus_faltante, fechas_faltante = [], []
            for f in filas:
                osa = _osa_pct(f["osa_pct"])
                if osa is not None and osa < umbral_osa:
                    skus_faltante.append(f["sku"])
                    fechas_faltante.append(f["fecha"])

            # 3. TABLEAU_INV_TIENDA / TABLEAU_VENTAS.
            if sku:
                # Un solo SKU: volumen mínimo, se trae el rango completo
                # (también los días sanos, que el detalle diario necesita).
                cur.execute(f"SELECT {_cols('TABLEAU_INV_TIENDA')} FROM tableau_inv_tienda "
                            "WHERE tienda = %s AND sku = %s AND fecha BETWEEN %s AND %s",
                            (tienda, sku, desde, hasta))
                filas = cur.fetchall()
            else:
                # Tienda completa: mismo criterio que el flujo de CSV
                # (orcmm_fuentes_csv.llaves_con_faltante) — traer sólo los
                # (sku, fecha) que BOPS_OSA ya marcó con faltante, no el
                # inventario/venta de TODOS los días de TODOS los SKU. Sin
                # esto se trae prácticamente la tabla entera — medido:
                # 2.66M de 2.72M filas, 97 de 119s de una corrida completa.
                filas = []
                if skus_faltante:
                    cur.execute(
                        f"SELECT {_cols('TABLEAU_INV_TIENDA', 't')} FROM tableau_inv_tienda t "
                        "JOIN unnest(%s::text[], %s::date[]) AS k(sku, fecha) "
                        "  ON t.sku = k.sku AND t.fecha = k.fecha "
                        "WHERE t.tienda = %s",
                        (skus_faltante, fechas_faltante, tienda))
                    filas = cur.fetchall()
            for f in filas:
                fu.inv_tienda[(_texto(f["sku"]), _texto(f["tienda"]), f["fecha"])] = f
            _registrar(fu, "TABLEAU_INV_TIENDA", filas, "fecha")

            if sku:
                cur.execute(f"SELECT {_cols('TABLEAU_VENTAS')} FROM tableau_ventas "
                            "WHERE tienda = %s AND sku = %s AND fecha BETWEEN %s AND %s",
                            (tienda, sku, desde, hasta))
                filas = cur.fetchall()
            else:
                filas = []
                if skus_faltante:
                    cur.execute(
                        f"SELECT {_cols('TABLEAU_VENTAS', 't')} FROM tableau_ventas t "
                        "JOIN unnest(%s::text[], %s::date[]) AS k(sku, fecha) "
                        "  ON t.sku = k.sku AND t.fecha = k.fecha "
                        "WHERE t.tienda = %s",
                        (skus_faltante, fechas_faltante, tienda))
                    filas = cur.fetchall()
            for f in filas:
                fu.ventas[(_texto(f["sku"]), _texto(f["tienda"]), f["fecha"])] = f
            _registrar(fu, "TABLEAU_VENTAS", filas, "fecha")

            # 5. CEDIS_INVENTARIO — scoped por cedis, no por tienda.
            filas = []
            if cedis_ids:
                sql = (f"SELECT {_cols('CEDIS_INVENTARIO')} FROM cedis_inventario "
                       "WHERE cedis = ANY(%s) AND fecha BETWEEN %s AND %s")
                params = [cedis_ids, desde, hasta]
                if sku:
                    sql += " AND sku = %s"
                    params.append(sku)
                cur.execute(sql, params)
                filas = cur.fetchall()
            for f in filas:
                fu.inv_cedis[(_texto(f["sku"]), _texto(f["cedis"]), f["fecha"])] = f
            _registrar(fu, "CEDIS_INVENTARIO", filas, "fecha")
            for (_, cedis, d) in fu.inv_cedis:
                fu.dias_cedis.setdefault(cedis, set()).add(d)

            # 6. CEDIS_TRANSFERENCIAS — tienda_destino directo, con lookback.
            sql = (f"SELECT {_cols('CEDIS_TRANSFERENCIAS')} FROM cedis_transferencias "
                   "WHERE tienda_destino = %s AND fecha_generacion BETWEEN %s AND %s")
            params = [tienda, desde_eventos, hasta]
            if sku:
                sql += " AND sku = %s"
                params.append(sku)
            cur.execute(sql, params)
            fu.transferencias = cur.fetchall()
            _registrar(fu, "CEDIS_TRANSFERENCIAS", fu.transferencias,
                       ["fecha_generacion", "fecha_salida_cedis", "fecha_recepcion_tienda"])

            # 7. SIMA_PEDIDOS_TIENDA — con lookback. `origen` es quién generó
            #    el pedido: se traen los de ESTA tienda y también los
            #    centralizados, que resurten a la tienda aunque no los haya
            #    generado ella (ver derivar_pedido_tienda).
            sql = (f"SELECT {_cols('SIMA_PEDIDOS_TIENDA')} FROM sima_pedidos_tienda "
                   "WHERE origen = ANY(%s) AND fecha_pedido BETWEEN %s AND %s")
            params = [[tienda, ORIGEN_CENTRALIZADO], desde_eventos, hasta]
            if sku:
                sql += " AND sku = %s"
                params.append(sku)
            cur.execute(sql, params)
            fu.pedidos_tienda = cur.fetchall()
            _registrar(fu, "SIMA_PEDIDOS_TIENDA", fu.pedidos_tienda,
                       ["fecha_pedido", "fecha_surtido"])

            # 8. COMPRAS_PEDIDOS_PROV — cedis_destino, no tienda, con lookback.
            fu.pedidos_prov = []
            if cedis_ids:
                sql = (f"SELECT {_cols('COMPRAS_PEDIDOS_PROV')} FROM compras_pedidos_prov "
                       "WHERE cedis_destino = ANY(%s) AND fecha_pedido BETWEEN %s AND %s")
                params = [cedis_ids, desde_eventos, hasta]
                if sku:
                    sql += " AND sku = %s"
                    params.append(sku)
                cur.execute(sql, params)
                fu.pedidos_prov = cur.fetchall()
            _registrar(fu, "COMPRAS_PEDIDOS_PROV", fu.pedidos_prov,
                       ["fecha_pedido", "fecha_recibo"])

            # 9. CITAS_PROV_CEDIS — cedis_destino, no tienda, con lookback.
            fu.citas_prov = []
            if cedis_ids:
                sql = (f"SELECT {_cols('CITAS_PROV_CEDIS')} FROM citas_prov_cedis "
                       "WHERE cedis_destino = ANY(%s) AND fecha_pedido BETWEEN %s AND %s")
                params = [cedis_ids, desde_eventos, hasta]
                if sku:
                    sql += " AND sku = %s"
                    params.append(sku)
                cur.execute(sql, params)
                fu.citas_prov = cur.fetchall()
            _registrar(fu, "CITAS_PROV_CEDIS", fu.citas_prov, ["fecha_pedido", "fecha_cita"])
    finally:
        conn.close()

    for c in fu.citas_prov:
        fu.citas_por_pedido.setdefault(
            (_texto(c.get("folio")), _texto(c.get("sku"))), []).append(c)

    _indexar_eventos(fu)
    _revisar_cobertura_de_citas(fu)
    _revisar_cobertura_de_sima(fu)

    aviso = aviso_prioridad_3(fu)
    if aviso:
        fu.advertencias.insert(0, aviso)

    return fu
