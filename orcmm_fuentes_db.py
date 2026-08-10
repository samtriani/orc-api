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

from orcmm_db import conectar
from orcmm_pipeline import (Fuentes, _indexar_eventos, _osa_pct, _revisar_cobertura_de_citas,
                             _texto, aviso_prioridad_3)

# Los 4 eventos (transferencias, pedidos de tienda, pedidos a proveedor,
# citas) se extraen ~un mes antes de la ventana de análisis en el flujo de
# Excel — ver HOJAS["CEDIS_TRANSFERENCIAS"]["ventana"] en
# orcmm_layout_spec.py: "un tránsito de febrero explica faltantes del 1 de
# marzo". Mismo criterio aquí: sin este margen, un pedido u orden generada
# antes de `desde` pero todavía vigente el día D no se encontraría.
LOOKBACK_EVENTOS_DIAS = 30


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
                     umbral_osa: float = 100.0, dsn: Optional[str] = None) -> Fuentes:
    fu = Fuentes()
    desde_eventos = desde - timedelta(days=LOOKBACK_EVENTOS_DIAS)

    conn = conectar(dsn)
    try:
        with conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
            # CEDIS que surte a la tienda. cedis_inventario, compras_pedidos_prov
            # y citas_prov_cedis están scoped por cedis, no por tienda.
            cur.execute("SELECT DISTINCT cedis_surtidor FROM catalogo WHERE tienda = %s",
                        (tienda,))
            cedis_ids = [r["cedis_surtidor"] for r in cur.fetchall() if r["cedis_surtidor"]]

            # 1. CATALOGO — estático, sin fecha.
            cur.execute("SELECT * FROM catalogo WHERE tienda = %s", (tienda,))
            filas = cur.fetchall()
            for f in filas:
                fu.catalogo[(_texto(f["sku"]), _texto(f["tienda"]))] = f
            fu.conteo["CATALOGO"] = len(filas)
            fu.rango["CATALOGO"] = (None, None)

            # 2. BOPS_OSA — define qué días entran al análisis. Se lee ANTES
            #    de inv_tienda/ventas a propósito: de aquí sale la lista de
            #    (sku, fecha) con faltante real, que es lo único que hace
            #    falta pedirle a esas dos tablas (ver punto 3).
            cur.execute("SELECT * FROM bops_osa "
                        "WHERE tienda = %s AND fecha BETWEEN %s AND %s",
                        (tienda, desde, hasta))
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

            # 3. TABLEAU_INV_TIENDA / TABLEAU_VENTAS — mismo criterio que el
            #    flujo de CSV (orcmm_fuentes_csv.llaves_con_faltante): traer
            #    sólo los (sku, fecha) que BOPS_OSA ya marcó con faltante, no
            #    el inventario/venta de TODOS los días de TODOS los SKU. Sin
            #    esto se trae prácticamente la tabla entera — medido: 2.66M
            #    de 2.72M filas, 97 de 119s de una corrida completa.
            filas = []
            if skus_faltante:
                cur.execute(
                    "SELECT t.* FROM tableau_inv_tienda t "
                    "JOIN unnest(%s::text[], %s::date[]) AS k(sku, fecha) "
                    "  ON t.sku = k.sku AND t.fecha = k.fecha "
                    "WHERE t.tienda = %s",
                    (skus_faltante, fechas_faltante, tienda))
                filas = cur.fetchall()
            for f in filas:
                fu.inv_tienda[(_texto(f["sku"]), _texto(f["tienda"]), f["fecha"])] = f
            _registrar(fu, "TABLEAU_INV_TIENDA", filas, "fecha")

            filas = []
            if skus_faltante:
                cur.execute(
                    "SELECT t.* FROM tableau_ventas t "
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
                cur.execute("SELECT * FROM cedis_inventario "
                            "WHERE cedis = ANY(%s) AND fecha BETWEEN %s AND %s",
                            (cedis_ids, desde, hasta))
                filas = cur.fetchall()
            for f in filas:
                fu.inv_cedis[(_texto(f["sku"]), _texto(f["cedis"]), f["fecha"])] = f
            _registrar(fu, "CEDIS_INVENTARIO", filas, "fecha")
            for (_, cedis, d) in fu.inv_cedis:
                fu.dias_cedis.setdefault(cedis, set()).add(d)

            # 6. CEDIS_TRANSFERENCIAS — tienda_destino directo, con lookback.
            cur.execute("SELECT * FROM cedis_transferencias "
                        "WHERE tienda_destino = %s AND fecha_generacion BETWEEN %s AND %s",
                        (tienda, desde_eventos, hasta))
            fu.transferencias = cur.fetchall()
            _registrar(fu, "CEDIS_TRANSFERENCIAS", fu.transferencias,
                       ["fecha_generacion", "fecha_salida_cedis", "fecha_recepcion_tienda"])

            # 7. SIMA_PEDIDOS_TIENDA — tienda directo, con lookback (0 filas hoy).
            cur.execute("SELECT * FROM sima_pedidos_tienda "
                        "WHERE tienda = %s AND fecha_pedido BETWEEN %s AND %s",
                        (tienda, desde_eventos, hasta))
            fu.pedidos_tienda = cur.fetchall()
            _registrar(fu, "SIMA_PEDIDOS_TIENDA", fu.pedidos_tienda,
                       ["fecha_pedido", "fecha_surtido"])

            # 8. COMPRAS_PEDIDOS_PROV — cedis_destino, no tienda, con lookback.
            fu.pedidos_prov = []
            if cedis_ids:
                cur.execute("SELECT * FROM compras_pedidos_prov "
                            "WHERE cedis_destino = ANY(%s) AND fecha_pedido BETWEEN %s AND %s",
                            (cedis_ids, desde_eventos, hasta))
                fu.pedidos_prov = cur.fetchall()
            _registrar(fu, "COMPRAS_PEDIDOS_PROV", fu.pedidos_prov,
                       ["fecha_pedido", "fecha_recibo"])

            # 9. CITAS_PROV_CEDIS — cedis_destino, no tienda, con lookback.
            fu.citas_prov = []
            if cedis_ids:
                cur.execute("SELECT * FROM citas_prov_cedis "
                            "WHERE cedis_destino = ANY(%s) AND fecha_pedido BETWEEN %s AND %s",
                            (cedis_ids, desde_eventos, hasta))
                fu.citas_prov = cur.fetchall()
            _registrar(fu, "CITAS_PROV_CEDIS", fu.citas_prov, ["fecha_pedido", "fecha_cita"])
    finally:
        conn.close()

    for c in fu.citas_prov:
        fu.citas_por_pedido.setdefault(
            (_texto(c.get("folio")), _texto(c.get("sku"))), []).append(c)

    _indexar_eventos(fu)
    _revisar_cobertura_de_citas(fu)

    aviso = aviso_prioridad_3(fu)
    if aviso:
        fu.advertencias.insert(0, aviso)

    return fu
