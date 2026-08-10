"""ORCMM — detalle diario de un SKU en una tienda, leído de Postgres.

Hermano del `orcmm_expediente.py` de línea de comandos (mismo propósito:
cruzar todas las fuentes para un SKU), pero éste lee de Postgres y devuelve
JSON en vez de imprimir a consola — lo usa GET /api/expediente.

Reutiliza las mismas funciones que ya usa el motor para decidir "¿había
tránsito/pedido/orden vigente el día D?" (derivar_transito_vigente,
derivar_envio_generado, derivar_pedido_tienda, derivar_orden_proveedor) —
así la gráfica explica exactamente lo que dice el veredicto, no una
aproximación aparte.
"""
from __future__ import annotations

from datetime import date
from typing import Optional

from orcmm_fuentes_db import leer_fuentes_db
from orcmm_pipeline import (OrdenProveedor, _osa_pct, derivar_envio_generado,
                             derivar_evidencias, derivar_orden_proveedor, derivar_pedido_tienda,
                             derivar_transito_vigente)
from orcmm_rca_periodo import clasificar


def expediente_sku(tienda: str, sku: str, desde: date, hasta: date,
                    umbral_osa: float = 100.0, dsn: Optional[str] = None) -> dict:
    fu = leer_fuentes_db(tienda, desde, hasta, umbral_osa, dsn, sku=sku)
    evidencias = derivar_evidencias(fu, umbral_osa)
    diagnosticos = clasificar(evidencias)
    veredicto_por_dia = {d["fecha"]: d for d in diagnosticos}  # sólo días con faltante

    cat = fu.catalogo.get((sku, tienda), {})
    cedis = cat.get("cedis_surtidor")

    fechas = sorted({d for (_, _, d) in fu.osa} | {d for (_, _, d) in fu.inv_tienda})
    dias = []
    for fecha in fechas:
        osa_row = fu.osa.get((sku, tienda, fecha), {})
        inv_row = fu.inv_tienda.get((sku, tienda, fecha), {})
        venta_row = fu.ventas.get((sku, tienda, fecha), {})
        cedis_row = fu.inv_cedis.get((sku, cedis, fecha), {}) if cedis else {}
        veredicto = veredicto_por_dia.get(fecha.isoformat())
        orden = derivar_orden_proveedor(fu, sku, cedis, fecha) if cedis else OrdenProveedor()

        dias.append({
            "fecha": fecha.isoformat(),
            "osa_pct": _osa_pct(osa_row.get("osa_pct")),
            "existencia_tienda": inv_row.get("existencia_piezas"),
            "unidades_vendidas": venta_row.get("unidades_vendidas"),
            "existencia_cedis": cedis_row.get("existencia_piezas"),
            "pedido_tienda_abierto": derivar_pedido_tienda(fu, sku, tienda, fecha),
            "transito_vigente": derivar_transito_vigente(fu, sku, tienda, fecha),
            "envio_generado": derivar_envio_generado(fu, sku, tienda, fecha),
            "orden_proveedor_vigente": orden.existe,
            "cajas_pedidas_proveedor": orden.cajas_pedidas,
            "cajas_entregadas_proveedor": orden.cajas_entregadas,
            "root_cause_id": veredicto["root_cause_id"] if veredicto else None,
            "causa_raiz": veredicto["causa_raiz"] if veredicto else None,
            "responsable": veredicto["responsable"] if veredicto else None,
        })

    return {
        "sku": sku,
        "tienda": tienda,
        "descripcion": cat.get("descripcion"),
        "desde": desde.isoformat(),
        "hasta": hasta.isoformat(),
        "dias": dias,
    }
