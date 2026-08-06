"""
ORCMM - Capa de servicio entre la web y los módulos de análisis.

Aquí no vive ninguna regla de negocio: todo se delega a los módulos que ya
existen y que se corren igual desde la línea de comandos.

    orcmm_validar_layout   ¿el archivo cumple el layout?
    orcmm_corregir_layout  arreglar lo que se puede arreglar solo
    orcmm_pipeline         leer, derivar y clasificar
    orcmm_rca_periodo      agregar los veredictos diarios

Lo único que se agrega es la traducción a JSON, para que la pantalla muestre
exactamente los mismos números que el Excel de resultados.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import List, Optional

RAIZ = Path(__file__).resolve().parent.parent
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from orcmm_corregir_layout import corregir                      # noqa: E402
from orcmm_layout_spec import HOJAS                             # noqa: E402
from orcmm_pipeline import (Fuentes, aviso_prioridad_3, citas_incumplidas,  # noqa: E402
                            derivar_evidencias, desempeno_proveedores,
                            discrepancias_pedido_cita, escribir_resultado,
                            leer_fuentes, osa_general)
from orcmm_rca_periodo import (clasificar, cobertura_modelo,    # noqa: E402
                               diagnosticar_periodo, resumen_por_causa,
                               resumen_por_responsable, resumen_por_subcausa)
from orcmm_validar_layout import validar_archivo                # noqa: E402

# A quién pedirle el dato que bloqueó la clasificación. Misma tabla que usa la
# hoja "Cobertura y fuentes" del Excel.
DE_QUIEN = {
    "inventario_tienda": "Tableau — inventario en tienda",
    "transito_vigente": "CEDIS / Logística — transferencias",
    "envio_cedis_generado": "CEDIS / Logística — transferencias",
    "pedido_tienda_generado": "SIMA — pedidos de tienda",
    "via_resurtido": "Catálogo — vía de resurtido del SKU",
    "inventario_cedis": "CEDIS — inventario diario",
    "pedido_proveedor_generado": "Compras — pedidos a proveedor",
    "proveedor_cajas_pedidas": "Compras — pedidos a proveedor",
    "proveedor_cajas_entregadas": "Compras — citas de proveedor",
    "proveedor_cajas_confirmadas_cita": "Compras — citas de proveedor",
    "regla_matriz_para_cita_de_proveedor_no_vencida":
        "Decisión de negocio — la matriz no cubre el faltante previo a la cita",
    "regla_matriz_para_entrega_completa_con_cedis_en_cero":
        "Decisión de negocio — la matriz no cubre CEDIS en cero con entrega completa",
}


# ===========================================================================
# 1. VALIDACIÓN
# ===========================================================================

def _a_dict(rep) -> dict:
    return {"errores": rep.errores, "faltan_datos": rep.faltan_datos,
            "advertencias": rep.advertencias, "ok": rep.info}


def diagnosticar_layout(ruta: Path, carpeta: Path) -> dict:
    """Valida el archivo y averigua si la corrección automática lo destraba.

    La única forma honesta de saber si el corrector sirve para ESTE archivo es
    correrlo sobre una copia y volver a validar. El archivo que subió el
    usuario no se toca.
    """
    reporte = validar_archivo(ruta)
    resultado = {
        "valido": not reporte.errores,
        "validacion": _a_dict(reporte),
        "corregible": False,
        "cambios_propuestos": [],
        "errores_tras_correccion": [],
        "ruta_corregida": None,
    }
    if not reporte.errores:
        return resultado

    corregido = carpeta / f"{ruta.stem} corregido.xlsx"
    try:
        cambios = corregir(ruta, corregido)
    except Exception as e:                       # archivo ilegible o roto
        resultado["validacion"]["errores"].append(
            f"No se pudo intentar la corrección automática: {e}")
        return resultado

    despues = validar_archivo(corregido)
    resultado.update({
        "corregible": len(despues.errores) < len(reporte.errores),
        "cambios_propuestos": cambios,
        "errores_tras_correccion": despues.errores,
        "ruta_corregida": str(corregido),
    })
    return resultado


# ===========================================================================
# 2. ANÁLISIS
# ===========================================================================

def _fuentes_a_dict(fu: Fuentes) -> List[dict]:
    filas = []
    for hoja in HOJAS:
        d0, d1 = fu.rango.get(hoja, (None, None))
        filas.append({
            "hoja": hoja,
            "filas": fu.conteo.get(hoja, 0),
            "desde": d0.isoformat() if d0 else None,
            "hasta": d1.isoformat() if d1 else None,
            "equipo": HOJAS[hoja]["equipo"],
            "owner": HOJAS[hoja]["owner"],
        })
    return filas


def _cobertura_a_dict(cob: dict) -> dict:
    vp_campo = cob["venta_perdida_por_campo_faltante"]
    return {
        "casos_totales": cob["casos_totales"],
        "casos_clasificados": cob["casos_clasificados"],
        "cobertura_casos_pct": cob["cobertura_casos_pct"],
        "venta_perdida_total": cob["venta_perdida_total"],
        "venta_perdida_clasificada": cob["venta_perdida_clasificada"],
        "cobertura_venta_perdida_pct": cob["cobertura_venta_perdida_pct"],
        "bloqueos": [
            {"campo": campo, "dias": n,
             "venta_perdida": vp_campo.get(campo, 0.0),
             "a_quien": DE_QUIEN.get(campo, "Revisar matriz")}
            for campo, n in cob["campos_que_bloquean"].items()
        ],
    }


def _proveedores_a_dict(fu: Fuentes) -> List[dict]:
    return [{
        "proveedor_id": d.proveedor_id,
        "nombre": d.nombre,
        "pedidos": d.pedidos,
        "cajas_pedidas": d.cajas_pedidas,
        "pct_surtido_pedido": d.pct_surtido_pedido,
        "citas": d.citas,
        "pedidos_sin_cita": d.pedidos_sin_cita,
        "cajas_pedidas_con_cita": d.cajas_pedidas_con_cita,
        "cajas_confirmadas": d.cajas_confirmadas,
        "cajas_entregadas": d.cajas_entregadas,
        "pct_confirmado": d.pct_confirmado,
        "pct_cumplimiento": d.pct_cumplimiento,
        "pct_efectivo": d.pct_efectivo,
        "citas_incumplidas": d.citas_incumplidas,
    } for d in desempeno_proveedores(fu)]


def analizar(ruta: Path, salida: Path, umbral_osa: float = 100.0) -> dict:
    """Corre el pipeline completo y devuelve el resumen que ve la pantalla.

    Es la misma pasada que produce el Excel: se lee una vez, se clasifica una
    vez y de ahí salen las dos salidas.
    """
    fu = leer_fuentes(ruta)
    evidencias = derivar_evidencias(fu, umbral_osa)

    if not evidencias:
        return {
            "hay_resultados": False,
            "osa_general": osa_general(fu),
            "motivo": ("No hay días con faltante que analizar. Revisar BOPS_OSA: o viene "
                       "vacía, o todos los días traen OSA al 100%."),
            "fuentes": _fuentes_a_dict(fu),
            "advertencias": fu.advertencias,
            "aviso_parcial": aviso_prioridad_3(fu),
        }

    diagnosticos = clasificar(evidencias)
    cob, _ = escribir_resultado(salida, fu, evidencias, diagnosticos)

    por_sku = [{
        "sku": d.sku,
        "tienda": d.tienda,
        "dias_con_faltante": d.dias_con_faltante,
        "dias_clasificados": d.dias_clasificados,
        "cobertura_pct": d.cobertura_pct,
        "venta_perdida": d.venta_perdida_total,
        "osa_promedio": d.osa_promedio,
        "root_cause_id": d.root_cause_id_dominante,
        "causa": d.causa_dominante,
        "responsable": d.responsable_dominante,
    } for d in diagnosticar_periodo(diagnosticos)]

    return {
        "hay_resultados": True,
        "osa_general": osa_general(fu),
        "umbral_osa": umbral_osa,
        "aviso_parcial": aviso_prioridad_3(fu),
        "advertencias": fu.advertencias,
        "fuentes": _fuentes_a_dict(fu),
        "cobertura": _cobertura_a_dict(cob),
        "por_causa": resumen_por_causa(diagnosticos),
        "por_responsable": resumen_por_responsable(diagnosticos),
        "por_subcausa": resumen_por_subcausa(diagnosticos),
        "por_sku_tienda": por_sku,
        "proveedores": _proveedores_a_dict(fu),
        "citas_falladas": [{
            **f, "fecha_cita": f["fecha_cita"].isoformat() if f["fecha_cita"] else None
        } for f in citas_incumplidas(fu)],
        "discrepancias": discrepancias_pedido_cita(fu),
    }
