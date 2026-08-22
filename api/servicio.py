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
from collections import defaultdict
from typing import Dict, List, Optional

RAIZ = Path(__file__).resolve().parent.parent
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from orcmm_corregir_layout import corregir                      # noqa: E402
from orcmm_layout_spec import CSV, HOJAS, origen_de             # noqa: E402
from orcmm_pipeline import (Fuentes, PaqueteFuentes, aviso_prioridad_3,  # noqa: E402
                            citas_incumplidas, derivar_evidencias,
                            desempeno_proveedores, discrepancias_pedido_cita,
                            escribir_resultado, fill_rate_proveedor, leer_fuentes,
                            nivel_servicio_tienda, osa_alcance, osa_general,
                            resumen_excluidos_sima,
                            universo_osa, waterfall_osa,
                            VIAS, clave_catalogo, es_dsd)
from orcmm_rca_engine import FUERA_DE_CATALOGO, ViaResurtido    # noqa: E402
from orcmm_rca_periodo import (clasificar, cobertura_modelo,    # noqa: E402
                               dentro_del_alcance, diagnosticar_periodo,
                               resumen_por_causa, resumen_por_responsable,
                               resumen_por_subcausa)
from orcmm_fuentes_db import _avisar                            # noqa: E402
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
    FUERA_DE_CATALOGO:
        "BOPS — SKU de una división que este análisis no cubre. BOPS entrega todas las divisiones y el catálogo sólo trae Abarrotes. No es un dato faltante: son días fuera del alcance",
}


# ===========================================================================
# 1. VALIDACIÓN
# ===========================================================================

def _a_dict(rep) -> dict:
    return {"errores": rep.errores, "faltan_datos": rep.faltan_datos,
            "advertencias": rep.advertencias, "ok": rep.info}


def diagnosticar_layout(ruta: Path, carpeta: Path, csvs=None) -> dict:
    """Valida el paquete y averigua si la corrección automática lo destraba.

    La única forma honesta de saber si el corrector sirve para ESTE archivo es
    correrlo sobre una copia y volver a validar. El archivo que subió el
    usuario no se toca.
    """
    csvs = list(csvs or [])
    reporte = validar_archivo(ruta, csvs)
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

    if not reporte.errores_corregibles:
        # Ninguno de los errores es de los que el corrector sabe arreglar
        # (celdas obligatorias vacías, cruces que no cuadran). Correrlo de
        # todas formas reescribiría un Excel de decenas de MB durante varios
        # minutos para llegar exactamente al mismo diagnóstico.
        resultado["motivo_no_corregible"] = (
            "Los errores que quedan no son de formato: son datos que faltan o que "
            "no cuadran entre hojas. La corrección automática no los puede inventar; "
            "hay que pedirle la reextracción al equipo dueño de la fuente.")
        return resultado

    corregido = carpeta / f"{ruta.stem} corregido.xlsx"
    try:
        cambios = corregir(ruta, corregido)
    except Exception as e:                       # archivo ilegible o roto
        resultado["validacion"]["errores"].append(
            f"No se pudo intentar la corrección automática: {e}")
        return resultado

    # La segunda pasada va SIN los CSV a propósito: el corrector sólo toca el
    # Excel, y volver a recorrer 2.7 millones de filas para obtener el mismo
    # veredicto duplicaría el tiempo de la validación.
    despues = validar_archivo(corregido)
    errores_csv = [e for e in reporte.errores
                   if e.split(":")[0] in {h for h in HOJAS if origen_de(h) == CSV}]
    resultado.update({
        "corregible": len(despues.errores) + len(errores_csv) < len(reporte.errores),
        "cambios_propuestos": cambios,
        "errores_tras_correccion": despues.errores + errores_csv,
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
        # Alcance: los días cuyo SKU sí está en el catálogo de la tienda. El
        # front encabeza con esta cobertura y deja la global como contraste.
        "casos_fuera_de_alcance": cob["casos_fuera_de_alcance"],
        "venta_perdida_fuera_de_alcance": cob["venta_perdida_fuera_de_alcance"],
        "casos_en_alcance": cob["casos_en_alcance"],
        "cobertura_casos_alcance_pct": cob["cobertura_casos_alcance_pct"],
        "venta_perdida_en_alcance": cob["venta_perdida_en_alcance"],
        "cobertura_venta_perdida_alcance_pct": cob["cobertura_venta_perdida_alcance_pct"],
        "bloqueos": [
            {"campo": campo, "dias": n,
             "venta_perdida": vp_campo.get(campo, 0.0),
             "a_quien": DE_QUIEN.get(campo, "Revisar matriz")}
            for campo, n in cob["campos_que_bloquean"].items()
        ],
    }


def _proveedores_a_dict(fu: Fuentes, umbral_osa: float = 100.0) -> List[dict]:
    return [{
        "proveedor_id": d.proveedor_id,
        # Cuando el mismo proveedor viene con varios IDs, aquí van todos.
        "ids": d.ids,
        "nombre": d.nombre,
        "pedidos": d.pedidos,
        "cajas_pedidas": d.cajas_pedidas,
        # Las que COMPRAS reporta como entregadas al cerrar el pedido. Es el
        # numerador del nivel de servicio: viene lleno en el 100% de los
        # pedidos, a diferencia de cajas_entregadas (de la cita), que sólo
        # existe para el 4.4% que llegó a agendar una.
        "cajas_surtidas": d.cajas_surtidas_pedido,
        "nivel_servicio": d.pct_surtido_pedido,
        "osa_periodo": d.osa_periodo,
        "dias_evaluados": d.dias_evaluados,
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
    } for d in desempeno_proveedores(fu, umbral_osa)]


class IndiceJerarquia:
    """Sección/categoría/subcategoría/marca y VÍA de cada SKU, comprimido.

    Va comprimido por necesidad, no por elegancia. Escribir los cuatro textos
    en cada renglón cuesta **1.3 MB** medidos sobre la tienda 287 —10,481
    SKU— encima de una respuesta que ya pesa 1.5 MB; con proveedor y formato
    eran 2.1 MB. Ese tamaño ya nos costó una vez el spinner infinito en
    Vercel, así que no se repite.

    En vez de eso se manda el catálogo de COMBINACIONES distintas —3,711 para
    esa misma tienda, porque miles de SKU comparten sección, categoría y
    marca— y cada renglón guarda su índice. Baja a 405 KB. Es el mismo truco
    que usa `_detalle_dias` con las causas.

    El formato de la tienda NO viaja: el análisis corre sobre una sola tienda
    (ver /api/analizar-tienda), así que sería la misma cadena repetida diez
    mil veces. El front lo saca de /api/tiendas si lo necesita.

    La VÍA va aquí y no en el renglón por la misma razón, y con una ventaja
    extra: los renglones del universo también indexan este catálogo, así que
    al filtrar por vía el denominador del OSA se recompone solo.

    Y es la vía que el MODELO usó, no la que dice el catálogo. Desde que los
    pedidos DSD se leen de COMPRAS, 330 SKU de Coyoacán que el catálogo marca
    "Vía 2" se clasifican por la rama directa. Ofrecer el filtro con la vía
    del catálogo dejaría al usuario eligiendo "Vía 2" y viendo los días
    resueltos por la rama de DSD, que se contradice solo.
    """

    def __init__(self, fu: Fuentes):
        self._fu = fu
        self.combos: List[list] = []
        self._indice: Dict[tuple, int] = {}

    def _via(self, sku: str, tienda: str) -> Optional[str]:
        if es_dsd(self._fu, sku, tienda):
            return ViaResurtido.DSD.value
        cat = self._fu.catalogo.get((sku, tienda)) or {}
        via = VIAS.get(clave_catalogo(cat.get("via_resurtido")))
        return via.value if via else None

    def de(self, sku: str, tienda: str) -> int:
        c = self._fu.comercial.get((sku, tienda)) or {}
        clave = (c.get("grupo_seccion"), c.get("categoria"),
                 c.get("subcategoria"), c.get("marca"),
                 self._via(sku, tienda))
        if clave not in self._indice:
            self._indice[clave] = len(self.combos)
            self.combos.append(list(clave))
        return self._indice[clave]


def _detalle_dias(fu: Fuentes, en_alcance: List[dict], jer: IndiceJerarquia) -> dict:
    """El detalle día por día, comprimido, para que el front pueda recalcular
    el waterfall y los Pareto cuando el usuario filtra.

    Sin esto no se puede: `por_sku_tienda` sólo trae la causa DOMINANTE de
    cada SKU —la que ganó en sus días—, así que un SKU con RC01 unos días y
    RC06 otros se ve como "100% RC01". Recalcular el Pareto desde ahí daría
    números equivocados, no sólo desactualizados.

    Va comprimido a propósito. En vez de repetir "Ejecución en Tienda" y
    "Tienda" en cada uno de los ~5,200 renglones, las causas van en un
    catálogo aparte y cada día guarda su índice. Quedan cuatro campos por
    día con nombres de una letra: pesa ~200 KB en vez de ~700 KB.

    `universo` es el denominador del waterfall —las filas de BOPS del
    alcance, no sólo los días con faltante— y va desglosado por SKU-tienda
    para que al filtrar se pueda recomponer el denominador correcto. Sin él,
    filtrar a un SKU dejaría los puntos de OSA calculados sobre el universo
    de la tienda entera, que es un número sin sentido.

    Cada renglón del universo lleva su `j` de jerarquía porque esta lista SÍ
    puede traer SKU que no están en `por_sku_tienda` —los que no tuvieron ni
    un día con faltante—, y sin el índice esos quedarían sin sección: al
    filtrar por categoría se caerían del denominador y el OSA saldría
    hundido. Los días (`dias`) no lo necesitan: salen de la misma lista
    `en_alcance` que `por_sku_tienda`, así que ahí siempre hay renglón.
    """
    universo: Dict[tuple, int] = defaultdict(int)
    for (sku, tienda, _) in fu.osa:
        if fu.en_alcance(sku, tienda):
            universo[(sku, tienda)] += 1

    causas: List[dict] = []
    indice: Dict[tuple, int] = {}
    filas = []
    for dg in en_alcance:
        # La subcausa entra en la LLAVE del catálogo, no en cada día: son un
        # puñado de combinaciones distintas contra decenas de miles de
        # renglones. Va aquí para que el front pueda recalcular el desglose
        # fino al filtrar, igual que ya hace con los dos Pareto.
        clave = (dg["root_cause_id"], dg["causa_raiz"], dg["responsable"],
                 dg.get("subcausa"))
        if clave not in indice:
            indice[clave] = len(causas)
            causas.append({"root_cause_id": clave[0], "causa": clave[1],
                           "responsable": clave[2], "subcausa": clave[3]})
        filas.append({
            "s": dg["sku"],
            "t": dg["tienda"],
            "c": indice[clave],
            "v": round(dg["venta_perdida"] or 0.0, 2),
        })

    return {
        "causas": causas,
        "dias": filas,
        "universo": [{"s": s, "t": t, "n": n, "j": jer.de(s, t)}
                     for (s, t), n in universo.items()],
    }


def escribir_excel(fu: Fuentes, salida: Path, umbral_osa: float = 100.0,
                   avisar=None) -> bool:
    """Escribe el Excel de resultados a partir de un `Fuentes` ya leído.

    Existe para poder sacar el Excel de la corrida. Medido sobre Coyoacán
    marzo: leer las fuentes son 54 s, clasificar 1 s y escribir el Excel
    ~287 s. O sea que el 82% de cada análisis se iba en generar un archivo de
    16 MB que muchas veces nadie abre, mientras la pantalla esperaba.

    Rehace la evidencia y la clasificación en vez de arrastrarlas desde
    `analizar`: es un segundo sobre 44 mil evidencias, y no vale la pena
    mantener 44 mil objetos vivos entre las dos llamadas para ahorrarlo.

    Devuelve False si no había nada que escribir.
    """
    evidencias = derivar_evidencias(fu, umbral_osa)
    if not evidencias:
        return False
    _avisar(avisar, "generando el Excel de resultados")
    escribir_resultado(salida, fu, evidencias, clasificar(evidencias), umbral_osa)
    return True


def analizar(ruta: Optional[Path], salida: Path, umbral_osa: float = 100.0,
             csvs=None, fu: Optional[Fuentes] = None, avisar=None,
             con_excel: bool = True) -> dict:
    """Corre el pipeline y devuelve el resumen que ve la pantalla.

    Si `fu` ya viene armado (p. ej. desde orcmm_fuentes_db.leer_fuentes_db,
    con datos de Postgres en vez de un archivo) se usa tal cual y `ruta`/
    `csvs` se ignoran — todo lo de aquí en adelante sólo lee de `Fuentes`.

    `con_excel=False` devuelve el resumen sin escribir el archivo, para que la
    pantalla no espere los ~287 s que cuesta. Quien lo pase tiene que llamar
    después a escribir_excel() si quiere la descarga.
    """
    fu = fu or leer_fuentes(PaqueteFuentes.desde(ruta, csvs or []), umbral_osa)
    _avisar(avisar, "derivando la evidencia de cada día")
    evidencias = derivar_evidencias(fu, umbral_osa)

    if not evidencias:
        return {
            "hay_resultados": False,
            "osa_alcance": osa_alcance(fu),
            "osa_general": osa_general(fu),
            "motivo": ("No hay días con faltante que analizar. Revisar BOPS_OSA: o viene "
                       "vacía, o todos los días traen OSA al 100%."),
            "fuentes": _fuentes_a_dict(fu),
            "advertencias": fu.advertencias,
            "aviso_parcial": aviso_prioridad_3(fu),
        }

    _avisar(avisar, "clasificando por causa raíz")
    diagnosticos = clasificar(evidencias)
    # La cobertura se calcula aquí y no se recibe del escritor del Excel: eran
    # el mismo número, pero pedírselo a él ataba el resumen a que el archivo
    # se generara, que es justo lo que con_excel viene a desatar.
    cob = cobertura_modelo(diagnosticos)
    if con_excel:
        _avisar(avisar, "generando el Excel de resultados")
        escribir_resultado(salida, fu, evidencias, diagnosticos, umbral_osa)

    # El Pareto y el detalle por SKU van sobre el alcance; la cobertura sigue
    # contando todo y reporta los días fuera de catálogo aparte. Mismo criterio
    # que el Excel, para que la pantalla y el archivo digan lo mismo.
    en_alcance = dentro_del_alcance(diagnosticos)
    jer = IndiceJerarquia(fu)

    por_sku = [{
        "sku": d.sku,
        "tienda": d.tienda,
        # Para poder buscar por nombre además de por código en el front —
        # no se usa para clasificar, sólo se arrastra al output.
        "descripcion": fu.catalogo.get((d.sku, d.tienda), {}).get("descripcion"),
        # Índice al catálogo `jerarquia` de abajo: sección, categoría,
        # subcategoría y marca. Viaja sólo para filtrar y leer el reporte,
        # ninguna regla del motor lo consume. Cuando el análisis corre por
        # archivo apunta a un combo de puros nulos, porque el catálogo
        # comercial sólo vive en Postgres.
        "j": jer.de(d.sku, d.tienda),
        "dias_con_faltante": d.dias_con_faltante,
        "dias_clasificados": d.dias_clasificados,
        "cobertura_pct": d.cobertura_pct,
        "venta_perdida": d.venta_perdida_total,
        # osa_promedio se conserva por compatibilidad, pero NO se muestra: da
        # 0 siempre (promedia sólo los días con faltante, que valen 0 por
        # definición). El bueno es osa_periodo, sobre los días evaluados.
        "osa_promedio": d.osa_promedio,
        "dias_evaluados": d.dias_evaluados,
        "osa_periodo": d.osa_periodo,
        "root_cause_id": d.root_cause_id_dominante,
        "causa": d.causa_dominante,
        "responsable": d.responsable_dominante,
    } for d in diagnosticar_periodo(en_alcance, universo_osa(fu, umbral_osa))]

    return {
        "hay_resultados": True,
        # Dos OSA, igual que las dos coberturas: el del alcance es el número de
        # portada, el general queda de contraste y mide la extracción.
        "osa_alcance": osa_alcance(fu),
        "osa_general": osa_general(fu),
        "waterfall": waterfall_osa(fu, diagnosticos),
        "fill_rate_proveedor": fill_rate_proveedor(fu),
        # La otra pata de la cadena: proveedor→CEDIS arriba, CEDIS→tienda aquí.
        "nivel_servicio_tienda": nivel_servicio_tienda(fu),
        # El letrero: cuántos SKU se dejaron fuera por no tener datos de
        # SIMA. Va SIEMPRE, aunque sea cero, para que el front no tenga
        # que adivinar si la exclusión está activa.
        "excluidos_sin_sima": resumen_excluidos_sima(fu, umbral_osa),
        "umbral_osa": umbral_osa,
        "aviso_parcial": aviso_prioridad_3(fu),
        "advertencias": fu.advertencias,
        "fuentes": _fuentes_a_dict(fu),
        "cobertura": _cobertura_a_dict(cob),
        "por_causa": resumen_por_causa(en_alcance),
        "por_responsable": resumen_por_responsable(en_alcance),
        "por_subcausa": resumen_por_subcausa(en_alcance),
        "por_sku_tienda": por_sku,
        "detalle_dias": _detalle_dias(fu, en_alcance, jer),
        # El catálogo que resuelve la `j` de cada renglón: una lista de
        # [sección, categoría, subcategoría, marca, vía]. Va después de las
        # dos listas que lo indexan porque se llena mientras se arman.
        "jerarquia": jer.combos,
        "proveedores": _proveedores_a_dict(fu, umbral_osa),
        "citas_falladas": [{
            **f, "fecha_cita": f["fecha_cita"].isoformat() if f["fecha_cita"] else None
        } for f in citas_incumplidas(fu)],
        "discrepancias": discrepancias_pedido_cita(fu),
    }
