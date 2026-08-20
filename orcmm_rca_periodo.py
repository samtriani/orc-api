"""
ORCMM - Agregación de diagnósticos RCA por periodo (ej. Marzo 2026)

El motor (orcmm_rca_engine.py) clasifica UN día a la vez, porque la evidencia
de inventario / tránsito / pedido cambia diariamente. Este módulo toma esos
veredictos diarios y los agrega en dos capas:

  1. Por SKU + Tienda: causa dominante del periodo, ponderada por venta perdida
  2. Global: Pareto de causas y responsables de todo el periodo

Todas las funciones trabajan sobre la lista de diagnósticos YA calculada, no
sobre las evidencias. El motor se corre una sola vez, con clasificar(), y su
resultado se reutiliza. Con el alcance real (abarrotes × 5 tiendas × 31 días)
la diferencia entre una pasada y tres es material.

Ambos Paretos se construyen desde el desglose diario, no desde la causa
dominante, para que causa y responsable reconcilien contra el mismo total.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Optional, Tuple

from orcmm_rca_engine import (FUERA_DE_CATALOGO, SIN_DATO_SIMA, EvidenciaSKUTienda, MotorRCA,
                              ViaResurtido)


# ---------------------------------------------------------------------------
# 0. LA ÚNICA PASADA DEL MOTOR
# ---------------------------------------------------------------------------

def clasificar(evidencias: List[EvidenciaSKUTienda]) -> List[dict]:
    """Corre la matriz sobre cada evidencia diaria. Una sola vez.

    Todo lo demás en este módulo consume el resultado de esta función.
    """
    motor = MotorRCA()
    return [motor.diagnosticar(ev) for ev in evidencias]


# ---------------------------------------------------------------------------
# 1. DIAGNÓSTICO POR SKU + TIENDA EN EL PERIODO
# ---------------------------------------------------------------------------

@dataclass
class DiagnosticoPeriodoSKUTienda:
    sku: str
    tienda: str
    dias_con_faltante: int
    dias_clasificados: int
    venta_perdida_total: float
    osa_promedio: Optional[float]
    causa_dominante: str
    root_cause_id_dominante: str
    responsable_dominante: str
    desglose_causas: Dict[str, float]        # causa -> venta perdida atribuida
    desglose_responsables: Dict[str, float]  # responsable -> venta perdida
    # OSA del SKU en TODO el periodo, no sólo en sus días malos. Requiere el
    # universo (ver universo_osa en orcmm_pipeline); sin él quedan en None.
    dias_evaluados: Optional[int] = None
    osa_periodo: Optional[float] = None

    @property
    def cobertura_pct(self) -> float:
        if not self.dias_con_faltante:
            return 0.0
        return round(self.dias_clasificados / self.dias_con_faltante * 100, 1)


def dentro_del_alcance(diagnosticos: List[dict]) -> List[dict]:
    """Quita los días que no le tocan al análisis. Hay DOS motivos:

      - el SKU no está en el catálogo de la tienda (FUERA_DE_CATALOGO): es
        una división que este análisis no cubre.
      - SIMA no trae ningún pedido del SKU (SIN_DATO_SIMA): sin ese dato la
        prioridad 3 no se puede contestar.

    Se usa para el Pareto y el detalle por SKU, que responden "¿de qué se
    está muriendo el negocio?". Meter ahí cualquiera de los dos inventa un
    bloque gigante que no es una causa raíz y esconde el Pareto real debajo.

    Los dos se conservan en la clasificación diaria, cada uno con su motivo,
    que es donde se auditan. Y el segundo se reporta aparte —cuántos SKU y
    cuántos días— porque a diferencia del primero, ése SÍ debía tener dato:
    ver resumen_excluidos_sima y EXCLUIR_SKU_SIN_SIMA.
    """
    return [dg for dg in diagnosticos
            if FUERA_DE_CATALOGO not in dg["datos_faltantes"]
            and SIN_DATO_SIMA not in dg["datos_faltantes"]]


def diagnosticar_periodo(diagnosticos: List[dict],
                          universo: Optional[Dict[Tuple[str, str], Tuple[int, int]]] = None
                          ) -> List[DiagnosticoPeriodoSKUTienda]:
    """Agrupa los veredictos diarios por SKU + Tienda.

    Recibe la salida de clasificar(); no vuelve a correr el motor.

    `universo` —(sku, tienda) -> (días medidos, días visibles)— sirve para el
    OSA del periodo. Va aparte porque aquí sólo llegan los días CON faltante,
    y el OSA de verdad necesita también los días sanos, que se quedaron en
    Fuentes. Lo arma universo_osa() en orcmm_pipeline. Sin él, osa_periodo
    queda en None y no se inventa nada.
    """
    grupos: Dict[Tuple[str, str], List[dict]] = defaultdict(list)
    for dg in diagnosticos:
        grupos[(dg["sku"], dg["tienda"])].append(dg)

    resultados: List[DiagnosticoPeriodoSKUTienda] = []

    for (sku, tienda), diags in grupos.items():
        vp_por_causa: Dict[str, float] = defaultdict(float)
        vp_por_responsable: Dict[str, float] = defaultdict(float)
        id_por_causa: Dict[str, str] = {}
        resp_por_causa: Dict[str, str] = {}
        osas: List[float] = []
        venta_perdida_total = 0.0
        dias_clasificados = 0

        for dg in diags:
            vp = dg["venta_perdida"] or 0.0
            causa = dg["causa_raiz"]

            vp_por_causa[causa] += vp
            vp_por_responsable[dg["responsable"]] += vp
            id_por_causa[causa] = dg["root_cause_id"]
            resp_por_causa[causa] = dg["responsable"]

            venta_perdida_total += vp
            dias_clasificados += int(dg["clasificado"])
            if dg["osa"] is not None:
                osas.append(dg["osa"])

        # Causa dominante = la que más venta perdida acumuló. Si no se registró
        # venta perdida, se decide por número de días con esa causa.
        if venta_perdida_total > 0:
            causa_dominante = max(vp_por_causa, key=vp_por_causa.get)
        else:
            causa_dominante = Counter(dg["causa_raiz"] for dg in diags).most_common(1)[0][0]

        medidos, visibles = (universo or {}).get((sku, tienda), (None, None))

        resultados.append(DiagnosticoPeriodoSKUTienda(
            sku=sku,
            tienda=tienda,
            dias_evaluados=medidos,
            # OSA real: días visibles sobre días medidos. NO se usa osa_promedio
            # para esto — ese promedia sólo los días con faltante, y como BOPS
            # reporta OSA binario (0 = no visible), esos días valen 0 por
            # definición: daba exactamente 0 para todos los SKU, siempre.
            osa_periodo=(round(visibles / medidos * 100, 1)
                         if medidos else None),
            dias_con_faltante=len(diags),
            dias_clasificados=dias_clasificados,
            venta_perdida_total=round(venta_perdida_total, 2),
            osa_promedio=round(sum(osas) / len(osas), 1) if osas else None,
            causa_dominante=causa_dominante,
            root_cause_id_dominante=id_por_causa[causa_dominante],
            responsable_dominante=resp_por_causa[causa_dominante],
            desglose_causas=dict(vp_por_causa),
            desglose_responsables=dict(vp_por_responsable),
        ))

    # Lo más caro primero: es el orden en el que se actúa
    resultados.sort(key=lambda d: d.venta_perdida_total, reverse=True)
    return resultados


# ---------------------------------------------------------------------------
# 2. COBERTURA DEL MODELO — qué tanto del desabasto quedó explicado
# ---------------------------------------------------------------------------

def cobertura_modelo(diagnosticos: List[dict]) -> dict:
    """Diagnóstico del diagnóstico: % de casos y de venta perdida que la
    matriz alcanzó a clasificar, y qué campos faltantes lo impidieron.

    Es el insumo para priorizar la integración de fuentes: el campo que más
    se repite es el que más caro está saliendo.

    Se reportan DOS coberturas, porque hay dos preguntas distintas:

      sobre el alcance  — de los días que al modelo le tocaba explicar,
                          ¿cuántos explicó? Es la que mide al modelo.
      global            — sobre todo lo que entregó BOPS, incluidos los SKU
                          que no están en el catálogo de la tienda. Es la que
                          mide la calidad de la extracción.

    Sin separarlas, un export de OSA que trae divisiones fuera del alcance
    hunde la cifra y parece una falla del modelo. Ver FUERA_DE_CATALOGO.
    """
    total_casos = 0
    casos_clasificados = 0
    fuera_casos = 0
    vp_total = 0.0
    vp_clasificada = 0.0
    vp_fuera = 0.0
    bloqueos: Counter = Counter()
    vp_bloqueada: Dict[str, float] = defaultdict(float)

    for dg in diagnosticos:
        vp = dg["venta_perdida"] or 0.0
        total_casos += 1
        vp_total += vp

        if dg["clasificado"]:
            casos_clasificados += 1
            vp_clasificada += vp
            continue

        # Los dos motivos de exclusión salen del alcance. Si sólo se restara
        # el de catálogo, la cobertura se calcularía sobre un denominador que
        # incluye días que el Pareto ya no cuenta, y los dos números dejarían
        # de hablar del mismo universo.
        if (FUERA_DE_CATALOGO in dg["datos_faltantes"]
                or SIN_DATO_SIMA in dg["datos_faltantes"]):
            fuera_casos += 1
            vp_fuera += vp

        for campo in dg["datos_faltantes"]:
            bloqueos[campo] += 1
            vp_bloqueada[campo] += vp

    def pct(parte: float, todo: float) -> float:
        return round(parte / todo * 100, 1) if todo else 0.0

    casos_alcance = total_casos - fuera_casos
    vp_alcance = vp_total - vp_fuera

    return {
        "casos_totales": total_casos,
        "casos_clasificados": casos_clasificados,
        "cobertura_casos_pct": pct(casos_clasificados, total_casos),
        "venta_perdida_total": round(vp_total, 2),
        "venta_perdida_clasificada": round(vp_clasificada, 2),
        "cobertura_venta_perdida_pct": pct(vp_clasificada, vp_total),
        # --- alcance: los días cuyo SKU sí está en el catálogo de la tienda
        "casos_fuera_de_alcance": fuera_casos,
        "venta_perdida_fuera_de_alcance": round(vp_fuera, 2),
        "casos_en_alcance": casos_alcance,
        "cobertura_casos_alcance_pct": pct(casos_clasificados, casos_alcance),
        "venta_perdida_en_alcance": round(vp_alcance, 2),
        "cobertura_venta_perdida_alcance_pct": pct(vp_clasificada, vp_alcance),
        "campos_que_bloquean": dict(bloqueos.most_common()),
        "venta_perdida_por_campo_faltante": {
            k: round(v, 2) for k, v in
            sorted(vp_bloqueada.items(), key=lambda kv: kv[1], reverse=True)
        },
    }


# ---------------------------------------------------------------------------
# 2b. PARETO EN FORMA DE TABLA
#
# pareto_periodo() da porcentajes; para pintar una tabla hace falta además el
# número de días, los pesos y el responsable. Vive aquí y no en quien escribe
# el Excel para que la hoja de resultados y la pantalla web muestren el mismo
# número: son dos vistas del mismo cálculo, no dos cálculos parecidos.
# ---------------------------------------------------------------------------

def _porcentajes(vp: Dict[str, float]) -> Dict[str, float]:
    total = sum(vp.values())
    return {k: round(v / total * 100, 1) if total else 0.0 for k, v in vp.items()}


def resumen_por_causa(diagnosticos: List[dict]) -> List[dict]:
    """Una fila por causa raíz, de la más cara a la más barata."""
    dias: Counter = Counter()
    vp: Dict[str, float] = defaultdict(float)
    responsable: Dict[str, str] = {}
    rc_id: Dict[str, str] = {}

    for dg in diagnosticos:
        causa = dg["causa_raiz"]
        dias[causa] += 1
        vp[causa] += dg["venta_perdida"] or 0.0
        responsable[causa] = dg["responsable"]
        rc_id[causa] = dg["root_cause_id"]

    pct = _porcentajes(vp)
    filas = [{"root_cause_id": rc_id[c], "causa": c, "dias": dias[c],
              "venta_perdida": round(vp[c], 2), "pct": pct[c],
              "responsable": responsable[c]} for c in dias]
    filas.sort(key=lambda f: (-f["venta_perdida"], -f["dias"]))
    return filas


def resumen_por_responsable(diagnosticos: List[dict]) -> List[dict]:
    """Una fila por responsable. Suma el mismo total que resumen_por_causa."""
    dias: Counter = Counter()
    vp: Dict[str, float] = defaultdict(float)

    for dg in diagnosticos:
        resp = dg["responsable"]
        dias[resp] += 1
        vp[resp] += dg["venta_perdida"] or 0.0

    pct = _porcentajes(vp)
    filas = [{"responsable": r, "dias": dias[r], "venta_perdida": round(vp[r], 2),
              "pct": pct[r]} for r in dias]
    filas.sort(key=lambda f: (-f["venta_perdida"], -f["dias"]))
    return filas


def resumen_por_subcausa(diagnosticos: List[dict]) -> List[dict]:
    """Detalle de las prioridades 3 y 8: qué falló y cuánto costó."""
    dias: Counter = Counter()
    vp: Dict[str, float] = defaultdict(float)

    for dg in diagnosticos:
        sub = dg.get("subcausa")
        if not sub:
            continue
        dias[sub] += 1
        vp[sub] += dg["venta_perdida"] or 0.0

    filas = [{"subcausa": s, "dias": dias[s], "venta_perdida": round(vp[s], 2)}
             for s in dias]
    filas.sort(key=lambda f: (-f["venta_perdida"], -f["dias"]))
    return filas


# ---------------------------------------------------------------------------
# 3. PARETO GLOBAL DEL PERIODO
# ---------------------------------------------------------------------------

def pareto_periodo(diagnosticos_periodo: List[DiagnosticoPeriodoSKUTienda]) -> dict:
    """Pareto de causas y responsables, ponderado por venta perdida.
    Ambos cortes suman 100% del mismo total."""

    venta_total = sum(d.venta_perdida_total for d in diagnosticos_periodo)

    por_causa: Dict[str, float] = defaultdict(float)
    por_responsable: Dict[str, float] = defaultdict(float)

    for d in diagnosticos_periodo:
        for causa, vp in d.desglose_causas.items():
            por_causa[causa] += vp
        for responsable, vp in d.desglose_responsables.items():
            por_responsable[responsable] += vp

    def a_porcentaje(dic: Dict[str, float]) -> Dict[str, float]:
        if venta_total == 0:
            return {k: 0.0 for k in dic}
        return {k: round(v / venta_total * 100, 1) for k, v in
                sorted(dic.items(), key=lambda kv: kv[1], reverse=True)}

    osa_valores = [d.osa_promedio for d in diagnosticos_periodo if d.osa_promedio is not None]

    return {
        "venta_perdida_total": round(venta_total, 2),
        "osa_promedio_periodo": round(sum(osa_valores) / len(osa_valores), 1) if osa_valores else None,
        "sku_tienda_analizados": len(diagnosticos_periodo),
        "dias_con_faltante": sum(d.dias_con_faltante for d in diagnosticos_periodo),
        "dias_clasificados": sum(d.dias_clasificados for d in diagnosticos_periodo),
        "pareto_por_responsable": a_porcentaje(por_responsable),
        "pareto_por_causa": a_porcentaje(por_causa),
    }


# ---------------------------------------------------------------------------
# 4. EJEMPLO
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import json

    evidencias_abril = [
        # Papel Higiénico / Tienda 287 — proveedor no surtió, 2 días
        EvidenciaSKUTienda(
            sku="7501059236776", tienda="287", fecha=date(2026, 4, 5),
            osa=24.1, venta_perdida=1200,
            inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=True,
            via_resurtido=ViaResurtido.VIA_1, inventario_cedis=0,
            pedido_proveedor_generado=True,
            proveedor_cajas_pedidas=40, proveedor_cajas_entregadas=0),
        EvidenciaSKUTienda(
            sku="7501059236776", tienda="287", fecha=date(2026, 4, 6),
            osa=22.0, venta_perdida=1356,
            inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=True,
            via_resurtido=ViaResurtido.VIA_1, inventario_cedis=0,
            pedido_proveedor_generado=True,
            proveedor_cajas_pedidas=40, proveedor_cajas_entregadas=0),

        # Café Soluble / Tienda 287 — CEDIS con stock y sin envío
        EvidenciaSKUTienda(
            sku="7501234567890", tienda="287", fecha=date(2026, 4, 10),
            osa=60.0, venta_perdida=800,
            inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=True,
            via_resurtido=ViaResurtido.VIA_1, inventario_cedis=663,
            envio_cedis_generado=False),

        # Tienda 512 — producto en piso, no en anaquel
        EvidenciaSKUTienda(
            sku="7501111222333", tienda="512", fecha=date(2026, 4, 12),
            osa=70.0, venta_perdida=450,
            inventario_tienda=15),

        # Tienda 317 — sin dato de tránsito: no se clasifica
        EvidenciaSKUTienda(
            sku="7500000000000", tienda="317", fecha=date(2026, 4, 18),
            osa=45.0, venta_perdida=610,
            inventario_tienda=0),
    ]

    diagnosticos = clasificar(evidencias_abril)          # una sola pasada
    por_sku_tienda = diagnosticar_periodo(diagnosticos)

    print("=== Diagnóstico por SKU + Tienda ===")
    for d in por_sku_tienda:
        print(f"  SKU {d.sku} / Tienda {d.tienda}: "
              f"[{d.root_cause_id_dominante}] {d.causa_dominante} "
              f"— ${d.venta_perdida_total:,.0f} "
              f"({d.dias_clasificados}/{d.dias_con_faltante} días, "
              f"cobertura {d.cobertura_pct}%)")

    print("\n=== Cobertura del modelo ===")
    print(json.dumps(cobertura_modelo(diagnosticos), indent=2, ensure_ascii=False))

    print("\n=== Pareto Abril 2026 ===")
    print(json.dumps(pareto_periodo(por_sku_tienda), indent=2, ensure_ascii=False))
