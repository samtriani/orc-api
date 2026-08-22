"""ORCMM — guardar y leer el resultado de cada análisis.

Hasta ahora una corrida vivía en memoria del proceso y moría con él: volver a
ver el análisis de la semana pasada obligaba a recalcularlo, y una corrida
completa de Coyoacán tarda ~5.7 minutos. Aquí se persiste el mismo resumen
que ya devuelve la API, para que la pantalla liste lo corrido y el detalle se
pinte de inmediato.

Nada de esto es crítico para clasificar: si la tabla no existe o la escritura
falla, el análisis ya terminó y su resultado sigue disponible en memoria. Por
eso `guardar` traga sus errores y sólo deja una advertencia — perder el
histórico es molesto, tumbar una corrida de seis minutos por no poder
escribirlo sería mucho peor.
"""
from __future__ import annotations

import json
import os
import subprocess
from types import SimpleNamespace
from datetime import date
from typing import List, Optional

from psycopg2.extras import Json, RealDictCursor, execute_values

from orcmm_db import conectar

# Los interruptores de negocio que estaban puestos al correr. Van con cada
# run por la misma razón que la versión del motor: son decisiones, no
# constantes, y cambian lo que dictamina el árbol.
INTERRUPTORES = (
    ("orcmm_rca_engine", "EVALUAR_PEDIDO_TIENDA"),
    ("orcmm_rca_engine", "EXCLUIR_SKU_SIN_SIMA"),
    ("orcmm_rca_engine", "CLASIFICAR_CITA_PENDIENTE"),
    ("orcmm_rca_engine", "CLASIFICAR_ENTREGA_COMPLETA_CEDIS_CERO"),
    ("orcmm_rca_engine", "REFINAR_RC01_CON_ALERTA"),
    ("orcmm_pipeline", "CEDIS_AUSENCIA_ES_CERO"),
    ("orcmm_pipeline", "INVENTARIO_CIERRE_NO_CONFIABLE"),
)

_version: Optional[str] = None


def version_motor() -> str:
    """El commit con el que corre este proceso.

    Se pregunta a git una sola vez y se recuerda: es la misma respuesta toda
    la vida del proceso. En un contenedor sin repo cae a la variable de
    entorno que inyecta el despliegue, y si tampoco está, a "desconocida" —
    que es feo pero honesto, y mejor que inventar un valor.
    """
    global _version
    if _version is not None:
        return _version
    try:
        sha = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=os.path.dirname(os.path.abspath(__file__)),
            capture_output=True, text=True, timeout=5,
        ).stdout.strip()
    except Exception:
        sha = ""
    _version = sha or os.getenv("ORCMM_VERSION") or "desconocida"
    return _version


def parametros() -> dict:
    """Los interruptores de negocio vigentes, leídos de los módulos."""
    import importlib
    puestos = {}
    for modulo, nombre in INTERRUPTORES:
        try:
            puestos[nombre] = getattr(importlib.import_module(modulo), nombre)
        except Exception:
            pass
    return puestos


def _portada(resumen: dict) -> dict:
    """Las cifras que la pantalla inicial necesita para pintar un renglón.

    Se desnormalizan a propósito: listar veinte corridas leyendo veinte
    resúmenes de 5 MB para sacar cuatro números sería absurdo.
    """
    cob = resumen.get("cobertura") or {}
    wf = resumen.get("waterfall") or {}
    return {
        "osa_alcance": resumen.get("osa_alcance"),
        "dias_faltante": cob.get("casos_en_alcance"),
        "venta_perdida": cob.get("venta_perdida_en_alcance"),
        "cobertura_pct": cob.get("cobertura_casos_alcance_pct"),
        # `wf` no se guarda aparte: ya va dentro del resumen. Se lee aquí sólo
        # para caer a su universo cuando la cobertura no trajo los casos.
        "_universo": wf.get("universo_filas"),
    }


def guardar(id_: str, tienda: str, desde: date, hasta: date, resumen: dict,
            umbral_osa: float = 100.0, segundos: Optional[float] = None,
            origen: str = "bd", archivo: Optional[str] = None,
            universo=None, dsn: Optional[str] = None) -> Optional[str]:
    """Guarda una corrida. Devuelve el aviso si algo falló, o None si todo bien.

    No levanta: ver el docstring del módulo.
    """
    p = _portada(resumen)
    try:
        conn = conectar(dsn)
        try:
            with conn, conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO runs (id, tienda, desde, hasta, umbral_osa,
                                      version_motor, parametros, osa_alcance,
                                      dias_faltante, venta_perdida, cobertura_pct,
                                      resumen, segundos, origen, archivo, universo)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (id) DO UPDATE SET
                        resumen = EXCLUDED.resumen,
                        universo = EXCLUDED.universo,
                        corrido_en = now()
                    """,
                    (id_, tienda, desde, hasta, umbral_osa,
                     version_motor(), Json(parametros()), p["osa_alcance"],
                     p["dias_faltante"], p["venta_perdida"], p["cobertura_pct"],
                     Json(json.loads(json.dumps(resumen, default=str))),
                     segundos, origen, archivo,
                     # Llave plana porque JSON no admite tuplas de llave.
                     Json({f"{k[0]}|{k[1]}": list(v) for k, v in (universo or {}).items()})),
                )
        finally:
            conn.close()
        return None
    except Exception as e:
        return (f"El análisis terminó bien, pero no se pudo guardar en el histórico "
                f"({e}). El resultado de esta pantalla es válido; lo que no va a "
                f"quedar es el registro para volver a consultarlo después.")


# Las 24 columnas de la hoja "Clasificación diaria", en su orden. Es el
# contrato entre lo que se guarda y lo que se vuelve a escribir: si cambia el
# orden aquí, cambia en las dos puntas o el Excel sale corrido.
COLUMNAS_DIA = (
    "sku", "tienda", "fecha", "osa", "venta_perdida",
    "inventario_tienda", "transito_vigente", "pedido_tienda_generado",
    "tipo_resurtido", "via_resurtido", "inventario_cedis", "envio_cedis_generado",
    "pedido_proveedor_generado", "cajas_pedidas",
    "cita_agendada", "cita_fecha", "cita_vencida",
    "cajas_confirmadas", "cajas_entregadas",
    "root_cause_id", "causa_raiz", "responsable", "subcausa",
    "prioridad_regla", "fuente", "detalle", "datos_faltantes",
)


def _fila_dia(ev, dg: dict) -> tuple:
    """Una evidencia y su dictamen, aplanados a las columnas de `run_dias`.

    `detalle` se guarda ya resuelto porque es lo único que el Excel usa de la
    lista de evidencia (su último renglón), y así no hay que guardar la lista
    entera. `datos_faltantes` sí va completo: lo necesitan dentro_del_alcance
    y la cobertura para reconstruir el resultado sin reclasificar.
    """
    detalle = dg["evidencia"][-1] if dg.get("evidencia") else ""
    if dg.get("datos_faltantes"):
        detalle = "FALTA EL DATO: " + ", ".join(dg["datos_faltantes"])

    return (
        ev.sku, ev.tienda, ev.fecha, ev.osa, ev.venta_perdida,
        ev.inventario_tienda, ev.transito_vigente, ev.pedido_tienda_generado,
        ev.tipo_resurtido.value if ev.tipo_resurtido else None,
        ev.via_resurtido.value if ev.via_resurtido else None,
        ev.inventario_cedis, ev.envio_cedis_generado,
        ev.pedido_proveedor_generado, ev.proveedor_cajas_pedidas,
        ev.proveedor_cita_agendada, ev.proveedor_fecha_cita, ev.proveedor_cita_vencida,
        ev.proveedor_cajas_confirmadas_cita, ev.proveedor_cajas_entregadas,
        dg["root_cause_id"], dg["causa_raiz"], dg["responsable"],
        dg.get("subcausa"), dg.get("prioridad_regla"), dg.get("fuente"), detalle,
        list(dg.get("datos_faltantes") or []),
    )


def guardar_dias(id_: str, evidencias, diagnosticos, dsn: Optional[str] = None
                 ) -> Optional[str]:
    """Guarda el detalle diario de una corrida. ~44 mil renglones por tienda-mes.

    Se inserta por lotes con execute_values: fila por fila serían 44 mil
    viajes de ida y vuelta a Neon. No levanta, por lo mismo que `guardar`.
    """
    filas = [_fila_dia(ev, dg) for ev, dg in zip(evidencias, diagnosticos)]
    if not filas:
        return None
    cols = ", ".join(COLUMNAS_DIA)
    try:
        conn = conectar(dsn)
        try:
            with conn, conn.cursor() as cur:
                # Rehacer, no acumular: si la corrida se vuelve a guardar, su
                # detalle viejo sobra.
                cur.execute("DELETE FROM run_dias WHERE run_id = %s", (id_,))
                execute_values(
                    cur,
                    f"INSERT INTO run_dias (run_id, {cols}) VALUES %s",
                    [(id_, *f) for f in filas],
                    page_size=2000,
                )
        finally:
            conn.close()
        return None
    except Exception as e:
        return (f"El análisis terminó bien, pero no se pudo guardar el detalle diario "
                f"({e}). El resultado es válido; lo que no va a poder hacerse es "
                f"regenerar el Excel sin volver a correr el análisis.")


def leer_dias(id_: str, dsn: Optional[str] = None) -> List[dict]:
    """El detalle diario guardado, en el orden en que va al Excel."""
    conn = conectar(dsn)
    try:
        with conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(f"SELECT {', '.join(COLUMNAS_DIA)} FROM run_dias "
                        f"WHERE run_id = %s ORDER BY sku, fecha", (id_,))
            return [dict(f) for f in cur.fetchall()]
    finally:
        conn.close()


def listar(limite: int = 50, tienda: Optional[str] = None,
           dsn: Optional[str] = None) -> List[dict]:
    """Las corridas hechas, lo más reciente primero. SIN el resumen: son
    5 MB cada uno y aquí sólo se pinta un renglón por corrida."""
    sql = ("SELECT id, tienda, desde, hasta, umbral_osa, version_motor, "
           "       osa_alcance, dias_faltante, venta_perdida, cobertura_pct, "
           "       corrido_en, segundos, origen, archivo "
           "FROM runs")
    params: list = []
    if tienda:
        sql += " WHERE tienda = %s"
        params.append(tienda)
    sql += " ORDER BY corrido_en DESC LIMIT %s"
    params.append(limite)

    conn = conectar(dsn)
    try:
        with conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(sql, params)
            return [dict(f) for f in cur.fetchall()]
    finally:
        conn.close()


def leer(id_: str, dsn: Optional[str] = None) -> Optional[dict]:
    """El resumen guardado de una corrida, tal cual lo devolvió el análisis."""
    conn = conectar(dsn)
    try:
        with conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT resumen, tienda, desde, hasta, version_motor, "
                        "corrido_en, umbral_osa, universo FROM runs WHERE id = %s", (id_,))
            f = cur.fetchone()
            if f is None:
                return None
            resumen = dict(f["resumen"])
            # Sellos de procedencia: la pantalla tiene que poder decir que
            # esto se leyó del histórico y con qué versión se calculó, no
            # hacerlo pasar por un análisis recién corrido.
            resumen["guardado"] = {
                "tienda": f["tienda"],
                "desde": f["desde"].isoformat(),
                "hasta": f["hasta"].isoformat(),
                "version_motor": f["version_motor"],
                "corrido_en": f["corrido_en"].isoformat(),
            }
            resumen["_umbral_osa"] = float(f["umbral_osa"] or 100.0)
            resumen["_universo"] = f["universo"] or {}
            return resumen
    finally:
        conn.close()


def borrar(id_: str, dsn: Optional[str] = None) -> bool:
    conn = conectar(dsn)
    try:
        with conn, conn.cursor() as cur:
            cur.execute("DELETE FROM runs WHERE id = %s", (id_,))
            return cur.rowcount > 0
    finally:
        conn.close()


# ===========================================================================
# Regenerar el Excel desde lo guardado
#
# El escritor del Excel (orcmm_pipeline.escribir_resultado) lee objetos de
# evidencia y diccionarios de dictamen, no la base. En vez de reescribirlo
# —son cinco hojas con su formato, y es un entregable de cliente— se
# reconstruyen esos objetos desde `run_dias`. Así el archivo regenerado sale
# byte por byte con el mismo criterio que el original, sin tocar una línea del
# código que lo produce.
# ===========================================================================

class _EvidenciaGuardada:
    """Lo que el escritor del Excel lee de una evidencia, traído de la tabla.

    Los nombres son los de EvidenciaSKUTienda a propósito: `_celda_cita` y la
    hoja diaria funcionan sin cambiarles nada.
    """

    __slots__ = ("sku", "tienda", "fecha", "osa", "venta_perdida",
                 "inventario_tienda", "transito_vigente", "pedido_tienda_generado",
                 "tipo_resurtido", "via_resurtido", "inventario_cedis",
                 "envio_cedis_generado", "pedido_proveedor_generado",
                 "proveedor_cajas_pedidas", "proveedor_cita_agendada",
                 "proveedor_fecha_cita", "proveedor_cita_vencida",
                 "proveedor_cajas_confirmadas_cita", "proveedor_cajas_entregadas")

    def __init__(self, f: dict):
        from orcmm_pipeline import TIPOS_RESURTIDO, VIAS, clave_catalogo

        self.sku = f["sku"]
        self.tienda = f["tienda"]
        self.fecha = f["fecha"]
        self.osa = _flotante(f["osa"])
        self.venta_perdida = _flotante(f["venta_perdida"])
        self.inventario_tienda = f["inventario_tienda"]
        self.transito_vigente = f["transito_vigente"]
        self.pedido_tienda_generado = f["pedido_tienda_generado"]
        # De vuelta al enum: el Excel escribe `.value`, y guardar el texto y
        # devolverlo como texto haría fallar ese acceso.
        self.tipo_resurtido = TIPOS_RESURTIDO.get(clave_catalogo(f["tipo_resurtido"]))
        self.via_resurtido = VIAS.get(clave_catalogo(f["via_resurtido"]))
        self.inventario_cedis = f["inventario_cedis"]
        self.envio_cedis_generado = f["envio_cedis_generado"]
        self.pedido_proveedor_generado = f["pedido_proveedor_generado"]
        self.proveedor_cajas_pedidas = f["cajas_pedidas"]
        self.proveedor_cita_agendada = f["cita_agendada"]
        self.proveedor_fecha_cita = f["cita_fecha"]
        self.proveedor_cita_vencida = f["cita_vencida"]
        self.proveedor_cajas_confirmadas_cita = f["cajas_confirmadas"]
        self.proveedor_cajas_entregadas = f["cajas_entregadas"]


class _FuentesGuardadas:
    """El pedacito de `Fuentes` que el escritor del Excel necesita.

    Son cuatro cosas —advertencias, conteo, rango y vacia()— y las tres
    primeras ya viajan dentro del resumen guardado. Las hojas de proveedor y
    citas NO se recalculan: se le pasan al escritor ya hechas, tal como
    quedaron en la corrida original.
    """

    def __init__(self, resumen: dict):
        self.advertencias = list(resumen.get("advertencias") or [])
        fuentes = resumen.get("fuentes") or []
        self.conteo = {f["hoja"]: f["filas"] for f in fuentes}
        self.rango = {f["hoja"]: (_fecha(f.get("desde")), _fecha(f.get("hasta")))
                      for f in fuentes}

    def vacia(self, hoja: str) -> bool:
        return self.conteo.get(hoja, 0) == 0


def _flotante(v):
    """Postgres devuelve `numeric` como Decimal, y el resto del pipeline suma
    floats: mezclarlos truena con "unsupported operand". Se normaliza aquí, en
    la frontera, y no salpica al resto."""
    return None if v is None else float(v)


def _fecha(v):
    from datetime import date as _d
    if not v:
        return None
    if isinstance(v, _d):
        return v
    return _d.fromisoformat(str(v)[:10])


def regenerar_excel(id_: str, salida, dsn: Optional[str] = None) -> bool:
    """Reescribe el Excel de una corrida guardada. Devuelve False si no está.

    No vuelve a leer las fuentes ni a clasificar: los ~56 s de análisis se
    ahorran completos y sólo queda el costo de escribir el archivo.
    """
    from orcmm_pipeline import escribir_resultado

    resumen = leer(id_, dsn)
    if resumen is None:
        return False
    filas = leer_dias(id_, dsn)
    if not filas:
        return False

    evidencias = [_EvidenciaGuardada(f) for f in filas]
    diagnosticos = [{
        "sku": f["sku"], "tienda": f["tienda"],
        "fecha": f["fecha"].isoformat() if f["fecha"] else None,
        "root_cause_id": f["root_cause_id"], "causa_raiz": f["causa_raiz"],
        "responsable": f["responsable"], "subcausa": f["subcausa"],
        "prioridad_regla": f["prioridad_regla"], "fuente": f["fuente"],
        "venta_perdida": _flotante(f["venta_perdida"]),
        "osa": _flotante(f["osa"]),
        "datos_faltantes": list(f["datos_faltantes"] or []),
        # Un día está clasificado si la matriz le puso causa. RC99 (sin
        # clasificar) y RC00 (fuera de alcance) salen los dos de un
        # Indeterminado, así que ninguno cuenta — mismo criterio que el motor.
        "clasificado": f["root_cause_id"] not in (None, "", "RC99", "RC00"),
        # El escritor sólo usa el último renglón de la evidencia, y es
        # justo lo que se guardó en `detalle`.
        "evidencia": [f["detalle"]] if f["detalle"] else [],
    } for f in filas]

    # El universo vuelve a su forma de tuplas: se guardó con llave plana
    # porque JSON no admite tuplas de llave.
    universo = {tuple(k.split("|", 1)): tuple(v)
                for k, v in (resumen.get("_universo") or {}).items()}

    escribir_resultado(
        salida, _FuentesGuardadas(resumen), evidencias, diagnosticos,
        resumen.get("_umbral_osa", 100.0),
        # El resumen los guarda como diccionarios y la hoja los lee por
        # atributo; las llaves son las mismas, así que un namespace basta.
        proveedores=[SimpleNamespace(**d) for d in (resumen.get("proveedores") or [])],
        # El resumen serializa la fecha de cita a texto y la hoja la vuelve a
        # formatear: se devuelve a fecha en la frontera, igual que _flotante.
        citas=[{**c, "fecha_cita": _fecha(c.get("fecha_cita"))}
               for c in (resumen.get("citas_falladas") or [])],
        discrepancias=resumen.get("discrepancias"),
        precalculado={
            "universo": universo,
            "waterfall": resumen.get("waterfall"),
            "osa_general": resumen.get("osa_general"),
            "osa_alcance": resumen.get("osa_alcance"),
        },
    )
    return True
