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
from datetime import date
from typing import List, Optional

from psycopg2.extras import Json, RealDictCursor

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
            dsn: Optional[str] = None) -> Optional[str]:
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
                                      resumen, segundos, origen, archivo)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (id) DO UPDATE SET
                        resumen = EXCLUDED.resumen,
                        corrido_en = now()
                    """,
                    (id_, tienda, desde, hasta, umbral_osa,
                     version_motor(), Json(parametros()), p["osa_alcance"],
                     p["dias_faltante"], p["venta_perdida"], p["cobertura_pct"],
                     Json(json.loads(json.dumps(resumen, default=str))),
                     segundos, origen, archivo),
                )
        finally:
            conn.close()
        return None
    except Exception as e:
        return (f"El análisis terminó bien, pero no se pudo guardar en el histórico "
                f"({e}). El resultado de esta pantalla es válido; lo que no va a "
                f"quedar es el registro para volver a consultarlo después.")


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
                        "corrido_en FROM runs WHERE id = %s", (id_,))
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
