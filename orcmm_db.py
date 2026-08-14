"""ORCMM — conexión a Postgres (Neon) y upsert genérico para el ETL.

No reescribe el parseo: los 4 conversores de tipo son los mismos que ya usa
orcmm_pipeline.py, y el mapeo de columnas se deriva de orcmm_layout_spec.HOJAS
en vez de retipear la lista de campos por tercera vez.
"""
from __future__ import annotations

import json
import os
import time
from typing import Iterable, Optional

import psycopg2
from psycopg2.extras import execute_values

from orcmm_layout_spec import BOL, DEC, ENT, FEC, HOJAS, HOR, LST, TXT
from orcmm_pipeline import _decimal, _entero, _fecha, _texto

# LST/BOL/HOR no tienen conversor dedicado en el pipeline hoy (nada los usa
# para clasificar): se guardan como texto tal cual llegan, sin normalizar.
CONVERTIDORES = {
    TXT: _texto, ENT: _entero, DEC: _decimal, FEC: _fecha,
    LST: _texto, BOL: _texto, HOR: _texto,
}

TABLAS = {
    "CATALOGO":             {"tabla": "catalogo",            "llave": ["sku", "tienda"]},
    "TABLEAU_INV_TIENDA":   {"tabla": "tableau_inv_tienda",   "llave": ["sku", "tienda", "fecha"]},
    "BOPS_OSA":             {"tabla": "bops_osa",             "llave": ["sku", "tienda", "fecha"]},
    "TABLEAU_VENTAS":       {"tabla": "tableau_ventas",       "llave": ["sku", "tienda", "fecha"]},
    "CEDIS_INVENTARIO":     {"tabla": "cedis_inventario",     "llave": ["sku", "cedis", "fecha"]},
    "CEDIS_TRANSFERENCIAS": {"tabla": "cedis_transferencias", "llave": ["folio", "sku"]},
    "SIMA_PEDIDOS_TIENDA":  {"tabla": "sima_pedidos_tienda",  "llave": ["folio", "sku"]},
    "COMPRAS_PEDIDOS_PROV": {"tabla": "compras_pedidos_prov", "llave": ["folio", "sku"]},
    # folio_cita solo no es único: una cita cubre varios SKU (verificado
    # contra datos reales — ver nota en sql/schema.sql).
    "CITAS_PROV_CEDIS":     {"tabla": "citas_prov_cedis",     "llave": ["folio_cita", "sku"]},
}

# Campos cuyo nombre de spec no calza como columna SQL con solo minúsculas
# (traen mayúsculas y/o espacios). Se mapean explícitamente aquí.
RENOMBRES = {
    "CATALOGO": {"Nombre_Tienda": "nombre_tienda"},
    "BOPS_OSA": {"Alerta Enviada": "alerta_enviada", "Alerta Ejecutada": "alerta_ejecutada"},
    "COMPRAS_PEDIDOS_PROV": {"Tienda_Destino": "tienda_destino"},
}


def columnas_de(hoja: str):
    """(campo_del_spec, columna_bd, convertidor) por cada campo — deriva de
    HOJAS en vez de retipear la lista de campos."""
    ren = RENOMBRES.get(hoja, {})
    return [(campo, ren.get(campo, campo.lower()), CONVERTIDORES[tipo])
            for campo, tipo, *_ in HOJAS[hoja]["campos"]]


def conectar(dsn: Optional[str] = None, reintentos: int = 3):
    """Conecta a Neon con reintento corto por si el compute serverless
    está dormido (cold start)."""
    dsn = dsn or os.environ["DATABASE_URL"]
    ultimo = None
    for i in range(reintentos):
        try:
            return psycopg2.connect(dsn, connect_timeout=10)
        except psycopg2.OperationalError as e:
            ultimo = e
            time.sleep(2 * (2 ** i))
    raise ultimo


def upsert_lote(cur, tabla: str, columnas_bd: list, llave: list,
                 filas: list, advertencias: Optional[list] = None) -> int:
    """INSERT ... ON CONFLICT (llave) DO UPDATE en un solo batch.

    Deduplica por llave natural DENTRO del batch antes de enviarlo: si la
    misma llave aparece dos veces en el mismo archivo (pasa con reentregas),
    Postgres tira "ON CONFLICT DO UPDATE command cannot affect row a second
    time". Se queda con la última ocurrencia y avisa en vez de tronar.
    """
    vistas = {}
    for f in filas:
        k = tuple(f.get(c) for c in llave)
        if k in vistas and advertencias is not None:
            advertencias.append(
                f"{tabla}: llave {k} repetida en el mismo archivo — se conserva la última.")
        vistas[k] = f
    filas = list(vistas.values())
    if not filas:
        return 0

    no_llave = [c for c in columnas_bd if c not in llave]
    set_clause = ", ".join(f"{c} = EXCLUDED.{c}" for c in no_llave) + ", cargado_en = now()"
    sql = (f"INSERT INTO {tabla} ({', '.join(columnas_bd)}) VALUES %s "
           f"ON CONFLICT ({', '.join(llave)}) DO UPDATE SET {set_clause}")
    valores = [tuple(f.get(c) for c in columnas_bd) for f in filas]
    execute_values(cur, sql, valores, page_size=1000)
    return len(filas)


def registrar_carga_inicio(cur, archivo_xlsx: str, archivos_csv: list) -> int:
    cur.execute(
        "INSERT INTO etl_cargas (archivo_xlsx, archivos_csv) VALUES (%s, %s::jsonb) RETURNING id",
        (archivo_xlsx, json.dumps(archivos_csv)))
    return cur.fetchone()[0]


def registrar_carga_fin(cur, id_: int, estado: str, filas_por_hoja: dict,
                         advertencias: list, validado: bool, forzado: bool,
                         duracion_s: float, error: Optional[str] = None) -> None:
    cur.execute(
        "UPDATE etl_cargas SET terminado_en = now(), estado = %s, filas_por_hoja = %s::jsonb, "
        "advertencias = %s::jsonb, validado = %s, forzado = %s, duracion_s = %s, error = %s "
        "WHERE id = %s",
        (estado, json.dumps(filas_por_hoja), json.dumps(advertencias), validado, forzado,
         round(duracion_s, 2), error, id_))
