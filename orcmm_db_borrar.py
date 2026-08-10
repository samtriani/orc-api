"""ORCMM — vacía las tablas de datos operativos de la base
(sql/borrar_datos.sql). Deja el esquema intacto, y por omisión NO toca
sucursales/catalogo_sku_tienda (informativas, ciclo de vida propio).
Después se puede volver a cargar con orcmm_etl_carga.py sin correr
orcmm_db_init.py de nuevo.

Es destructivo e irreversible, así que exige --si explícito:

    python orcmm_db_borrar.py --si
    python orcmm_db_borrar.py --si --con-catalogos   # también sucursales/catalogo_sku_tienda
"""
import argparse
from pathlib import Path

from dotenv import load_dotenv

from orcmm_db import conectar

TABLAS_OPERATIVAS = ("catalogo", "tableau_inv_tienda", "bops_osa", "tableau_ventas",
                     "cedis_inventario", "cedis_transferencias", "sima_pedidos_tienda",
                     "compras_pedidos_prov", "citas_prov_cedis", "etl_cargas")
TABLAS_CATALOGOS = ("sucursales", "catalogo_sku_tienda")


def main() -> int:
    ap = argparse.ArgumentParser(description="Vacía las tablas de datos operativos de ORCMM en Postgres.")
    ap.add_argument("--si", action="store_true", help="Confirma el borrado. Sin esto, no hace nada.")
    ap.add_argument("--con-catalogos", action="store_true",
                    help="También vacía sucursales/catalogo_sku_tienda (por omisión se conservan).")
    args = ap.parse_args()

    if not args.si:
        print("Esto BORRA los datos operativos cargados (no el esquema, no sucursales/"
              "catalogo_sku_tienda salvo --con-catalogos). Vuelve a correr con --si para confirmar.")
        return 1

    load_dotenv()
    tablas = TABLAS_OPERATIVAS + (TABLAS_CATALOGOS if args.con_catalogos else ())
    sql = f"TRUNCATE TABLE {', '.join(tablas)} RESTART IDENTITY;"

    conn = conectar()
    try:
        with conn, conn.cursor() as cur:
            cur.execute(sql)
        print(f"Tablas vaciadas: {', '.join(tablas)}.")
        if not args.con_catalogos:
            print("sucursales/catalogo_sku_tienda NO se tocaron (usa --con-catalogos para incluirlas).")
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
