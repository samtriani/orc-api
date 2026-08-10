"""ORCMM — vacía las 10 tablas de la base (sql/borrar_datos.sql). Deja el
esquema intacto; después se puede volver a cargar con orcmm_etl_carga.py
sin correr orcmm_db_init.py de nuevo.

Es destructivo e irreversible, así que exige --si explícito:

    python orcmm_db_borrar.py --si
"""
import argparse
from pathlib import Path

from dotenv import load_dotenv

from orcmm_db import conectar


def main() -> int:
    ap = argparse.ArgumentParser(description="Vacía todas las tablas de ORCMM en Postgres.")
    ap.add_argument("--si", action="store_true", help="Confirma el borrado. Sin esto, no hace nada.")
    args = ap.parse_args()

    if not args.si:
        print("Esto BORRA todos los datos cargados (no el esquema). "
              "Vuelve a correr con --si para confirmar.")
        return 1

    load_dotenv()
    sql = Path(__file__).parent.joinpath("sql", "borrar_datos.sql").read_text(encoding="utf-8")
    # El archivo trae un DROP TABLE comentado como alternativa manual; sólo
    # se ejecuta el TRUNCATE de la parte activa.
    conn = conectar()
    try:
        with conn, conn.cursor() as cur:
            cur.execute(sql)
        print("Tablas vaciadas.")
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
