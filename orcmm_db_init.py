"""ORCMM — aplica sql/schema.sql a la base configurada en DATABASE_URL.

    python orcmm_db_init.py

Seguro de volver a correr: todo el DDL usa IF NOT EXISTS.
"""
from pathlib import Path

from dotenv import load_dotenv

from orcmm_db import conectar


def main() -> None:
    load_dotenv()
    sql = Path(__file__).parent.joinpath("sql", "schema.sql").read_text(encoding="utf-8")
    conn = conectar()
    try:
        with conn, conn.cursor() as cur:
            cur.execute(sql)
        print("Esquema aplicado.")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
