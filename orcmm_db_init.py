"""ORCMM — aplica sql/orcmm_ddl_completo.sql a la base de DATABASE_URL.

    python orcmm_db_init.py

Seguro de volver a correr: todo el DDL usa IF NOT EXISTS.

Apuntaba a sql/historico/schema.sql, que se quedó en la estructura de agosto: no
creaba `runs` ni `run_dias` ni los índices que se agregaron después. El
archivo consolidado sí trae las 14 tablas, y es el mismo que se le
entrega al cliente para instalar desde cero.
"""
from pathlib import Path

from dotenv import load_dotenv

from orcmm_db import conectar


def main() -> None:
    load_dotenv()
    sql = Path(__file__).parent.joinpath("sql", "orcmm_ddl_completo.sql").read_text(encoding="utf-8")
    conn = conectar()
    try:
        with conn, conn.cursor() as cur:
            cur.execute(sql)
        print("Esquema aplicado.")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
