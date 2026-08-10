-- =========================================================================
-- ORCMM — borrar la información cargada, SIN tocar el esquema.
--
-- Vacía las 10 tablas (los datos de las 9 hojas + la bitácora de cargas)
-- pero deja las tablas, índices y PRIMARY KEY tal como están. Después de
-- correr esto, orcmm_etl_carga.py puede volver a cargar desde cero sin
-- necesidad de correr orcmm_db_init.py otra vez.
--
-- No hay FOREIGN KEY entre tablas (diseño a propósito, ver sql/schema.sql),
-- así que el orden de los TRUNCATE no importa.
--
-- Uso:
--   psql "$DATABASE_URL" -f sql/borrar_datos.sql
-- o, sin psql instalado (como en esta máquina):
--   python -c "from dotenv import load_dotenv; load_dotenv(); from pathlib import Path; from orcmm_db import conectar; c=conectar(); \
--              exec_sql=Path('sql/borrar_datos.sql').read_text(encoding='utf-8'); \
--              conn=c; \
--              [conn.cursor().execute(exec_sql)]; conn.commit(); conn.close()"
-- (o, más simple, el mismo patrón que orcmm_db_init.py pero apuntando a
-- este archivo).
-- =========================================================================

TRUNCATE TABLE
    catalogo,
    tableau_inv_tienda,
    bops_osa,
    tableau_ventas,
    cedis_inventario,
    cedis_transferencias,
    sima_pedidos_tienda,
    compras_pedidos_prov,
    citas_prov_cedis,
    sucursales,
    catalogo_sku_tienda,
    etl_cargas
RESTART IDENTITY;

-- -------------------------------------------------------------------------
-- Alternativa: borrar las TABLAS por completo (esquema incluido), no sólo
-- su contenido. Sólo hace falta si vas a rediseñar el DDL desde cero;
-- para eso vuelve a correr sql/schema.sql (o orcmm_db_init.py) después.
-- Comentado a propósito — descomentar sólo si es lo que se quiere.
-- -------------------------------------------------------------------------
-- DROP TABLE IF EXISTS
--     catalogo,
--     tableau_inv_tienda,
--     bops_osa,
--     tableau_ventas,
--     cedis_inventario,
--     cedis_transferencias,
--     sima_pedidos_tienda,
--     compras_pedidos_prov,
--     citas_prov_cedis,
--     sucursales,
--     catalogo_sku_tienda,
--     etl_cargas;
