-- =========================================================================
-- ORCMM — borrar la información cargada, SIN tocar el esquema.
--
-- Vacía las tablas de datos operativos (las 9 hojas del layout + la
-- bitácora de cargas) pero deja las tablas, índices y PRIMARY KEY tal como
-- están, y NO toca sucursales/catalogo_sku_tienda (son informativas, de
-- ciclo de vida propio — no se recargan cada vez que llega una versión
-- nueva de los datos operativos). Después de correr esto,
-- orcmm_etl_carga.py puede volver a cargar desde cero sin necesidad de
-- correr orcmm_db_init.py otra vez.
--
-- No hay FOREIGN KEY entre tablas (diseño a propósito, ver sql/schema.sql),
-- así que el orden de los TRUNCATE no importa.
--
-- Uso rápido: python orcmm_db_borrar.py --si
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
    etl_cargas
RESTART IDENTITY;

-- -------------------------------------------------------------------------
-- Catálogos informativos (sucursales, catalogo_sku_tienda) — NO se tocan
-- arriba a propósito. Sólo bórralos si de verdad hace falta (p. ej. llegó
-- un catálogo de SKU corregido); usa `python orcmm_db_borrar.py --si
-- --con-catalogos` en vez de descomentar esto a mano.
-- -------------------------------------------------------------------------
-- TRUNCATE TABLE sucursales, catalogo_sku_tienda RESTART IDENTITY;

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
