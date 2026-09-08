-- ============================================================================
-- ORCMM · Migración V8 — columnas nuevas del layout
-- Ejecutar UNA vez sobre la base (Neon/Postgres) antes de recargar con
--   python orcmm_etl_carga.py <layout_V8>.xlsx ...
-- Todas son IF NOT EXISTS: seguro de volver a correr.
-- Son columnas informativas (ninguna regla del motor las consume aún), salvo
-- las banderas de alerta, que refinan la SUBCAUSA de RC01 sin cambiar la causa
-- ni el responsable (ver REFINAR_RC01_CON_ALERTA en orcmm_rca_engine.py).
-- ============================================================================

-- CATALOGO: proveedor principal del SKU (duplica lo que ya vive en compras/citas,
-- pero permite conocerlo sin depender de una orden abierta).
ALTER TABLE catalogo
    ADD COLUMN IF NOT EXISTS proveedor_id      text,
    ADD COLUMN IF NOT EXISTS proveedor_nombre  text;

-- BOPS_OSA: banderas 0/1 del sistema de alertas (insumo futuro para RC01).
ALTER TABLE bops_osa
    ADD COLUMN IF NOT EXISTS alerta_enviada    integer,
    ADD COLUMN IF NOT EXISTS alerta_ejecutada  integer;

-- COMPRAS_PEDIDOS_PROV: tienda que originó el pedido a proveedor.
ALTER TABLE compras_pedidos_prov
    ADD COLUMN IF NOT EXISTS tienda_destino    text;

-- TABLEAU_INV_TIENDA: el mínimo intradía, cuando Tableau lo entregue de verdad.
-- Mientras tanto llega duplicando la foto de cierre — ver el parche
-- INVENTARIO_CIERRE_NO_CONFIABLE en orcmm_pipeline.py.
--
-- existencia_piezas deja de ser NOT NULL: con el V8 puede venir vacía, y el
-- modelo distingue "no había" de "no se midió". Forzar NOT NULL obligaría a
-- inventar un cero, que es justo lo que el pipeline evita en todas partes.
ALTER TABLE tableau_inv_tienda
    ADD COLUMN IF NOT EXISTS existencia_minima_dia integer;

ALTER TABLE tableau_inv_tienda
    ALTER COLUMN existencia_piezas DROP NOT NULL;
