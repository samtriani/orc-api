-- ============================================================================
-- ORCMM · Migración — CATALOGO.decil
--
-- Ejecutar UNA vez ANTES de cargar el layout que trae la columna nueva: si el
-- ETL corre primero, truena al escribir una columna que no existe.
--
-- Es idempotente (IF NOT EXISTS), así que volver a correrla no hace nada.
--
-- Nullable a propósito: los catálogos ya cargados no la traen, y forzar un
-- valor obligaría a inventar un decil. Mientras venga vacía, el filtro
-- simplemente no ofrece opciones — no rompe nada.
-- ============================================================================

ALTER TABLE catalogo
    ADD COLUMN IF NOT EXISTS decil TEXT;

-- El filtro agrupa por decil sobre el catálogo de una tienda.
CREATE INDEX IF NOT EXISTS ix_catalogo_decil ON catalogo (tienda, decil);
