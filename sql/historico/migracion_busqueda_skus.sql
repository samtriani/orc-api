-- ============================================================================
-- ORCMM · Migración — búsqueda de SKU por nombre
--
-- El autocompletar de la pantalla sólo podía ofrecer los SKU que tuvieron
-- faltante, porque son los únicos cuya descripción viaja en la respuesta. Los
-- SANOS —6,930 de 10,454 en Coyoacán— no se podían buscar por nombre, y
-- mandar sus descripciones costaba 552 KB. Se resuelve consultando aquí.
--
-- `unaccent` porque el catálogo escribe "CAFÉ" y la gente teclea "cafe".
-- ============================================================================

CREATE EXTENSION IF NOT EXISTS unaccent;

-- Trigramas: la búsqueda lleva comodín al principio ('%cafe%'), y sin esto
-- Postgres no puede usar ningún índice y recorre el catálogo entero. Con 26
-- mil filas por tienda se aguanta, pero no cuando entren las 94 sucursales.
CREATE EXTENSION IF NOT EXISTS pg_trgm;
CREATE INDEX IF NOT EXISTS ix_catalogo_descripcion_trgm
    ON catalogo USING gin (descripcion gin_trgm_ops);
