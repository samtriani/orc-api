-- ============================================================================
-- ORCMM · Catálogo comercial — jerarquía completa y proveedor
--
-- Ejecutar UNA vez antes de recargar con:
--   python orcmm_etl_catalogos.py --sku <catalogo_con_las_12_columnas>.xlsx
--
-- Todas son IF NOT EXISTS: seguro de volver a correr.
--
-- Esta tabla es el catálogo GLOBAL de referencia (nombre del SKU, nombre de
-- tienda, jerarquía comercial). NO sustituye a `catalogo`, que es la hoja
-- transaccional del layout y es la que el motor lee para clasificar
-- (cedis_surtidor, via_resurtido). Conviven a propósito.
-- ============================================================================

-- La jerarquía venía a medias: había División y Sección, faltaban los dos
-- niveles de abajo. Sin ellos no se puede filtrar por Categoría ni
-- Subcategoría, que es donde vive el detalle que pide negocio.
ALTER TABLE catalogo_sku_tienda
    ADD COLUMN IF NOT EXISTS categoria    text,
    ADD COLUMN IF NOT EXISTS subcategoria text,
    ADD COLUMN IF NOT EXISTS proveedor_id text,
    ADD COLUMN IF NOT EXISTS marca        text;

-- Por donde se va a filtrar. division y proveedor_nombre ya tenían índice.
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_categoria
    ON catalogo_sku_tienda (categoria);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_subcategoria
    ON catalogo_sku_tienda (subcategoria);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_marca
    ON catalogo_sku_tienda (marca);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_proveedor_id
    ON catalogo_sku_tienda (proveedor_id);

-- ----------------------------------------------------------------------------
-- Normalizar los prefijos numéricos que traía la carga anterior.
--
-- Venían como "1 - ABARROTES" y "13 - VINOS Y LICORES (MAS DE 20 GL)". El
-- número no aporta nada al filtro y ensucia el desplegable; la entrega nueva
-- ya trae el texto limpio, así que sin esto convivirían las dos formas del
-- MISMO valor y aparecerían como dos opciones distintas — que es peor que
-- cualquiera de las dos.
--
-- El patrón exige dígitos + espacios + guion + espacios al inicio, para no
-- morder un nombre que legítimamente empiece con número.
-- ----------------------------------------------------------------------------
UPDATE catalogo_sku_tienda
   SET division = regexp_replace(division, '^\s*\d+\s*-\s*', '')
 WHERE division ~ '^\s*\d+\s*-\s*';

UPDATE catalogo_sku_tienda
   SET grupo_seccion = regexp_replace(grupo_seccion, '^\s*\d+\s*-\s*', '')
 WHERE grupo_seccion ~ '^\s*\d+\s*-\s*';
