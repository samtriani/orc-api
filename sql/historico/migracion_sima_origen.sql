-- ============================================================================
-- ORCMM · Migración — SIMA_PEDIDOS_TIENDA.tienda pasa a llamarse `origen`
--
-- Ejecutar UNA vez sobre la base antes de recargar con la entrega de SIMA del
-- 2026-08-19 en adelante. Es idempotente: si la columna ya se renombró, no
-- hace nada.
--
-- POR QUÉ: la columna dejó de ser siempre una tienda. Ahora dice quién GENERÓ
-- el pedido — el id de la sucursal, o '300' cuando lo generó el proceso
-- central. Dejarla llamándose `tienda` con un valor que no es una tienda es
-- justo el tipo de nombre que hace que alguien filtre mal seis meses después.
--
-- El dato viejo no se pierde ni se reinterpreta: las filas que ya estaban
-- traían el id de la tienda, que sigue siendo un `origen` válido.
-- ============================================================================

DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM information_schema.columns
               WHERE table_name = 'sima_pedidos_tienda' AND column_name = 'tienda')
       AND NOT EXISTS (SELECT 1 FROM information_schema.columns
                       WHERE table_name = 'sima_pedidos_tienda' AND column_name = 'origen')
    THEN
        ALTER TABLE sima_pedidos_tienda RENAME COLUMN tienda TO origen;
    END IF;
END $$;

-- El índice viejo apuntaba a la columna por su nombre anterior.
DROP INDEX IF EXISTS ix_pedidos_tienda_tienda;
CREATE INDEX IF NOT EXISTS ix_pedidos_tienda_origen ON sima_pedidos_tienda (origen);
