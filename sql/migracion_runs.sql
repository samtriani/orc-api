-- ============================================================================
-- ORCMM · Migración — tabla `runs`: el resultado de cada análisis, guardado
--
-- POR QUÉ: hasta ahora cada corrida vivía en memoria del proceso y moría con
-- él. Volver a ver un análisis de la semana pasada obligaba a recalcularlo, y
-- una corrida completa de Coyoacán tarda ~5.7 minutos. Con esto la pantalla
-- inicial lista lo que ya se corrió y el detalle se pinta de inmediato.
--
-- El resumen se guarda como JSONB porque es EXACTAMENTE lo que hoy devuelve
-- /api/analizar/{id}/resumen: el front no necesita saber si el JSON se acaba
-- de calcular o se leyó de aquí. Pesa 5.74 MB en claro; Postgres lo comprime
-- solo (TOAST) y en disco quedan ~500 KB.
--
-- `version_motor` NO es opcional. Entre el 18 y el 22 de agosto las reglas
-- cambiaron cinco veces —el parche de inventario, RC03 contra SIMA, los
-- pedidos DSD, el pedido no surtido—, así que dos corridas del mismo periodo
-- hechas con días de diferencia no son comparables. Sin este campo, dentro de
-- tres meses nadie va a poder explicar por qué no cuadran. `parametros`
-- guarda por la misma razón los interruptores de negocio que estaban puestos.
--
-- Idempotente: se puede ejecutar varias veces.
-- ============================================================================

CREATE TABLE IF NOT EXISTS runs (
    id              text PRIMARY KEY,
    tienda          text NOT NULL,
    desde           date NOT NULL,
    hasta           date NOT NULL,
    umbral_osa      numeric NOT NULL DEFAULT 100,

    -- Qué versión del motor produjo esto. Ver arriba.
    version_motor   text,
    parametros      jsonb,

    -- Cifras de portada, desnormalizadas para poder listar las corridas sin
    -- abrir el resumen completo: leer 5 MB de JSONB para pintar un renglón
    -- de tabla sería absurdo.
    osa_alcance     numeric,
    dias_faltante   integer,
    venta_perdida   numeric,
    cobertura_pct   numeric,

    resumen         jsonb NOT NULL,

    corrido_en      timestamptz NOT NULL DEFAULT now(),
    segundos        numeric,
    origen          text NOT NULL DEFAULT 'bd',   -- 'bd' | 'archivo'
    archivo         text
);

-- El listado de la pantalla inicial: lo más reciente primero.
CREATE INDEX IF NOT EXISTS ix_runs_corrido_en ON runs (corrido_en DESC);
-- "¿ya corrí esta tienda y periodo?" — para ofrecer la corrida existente en
-- vez de recalcular 5.7 minutos de lo mismo.
CREATE INDEX IF NOT EXISTS ix_runs_tienda_periodo ON runs (tienda, desde, hasta);
