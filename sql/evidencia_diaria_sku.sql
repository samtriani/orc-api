-- ============================================================================
-- ORCMM · Evidencia diaria de un SKU — para auditar a mano en DBeaver
--
-- Arma en UN renglón por día lo que hoy se revisa tabla por tabla: OSA,
-- inventario en tienda, inventario en CEDIS, tránsito, pedido de tienda y
-- orden a proveedor con su cita.
--
-- QUÉ ES Y QUÉ NO ES
--
-- Devuelve la EVIDENCIA, no el veredicto. La causa raíz la dictamina el motor
-- en Python (orcmm_rca_engine) con 10 reglas en orden de prioridad, y
-- reimplementarlo aquí crearía una segunda verdad que se desincroniza — y al
-- auditar estarías comparando el modelo contra una copia del modelo.
--
-- Con estas columnas se camina el árbol a mano:
--   1. ¿inv_tienda > 0?            -> RC01 Ejecución en Tienda
--   2. ¿transito_vigente?          -> RC02 Transporte / Tránsito
--   3. ¿pedido_tienda_abierto = N? -> RC03 Pedido de tienda no generado
--   4. vía de resurtido            -> bifurca a CEDIS o DSD
--   5-6. ¿inv_cedis = 0?           -> RC04 CEDIS No Surtió
--   7-8. orden a proveedor y cita  -> RC05 / RC06
--
-- LAS REGLAS DE VIGENCIA son las mismas del motor, copiadas al pie:
--   tránsito vigente  : salida <= D AND (recepción IS NULL OR D < recepción)
--   envío generado    : generación <= D AND (recepción IS NULL OR recepción > D)
--   pedido de tienda  : fecha_pedido <= D AND (surtido IS NULL OR surtido > D)
--   orden a proveedor : fecha_pedido <= D AND (recibo IS NULL OR recibo > D)
--
-- Un pedido de tienda con origen '300' es CENTRALIZADO y cubre a todas las
-- sucursales, por eso entra igual que el de la tienda.
--
-- CÓMO USARLO: cambiar los tres valores del bloque `parametros`. Las fechas
-- van AAAA-MM-DD (Postgres no acepta dd/mm/yyyy sin convertir).
-- ============================================================================

WITH parametros AS (
    SELECT
        '7898024390107'::text  AS p_sku,
        '287'::text            AS p_tienda,
        DATE '2026-03-01'      AS p_desde,
        DATE '2026-03-30'      AS p_hasta
),

-- Los eventos se buscan un mes antes del periodo: un pedido de febrero puede
-- seguir vigente en marzo. Mismo criterio que LOOKBACK_EVENTOS_DIAS.
rango AS (
    SELECT *, p_desde - INTERVAL '30 days' AS p_desde_eventos FROM parametros
),

-- El CEDIS que surte a la tienda: el inventario de CEDIS y los pedidos a
-- proveedor se llavean por CEDIS, no por tienda.
cat AS (
    SELECT c.sku, c.tienda, c.cedis_surtidor, c.via_resurtido,
           c.tipo_resurtido, c.descripcion, c.proveedor_nombre
    FROM catalogo c, parametros p
    WHERE c.sku = p.p_sku AND c.tienda = p.p_tienda
),

-- Un renglón por día del periodo, aunque BOPS no lo haya reportado: así se
-- ve el hueco de extracción en vez de que el día desaparezca.
dias AS (
    SELECT generate_series(p_desde, p_hasta, INTERVAL '1 day')::date AS fecha
    FROM parametros
)

SELECT
    d.fecha,
    TO_CHAR(d.fecha, 'DD/MM/YYYY')                      AS fecha_ddmmyyyy,
    p.p_sku                                             AS sku,
    cat.descripcion,
    p.p_tienda                                          AS tienda,
    cat.via_resurtido                                   AS via,
    cat.tipo_resurtido                                  AS tipo_resurtido,

    -- BOPS: define si el día entra al análisis. osa_pct es BINARIO (0 = no
    -- visible en anaquel, 1 = visible); el motor lo normaliza a 0-100.
    b.osa_pct,
    CASE WHEN b.sku IS NULL THEN 'SIN FILA EN BOPS'
         WHEN b.osa_pct < 1 THEN 'CON FALTANTE'
         ELSE 'ok' END                                  AS dia,
    b.venta_perdida_estimada,
    b.alerta_enviada,
    b.alerta_ejecutada,

    -- Prioridad 1. OJO: es la foto de CIERRE (23:59), no el mínimo del día.
    -- existencia_minima_dia viene nula en toda la tabla por ahora.
    it.existencia_piezas                                AS inv_tienda,
    it.existencia_minima_dia,

    -- Prioridad 2. Ya salió de CEDIS y aún no se recibe.
    (SELECT count(*) > 0 FROM cedis_transferencias t
      WHERE t.sku = p.p_sku AND t.tienda_destino = p.p_tienda
        AND t.fecha_salida_cedis <= d.fecha
        AND (t.fecha_recepcion_tienda IS NULL OR d.fecha < t.fecha_recepcion_tienda)
    )                                                   AS transito_vigente,

    -- Generada pero todavía no recibida (prioridad 5-6, rama CEDIS).
    (SELECT count(*) > 0 FROM cedis_transferencias t
      WHERE t.sku = p.p_sku AND t.tienda_destino = p.p_tienda
        AND t.fecha_generacion <= d.fecha
        AND (t.fecha_recepcion_tienda IS NULL OR t.fecha_recepcion_tienda > d.fecha)
    )                                                   AS envio_cedis_generado,

    -- ---------------------------------------------------------------- SIMA
    -- Prioridad 3. `pedido_tienda_abierto` es lo que el motor consume; las
    -- demás columnas son para poder auditarlo sin salir del renglón: con un
    -- true/false no se puede ver CUÁL pedido lo sostiene ni por qué.
    -- Incluye el centralizado (origen '300'), que resurte a todas las tiendas.
    (SELECT count(*) > 0 FROM sima_pedidos_tienda s
      WHERE s.sku = p.p_sku AND s.origen IN (p.p_tienda, '300')
        AND s.fecha_pedido <= d.fecha
        AND (s.fecha_surtido IS NULL OR s.fecha_surtido > d.fecha)
    )                                                   AS pedido_tienda_abierto,
    sm.folio                                            AS sima_folio,
    sm.origen                                           AS sima_origen,
    CASE WHEN sm.origen = '300' THEN 'CENTRALIZADO'
         WHEN sm.origen IS NOT NULL THEN 'de la tienda' END AS sima_quien_pidio,
    sm.fecha_pedido                                     AS sima_fecha_pedido,
    sm.fecha_requerida                                  AS sima_fecha_requerida,
    sm.fecha_surtido                                    AS sima_fecha_surtido,
    sm.cantidad_pedida_piezas                           AS sima_pzas_pedidas,
    sm.cantidad_surtida_piezas                          AS sima_pzas_surtidas,
    -- Cuántos pedidos existen en TODO el periodo, aunque no estén vigentes
    -- este día: distingue "este SKU no se pide nunca" de "se pide, pero no
    -- había pedido abierto ese día". Los dos dan RC03 y son cosas distintas.
    (SELECT count(*) FROM sima_pedidos_tienda s
      WHERE s.sku = p.p_sku AND s.origen IN (p.p_tienda, '300')
    )                                                   AS sima_pedidos_del_periodo,

    -- Existencia en CEDIS, ya neta de reservas. Si NO hay fila y el día sí se
    -- extrajo para ese CEDIS, el modelo lo lee como cero confirmado
    -- (CEDIS_AUSENCIA_ES_CERO): la columna de al lado dice cuál de los dos es.
    GREATEST(ic.existencia_piezas - COALESCE(ic.piezas_reservadas, 0), 0) AS inv_cedis,
    -- Descomentar si hace falta distinguir un cero real de una ausencia: dice
    -- si la fila existe, o si no existe pero el día SÍ se extrajo para ese
    -- CEDIS —que es cuando el modelo la lee como cero confirmado
    -- (CEDIS_AUSENCIA_ES_CERO)— o si el día no se extrajo y no hay dato.
    -- CASE WHEN ic.sku IS NOT NULL THEN 'con fila'
    --      WHEN EXISTS (SELECT 1 FROM cedis_inventario x
    --                    WHERE x.cedis = cat.cedis_surtidor AND x.fecha = d.fecha)
    --           THEN 'sin fila, dia SI extraido -> se lee como CERO'
    --      ELSE 'sin fila, dia NO extraido -> sin dato'
    -- END                                              AS inv_cedis_nota,

    -- Prioridad 7-8. La orden vigente es la de compromiso más próximo, igual
    -- que derivar_orden_proveedor: fecha_cita, si no recibo, si no pedido.
    op.folio                                            AS folio_pedido_prov,
    op.cajas_pedidas,
    op.cajas_entregadas                                 AS cajas_entregadas_pedido,
    op.fecha_cita                                       AS fecha_cita_pedido,
    op.fecha_recibo,

    -- La cita del mismo folio: es la que juzga al proveedor.
    ct.folio_cita,
    ct.fecha_cita,
    ct.cajas_confirmadas_cita,
    ct.cajas_entregadas                                 AS cajas_entregadas_cita,
    ct.estatus_cita

FROM dias d
CROSS JOIN rango p
LEFT JOIN cat ON TRUE
LEFT JOIN bops_osa b
       ON b.sku = p.p_sku AND b.tienda = p.p_tienda AND b.fecha = d.fecha
LEFT JOIN tableau_inv_tienda it
       ON it.sku = p.p_sku AND it.tienda = p.p_tienda AND it.fecha = d.fecha
LEFT JOIN cedis_inventario ic
       ON ic.sku = p.p_sku AND ic.cedis = cat.cedis_surtidor AND ic.fecha = d.fecha

-- El pedido de tienda vigente al día, con su detalle. Misma regla que
-- derivar_pedido_tienda; el más reciente primero para que se vea el que manda.
LEFT JOIN LATERAL (
    SELECT s.*
    FROM sima_pedidos_tienda s
    WHERE s.sku = p.p_sku
      AND s.origen IN (p.p_tienda, '300')
      AND s.fecha_pedido <= d.fecha
      AND (s.fecha_surtido IS NULL OR s.fecha_surtido > d.fecha)
    ORDER BY s.fecha_pedido DESC, s.folio
    LIMIT 1
) sm ON TRUE

-- La orden a proveedor vigente al día: se ordena por compromiso y se toma la
-- primera, con el folio como desempate para que sea determinista.
LEFT JOIN LATERAL (
    SELECT o.*
    FROM compras_pedidos_prov o
    WHERE o.sku = p.p_sku
      AND (o.cedis_destino = cat.cedis_surtidor OR o.tienda_destino = p.p_tienda)
      AND o.fecha_pedido <= d.fecha
      AND (o.fecha_recibo IS NULL OR o.fecha_recibo > d.fecha)
    ORDER BY COALESCE(o.fecha_cita, o.fecha_recibo, o.fecha_pedido), o.folio
    LIMIT 1
) op ON TRUE

LEFT JOIN LATERAL (
    SELECT c.* FROM citas_prov_cedis c
    WHERE c.folio = op.folio AND c.sku = p.p_sku
    ORDER BY c.fecha_cita DESC
    LIMIT 1
) ct ON TRUE

ORDER BY d.fecha;
