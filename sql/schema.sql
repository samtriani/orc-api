-- =========================================================================
-- ORCMM — esquema de datos crudos (Neon Postgres)
--
-- Sin FOREIGN KEY entre hojas a propósito: la cobertura cruzada real trae
-- huecos legítimos (ver orcmm_validar_layout.py, que valida eso en la capa
-- de aplicación — no aquí). Cargas idempotentes por UPSERT sobre la llave
-- natural de cada hoja. Todo usa IF NOT EXISTS: correr este archivo de
-- nuevo es seguro.
-- =========================================================================

-- -------------------------------------------------------------------------
-- 1. CATALOGO — estático, una fila por sku+tienda
-- -------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS catalogo (
    sku               TEXT NOT NULL,
    descripcion       TEXT,
    tienda            TEXT NOT NULL,
    cedis_surtidor    TEXT NOT NULL,
    division          TEXT NOT NULL,
    via_resurtido     TEXT NOT NULL,
    piezas_por_caja   INTEGER NOT NULL,
    tipo_resurtido    TEXT,
    rol_frecuencia    TEXT,
    linea_vs_io       TEXT,
    estatus_activo    TEXT NOT NULL,
    nombre_tienda     TEXT,
    cargado_en        TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (sku, tienda)
);
CREATE INDEX IF NOT EXISTS ix_catalogo_tienda ON catalogo (tienda);
CREATE INDEX IF NOT EXISTS ix_catalogo_cedis_surtidor ON catalogo (cedis_surtidor);

-- -------------------------------------------------------------------------
-- 2. TABLEAU_INV_TIENDA — diario, sku+tienda+fecha (origen: CSV o Excel)
-- -------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS tableau_inv_tienda (
    sku                    TEXT NOT NULL,
    tienda                 TEXT NOT NULL,
    fecha                  DATE NOT NULL,
    existencia_piezas      INTEGER NOT NULL,
    hora_de_corte          TEXT,
    existencia_minima_dia  INTEGER,
    cargado_en             TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (sku, tienda, fecha)
);
CREATE INDEX IF NOT EXISTS ix_inv_tienda_fecha ON tableau_inv_tienda (fecha);
CREATE INDEX IF NOT EXISTS ix_inv_tienda_tienda_fecha ON tableau_inv_tienda (tienda, fecha);

-- -------------------------------------------------------------------------
-- 3. BOPS_OSA — diario, sku+tienda+fecha
-- -------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS bops_osa (
    sku                     TEXT NOT NULL,
    tienda                  TEXT NOT NULL,
    fecha                   DATE NOT NULL,
    osa_pct                 NUMERIC(5,2) NOT NULL CHECK (osa_pct BETWEEN 0 AND 100),
    venta_perdida_estimada  NUMERIC(14,2) NOT NULL,
    numero_rupturas         INTEGER,
    minutos_sin_producto    INTEGER,
    cargado_en              TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (sku, tienda, fecha)
);
CREATE INDEX IF NOT EXISTS ix_bops_osa_fecha ON bops_osa (fecha);
CREATE INDEX IF NOT EXISTS ix_bops_osa_tienda_fecha ON bops_osa (tienda, fecha);

-- -------------------------------------------------------------------------
-- 4. TABLEAU_VENTAS — diario, sku+tienda+fecha (origen: CSV o Excel)
-- -------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS tableau_ventas (
    sku                     TEXT NOT NULL,
    tienda                  TEXT NOT NULL,
    fecha                   DATE NOT NULL,
    importe_venta           NUMERIC(14,2) NOT NULL,
    unidades_vendidas       INTEGER,
    venta_perdida_estimada  NUMERIC(14,2),
    metodo_estimacion       TEXT,
    cargado_en              TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (sku, tienda, fecha)
);
CREATE INDEX IF NOT EXISTS ix_ventas_fecha ON tableau_ventas (fecha);
CREATE INDEX IF NOT EXISTS ix_ventas_tienda_fecha ON tableau_ventas (tienda, fecha);

-- -------------------------------------------------------------------------
-- 5. CEDIS_INVENTARIO — diario, sku+cedis+fecha
-- -------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS cedis_inventario (
    sku                 TEXT NOT NULL,
    cedis               TEXT NOT NULL,
    fecha               DATE NOT NULL,
    existencia_piezas   INTEGER NOT NULL,
    piezas_reservadas   INTEGER,
    cargado_en          TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (sku, cedis, fecha)
);
CREATE INDEX IF NOT EXISTS ix_cedis_inv_fecha ON cedis_inventario (fecha);
CREATE INDEX IF NOT EXISTS ix_cedis_inv_cedis_fecha ON cedis_inventario (cedis, fecha);

-- -------------------------------------------------------------------------
-- 6. CEDIS_TRANSFERENCIAS — evento, llave folio+sku
-- -------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS cedis_transferencias (
    folio                     TEXT NOT NULL,
    sku                       TEXT NOT NULL,
    cedis_origen              TEXT NOT NULL,
    tienda_destino            TEXT NOT NULL,
    fecha_generacion          DATE NOT NULL,
    fecha_salida_cedis        DATE NOT NULL,
    -- Nullable pese al * del spec: "VACÍO si no se había recibido al corte".
    fecha_recepcion_tienda    DATE,
    cantidad_enviada_piezas   INTEGER NOT NULL,
    cantidad_recibida_piezas  INTEGER,
    estatus                   TEXT,
    cargado_en                TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (folio, sku)
);
CREATE INDEX IF NOT EXISTS ix_transferencias_fecha_gen ON cedis_transferencias (fecha_generacion);
CREATE INDEX IF NOT EXISTS ix_transferencias_tienda_destino ON cedis_transferencias (tienda_destino);
CREATE INDEX IF NOT EXISTS ix_transferencias_cedis_origen ON cedis_transferencias (cedis_origen);

-- -------------------------------------------------------------------------
-- 7. SIMA_PEDIDOS_TIENDA — evento, llave folio+sku
-- -------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS sima_pedidos_tienda (
    folio                    TEXT NOT NULL,
    sku                      TEXT NOT NULL,
    tienda                   TEXT NOT NULL,
    fecha_pedido             DATE NOT NULL,
    fecha_requerida          DATE,
    cantidad_pedida_piezas   INTEGER NOT NULL,
    cantidad_surtida_piezas  INTEGER NOT NULL,
    -- Nullable pese al * del spec: "VACÍO si sigue abierto".
    fecha_surtido            DATE,
    estatus                  TEXT,
    cargado_en               TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (folio, sku)
);
CREATE INDEX IF NOT EXISTS ix_pedidos_tienda_fecha_pedido ON sima_pedidos_tienda (fecha_pedido);
CREATE INDEX IF NOT EXISTS ix_pedidos_tienda_tienda ON sima_pedidos_tienda (tienda);

-- -------------------------------------------------------------------------
-- 8. COMPRAS_PEDIDOS_PROV — evento, llave folio+sku
-- -------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS compras_pedidos_prov (
    folio               TEXT NOT NULL,
    sku                 TEXT NOT NULL,
    proveedor_id        TEXT NOT NULL,
    proveedor_nombre    TEXT,
    cedis_destino       TEXT NOT NULL,
    fecha_pedido        DATE NOT NULL,
    fecha_cita          DATE,
    -- Nullable a propósito: 10-15% vacío es normal ("aún no se recibe"),
    -- pese a que el spec lo marca obligatorio.
    fecha_recibo        DATE,
    cajas_pedidas       INTEGER NOT NULL,
    cajas_entregadas    INTEGER NOT NULL,
    estatus             TEXT,
    cargado_en          TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (folio, sku)
);
CREATE INDEX IF NOT EXISTS ix_compras_fecha_pedido ON compras_pedidos_prov (fecha_pedido);
CREATE INDEX IF NOT EXISTS ix_compras_proveedor ON compras_pedidos_prov (proveedor_id);
CREATE INDEX IF NOT EXISTS ix_compras_cedis_destino ON compras_pedidos_prov (cedis_destino);

-- -------------------------------------------------------------------------
-- 9. CITAS_PROV_CEDIS — llave (folio_cita, sku). Una cita es una ventana de
--    entrega en CEDIS y puede traer varios SKU (verificado contra datos
--    reales: folio_cita 957821 trae 7 SKU distintos, cada uno con sus
--    propias cajas_confirmadas_cita). folio_cita SOLO no es único.
--    folio (el pedido) es referencia informativa SIN FK — ~61% de las
--    citas llegan sin match en compras_pedidos_prov.
-- -------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS citas_prov_cedis (
    folio                    TEXT NOT NULL,
    folio_cita               TEXT NOT NULL,
    sku                      TEXT NOT NULL,
    proveedor_id             TEXT NOT NULL,
    proveedor_nombre         TEXT,
    cedis_destino            TEXT NOT NULL,
    fecha_pedido             DATE NOT NULL,
    fecha_cita               DATE NOT NULL,
    cajas_confirmadas_cita   INTEGER NOT NULL,
    cajas_entregadas         INTEGER NOT NULL,
    estatus_cita             TEXT,
    fecha_entrega_real       DATE,
    cargado_en               TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (folio_cita, sku)
);
CREATE INDEX IF NOT EXISTS ix_citas_folio ON citas_prov_cedis (folio);
CREATE INDEX IF NOT EXISTS ix_citas_fecha_cita ON citas_prov_cedis (fecha_cita);
CREATE INDEX IF NOT EXISTS ix_citas_proveedor ON citas_prov_cedis (proveedor_id);

-- -------------------------------------------------------------------------
-- 10. SUCURSALES — maestro de tiendas de todo el grupo (La Comer, Fresko,
--     City Market, City Café, Sumesa). INFORMATIVA: el motor de
--     clasificación sigue usando catalogo.cedis_surtidor / Nombre_Tienda,
--     no esta tabla — le falta el CEDIS que surte a cada tienda.
-- -------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS sucursales (
    tienda      TEXT NOT NULL,
    formato     TEXT,
    nombre      TEXT,
    direccion   TEXT,
    cp          TEXT,
    cargado_en  TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (tienda)
);
CREATE INDEX IF NOT EXISTS ix_sucursales_formato ON sucursales (formato);

-- -------------------------------------------------------------------------
-- 11. CATALOGO_SKU_TIENDA — catálogo completo de SKU por tienda tal como lo
--     entrega Compras (un archivo por tienda). INFORMATIVA: trae categoría/
--     proveedor/frecuencia que no existen en `catalogo`, pero le falta
--     cedis_surtidor, así que el motor de clasificación sigue usando
--     `catalogo`, no esta tabla.
-- -------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS catalogo_sku_tienda (
    sku                 TEXT NOT NULL,
    tienda              TEXT NOT NULL,
    tienda_nombre       TEXT,
    articulo_nombre     TEXT,
    division            TEXT,
    grupo_seccion       TEXT,
    proveedor_nombre    TEXT,
    resurtido_tipo      TEXT,
    resurtido_frec      TEXT,
    unidades_empaque    INTEGER,
    resurtido           INTEGER,
    catalogo_activo     INTEGER,
    linea_io            TEXT,
    via_resurtido       TEXT,
    fecha_inicial       DATE,
    cargado_en          TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (sku, tienda)
);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_tienda ON catalogo_sku_tienda (tienda);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_division ON catalogo_sku_tienda (division);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_proveedor ON catalogo_sku_tienda (proveedor_nombre);

-- =========================================================================
-- 12. etl_cargas — bitácora de cada corrida del ETL (linaje, no resultados
--     de clasificación).
-- =========================================================================
CREATE TABLE IF NOT EXISTS etl_cargas (
    id                BIGSERIAL PRIMARY KEY,
    iniciado_en       TIMESTAMPTZ NOT NULL DEFAULT now(),
    terminado_en      TIMESTAMPTZ,
    archivo_xlsx      TEXT,
    archivos_csv      JSONB NOT NULL DEFAULT '[]'::jsonb,
    validado          BOOLEAN NOT NULL DEFAULT false,
    forzado           BOOLEAN NOT NULL DEFAULT false,
    estado            TEXT NOT NULL DEFAULT 'en_proceso'
                       CHECK (estado IN ('en_proceso','ok','parcial','error')),
    filas_por_hoja    JSONB NOT NULL DEFAULT '{}'::jsonb,
    advertencias      JSONB NOT NULL DEFAULT '[]'::jsonb,
    error             TEXT,
    duracion_s        NUMERIC(10,2)
);
CREATE INDEX IF NOT EXISTS ix_etl_cargas_iniciado_en ON etl_cargas (iniciado_en DESC);
