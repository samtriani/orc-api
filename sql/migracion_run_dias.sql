-- ============================================================================
-- ORCMM · Migración — `run_dias`: el veredicto de cada día, guardado
--
-- Es exactamente la hoja "Clasificación diaria" del Excel, fila por fila: las
-- mismas 24 columnas, con la evidencia que llevó a cada dictamen. Por eso
-- sirve para dos cosas a la vez:
--
--   1. Regenerar el Excel sin volver a leer las fuentes ni reclasificar. Una
--      corrida completa cuesta ~288 s; escribir el archivo desde aquí se
--      salta los 56 s de análisis.
--   2. Consultar el resultado. Hasta ahora el detalle diario sólo existía
--      dentro de un JSON de 5 MB o de un .xlsx de 16 MB: no se podía
--      preguntar "todos los RC03 de marzo" sin abrir uno de los dos.
--
-- Son ~44 mil renglones por corrida de una tienda y un mes. Se borra en
-- cascada con la corrida: un detalle sin su `runs` no dice de dónde salió ni
-- con qué versión del motor, y eso ya no es un dato, es basura.
--
-- Idempotente: se puede ejecutar varias veces.
-- ============================================================================

CREATE TABLE IF NOT EXISTS run_dias (
    run_id          text NOT NULL REFERENCES runs(id) ON DELETE CASCADE,

    sku             text NOT NULL,
    tienda          text NOT NULL,
    fecha           date NOT NULL,
    osa             numeric,
    venta_perdida   numeric,

    -- La evidencia, en el mismo orden en que la lee el árbol. Los booleanos
    -- son de TRES estados a propósito: NULL es "no se sabe", que no es lo
    -- mismo que false. Todo el modelo se apoya en esa distinción.
    inventario_tienda           integer,
    transito_vigente            boolean,
    pedido_tienda_generado      boolean,
    tipo_resurtido              text,
    via_resurtido               text,
    inventario_cedis            integer,
    envio_cedis_generado        boolean,
    pedido_proveedor_generado   boolean,
    cajas_pedidas               integer,
    -- La cita, en crudo y no como el texto que pinta el Excel: así el
    -- escritor la vuelve a formatear igual que en la corrida original, y de
    -- paso se puede preguntar "cuántas citas vencidas hubo en marzo".
    cita_agendada               boolean,
    cita_fecha                  date,
    cita_vencida                boolean,
    cajas_confirmadas           integer,
    cajas_entregadas            integer,

    -- El dictamen.
    root_cause_id   text,
    causa_raiz      text,
    responsable     text,
    subcausa        text,
    prioridad_regla integer,
    fuente          text,
    detalle         text,
    -- Qué dato faltó, cuando faltó. Es lista porque una regla puede quedarse
    -- esperando más de uno, y la necesitan dentro_del_alcance y la cobertura
    -- para reconstruir el resultado sin volver a clasificar.
    datos_faltantes text[]
);

-- Regenerar el Excel: se leen todos los días de una corrida en orden.
CREATE INDEX IF NOT EXISTS ix_run_dias_run ON run_dias (run_id, sku, fecha);
-- Las consultas que motivaron la tabla: "los RC03 de esta corrida", "este
-- SKU a lo largo del periodo".
CREATE INDEX IF NOT EXISTS ix_run_dias_causa ON run_dias (run_id, root_cause_id);
CREATE INDEX IF NOT EXISTS ix_run_dias_sku ON run_dias (sku, tienda, fecha);
