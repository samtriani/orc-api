-- ###########################################################################
--
--  ORCMM — Modelo de clasificación de causa raíz de faltantes (OSA)
--  ESTRUCTURA COMPLETA DE BASE DE DATOS — PostgreSQL
--
--  Este archivo crea TODA la estructura desde cero. Consolida el esquema
--  base y las siete migraciones que se aplicaron durante el desarrollo, así
--  que reemplaza a todos los archivos anteriores de la carpeta sql/.
--
--  ---------------------------------------------------------------------
--  CÓMO EJECUTARLO
--  ---------------------------------------------------------------------
--
--    1. Crear la base de datos y el usuario de la aplicación (fuera de este
--       archivo, porque CREATE DATABASE no puede ir dentro de una
--       transacción):
--
--         CREATE DATABASE orcmm ENCODING 'UTF8';
--         CREATE USER orcmm_app WITH PASSWORD '<contraseña>';
--         GRANT ALL PRIVILEGES ON DATABASE orcmm TO orcmm_app;
--
--    2. Conectarse a esa base y ejecutar este archivo:
--
--         psql -h <servidor> -U <usuario> -d orcmm -f orcmm_ddl_completo.sql
--
--    3. Al final imprime un resumen de lo creado para verificar.
--
--  Requisitos: PostgreSQL 12 o superior (probado en 14 y 16).
--
--  ES IDEMPOTENTE. Todo usa IF NOT EXISTS: volver a ejecutarlo sobre una
--  base que ya tiene la estructura no hace nada y no borra información.
--  No contiene ningún DELETE, DROP TABLE ni TRUNCATE.
--
--  ---------------------------------------------------------------------
--  NOTA PARA AZURE DATABASE FOR POSTGRESQL
--  ---------------------------------------------------------------------
--
--  Las extensiones `unaccent` y `pg_trgm` (sección 0) deben habilitarse
--  antes en el parámetro de servidor `azure.extensions`, desde el portal:
--
--      Configuración del servidor -> azure.extensions -> agregar
--      UNACCENT y PG_TRGM -> Guardar
--
--  Si no se habilitan, este archivo NO falla: emite una advertencia y se
--  salta el índice de búsqueda. El sistema funciona igual, pero el buscador
--  de productos por nombre recorre el catálogo completo en cada consulta.
--  Con muchas tiendas eso se vuelve lento, así que conviene habilitarlas.
--
--  ---------------------------------------------------------------------
--  CRITERIOS DE DISEÑO — para quien administre esta base
--  ---------------------------------------------------------------------
--
--  1. NO HAY LLAVES FORÁNEAS entre las tablas de datos de origen, y es a
--     propósito. La información real trae huecos legítimos: hay citas de
--     proveedor sin pedido que las respalde, y SKU con faltante que no
--     están en el catálogo. Una FK rechazaría esas filas, y justamente
--     esos huecos son parte de lo que el análisis tiene que reportar. La
--     validación se hace en la capa de aplicación, que sabe distinguir un
--     hueco esperado de un error de extracción.
--
--     La única FK del modelo es run_dias -> runs, porque ahí sí un detalle
--     sin su corrida es basura: no se sabría de qué periodo salió ni con
--     qué versión del motor.
--
--  2. MUCHAS COLUMNAS SON NULLABLE A PROPÓSITO. El modelo distingue "no
--     había producto" (cero) de "no se midió" (vacío), y esa distinción
--     lleva a conclusiones opuestas. Forzar NOT NULL obligaría a inventar
--     ceros. Donde una columna es nullable pese a que el layout la marca
--     obligatoria, va comentado el porqué.
--
--  3. LAS CARGAS SON IDEMPOTENTES. Cada tabla tiene llave primaria natural
--     y la aplicación inserta con ON CONFLICT DO UPDATE. Volver a cargar la
--     misma entrega, o una reentrega corregida, actualiza en vez de
--     duplicar.
--
-- ###########################################################################


-- ===========================================================================
-- 0. EXTENSIONES
--
-- Sólo las necesita el buscador de productos por nombre: `unaccent` para que
-- "cafe" encuentre "CAFÉ", y `pg_trgm` para que la búsqueda con comodín al
-- inicio ('%cafe%') pueda usar un índice en vez de recorrer todo.
--
-- Si el usuario no tiene permisos para crearlas, se avisa y se continúa.
-- ===========================================================================

DO $$
BEGIN
    CREATE EXTENSION IF NOT EXISTS unaccent;
EXCEPTION WHEN OTHERS THEN
    RAISE WARNING 'No se pudo crear la extensión unaccent (%). El sistema '
                  'funciona, pero la búsqueda de productos no ignorará los '
                  'acentos.', SQLERRM;
END $$;

DO $$
BEGIN
    CREATE EXTENSION IF NOT EXISTS pg_trgm;
EXCEPTION WHEN OTHERS THEN
    RAISE WARNING 'No se pudo crear la extensión pg_trgm (%). El sistema '
                  'funciona, pero la búsqueda de productos por nombre será '
                  'lenta al crecer el catálogo.', SQLERRM;
END $$;


-- ###########################################################################
-- BLOQUE 1 — DATOS DE ORIGEN
--
-- Las nueve tablas que alimenta el proceso de carga con los archivos que La
-- Comer deposita. Son almacenamiento crudo: se guarda cada fila tal como
-- llega, sin transformar.
-- ###########################################################################


-- ===========================================================================
-- 1.1 CATALOGO — maestro operativo de SKU por tienda
--
-- Es la tabla que define QUÉ SKU le tocan al análisis en cada tienda, y por
-- dónde se resurte cada uno. El motor de clasificación la lee para saber la
-- vía de resurtido y qué CEDIS surte a la tienda: sin eso no puede decidir
-- a quién atribuir un faltante.
--
-- Estática: una fila por sku + tienda, se recarga cuando cambia el catálogo.
-- ===========================================================================

CREATE TABLE IF NOT EXISTS catalogo (
    sku               TEXT NOT NULL,
    descripcion       TEXT,
    tienda            TEXT NOT NULL,
    -- Qué CEDIS surte a esta tienda para este SKU. Lo usa el motor para
    -- buscar el inventario y los pedidos a proveedor que corresponden.
    cedis_surtidor    TEXT NOT NULL,
    division          TEXT NOT NULL,
    -- VIA 1 (por CEDIS), VIA 2 o DSD (entrega directa del proveedor a
    -- tienda). Determina qué rama del árbol de causas se recorre.
    via_resurtido     TEXT NOT NULL,
    piezas_por_caja   INTEGER NOT NULL,
    tipo_resurtido    TEXT,
    rol_frecuencia    TEXT,
    linea_vs_io       TEXT,
    estatus_activo    TEXT NOT NULL,
    nombre_tienda     TEXT,
    -- Proveedor principal del SKU. Duplica lo que ya vive en las tablas de
    -- compras y citas, pero permite conocerlo sin depender de que haya una
    -- orden abierta.
    proveedor_id      TEXT,
    proveedor_nombre  TEXT,
    -- Decil de venta del SKU ("Decil 1" a "Decil 10"). Informativa: ninguna
    -- regla del motor la consume, sirve para filtrar el tablero y poder ver
    -- el análisis sólo de los productos de mayor rotación.
    decil             TEXT,
    cargado_en        TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (sku, tienda)
);

CREATE INDEX IF NOT EXISTS ix_catalogo_tienda
    ON catalogo (tienda);
CREATE INDEX IF NOT EXISTS ix_catalogo_cedis_surtidor
    ON catalogo (cedis_surtidor);
CREATE INDEX IF NOT EXISTS ix_catalogo_decil
    ON catalogo (tienda, decil);

-- Búsqueda de producto por nombre desde el tablero. Requiere pg_trgm; si la
-- extensión no está disponible se omite y la búsqueda sigue funcionando,
-- sólo que sin índice.
DO $$
BEGIN
    CREATE INDEX IF NOT EXISTS ix_catalogo_descripcion_trgm
        ON catalogo USING gin (descripcion gin_trgm_ops);
EXCEPTION WHEN OTHERS THEN
    -- No basta con preguntar si la extensión existe: también tiene que estar
    -- en el search_path para que se resuelva gin_trgm_ops. Se intenta y, si
    -- no se puede, se avisa y se sigue: el índice es una optimización.
    RAISE WARNING 'Se omite el índice de búsqueda ix_catalogo_descripcion_trgm '
                  '(%). El sistema funciona; la búsqueda de productos por '
                  'nombre será más lenta al crecer el catálogo. Ver la nota '
                  'de Azure al inicio de este archivo.', SQLERRM;
END $$;


-- ===========================================================================
-- 1.2 BOPS_OSA — disponibilidad en anaquel, día por día
--
-- Es la tabla que dispara todo el análisis: dice qué días un producto no
-- estuvo disponible en el anaquel, y cuánta venta se estima que se perdió.
--
-- Es la tabla más grande del modelo: una fila por SKU, tienda y día.
-- ===========================================================================

CREATE TABLE IF NOT EXISTS bops_osa (
    sku                     TEXT NOT NULL,
    tienda                  TEXT NOT NULL,
    fecha                   DATE NOT NULL,
    -- En la práctica es binaria: 0 = el producto no estuvo visible ese día,
    -- 1 = sí estuvo. El CHECK admite el rango completo por si en el futuro
    -- se entrega como porcentaje real.
    osa_pct                 NUMERIC(5,2) NOT NULL
                            CHECK (osa_pct BETWEEN 0 AND 100),
    venta_perdida_estimada  NUMERIC(14,2) NOT NULL,
    numero_rupturas         INTEGER,
    minutos_sin_producto    INTEGER,
    -- Banderas 0/1 del sistema de alertas de piso. Sólo afinan el detalle
    -- de la causa "Ejecución en Tienda"; no cambian la causa ni el
    -- responsable. Nullable a propósito: sin bandera no se afina nada.
    alerta_enviada          INTEGER,
    alerta_ejecutada        INTEGER,
    cargado_en              TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (sku, tienda, fecha)
);

CREATE INDEX IF NOT EXISTS ix_bops_osa_fecha
    ON bops_osa (fecha);
CREATE INDEX IF NOT EXISTS ix_bops_osa_tienda_fecha
    ON bops_osa (tienda, fecha);


-- ===========================================================================
-- 1.3 TABLEAU_INV_TIENDA — inventario en tienda, día por día
--
-- Responde la primera pregunta del árbol: ¿había producto en la tienda? Si
-- lo había y aun así el anaquel estuvo vacío, el problema es de ejecución en
-- piso y no de abasto.
-- ===========================================================================

CREATE TABLE IF NOT EXISTS tableau_inv_tienda (
    sku                    TEXT NOT NULL,
    tienda                 TEXT NOT NULL,
    fecha                  DATE NOT NULL,
    -- Nullable a propósito: puede venir vacía, y el modelo distingue "no
    -- había" de "no se midió". Forzar NOT NULL obligaría a inventar un cero.
    existencia_piezas      INTEGER,
    hora_de_corte          TEXT,
    -- Existencia mínima del día. Es mejor insumo que la foto de cierre,
    -- porque un producto puede agotarse a media mañana y reponerse en la
    -- tarde, y el cierre no lo mostraría.
    existencia_minima_dia  INTEGER,
    cargado_en             TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (sku, tienda, fecha)
);

CREATE INDEX IF NOT EXISTS ix_inv_tienda_fecha
    ON tableau_inv_tienda (fecha);
CREATE INDEX IF NOT EXISTS ix_inv_tienda_tienda_fecha
    ON tableau_inv_tienda (tienda, fecha);


-- ===========================================================================
-- 1.4 TABLEAU_VENTAS — venta diaria por producto y tienda
-- ===========================================================================

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

CREATE INDEX IF NOT EXISTS ix_ventas_fecha
    ON tableau_ventas (fecha);
CREATE INDEX IF NOT EXISTS ix_ventas_tienda_fecha
    ON tableau_ventas (tienda, fecha);


-- ===========================================================================
-- 1.5 CEDIS_INVENTARIO — existencia en centro de distribución, día por día
--
-- Si la tienda no tenía producto, la siguiente pregunta es si el CEDIS lo
-- tenía. Si lo tenía y no lo envió, la causa es del CEDIS; si no lo tenía,
-- hay que mirar hacia el proveedor.
-- ===========================================================================

CREATE TABLE IF NOT EXISTS cedis_inventario (
    sku                 TEXT NOT NULL,
    cedis               TEXT NOT NULL,
    fecha               DATE NOT NULL,
    existencia_piezas   INTEGER NOT NULL,
    piezas_reservadas   INTEGER,
    cargado_en          TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (sku, cedis, fecha)
);

CREATE INDEX IF NOT EXISTS ix_cedis_inv_fecha
    ON cedis_inventario (fecha);
CREATE INDEX IF NOT EXISTS ix_cedis_inv_cedis_fecha
    ON cedis_inventario (cedis, fecha);


-- ===========================================================================
-- 1.6 CEDIS_TRANSFERENCIAS — envíos del CEDIS a la tienda
--
-- Permite saber si había mercancía en tránsito el día del faltante. Si la
-- había, la causa es de transporte y no de abasto.
-- ===========================================================================

CREATE TABLE IF NOT EXISTS cedis_transferencias (
    folio                     TEXT NOT NULL,
    sku                       TEXT NOT NULL,
    cedis_origen              TEXT NOT NULL,
    tienda_destino            TEXT NOT NULL,
    fecha_generacion          DATE NOT NULL,
    fecha_salida_cedis        DATE NOT NULL,
    -- Nullable pese a que el layout la marca obligatoria: viene vacía
    -- cuando el envío todavía no se había recibido al momento del corte.
    fecha_recepcion_tienda    DATE,
    cantidad_enviada_piezas   INTEGER NOT NULL,
    cantidad_recibida_piezas  INTEGER,
    estatus                   TEXT,
    cargado_en                TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (folio, sku)
);

CREATE INDEX IF NOT EXISTS ix_transferencias_fecha_gen
    ON cedis_transferencias (fecha_generacion);
CREATE INDEX IF NOT EXISTS ix_transferencias_tienda_destino
    ON cedis_transferencias (tienda_destino);
CREATE INDEX IF NOT EXISTS ix_transferencias_cedis_origen
    ON cedis_transferencias (cedis_origen);


-- ===========================================================================
-- 1.7 SIMA_PEDIDOS_TIENDA — pedidos de resurtido que hace la tienda
--
-- Permite saber si alguien pidió el producto. Si nadie lo pidió, ninguna
-- parte de la cadena de abasto pudo haberlo surtido.
-- ===========================================================================

CREATE TABLE IF NOT EXISTS sima_pedidos_tienda (
    folio                    TEXT NOT NULL,
    sku                      TEXT NOT NULL,
    -- OJO CON EL NOMBRE: dice quién GENERÓ el pedido, no para quién es. Trae
    -- el id de la tienda, o el valor '300' cuando lo generó el proceso
    -- centralizado, en cuyo caso el pedido cubre a TODAS las tiendas. Se
    -- llamaba `tienda` en versiones anteriores del layout; se renombró
    -- porque un valor que no siempre es una tienda, llamado `tienda`, es
    -- justo lo que hace que alguien filtre mal seis meses después.
    origen                   TEXT NOT NULL,
    fecha_pedido             DATE NOT NULL,
    fecha_requerida          DATE,
    cantidad_pedida_piezas   INTEGER NOT NULL,
    cantidad_surtida_piezas  INTEGER NOT NULL,
    -- Nullable pese al layout: viene vacía si el pedido sigue abierto.
    fecha_surtido            DATE,
    estatus                  TEXT,
    cargado_en               TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (folio, sku)
);

CREATE INDEX IF NOT EXISTS ix_pedidos_tienda_fecha_pedido
    ON sima_pedidos_tienda (fecha_pedido);
CREATE INDEX IF NOT EXISTS ix_pedidos_tienda_origen
    ON sima_pedidos_tienda (origen);


-- ===========================================================================
-- 1.8 COMPRAS_PEDIDOS_PROV — pedidos de compra al proveedor
--
-- Con esta tabla se juzga al proveedor: cuántas cajas se le pidieron y
-- cuántas entregó.
--
-- El pedido es del CEDIS, no de una tienda: un mismo pedido abastece a
-- todas las tiendas que ese CEDIS surte. Por eso `tienda_destino` casi
-- siempre viene vacía, y por eso estas cifras NO deben sumarse entre
-- tiendas: se contarían varias veces.
-- ===========================================================================

CREATE TABLE IF NOT EXISTS compras_pedidos_prov (
    folio               TEXT NOT NULL,
    sku                 TEXT NOT NULL,
    proveedor_id        TEXT NOT NULL,
    proveedor_nombre    TEXT,
    -- Nullable pese a que el layout la marca obligatoria: medido sobre la
    -- información real, 4,967 de 852,994 pedidos (0.6%) llegan sin CEDIS
    -- destino. Con NOT NULL esas filas se rechazarían en la carga.
    cedis_destino       TEXT,
    fecha_pedido        DATE NOT NULL,
    fecha_cita          DATE,
    -- Nullable a propósito: que venga vacía es normal y significa "todavía
    -- no se recibe". El modelo usa precisamente eso para saber qué pedidos
    -- seguían vigentes el día del faltante.
    fecha_recibo        DATE,
    cajas_pedidas       INTEGER NOT NULL,
    cajas_entregadas    INTEGER NOT NULL,
    estatus             TEXT,
    -- Informativa. Ver la nota de arriba: casi siempre viene vacía.
    tienda_destino      TEXT,
    cargado_en          TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (folio, sku)
);

CREATE INDEX IF NOT EXISTS ix_compras_fecha_pedido
    ON compras_pedidos_prov (fecha_pedido);
CREATE INDEX IF NOT EXISTS ix_compras_proveedor
    ON compras_pedidos_prov (proveedor_id);
CREATE INDEX IF NOT EXISTS ix_compras_cedis_destino
    ON compras_pedidos_prov (cedis_destino);


-- ===========================================================================
-- 1.9 CITAS_PROV_CEDIS — citas de entrega del proveedor en el CEDIS
--
-- Es la fecha comprometida de entrega. Sin cita no hay plazo vencido, y sin
-- plazo vencido no se puede afirmar que el proveedor haya incumplido.
--
-- LA LLAVE ES (folio_cita, sku), no folio_cita sola: una cita es una ventana
-- de entrega y puede cubrir varios productos, cada uno con sus propias cajas
-- confirmadas. Verificado contra información real.
--
-- `folio` (el pedido de compra al que corresponde) es una referencia
-- informativa y NO tiene llave foránea: alrededor del 61% de las citas
-- llegan sin un pedido que les corresponda en compras_pedidos_prov.
-- ===========================================================================

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

CREATE INDEX IF NOT EXISTS ix_citas_folio
    ON citas_prov_cedis (folio);
CREATE INDEX IF NOT EXISTS ix_citas_fecha_cita
    ON citas_prov_cedis (fecha_cita);
CREATE INDEX IF NOT EXISTS ix_citas_proveedor
    ON citas_prov_cedis (proveedor_id);


-- ###########################################################################
-- BLOQUE 2 — CATÁLOGOS DE REFERENCIA
--
-- No los consume el motor de clasificación: sirven para presentar y filtrar
-- los resultados en el tablero.
-- ###########################################################################


-- ===========================================================================
-- 2.1 SUCURSALES — maestro de tiendas del grupo
--
-- La Comer, Fresko, City Market, City Café y Sumesa. Informativa: el motor
-- sigue usando catalogo.cedis_surtidor y catalogo.nombre_tienda, porque a
-- esta tabla le falta qué CEDIS surte a cada tienda.
-- ===========================================================================

CREATE TABLE IF NOT EXISTS sucursales (
    tienda      TEXT NOT NULL,
    formato     TEXT,
    nombre      TEXT,
    direccion   TEXT,
    cp          TEXT,
    cargado_en  TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (tienda)
);

CREATE INDEX IF NOT EXISTS ix_sucursales_formato
    ON sucursales (formato);


-- ===========================================================================
-- 2.2 CATALOGO_SKU_TIENDA — catálogo comercial con la jerarquía completa
--
-- Trae división, sección, categoría, subcategoría y marca, que es por donde
-- el tablero deja filtrar el análisis. NO sustituye a `catalogo`: a esta le
-- falta el CEDIS surtidor, así que el motor de clasificación sigue leyendo
-- la otra. Conviven a propósito.
--
-- Nota de calidad de datos: los valores de división y sección deben cargarse
-- sin el prefijo numérico ("ABARROTES", no "1 - ABARROTES"). Si conviven las
-- dos formas del mismo valor, el tablero las muestra como dos opciones
-- distintas y cada una filtra la mitad de los productos.
-- ===========================================================================

CREATE TABLE IF NOT EXISTS catalogo_sku_tienda (
    sku                 TEXT NOT NULL,
    tienda              TEXT NOT NULL,
    tienda_nombre       TEXT,
    articulo_nombre     TEXT,
    division            TEXT,
    grupo_seccion       TEXT,
    categoria           TEXT,
    subcategoria        TEXT,
    proveedor_id        TEXT,
    proveedor_nombre    TEXT,
    marca               TEXT,
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

CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_tienda
    ON catalogo_sku_tienda (tienda);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_division
    ON catalogo_sku_tienda (division);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_categoria
    ON catalogo_sku_tienda (categoria);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_subcategoria
    ON catalogo_sku_tienda (subcategoria);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_marca
    ON catalogo_sku_tienda (marca);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_proveedor
    ON catalogo_sku_tienda (proveedor_nombre);
CREATE INDEX IF NOT EXISTS ix_catalogo_sku_tienda_proveedor_id
    ON catalogo_sku_tienda (proveedor_id);


-- ###########################################################################
-- BLOQUE 3 — RESULTADOS DEL ANÁLISIS
--
-- Aquí se guarda lo que produce el modelo. Permite volver a consultar un
-- análisis y regenerar su reporte sin recalcularlo: una corrida completa de
-- una tienda y un mes toma alrededor de un minuto de proceso.
-- ###########################################################################


-- ===========================================================================
-- 3.1 RUNS — una fila por análisis ejecutado
--
-- `version_motor` NO es un campo decorativo. Las reglas de negocio cambian
-- cuando La Comer ratifica una decisión, y dos análisis del mismo periodo
-- hechos con semanas de diferencia pueden dar cifras distintas por eso. Sin
-- este campo nadie podría explicar la diferencia meses después. `parametros`
-- guarda, por la misma razón, los interruptores de negocio vigentes en esa
-- corrida.
-- ===========================================================================

CREATE TABLE IF NOT EXISTS runs (
    id              TEXT PRIMARY KEY,
    tienda          TEXT NOT NULL,
    desde           DATE NOT NULL,
    hasta           DATE NOT NULL,
    umbral_osa      NUMERIC NOT NULL DEFAULT 100,

    -- Trazabilidad. Ver el comentario de arriba.
    version_motor   TEXT,
    parametros      JSONB,

    -- Cifras de portada, desnormalizadas para poder listar los análisis sin
    -- abrir el resumen completo: leer varios MB de JSON para pintar un
    -- renglón de tabla sería absurdo.
    osa_alcance     NUMERIC,
    dias_faltante   INTEGER,
    venta_perdida   NUMERIC,
    cobertura_pct   NUMERIC,

    -- El resultado completo que consume el tablero. Postgres lo comprime
    -- solo (TOAST): unos 6 MB en claro quedan en cerca de 500 KB en disco.
    resumen         JSONB NOT NULL,

    -- (sku|tienda) -> [días medidos, días visibles]. Es el denominador del
    -- indicador por producto y NO se puede reconstruir desde el detalle
    -- diario, porque los días sanos no llegan ahí.
    universo        JSONB,

    corrido_en      TIMESTAMPTZ NOT NULL DEFAULT now(),
    segundos        NUMERIC,
    origen          TEXT NOT NULL DEFAULT 'bd',   -- 'bd' | 'archivo'
    archivo         TEXT
);

-- El listado de la pantalla inicial: lo más reciente primero.
CREATE INDEX IF NOT EXISTS ix_runs_corrido_en
    ON runs (corrido_en DESC);
-- "¿ya se analizó esta tienda en este periodo?", para ofrecer el análisis
-- existente en vez de volver a calcular lo mismo.
CREATE INDEX IF NOT EXISTS ix_runs_tienda_periodo
    ON runs (tienda, desde, hasta);


-- ===========================================================================
-- 3.2 RUN_DIAS — el veredicto de cada día, con su evidencia
--
-- Es la hoja "Clasificación diaria" del reporte, fila por fila: la causa
-- raíz de cada día con faltante y toda la evidencia que llevó a ella.
--
-- Sirve para dos cosas: regenerar el reporte sin volver a leer los archivos
-- de origen, y consultar resultados directamente ("todos los faltantes por
-- incumplimiento de proveedor de marzo"), que dentro de un archivo de varios
-- MB no se podía.
--
-- ES LA TABLA QUE MÁS CRECE junto con bops_osa. Se borra en cascada con su
-- análisis: un detalle sin la corrida que lo produjo no dice de qué periodo
-- salió ni con qué versión del motor, y deja de ser información.
-- ===========================================================================

CREATE TABLE IF NOT EXISTS run_dias (
    run_id          TEXT NOT NULL REFERENCES runs(id) ON DELETE CASCADE,

    sku             TEXT NOT NULL,
    tienda          TEXT NOT NULL,
    fecha           DATE NOT NULL,
    osa             NUMERIC,
    venta_perdida   NUMERIC,

    -- La evidencia, en el mismo orden en que la evalúa el árbol de causas.
    -- Los booleanos son de TRES estados a propósito: NULL significa "no se
    -- sabe", que no es lo mismo que false. Todo el modelo se apoya en esa
    -- distinción.
    inventario_tienda           INTEGER,
    transito_vigente            BOOLEAN,
    pedido_tienda_generado      BOOLEAN,
    tipo_resurtido              TEXT,
    via_resurtido               TEXT,
    inventario_cedis            INTEGER,
    envio_cedis_generado        BOOLEAN,
    pedido_proveedor_generado   BOOLEAN,
    cajas_pedidas               INTEGER,
    -- La cita en crudo, no como el texto que se pinta en el reporte: así se
    -- puede volver a dar formato igual que en la corrida original, y de paso
    -- consultar cosas como "cuántas citas vencidas hubo en marzo".
    cita_agendada               BOOLEAN,
    cita_fecha                  DATE,
    cita_vencida                BOOLEAN,
    cajas_confirmadas           INTEGER,
    cajas_entregadas            INTEGER,

    -- El dictamen.
    root_cause_id   TEXT,
    causa_raiz      TEXT,
    responsable     TEXT,
    subcausa        TEXT,
    prioridad_regla INTEGER,
    fuente          TEXT,
    detalle         TEXT,
    -- Qué dato faltó, cuando faltó. Es lista porque una regla puede quedarse
    -- esperando más de uno.
    datos_faltantes TEXT[]
);

-- Regenerar el reporte: se leen todos los días de un análisis en orden.
CREATE INDEX IF NOT EXISTS ix_run_dias_run
    ON run_dias (run_id, sku, fecha);
-- Las consultas que motivaron la tabla: "las causas de este análisis",
-- "este producto a lo largo del periodo".
CREATE INDEX IF NOT EXISTS ix_run_dias_causa
    ON run_dias (run_id, root_cause_id);
CREATE INDEX IF NOT EXISTS ix_run_dias_sku
    ON run_dias (sku, tienda, fecha);


-- ###########################################################################
-- BLOQUE 4 — LINAJE
-- ###########################################################################


-- ===========================================================================
-- 4.1 ETL_CARGAS — bitácora de cada carga de información
--
-- Registra qué archivo se cargó, cuándo, cuántas filas por hoja y qué
-- advertencias salieron. Es lo que permite responder "¿por qué el análisis
-- de esta tienda salió raro?" mirando si la última carga vino completa.
-- ===========================================================================

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

CREATE INDEX IF NOT EXISTS ix_etl_cargas_iniciado_en
    ON etl_cargas (iniciado_en DESC);


-- ###########################################################################
-- VERIFICACIÓN
--
-- Imprime lo que quedó creado. Deben aparecer 14 tablas.
-- ###########################################################################

DO $$
DECLARE
    n_tablas   INTEGER;
    n_indices  INTEGER;
    faltantes  TEXT;
BEGIN
    SELECT count(*) INTO n_tablas
      FROM information_schema.tables
     WHERE table_schema = 'public'
       AND table_name IN ('catalogo','bops_osa','tableau_inv_tienda',
                          'tableau_ventas','cedis_inventario',
                          'cedis_transferencias','sima_pedidos_tienda',
                          'compras_pedidos_prov','citas_prov_cedis',
                          'sucursales','catalogo_sku_tienda','runs',
                          'run_dias','etl_cargas');

    SELECT count(*) INTO n_indices
      FROM pg_indexes WHERE schemaname = 'public';

    SELECT string_agg(t, ', ') INTO faltantes
      FROM unnest(ARRAY['catalogo','bops_osa','tableau_inv_tienda',
                        'tableau_ventas','cedis_inventario',
                        'cedis_transferencias','sima_pedidos_tienda',
                        'compras_pedidos_prov','citas_prov_cedis',
                        'sucursales','catalogo_sku_tienda','runs',
                        'run_dias','etl_cargas']) AS t
     WHERE NOT EXISTS (SELECT 1 FROM information_schema.tables
                        WHERE table_schema = 'public' AND table_name = t);

    RAISE NOTICE '';
    RAISE NOTICE '=====================================================';
    RAISE NOTICE ' ORCMM — estructura instalada';
    RAISE NOTICE '=====================================================';
    RAISE NOTICE ' Tablas creadas : % de 14', n_tablas;
    RAISE NOTICE ' Indices totales: %', n_indices;

    IF n_tablas = 14 THEN
        RAISE NOTICE ' Resultado      : CORRECTO';
    ELSE
        RAISE NOTICE ' Resultado      : INCOMPLETO — faltan: %', faltantes;
    END IF;
    RAISE NOTICE '=====================================================';
    RAISE NOTICE '';
END $$;
