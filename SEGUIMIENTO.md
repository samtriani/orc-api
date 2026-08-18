# Seguimiento de sesiones — Claude Code

Este proyecto se trabaja desde dos máquinas distintas. Este archivo es la
memoria entre sesiones: **al empezar, léelo**; **al cerrar, agrega una
entrada nueva arriba** con qué cambió y qué quedó pendiente. No se borran
las entradas viejas — así queda el historial de decisiones.

> **Regla nueva de esta sesión**: antes de comitear/subir cualquier cambio,
> avisar primero qué se va a subir (archivos + resumen) y esperar luz
> verde — aunque el trabajo ya venga pedido. El usuario lo pidió explícito
> porque pedir "haz X" no es lo mismo que autorizar subirlo a git.

## Cómo está repartido el proyecto

- **`orc-api`** (este repo) — motor de clasificación (`orcmm_rca_engine.py`),
  pipeline (`orcmm_pipeline.py`), API FastAPI (`api/`). Desplegado en Fly.io:
  https://orc-api.fly.dev
- **Postgres (Neon)** — capa raw persistente de las 9 hojas del layout +
  2 catálogos informativos. `DATABASE_URL` vive en `orc-api/.env` (git-
  ignorado, nunca se sube) y, en Fly.io, como `fly secrets set
  DATABASE_URL=...`. Esquema en `sql/schema.sql`, se aplica con `python
  orcmm_db_init.py`. El motor de clasificación **ya puede leer de aquí**:
  `POST /api/analizar-tienda` (tienda + periodo, sin subir archivo) además
  del `POST /api/analizar` de siempre (sube archivo). Ver sesión 2026-08-10.
- **`orc-gui`** ([github.com/samtriani/orc-gui](https://github.com/samtriani/orc-gui)) —
  front Angular. Desplegado en Vercel: https://orc-gui.vercel.app — hace
  *rewrite* server-side de `/api/*` hacia Fly.io, así que no hay CORS de por
  medio entre el navegador y el back.

---

## Sesión 2026-08-18 — el dashboard se vendió, prioridad 3 prendida y el OSA que siempre daba 0

El cliente aprobó el dashboard y mandó **20 peticiones de mejora**. Se
hicieron 15 (orc-api `d5f5be5`, `74fb1a5` · orc-gui `a2d7b23`).

**Lo que mueve números — avisar antes de comparar contra un reporte viejo:**

1. **Prioridad 3 PRENDIDA** (`EVALUAR_PEDIDO_TIENDA = True`). SIMA ya
   entrega. Los días que le tocaban a RC03 se estaban repartiendo entre
   RC04/RC05/RC06: se le cobraba a CEDIS y al proveedor un faltante que se
   explicaba porque la tienda no pidió. **El Pareto se mueve.** Probado:
   sin pedido → RC03/Tienda; sin dato → RC99 bloqueado por
   `pedido_tienda_generado`, **no** culpa a la tienda.
2. **OSA PROM salía 0.0 en TODOS los renglones** — lo cachó el cliente al
   descargar el Excel que se le mandó a Flor. No era dato malo:
   `osa_promedio` promedia los días CON faltante, y como BOPS reporta OSA
   **binario** (0 = no visible), esos días valen 0 por definición.
   Verificado: 332 de 332 SKU en 0.0. Ahora `osa_periodo` = días visibles /
   días **medidos** (no del calendario: BOPS trae ~31.6 de 43 días por SKU,
   y afirmar algo de un día sin fila sería inventarlo).
   **El Excel cambió de columnas**: "OSA prom." → "días eval." + "OSA del
   periodo".
3. **Proveedores: de 648 a 643.** Nestlé salía dos veces —dos IDs
   distintos, 16661 y 16653—. La clave de consolidación normaliza a sólo
   letras y números: quitar puntuación no basta, `S.A. DE C.V.` se vuelve
   `S A DE C V` y no empata con `SA DE CV`. Aparecieron 5 más que nadie
   había notado: Herdez, La Costeña, PepsiCo, Mondelez.

**Datos y esquema:**

- SIMA entrega la columna como **`numero_pedido`**; el spec la llama
  `folio`. Entra por `ALIAS_ENCABEZADOS`, así el archivo carga tal como
  sale de SIMA. **Confirmado con el cliente: son PIEZAS individuales, no
  cajas.**
- `schema.sql` estaba **desfasado** de `migracion_v8.sql`: faltaban las
  banderas de alerta, `proveedor_*` en catálogo, `tienda_destino` en
  compras, y `existencia_piezas` seguía `NOT NULL`. Ya calza con el spec
  en las 9 tablas.

**PENDIENTE CRÍTICO — el OOM.** La corrida completa de Coyoacán **mata la
máquina** (`exit_code=137, oom_killed=true`), y los 4 GB de `fffefd6` no
alcanzaron. Medido por etapas: `leer_fuentes_db` se lleva el **78%** (1.9
GB) y el Excel de openpyxl el resto. La tabla grande es **BOPS_OSA con
886,313 filas** para una tienda en 43 días (no compras, que son 186 mil).
El recorte de columnas quedó **verificado byte-idéntico** (JSON de
1,180,978 bytes con y sin recorte) pero **NO está aplicado** — hay que
rehacerlo. Y ojo: el recorte solo no basta, el Excel necesita su propia
solución.

**Lo que sigue, con decisión del cliente pendiente:**

- **Filtros División / Sección / Categoría / Subcategoría**: sólo existe
  `division` (y `grupo_seccion` en `catalogo_sku_tienda`). **Categoría y
  Subcategoría no están en ninguna tabla** — hay que pedirlas.
- **Bottom 10 de proveedores + "Nivel de Servicio"**: hoy hay **tres**
  definiciones distintas (`pct_surtido_pedido`, `pct_efectivo`,
  `pct_cumplimiento`) y el ranking cambia según cuál mande. Falta que el
  cliente elija. "OSA del proveedor" no existe: hay que construirlo con el
  link SKU→proveedor que el V8 trajo al catálogo.
- **Pedido de tienda por día con cantidad** en el expediente: el índice
  `pedidos_tienda_por` sólo guarda fechas, no cantidades; agregarlas toca
  la ruta de clasificación. Se dejó para cuando SIMA esté cargada y se
  pueda verificar de punta a punta.

---

## Sesión 2026-08-11 (noche) — filtros que sí llegan al Pareto, y la grafica arreglada

**Se hizo:**

1. **El waterfall y los dos Pareto ya reaccionan a los filtros** — era el
   pendiente de la entrada de abajo. El backend manda `detalle_dias`
   comprimido (causas en catálogo aparte, cada día con su índice y campos de
   una letra): ~5,200 renglones en ~200 KB en vez de ~700 KB. Se pide una
   sola vez con el resumen; el poll sigue en 85 bytes.
   → orc-api `2cde26b` (Fly **v18**) · orc-gui `5351239`
   - El denominador del waterfall se recompone con el universo de BOPS del
     SKU/tienda filtrado, que va desglosado en la respuesta. Sin eso, filtrar
     a un SKU dejaría los puntos de OSA sobre el universo de la tienda.
   - **Filtrar por causa NO cambia el denominador**: el universo lo define
     qué SKU estás mirando. Si se encogiera a los días de esa causa, siempre
     daría 100%.
   - Sin filtros se usan tal cual las listas del backend, para no reproducir
     sus redondeos con otra aritmética.
2. **Evolución diaria: las escalas estaban mal de raíz.** `maximoDia()`
   escalaba cada serie contra su propio máximo — tres reglas en el mismo
   dibujo, así que un inventario de 3 podía verse más alto que una venta de
   15. Ahora tienda y venta comparten escala (mismo lugar, misma unidad) y
   **CEDIS va en su propia fila**, porque medido sobre los datos reales es de
   10 a 500 veces mayor (un SKU con 4,421 en tienda tiene 105,504 en CEDIS);
   en la misma regla la venta quedaría en 0.18% de altura. Cada fila escribe
   su máximo. → orc-gui `38d164d`
3. **Colores de la tira de causa raíz**, que no se veían: los pasteles del
   Excel en cuadros de 8 px sobre blanco. Ahora saturados y reasignados **por
   frecuencia**, para que las dos que de verdad coinciden queden lo más
   separadas posible: RC01 (94.1%) naranja contra RC06 (5.3%) azul miden
   ΔE 34.8. **RC99 pasa a rojo oscuro `#a10c22`** —es una alarma, no una
   causa— y no al rojo de siempre, que queda a ΔE 12.7 del naranja de RC01,
   bajo el piso de 15. "Sin faltante" pasa a verde pálido: antes era el mismo
   gris que RC99 y no se distinguían.
   **La leyenda de causas del expediente es obligatoria, no decorativa**: con
   siete causas ninguna combinación clarea la separación para daltonismo con
   todos los pares. Si se quita, la gráfica queda mal.
4. Autocompletar: **elegir una sugerencia con el mouse ya agrega la ficha.**
   Sólo se agregaba con Enter, y el `<datalist>` dispara `input`, nunca
   `keydown`. → orc-gui `a23c579`
5. Los dos paneles de tendencia quedan **apagados tras
   `MOSTRAR_PANELES_SIN_DATOS`** (orc-gui `2bdc0b6`). El maquetado sigue vivo;
   volver a encenderlos es una línea cuando lleguen los meses.

**Medido — la base NO es el cuello, y no hacen falta índices:**

```
compras_pedidos_prov WHERE cedis_destino='280'  →  Seq Scan · 1,469 ms · 848,027 filas
citas_prov_cedis     WHERE cedis_destino='280'  →  Seq Scan ·    71 ms ·  37,625 filas
```

`ix_compras_cedis_destino` **ya existe**; Postgres elige Seq Scan a propósito
porque las 848,027 filas comparten CEDIS y el filtro no descarta nada. Segundo
y medio en total: los minutos de una corrida se van en **Python**, no en SQL.
Si algún día hay que acelerar esto, es por ahí — empezando por medir
`_es_vacio()`, que corre en cada campo de cada fila y sigue sin cronometrarse.

**Pendiente para mañana — el scorecard de proveedor con filtros:**

Se puede y NO sería mentir: cada pedido trae columna `sku` y el cruce con
citas ya es por `(folio, sku)` (`orcmm_pipeline.py:911`). Haría falta mandar
el desglose por **(proveedor, sku)** —los mismos ocho contadores de
`desempeno_proveedores`— acotado a los 496 SKU con faltante, que son los
únicos que ofrece el autocompletar: ~50 KB.

Lo que NO se puede: filtrar por **causa raíz o responsable** (un pedido no
tiene causa; nace de un día con faltante en anaquel, otro universo) ni por
**tienda** (los pedidos van al CEDIS, columna `cedis_destino`).

**Y hay que decidir algo antes de hacerlo**: hoy el scorecard mide al
proveedor sobre TODOS sus pedidos a propósito —su docstring dice que un
proveedor puede incumplir sin que se note en anaquel—. Filtrado por SKU pasa
a responder otra pregunta, y el fill rate de portada seguiría siendo el
global: quedarían dos números que no cuadran en la misma pantalla. La tabla
tendría que **anunciar** que está recortada, para que nadie lea un 40% de un
proveedor creyendo que es su desempeño general.

**Otro pendiente:** el filtrado nuevo **no se verificó de punta a punta** —
compila y la lógica está revisada, pero no se corrió un análisis completo para
confirmar que `detalle_dias` llega bien. Si al mover un filtro el waterfall no
se inmuta, ahí está la pista.

---

## Sesión 2026-08-11 — moneda, filtros con varios valores, y qué falta por conectar

Sesión corta de ajustes al front, sobre lo que dejó la sesión anterior
(cola/cancelación, split de `/resumen`, identidad de marca).

**Se hizo:**

1. **Venta perdida con `$`** en las 7 tarjetas/tablas donde aparece
   (portada, ficha de SKU, waterfall, Pareto, subcausas, bloqueos, detalle
   por SKU-tienda). Prefijo literal en la plantilla, no `CurrencyPipe` —
   más simple y sin depender de datos de locale que no están registrados.
2. **Filtros de SKU y proveedor aceptan varios valores a la vez** (fichas
   removibles, "cualquiera de estos"), con autocompletar nativo
   (`<datalist>`, sin librería nueva) sugiriendo de lo que ya trae el
   resultado. Sigue aceptando texto libre que no esté en la lista.
3. **El filtro de SKU ahora busca por nombre además de por código.** El
   backend no mandaba el nombre en `por_sku_tienda` — se agregó
   `descripcion` (de `CATALOGO`, ya estaba en memoria, cero consultas
   nuevas). Las sugerencias del datalist muestran "código — nombre"; al
   elegir una (o Enter) la ficha se queda sólo con el código.
   → orc-api `0d2c0a5`, desplegado y verificado en producción.
4. **"SKU que más costaron" (portada) ahora sí reacciona a los filtros**
   — antes leía la lista cruda del backend, ignorando lo que estuviera
   filtrado. Sin riesgo: cada renglón ya trae su propia venta_perdida, no
   había nada que recalcular mal.

**Pendiente — decisión para esta noche, en la otra máquina:**

- **El waterfall ("Causas raíz de faltantes") y el Pareto por
  causa/responsable siguen SIN reaccionar a los filtros**, y a propósito
  no se tocaron hoy. Se verificó que no se pueden recalcular en el
  navegador desde `por_sku_tienda`: esa tabla sólo trae **la causa
  dominante por SKU** (la que ganó en sus días), no el detalle día por
  día. Un SKU con RC01 unos días y RC06 otros se ve como "100% RC01" en
  esa tabla — recalcular el Pareto desde ahí daría números
  **equivocados**, no sólo desactualizados.
  Arreglarlo bien exige que el backend mande el detalle diario al front,
  y eso choca de frente con el trabajo de la sesión anterior de
  **aligerar** esa misma respuesta (se partió en `/resumen` porque pesaba
  1.5MB y tardaba 17-30s por el proxy de Vercel). Hay que decidir el
  trade-off a propósito — mandar el detalle sólo para el periodo/tienda
  ya analizado, paginarlo, un endpoint aparte bajo demanda, etc. — no
  sumarlo sin más.

---

## Sesión 2026-08-10 (continuación 3) — la serie de venta del expediente salía vacía

Revisando el expediente ya desplegado se vio que **`unidades_vendidas` venía
`None` en los 43 días de los dos SKU de prueba**, o sea que la barra de venta
de la gráfica estaba permanentemente en blanco y `altoBarraDia` dividía contra
un máximo de 0. No se había notado porque sólo se había verificado el conteo
de días con causa raíz, que sí cuadraba.

**Consultado directo a Neon** (`flyctl ssh console`, la máquina ya tiene el
`DATABASE_URL`):

```
tableau_ventas    filas: 254,986
  con importe_venta:    254,986   (100%)
  con unidades_vendidas:      0   (0%)
```

La columna existe en `sql/schema.sql` y es nullable, pero **nunca se llena**:
el export de Tableau manda la columna de la métrica **sin encabezado** y
`orcmm_fuentes_csv.py:90` la aterriza en `importe_venta`. El nombre no
corresponde — **los valores son piezas, no pesos**: 3.00 de un papel higiénico
Petalo de 18 rollos son tres paquetes, no tres pesos. Confirmado con el
usuario, que lee esa columna como unidades vendidas del día.

**Se hizo:** `_vendidas()` en `orcmm_expediente_db.py` — prefiere
`unidades_vendidas` para el día que el export sí la traiga, y mientras tanto
cae a `importe_venta`, que es donde el dato realmente está. **El front no se
tocó**: su etiqueta "Unidades vendidas" ya era la correcta. Un archivo,
20 líneas, sin migración ni recarga — el dato ya viajaba desde Postgres
(la consulta es `SELECT *`), sólo no se copiaba al JSON.

Desplegado y verificado en producción (**v11**): el SKU sano pasó de 0 a
**33 de 43 días con venta**, y se lee coherente — el inventario baja conforme
se vende (23 → 21 → 21 → 20). Los 10 días sin dato son días sin registro de
venta, no un error. El contraste con el vino sigue siendo el mismo:
existencia congelada en 3.00 y **una sola fila de venta en todo el periodo**
(31-mar, 0.00), lo que refuerza la sospecha de inventario fantasma.

**Pendiente:** el arreglo de fondo es renombrar el mapeo en
`orcmm_fuentes_csv.py:90` para que el dato aterrice en `unidades_vendidas`,
pero eso obliga a **marcar `importe_venta` opcional en el spec** (hoy está
como obligatoria, `orcmm_layout_spec.py:132`) y a **recargar las 254,986
filas**. Hoy funciona con el fallback. Lo que hay que pedirle a Andrés es que
el export **nombre la columna**; mientras llegue sin encabezado, cualquier
mapeo es una suposición.

---

## Sesión 2026-08-10 (continuación 2) — detalle diario por SKU (expediente)

El cliente vendió un dashboard ejecutivo de OSA (bocetos compartidos por el
usuario) que incluye una vista "Evolución Diaria" por SKU. De las piezas que
faltaban para esa historia, ésta era la única que ya se podía construir con
los datos que hay hoy (las otras — tendencia de 6 meses, comparación entre
tiendas — necesitan datos que todavía no llegan). El usuario la eligió como
la primera a construir.

**Se hizo:**

1. **`orcmm_fuentes_db.leer_fuentes_db` gana un parámetro `sku` opcional.**
   Con él, todas las consultas se acotan también por SKU, y
   `TABLEAU_INV_TIENDA`/`TABLEAU_VENTAS` cambian de estrategia: en vez del
   filtro "sólo días con faltante" (necesario para una tienda completa),
   traen TODOS los días del rango — un solo SKU es volumen mínimo, y el
   detalle diario necesita también los días sanos.
2. **`orcmm_expediente_db.py`** (nuevo) — arma el JSON día por día
   reutilizando las MISMAS funciones que ya usa el motor para clasificar
   (`derivar_transito_vigente`, `derivar_envio_generado`,
   `derivar_pedido_tienda`, `derivar_orden_proveedor`, todas de
   `orcmm_pipeline.py`) — así la gráfica explica exactamente lo que dice el
   veredicto, no una aproximación aparte. Es el hermano JSON/Postgres de
   `orcmm_expediente.py` (CLI, Excel/CSV, consola) — no se tocó ese archivo.
3. **`GET /api/expediente?tienda=&sku=&desde=&hasta=&umbral_osa=`** — sin
   cola (`run_in_threadpool`, igual que `/api/tiendas`): un solo SKU
   responde en ~2s, no hace falta poll.
4. **Verificado contra datos reales**: comparé el conteo de días con causa
   raíz del expediente de un SKU contra su fila en `por_sku_tienda` de un
   análisis completo de la misma tienda/periodo — coincide exacto (42 días
   con faltante, 42 clasificados, RC01 en los dos caminos).
5. **Front**: botón "Evolución diaria" en la tabla "Por SKU-Tienda" (sólo en
   modo "tienda" — un resultado por archivo puede traer un SKU que ni
   siquiera esté en Postgres). Gráfica nueva sin librería (mismo criterio
   que el waterfall ya existente: barras vía `[style.height]` escaladas
   contra el máximo local de cada métrica), con la tira de causa raíz por
   día reutilizando la paleta `COLOR_CAUSA` ya existente — nada de colores
   nuevos.
6. `angular.json`: subido el presupuesto de CSS por componente de 8kB a
   12kB — `app.css` ya lo rebasaba por la acumulación legítima de esta
   sesión (marca, selector de tienda, ahora esta gráfica); el techo de
   error sigue en 16kB.
7. **Entre medio**: llegó una entrega V6 (`OneDrive_1_8-10-2026_v6/`),
   misma tienda/periodo que la V5, conteos idénticos — se recargó con
   `orcmm_etl_carga.py` (esta vez sin necesitar `--forzar`: V6 ya trae
   `fecha_recibo` sin el hueco que documentaba la sesión anterior). Antes
   de recargar se vació la base con `orcmm_db_borrar.py --si`, que ahora
   por omisión **no** toca `sucursales`/`catalogo_sku_tienda` (tienen
   ciclo de vida propio) — `--con-catalogos` los incluye si hace falta.

**`orc-api` desplegado y verificado en producción** (mismo día):
`/api/expediente`, `/api/tiendas` y `/api/analizar-tienda` dan en
`https://orc-api.fly.dev` exactamente los mismos números que en local
(43 días, 42 con causa, para el SKU de prueba). El front (`orc-gui`)
**sigue sin subirse a Vercel** — el push a `main` ya activaría el deploy
automático, sólo falta confirmarlo/probarlo ahí.

**Pendiente para la siguiente sesión:**

- Confirmar que Vercel desplegó `orc-gui` solo con el push a `main` y
  probar la gráfica ahí.
- Falta ver esto en un navegador real (sin herramienta de screenshot en
  este entorno, sólo se verificó por API/build) — revisar que la gráfica
  se vea bien, que el scroll horizontal funcione con periodos largos, y que
  las marcas PT/CD/PV se entiendan sin la leyenda a la mano.
- De las piezas que le faltan a la historia completa del dashboard
  (bocetos del cliente): sigue pendiente la tendencia de 6 meses (necesita
  más historial cargado) y la comparación entre tiendas (necesita las
  otras 4 tiendas con datos operativos, no sólo su catálogo de SKU).

---

## Sesión 2026-08-10 (continuación) — analizar directo desde Postgres

Fase 2 de la entrada de abajo: el motor de clasificación ya puede correr
contra la base en vez de un archivo subido. Se hizo en la misma sesión,
apoyado en el esquema ya cargado.

**Se hizo:**

1. **`orcmm_fuentes_db.py`** — arma un `Fuentes` con `SELECT` en vez de
   `openpyxl`. Reutiliza `_indexar_eventos`, `_revisar_cobertura_de_citas`,
   `aviso_prioridad_3` tal cual (son funciones planas sobre `Fuentes`, no
   Excel-específicas). `catalogo`/`bops_osa`/`tableau_inv_tienda`/
   `tableau_ventas` se filtran por tienda; `cedis_inventario`,
   `compras_pedidos_prov` y `citas_prov_cedis` NO tienen columna `tienda` —
   se filtran por el `cedis_surtidor` de esa tienda (`catalogo`). Las 4
   tablas de evento llevan 30 días de lookback antes de `desde` (mismo
   criterio que ya documentaba el layout de Excel: "un tránsito de febrero
   explica faltantes del 1 de marzo").
2. **`api/servicio.py::analizar()`** gana un parámetro opcional `fu`: si ya
   viene armado, se usa tal cual y se salta `leer_fuentes`. Todo lo de ahí
   en adelante (`derivar_evidencias`, `clasificar`, `escribir_resultado`)
   es 100% agnóstico del origen — cero cambios necesarios.
3. **`POST /api/analizar-tienda`** (tienda + desde + hasta, JSON) y
   **`GET /api/tiendas`** (catálogo para el selector del front, sólo
   tiendas que ya tienen `BOPS_OSA` cargado — hoy nada más 287/Coyoacán).
   Reutiliza `Trabajo`/`TRABAJOS`/`CANDADO`/`EJECUTOR`/`_limpiar_viejos`
   sin tocarlos; `GET /api/analizar/{id}` y `GET /api/resultado/{id}`
   tampoco cambiaron. Sin `diagnosticar_layout`/`corregir`/`forzar`: no hay
   archivo que validar, el ETL ya lo hizo al cargar.
4. **Bug de rendimiento encontrado y corregido tras probarlo con datos
   reales**: la primera versión traía `TABLEAU_INV_TIENDA`/
   `TABLEAU_VENTAS` filtradas sólo por tienda+fecha, sin aplicar el mismo
   filtro de "sólo los días con faltante real" que ya usa el flujo de CSV
   (`orcmm_fuentes_csv.llaves_con_faltante`). Resultado medido: 2.66M de
   2.72M filas traídas para nada, 97 de 119s de una corrida completa.
   Corregido con un `JOIN unnest(...)` contra las llaves `(sku, fecha)` que
   `BOPS_OSA` ya marcó con faltante — **de 119s a 16s**, mismo resultado
   exacto (verificado número por número). Lección: "SQL ya filtra por
   tienda/fecha" no es lo mismo que "SQL filtra por lo que el modelo
   realmente necesita" — el filtro que importaba era el segundo.
5. **`api/main.py` ahora carga `.env`** (`load_dotenv()` al importar) — sin
   esto `DATABASE_URL` no existía al correr `uvicorn` en local. En Fly.io
   no hace nada (no hay `.env`, la variable ya viene del secreto).
6. Front (`orc-gui`): nuevo modo "Elegir tienda y periodo" junto al de
   subir archivo (aditivo, no lo reemplaza). Tienda de selección única y
   obligatoria — el análisis siempre corre sobre una sola. Ver
   `SEGUIMIENTO.md` de `orc-gui` / los commits del front para el detalle.
7. **Ojo con `--reload` de uvicorn en Windows**: ralentiza mucho las
   corridas (vigila archivos todo el tiempo, compite por CPU con el hilo
   que arma el `Fuentes`). En local, correr sin `--reload` salvo que se
   esté editando código activamente.

**Desplegado y verificado en producción** (mismo día): `fly secrets set
DATABASE_URL=...` + `flyctl deploy --ha=false`. `GET /api/tiendas` y
`POST /api/analizar-tienda` contra `https://orc-api.fly.dev` dan
exactamente los mismos números que en local (`osa_general` 74.5%,
`casos_totales` 30565, 94.1% RC01) y la descarga del Excel funciona.

**Pendiente para la siguiente sesión:**

- Sólo hay una tienda cargada (287). Cuando haya más, vale la pena volver a
  medir tiempos: hoy `catalogo`/`compras_pedidos_prov`/`citas_prov_cedis`
  casi no se benefician del filtro por tienda porque todo lo cargado es de
  una sola tienda — con varias tiendas cargadas esas consultas también
  deberían acelerarse solas.
- `compras_pedidos_prov` con llave `(folio, sku)` sigue con la limitación
  conocida (~0.6% de filas con múltiples fechas de recibo, se queda con la
  última) — documentada en la sesión anterior, no se tocó.

---

## Sesión 2026-08-10 — persistencia en Postgres (Neon): esquema + ETL

El usuario ya tiene una instancia de Postgres en Neon y quiere dejar de
re-subir el Excel en cada corrida. Esta sesión se acotó a diseñar el
esquema y el ETL que carga las fuentes ahí — **el motor de clasificación
sigue leyendo del Excel/CSV subido, no de la base** (eso es la fase 2, ver
pendientes). El upload en `orc-gui` sigue igual; esto es aditivo.

**Se hizo:**

1. **Esquema raw** (`sql/schema.sql`, aplicado con `orcmm_db_init.py`,
   idempotente — todo `IF NOT EXISTS`): una tabla por cada una de las 9
   hojas del layout, llave primaria = llave natural de cada una (sin ID
   sustituto), sin `FOREIGN KEY` entre hojas a propósito (huecos de
   cobertura legítimos y ya medidos, p. ej. 61% de citas sin match en
   compras). Más `etl_cargas` (bitácora de cada corrida).
2. **`orcmm_etl_carga.py`** — CLI que carga el `.xlsx` + los CSV sueltos a
   Postgres por `UPSERT` (idempotente). Reutiliza el mismo parseo que ya
   usa el pipeline (`leer_hoja`, `leer_csv`, los 4 conversores de tipo,
   `validar_archivo`) — no se reescribió nada de eso. A diferencia del
   pipeline de clasificación, llama a `leer_csv` con `llaves=None`: guarda
   TODAS las filas, no sólo los días con faltante.
3. **Bug real encontrado y corregido en el diseño**: la llave de
   `CITAS_PROV_CEDIS` se había definido como `folio_cita` solo, por una
   lectura literal de "único por cita" en el spec. Al cargar la entrega V5
   real (39,840 filas) colapsó a 3,922 — **90% de pérdida silenciosa**,
   porque una cita sí cubre varios SKU (verificado: folio_cita 957821 trae
   7 SKU distintos). Corregido a `(folio_cita, sku)` → 37,625 filas reales.
   Lección: no confiar en la redacción del spec para la llave sin
   verificarla contra el dato real cuando hay ambigüedad.
4. **Límite conocido, no corregido a propósito**: `COMPRAS_PEDIDOS_PROV`
   con llave `(folio, sku)` pierde ~0.6% de filas (4,967 de 848,027) por
   entregas parciales del mismo pedido en fechas de recibo distintas. No
   se agregó `fecha_recibo` a la llave porque los ~120,816 pedidos sin
   recibir comparten `NULL`, y Postgres no considera dos `NULL` iguales
   para el `UNIQUE` — eso habría roto la idempotencia (cada re-carga
   generaría filas nuevas en vez de pisar las existentes). Se documenta
   como límite aceptado, no como bug.
5. **Carga real contra la entrega V5** (`OneDrive_1_8-10-2026/`, Excel de
   61 MB + 7 CSV, ~250 MB): CATALOGO 26,407 · TABLEAU_INV_TIENDA 2,719,780 ·
   BOPS_OSA 119,958 · TABLEAU_VENTAS 254,986 · CEDIS_INVENTARIO 23,349 ·
   CEDIS_TRANSFERENCIAS 30,000 · SIMA_PEDIDOS_TIENDA 0 (sigue vacía) ·
   COMPRAS_PEDIDOS_PROV 848,027 · CITAS_PROV_CEDIS 37,625. Verificada
   idempotencia corriendo el CLI completo dos veces (conteos idénticos).
6. **Dos tablas informativas nuevas** (`sucursales`, `catalogo_sku_tienda`)
   + `orcmm_etl_catalogos.py`, a partir de archivos que llegaron aparte
   (`Catalogos SKU/`): listado de las 94 sucursales de todo el grupo (La
   Comer, Fresko, City Market, City Café, Sumesa) y el catálogo completo de
   SKU de 5 tiendas (107,339 filas). **Explícitamente informativas**: el
   motor sigue usando `catalogo` (la hoja del layout), porque a estos
   catálogos nuevos les falta `cedis_surtidor`. Los 5 archivos de SKU NO
   comparten el mismo encabezado entre sí (una trae "Codigo" y "Código de
   Barras" duplicados, otra un "code" extra, tres no traen Línea/Vía) — el
   lector empareja columnas por nombre con alias, no por posición fija, y
   sólo `sku`/`tienda` son obligatorios. También manejan `"NA"` (texto
   literal, no `#N/A` de fórmula) como dato ausente en columnas numéricas —
   filtro local en `orcmm_etl_catalogos.py`, sin tocar el conversor
   compartido de `orcmm_pipeline.py`.
7. **`sql/borrar_datos.sql`** + **`orcmm_db_borrar.py`** — vacía las 12
   tablas (`TRUNCATE`, deja el esquema) sin re-cargar nada. Exige `--si`
   explícito. Trae también un `DROP TABLE` comentado por si algún día hace
   falta tirar el esquema completo.

**Pendiente para la siguiente sesión:**

- **Fase 2, no empezada**: que el motor de clasificación lea de Postgres
  en vez de archivos subidos. Las 9 tablas raw ya están pensadas para
  reconstruir el mismo `Fuentes` dataclass que arma `leer_fuentes` (mismas
  llaves naturales), así que sería mecánico — un `orcmm_fuentes_db.py`
  nuevo, sin tocar `derivar_evidencias`/`clasificar`/`escribir_resultado`.
  Ahí también se decide: nuevo endpoint en la API, tabla de resultados
  (`diagnosticos`), y si conviene una columna `carga_id` de linaje fila por
  fila en las tablas raw (se dejó fuera por simplicidad).
- `SIMA_PEDIDOS_TIENDA` sigue vacía en la entrega V5 — mismo pendiente de
  siempre con SIMA.
- Revisar con Compras si `catalogo_sku_tienda` (107,339 filas, 5 tiendas)
  eventualmente debe reemplazar a `catalogo` (26,407 filas) como fuente
  para el motor — hoy no puede porque le falta `cedis_surtidor`.
- `requirements.txt` creció con `psycopg2-binary` y `python-dotenv` — no
  hace falta nada en Fly.io todavía porque esta fase no toca `api/`, pero
  cuando la fase 2 exponga esto en la API, `DATABASE_URL` se pone igual que
  `ORCMM_ORIGENES` hoy: `fly secrets set`.

---

## Sesión 2026-08-09 (cierre) — deploy, front con filtros, y los dos "errores de layout"

**Se hizo:**

1. **`orc-api` desplegado en Fly.io** — versión 7, imagen
   `deployment-01KZMTWGW9JH…`, una sola máquina (`--ha=false`), health check
   pasando y verificado en vivo contra `https://orc-api.fly.dev/api/salud`.
   Con esto se salda el pendiente que arrastraba dos sesiones: producción ya
   corre el fix de `#N/A`. El deploy avisa que la app "no escucha en
   0.0.0.0:8080" — es falso positivo de temporización, revisa antes de que
   uvicorn termine de bindear; el health check posterior y el 200 desde fuera
   lo desmienten.
2. **Front: filtros y paginación** (orc-gui `6f4b6b4`). Barra sticky con SKU,
   tienda, causa raíz, responsable y proveedor; paginación a 20 en las siete
   tablas; SKU clicables; ficha del SKU buscado. `paginacion.ts` es nuevo y
   reutilizable.
3. Se mataron dos uvicorn huérfanos (del 5 y el 7 de agosto) que tenían
   tomados los puertos 8000 y 8080. **El del 8000 servía código de dos días
   antes**, así que lo que se probara contra la API en ese rato no reflejaba
   los cambios. Vale la pena revisar que no queden procesos viejos antes de
   dar por bueno un resultado.

**Los dos "errores de layout" que reportó Compras — medidos, y ninguno es
problema de rango:**

- **`fecha_recibo` con 120,816 vacíos.** No es censura por corte de ventana:
  el porcentaje de vacíos está **plano entre 10% y 15% en los 14 meses**
  (jun-2025 11.2%, dic-2025 20.4%, jun-2026 15.2%). Si fuera "todavía no
  llega", los meses viejos estarían en ~0%. Lo que pasa es que **el spec se
  contradice**: `orcmm_layout_spec.py:222` marca el campo obligatorio y su
  propia descripción dice "VACÍO si no se recibió". El validador hace lo que
  se le pidió; el que está mal es el spec. Hay que ponerlo opcional — **pero
  junto con acotar la vigencia de los pedidos**, porque ese mismo vacío es lo
  que produce los pedidos zombi.
- **62,102 citas huérfanas (60.9%).** Tampoco es rango: su `fecha_cita` cae
  entre 2026-02-01 y 2026-03-31, dentro de la ventana de COMPRAS
  (2025-05 a 2026-06), y los 8,509 folios huérfanos están **100% dentro del
  rango numérico** de los folios de COMPRAS — intercalados, no en un tramo
  faltante. Además hay **0 casos** de "folio sí, SKU no": es puramente a nivel
  folio. Y el conteo de `#N/A` es **exactamente 62,102**, el mismo número: el
  `XLOOKUP` falla justo en las huérfanas, son dos síntomas de lo mismo.
  **La causa se ve en el proveedor**: el 79.2% de las huérfanas trae un
  `proveedor_id` que no existe en COMPRAS — TRUPER, REVLON, GRISI, BDF,
  textiles, juguetes. **COMPRAS llegó filtrado a las divisiones del alcance y
  CITAS llegó completo, con las compras de toda la tienda.** No es un hueco:
  son dos universos.

**Cómo tratarlos (propuesto, sin implementar):**

1. `fecha_recibo` → opcional en el spec, **más** acotar la vigencia en
   `derivar_orden_proveedor` con una constante configurable arriba del módulo
   (como `CEDIS_AUSENCIA_ES_CERO`). **Falta decidir la ventana** — ¿lead time
   del SKU, o fijo de 30-45 días? Es criterio de negocio.
2. Partir la advertencia de citas huérfanas en sus dos poblaciones: las 49,200
   de proveedores fuera de alcance se filtran al leer y se reportan como nota;
   las **12,902 de proveedores que sí están en COMPRAS** (P&G, Colgate,
   Danone) son hueco real y **eso** es lo que hay que pedirle a Compras.
   Pedir 12,902 con folio y fecha es una petición concreta; pedir 62 mil, no.
3. Quitar `fecha_pedido` del layout de CITAS: es derivable por join y hoy sólo
   genera 62 mil `#N/A`.

**Pendiente para la siguiente sesión:**

- Decidir la ventana de vigencia y aplicar los tres puntos de arriba.
- **Endpoint `/api/analisis/{id}/sku/{sku}`** para ver el día a día de un SKU
  desde el front. Hoy la respuesta sólo trae `por_sku_tienda` (un renglón
  agregado) y la historia diaria vive únicamente en el Excel. Ojo con el caso
  del SKU sano: el pipeline descarta los días sin faltante desde el principio
  (`llaves_con_faltante`), así que enseñar "estaba sano" con su inventario
  diario obliga a retener también esos días — decisión de diseño con costo en
  memoria, no está tomada.
- Sigue sin CI en `orc-api` (`superfly/flyctl-actions`).

---

## Sesión 2026-08-09 — corrida con el V5 reentregado

Llegó una reentrega del layout V5 (el Excel pasó de 35 MB a 64 MB). **La
conclusión es que no aportó nada al análisis**, y vale la pena dejar escrito
por qué, para no volver a pedir lo mismo.

**Se hizo:**

1. **Fix de lectura — errores de Excel.** El pipeline reventaba con
   `ValueError: Invalid isoformat string: '#N/A'`. La columna `fecha_pedido`
   de `CITAS_PROV_CEDIS` no viene como dato: viene con la fórmula
   `=IFERROR(_xlfn.XLOOKUP(...))`, y el prefijo `_xlfn.` significa que Excel
   la guardó como función desconocida, así que nunca resolvió y dejó el error
   en caché. Se agregó `ERRORES_EXCEL` + `_es_vacio()`, y los cuatro
   convertidores (`_fecha`, `_texto`, `_entero`, `_decimal`) los leen como
   dato ausente. Ojo con `_texto`: antes habría metido `"#N/A"` como si fuera
   un folio válido. `registrar()` ahora cuenta las celdas con error y las
   reporta por columna — en esta corrida salieron **62,102** en `fecha_pedido`.
2. **Corrida completa** con los 7 CSV de Tableau. OSA del periodo **85.9%**,
   5,192 de 5,201 días clasificados (99.8%), venta perdida clasificada al
   99.9%. Pareto: **94.1% Ejecución en Tienda** · 5.3% Proveedor · 0.3%
   Compras · 0.1% Transporte · 0.1% CEDIS · 0.1% sin clasificar. El titular
   **no se movió** respecto de la corrida anterior.

**Lo que se midió de la reentrega (para no volver a pedirla igual):**

- `COMPRAS_PEDIDOS_PROV` creció de 151,138 a **853,073** filas, pero la
  ventana va de **2025-05-07 a 2026-06-25**: sólo el **16.4%** (140,122
  filas) cae dentro de feb-mar 2026. 60.3% es histórico anterior y 23.3% es
  posterior al cierre, hasta junio. Las ~700 K filas nuevas están **todas**
  fuera del periodo. Por eso los huérfanos de citas apenas bajaron de 62.0%
  a 60.9%.
- `CEDIS_INVENTARIO` llegó **idéntica**: 23,349 filas, 1,553 ceros, 2,979
  SKU, `piezas_reservadas` vacía en el 100%. Son ~602 SKU con serie contra
  los 26,407 del catálogo (**2.2% de cobertura**), y sigue la ventana rota
  del 8 al 14 de marzo, donde el panel se renueva casi por completo cada día
  (del 7 al 8 de marzo sólo sobreviven 2 SKU).
- `SIMA_PEDIDOS_TIENDA` sigue **vacía**.

**Hallazgo: `cajas_confirmadas_cita` no sirve.** Es una copia exacta de
`cajas_pedidas` en **38,858 de 38,858** pares (folio, sku) que cruzan — 100%,
sin una sola excepción. Es justo lo que el propio layout pide avisar en la
descripción de la columna. De ahí sale el `pct_confirmado` por encima de 100%
(113.0% global, BONAFONT 187.7%, Cía. Vinícola del Norte 193.0%): cuando un
pedido tiene varias citas, cada una carga el total del pedido y
`_aplicar_citas` las suma. 775 pares con cita múltiple generan 162,520 cajas
de exceso, y el hueco total sobre el 100% es de 154,769 — o sea, **ese 2% de
pares explica el desvío completo**. `cajas_entregadas` en cambio **sí es dato
real** (8.0% no se presentó, 5.9% parcial, nunca excede), así que
`cumplimiento` y `efectivo` se sostienen y la prioridad 8 se salva.

**Herramienta nueva: `orcmm_expediente.py`.** Cruza las ocho fuentes para un
SKU y deja el dato crudo al lado del veredicto. Se armó para auditar el caso
mudo —un SKU que no produce ni una fila— que desde el Excel de salida no se
puede diagnosticar. Uso en el README.

**Dos auditorías que salieron de ahí, y las dos preocupan:**

1. `7506425626212` (papel higiénico Petalo, tienda 287) — **0 días
   clasificados, y está bien**: nunca se agotó. 44 días de existencia, mínimo
   7 piezas, resurtido antes de romper. BOPS nunca lo reportó porque no hubo
   faltante. Sirve como control negativo del modelo. Pero destapó que
   **COMPRAS y CITAS se contradicen**: el folio 26300919602 aparece en COMPRAS
   con 148 de 148 cajas entregadas y recibo el 3-feb, y en CITAS con 28
   entregadas de 148 confirmadas. 120 cajas de diferencia sobre el mismo folio
   y el mismo SKU. Además COMPRAS trae `fecha_cita` vacía aunque la cita
   existe.
2. `663985002478` (vino Malleolus Emilio Moro, tienda 287) — **el SKU de mayor
   impacto de todo el Pareto**, $6,177 en 39 días como RC01 Ejecución en
   Tienda. La regla es correcta —BOPS dice OSA 0 y el sistema dice que hay
   inventario, así que el producto está en tienda y no en anaquel— pero el
   dato de abajo **no se sostiene**: la existencia está congelada en 3.00
   exactas del 16-feb al 19-mar, salta a 9.00 y se queda congelada hasta el
   31-mar. Cero movimiento en 44 días, y `TABLEAU_VENTAS` sólo trae un
   registro (31-mar, importe 0). Un vino que no vende una sola botella en seis
   semanas con "3 en existencia" es **inventario fantasma**, no una falla de
   acomodo. Si se confirma, una parte del 94.1% de RC01 no es ejecución sino
   exactitud de inventario — que es otro dueño y otra solución (conteo
   cíclico, ajuste), no "surtir el anaquel".

**Hallazgo de modelo: pedidos zombi.** En ese mismo expediente, el motor eligió
como "orden vigente que explica el faltante" un pedido del **14-jul-2025** (10
cajas, sin `fecha_recibo`, 0 entregadas) para explicar días de febrero y marzo
de 2026. `derivar_orden_proveedor` deja vigente todo pedido sin recibo, y en
COMPRAS hay **120,816 filas con `cajas_entregadas` = 0**, muchas sin recibo:
nunca se cierran y se acumulan. Aquí no cambió el veredicto porque RC01 dispara
en prioridad 1, antes de llegar a la rama de proveedor — pero en los días que
sí bajan a proveedor, la orden elegida puede tener meses de antigüedad. Vale la
pena acotar la vigencia por una ventana razonable (¿el lead time del SKU?).

**Pendiente para la siguiente sesión:**

- **Validar el inventario fantasma** con La Comer sobre el caso del vino: es la
  línea más grande del Pareto y el diagnóstico cambia de dueño según la
  respuesta.
- **Acotar la vigencia de los pedidos** en `derivar_orden_proveedor`.
- **No publicar `pct_confirmado`** mientras `cajas_confirmadas_cita` sea copia
  del pedido. Decidir además si `_aplicar_citas` debe seguir sumando ese campo
  entre citas vencidas — es cambio de criterio, quedó sin tocar a propósito.
- **Pedir a Compras**: (a) `cajas_confirmadas_cita` con el compromiso real al
  agendar; (b) `estatus_cita`, hoy **vacía en las 101,942 filas**, sin ella no
  hay cómo descartar canceladas ni reprogramadas; (c) `fecha_pedido` en duro,
  no calculada; (d) confirmar si la hoja trae TODAS las citas del periodo —
  **809,169 de 848,027 pedidos (95.4%) no tienen cita** y el modelo los
  dictamina RC06 "el proveedor nunca agendó". Con ese porcentaje, lo más
  probable es que sea hueco de captura y no del proveedor.
- **Pedir a CEDIS** la foto diaria sobre los 26,407 SKU del catálogo, no sobre
  602, y aclarar qué pasó del 8 al 14 de marzo. Ojo con la premisa de
  `CEDIS_AUSENCIA_ES_CERO`: dice que la extracción omite los SKU en cero, pero
  el archivo trae 1,553 filas con `existencia_piezas = 0` — si de verdad los
  omitiera, no debería haber ninguna. Conviene reconfirmarlo, porque de ese
  supuesto depende que la rama CEDIS quede en 0.1% y el peso caiga en
  proveedor.
- **Revisar el 83% excluido**: de 30,565 días con faltante que entregó BOPS,
  25,364 ($285,907) quedan fuera por `sku_fuera_del_catalogo_de_la_tienda`.
  La venta perdida excluida casi duplica la analizada ($155,959).
- Sigue pendiente de la sesión anterior: **desplegar `orc-api`** (el deploy en
  Fly.io está atrás del layout V5) y **armar el CI** con
  `superfly/flyctl-actions`.

---

## Sesión 2026-08-06 (noche) — layout V5: Excel + CSV

El layout dejó de caber en un solo Excel. `TABLEAU_INV_TIENDA` (2.7 millones
de filas, en 6 archivos) y `TABLEAU_VENTAS` ahora se entregan como CSV
aparte; el resto sigue en el .xlsx. **El back ya lee las dos cosas como una
sola fuente.** Todo el análisis y el plan están en
[`PLAN_MULTIFUENTE.md`](PLAN_MULTIFUENTE.md) — leerlo antes de tocar nada.

**Se hizo:**

1. `orcmm_fuentes_csv.py` nuevo: lee los exportes de Tableau (UTF-16 con BOM,
   separados por TAB, encabezados con nombre de negocio y la columna de la
   métrica **sin encabezado**, fechas en texto y en dos idiomas). Streaming y
   filtrado por llave contra los días con faltante: 2.7 M de filas en 7.5 s
   con 22 MB de pico. Corre como inspector con `--contra <layout.xlsx>` para
   verificar un export nuevo en 12 segundos.
2. `leer_fuentes` recibe un `PaqueteFuentes` (xlsx + CSV agrupados por
   prefijo del nombre de archivo) y el `umbral_osa`, que es lo que define qué
   llaves vale la pena guardar.
3. **Indexación de eventos** (`_indexar_eventos`). Las tres derivaciones del
   día D recorrían la lista completa de eventos una vez por cada día con
   faltante: 6.4 mil millones de vueltas, horas de corrida. Ahora cada día
   toca sólo los suyos: **0.3 s**. Con los datos de ejemplo nunca se notó.
4. La venta perdida se lee de `BOPS_OSA` (columna nueva en el V5), con
   `TABLEAU_VENTAS` de respaldo. Sube la cobertura sobre impacto de 0% a 100%.
5. `CEDIS_AUSENCIA_ES_CERO`: La Comer confirmó que el reporte de CEDIS omite
   los SKU en cero, así que la ausencia de fila **es** un cero. Única
   excepción a "vacío no es cero" en todo el modelo; sólo aplica a los días
   que la extracción cubre. Los días con dato de CEDIS pasan de 17 a 5,201.
6. **Separación de alcance** (`FUERA_DE_CATALOGO`, `R0_DentroDelCatalogo`):
   BOPS_OSA entrega SKU de divisiones fuera del alcance (Nivea, Sony,
   L'Oréal) mientras el catálogo es 100% Abarrotes. Esos días se cuentan
   aparte, no como dato faltante. Decisión del usuario: **separar, no
   descartar**.
7. Validador: entiende los CSV, valida en streaming y **cruza las fuentes**
   (¿el inventario es de la misma tienda y periodo que OSA?). Refactorizado a
   una sola pasada por hoja en `read_only`: de 125 s a 40 s.
8. API v3: subida multi-archivo o **ZIP** (272 MB → 37.9 MB), análisis
   **asíncrono** con `POST /api/analizar` → `GET /api/analizar/{id}`, un solo
   worker, y banderas `corregir` / `forzar`. Fly a **2 GB** (medido: 843 MB
   de pico).

**Resultado sobre los datos reales** (83 s por CLI, 204 s por API):

```
OSA general 74.5%
BOPS entregó 30,565 días con faltante
  25,364 fuera del catálogo ($285,907) — no entran al análisis
   5,201 dentro del alcance · clasificados 5,181 (99.6%)
Pareto: 94.1% Ejecución en Tienda · 4.0% Proveedor · 1.4% Compras · 0.2% sin clasificar
```

**Pendiente para la siguiente sesión:**

- **Desplegar** (`flyctl deploy --ha=false`). Todavía NO se hizo: el cambio de
  VM a 2 GB va en ese mismo deploy y sin él la máquina muere a mitad del
  análisis.
- **El front tiene que cambiar sí o sí**: el contrato es otro (varios
  archivos o zip, y poll en vez de respuesta directa). La respuesta ya trae
  las dos coberturas (`cobertura_casos_alcance_pct` y `cobertura_casos_pct`)
  para encabezar con la del alcance.
- **Pedirle a La Comer**: (a) el export de BOPS_OSA filtrado a Abarrotes;
  (b) `CITAS_PROV_CEDIS` está rota — 101,933 de 101,943 filas sin
  `cedis_destino`, 63,197 sin `fecha_pedido` y 63,197 citas con folio que no
  existe entre los pedidos. **El scorecard de proveedores no se puede firmar**
  hasta que se aclare (se ven tasas imposibles, como confirmar 189.5% de lo
  pedido). El Pareto sí, porque el 94.1% no depende de esa hoja.
- Que los exportes de Tableau salgan **con encabezado en la cuarta columna**;
  hoy se asigna por posición y el lector lo advierte en cada corrida.
- Quedó sin responder: ¿la Vía 2 debe saltarse la pregunta de inventario en
  CEDIS? Hoy no la salta (comparte reglas 5-8 con la Vía 1, confirmado el
  2026-08-05), aunque en cross-dock el inventario da cero y pasa de largo
  igual. Sólo cambia el caso en que el producto está físicamente en CEDIS
  esperando despacho: hoy eso dictamina RC04 contra CEDIS.

---

## Sesión 2026-08-06 (continuación, tarde)

**Se hizo:**

1. `npm install` en `orc-gui` (ahora sobre disco local, ya no Google Drive)
   — 539 paquetes, sin errores. `npm run build` corre limpio
   (`ng build`, bundle inicial 202.53 kB, 4.2s). Queda pendiente probar
   `npm start` (`ng serve`) de forma interactiva, pero el build confirma
   que el código compila sin errores de TypeScript/plantillas.
2. **Bug encontrado y corregido:** el front mostraba "OSA general del
   periodo" como "%" vacío (sin número). Causa: el deploy a Fly.io
   (`3293f52`, 2026-08-05 23:19) quedó *antes* del commit que agrega
   `osa_general` (`3f7bc2b`, 2026-08-06 16:02) — la API en producción
   nunca tuvo el campo nuevo, así que `{{ r.osa_general }}%` en
   `app.html` interpolaba `undefined`. Se corrió `flyctl deploy --ha=false`
   de nuevo con el código actual y se verificó contra la API en vivo con
   el Excel de ejemplo: `osa_general: 76.4`, correcto. **Ojo:** ese deploy
   quedó atrás de nuevo con el trabajo de layout V5 de la sesión de la
   noche (ver entrada de arriba) — hace falta un deploy más.
3. Estilo La Comer en `orc-gui`: colores de marca (naranja `#F0501E` /
   `#C43C10`, extraídos del favicon oficial de lacomer.com.mx) aplicados
   al acento visual — franja del header, borde de foco, spinner, botón
   principal y número de "Descargar". Variable CSS renombrada de
   `--amarillo` a `--acento`/`--acento-oscuro`/`--acento-suave`.
   Verificado contraste AA (5.25:1) del botón. → orc-gui `28530f3`.
   Sigue pendiente el logo (no hay asset oficial descargado todavía).

**Pendiente para la siguiente sesión:**

- Armar CI para `orc-api`: no hay `.github/workflows/` en el repo, así
  que a diferencia de `orc-gui` (Vercel, que sí redespliega solo en cada
  push) todo deploy a Fly.io es manual con `flyctl deploy`. Esto fue la
  causa raíz del bug del punto 2 — un commit se quedó sin desplegar por
  horas, y ya volvió a pasar con el trabajo de la noche. Propuesta:
  workflow con `superfly/flyctl-actions` + `FLY_API_TOKEN` como secret,
  disparado en push a `main`.
- **Desplegar** `orc-api` con el layout V5 + VM de 2 GB (ver entrada de
  arriba) — sigue sin hacerse.
- Logo de La Comer en `orc-gui` (el color de marca ya quedó aplicado,
  falta el asset gráfico oficial).
- Cuando SIMA entregue los pedidos de tienda: `EVALUAR_PEDIDO_TIENDA = True`
  en `orcmm_rca_engine.py`.
- `RESPONSABLE_PEDIDO_NO_GENERADO` y `RESPONSABLE_SIN_CITA` siguen sin
  ratificar con La Comer.

---

## Sesión 2026-08-06

1. Validación end-to-end del Excel con datos corregidos contra el pipeline
   completo — 92.3% de cobertura, coincide con lo ya documentado en el
   README.
2. `CATALOGO.tipo_resurtido` (Manual/Automático) estaba capturado en el
   layout pero ninguna regla lo leía. Se incorporó a la prioridad 3 (RC03
   "Pedido No Generado"): `Automático` → Compras/Abasto, `Manual` → Tienda,
   sin dato → Tienda/Abasto (el comportamiento de antes, sin romper nada).
   Nuevo `TipoResurtido`, `SubcausaPedidoTienda` y
   `RESPONSABLE_PEDIDO_NO_GENERADO` en `orcmm_rca_engine.py`.
   **Sigue dormido** mientras `EVALUAR_PEDIDO_TIENDA = False` (SIMA no
   entrega pedidos de tienda). Es un supuesto de negocio **pendiente de
   ratificar con La Comer**, igual que `RESPONSABLE_SIN_CITA`.
   → orc-api `cecffe5`
3. Deploy a Fly.io (`orc-api`, región `dfw`, una sola máquina,
   `flyctl deploy --ha=false`). Confirmado vivo en `/api/salud` y `/docs`.
4. Nuevo %OSA general del periodo. Antes no existía ningún OSA agregado
   real: lo más parecido (`osa_promedio_periodo` en `orcmm_rca_periodo.py`)
   promediaba sólo los días que YA tenían faltante, así que siempre salía
   sesgado a la baja, y ni siquiera llegaba a la API. Ahora `osa_general()`
   en `orcmm_pipeline.py` promedia TODAS las filas de `BOPS_OSA` (todo SKU,
   tienda y día leído). Expuesto en la API (`osa_general` en la respuesta de
   `/api/analizar`), en el Excel ("Cobertura y fuentes", primera fila) y en
   el CLI.
   → orc-api `3f7bc2b`
5. Front: la pantalla de resultados ahora encabeza con el %OSA general
   (antes arrancaba directo en advertencias y el Pareto por causa, sin dar
   la foto general primero). El aviso de "Fuentes que no llegaron
   completas" se movió después de la franja de cifras generales. Se
   corrigió una etiqueta que quedó desactualizada del cambio anterior
   ("Qué falla del proveedor costó más" → "Qué detalle de la causa costó
   más", porque la tabla de subcausas ya puede traer tanto proveedor como
   pedido de tienda).
   → orc-gui `396b8c8`

**Pendiente para la siguiente sesión:**

- Verificar que `orc-gui` compile limpio (`npm install && npm start`). No
  se alcanzó a correr en esta sesión: el `npm ci` se cortó a medias porque
  el repo vivía en una unidad de Google Drive sincronizada (muy lento para
  escribir `node_modules`) y se decidió mover todo a disco local a mitad
  de la sesión.
- Estilo La Comer (logo, colores de marca oficiales) en `orc-gui`. El
  usuario pidió dejarlo para después — no hay assets de marca en ningún
  repo todavía. Esta sesión sólo tocó la jerarquía de información
  (general → particular), nada de estilo visual.
- Cuando SIMA entregue los pedidos de tienda: poner
  `EVALUAR_PEDIDO_TIENDA = True` en `orcmm_rca_engine.py` y volver a
  correr. Eso activa RC03 y, con él, el refinamiento de `tipo_resurtido`
  del punto 2.
- `RESPONSABLE_PEDIDO_NO_GENERADO` (Automático→Compras/Abasto,
  Manual→Tienda) es un supuesto de negocio, no confirmado con La Comer.
- `ORCMM_ORIGENES` sigue sin configurarse en Fly.io. No hace falta hoy
  porque Vercel usa *rewrite* server-side (ver arriba), pero si el front
  cambia de arquitectura y deja de usar ese proxy, hay que revisar CORS.

---

## Cómo retomar en una máquina nueva

```bash
git clone https://github.com/samtriani/orc-api.git
git clone https://github.com/samtriani/orc-gui.git
```

- `orc-api`: `pip install -r requirements.txt` — ver su `README.md` para
  correr el CLI o levantar la API local.
- `orc-gui`: `npm install` — ver su `README.md` para levantar el front
  local (proxea `/api` al puerto 8000 en desarrollo).

No poner el checkout en una carpeta sincronizada por OneDrive/Google
Drive/etc.: `npm install` escribe miles de archivos pequeños y en un
drive así puede volverse desesperantemente lento.
