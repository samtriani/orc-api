# Seguimiento de sesiones — Claude Code

Este proyecto se trabaja desde dos máquinas distintas. Este archivo es la
memoria entre sesiones: **al empezar, léelo**; **al cerrar, agrega una
entrada nueva arriba** con qué cambió y qué quedó pendiente. No se borran
las entradas viejas — así queda el historial de decisiones.

## Cómo está repartido el proyecto

- **`orc-api`** (este repo) — motor de clasificación (`orcmm_rca_engine.py`),
  pipeline (`orcmm_pipeline.py`), API FastAPI (`api/`). Desplegado en Fly.io:
  https://orc-api.fly.dev
- **`orc-gui`** ([github.com/samtriani/orc-gui](https://github.com/samtriani/orc-gui)) —
  front Angular. Desplegado en Vercel: https://orc-gui.vercel.app — hace
  *rewrite* server-side de `/api/*` hacia Fly.io, así que no hay CORS de por
  medio entre el navegador y el back.

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
