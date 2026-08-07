# Plan — lectura multi-fuente (Excel layout + CSV de Tableau)

Fecha: 2026-08-06. Estado: **back terminado** — §3.1 a §3.5 implementados y
probados contra el V5, por CLI y por API. Falta desplegar y el front.

Corrida de referencia:

```
python orcmm_pipeline.py "datos_reales_v5/...V5.xlsx" datos_reales_v5/*.csv
```

| | |
|---|---|
| Filas leídas | 3.1 M (2.7 M de inventario + 255 K de ventas + el Excel) |
| Tiempo total | 83 s (50 s lectura, 0.3 s derivación, 32 s escritura del Excel) |
| OSA general | 74.5% |
| Días con faltante | 30,565 · venta perdida $441,866.16 |
| Fuera del catálogo | 25,364 ($285,907.33) — ver §2.2 |
| Dentro del alcance | 5,201 · clasificados 5,181 = **99.6%** (99.8% del impacto) |

Pareto resultante, ya sobre el alcance:

```
94.1%  Ejecución en Tienda          ← había producto y no llegó al anaquel
 4.0%  Incumplimiento Proveedor
 1.4%  Pedido Proveedor No Generado
 0.2%  Sin clasificar
 0.1%  Transporte / Tránsito
 0.1%  CEDIS No Surtió
```

RC01 se dictamina en la prioridad 1, antes del interruptor de SIMA, así que
ese 94.1% no depende de `EVALUAR_PEDIDO_TIENDA`. El reparto entre las causas
de abasto sí.

El layout V5 ya no cabe en un solo Excel: `TABLEAU_INV_TIENDA` y
`TABLEAU_VENTAS` pasan del millón de filas que aguanta una hoja, y salieron a
CSV. El resto de las hojas —incluida `CEDIS_INVENTARIO`— siguen dentro del
Excel. El back tiene que leer las dos cosas como una sola fuente.

Este documento es el análisis de lo que llegó y el plan de back. El front va
después, en su propio documento.

---

## 1. Qué llegó, medido

Carpeta `code/datos_reales_v5/`.

### 1.1 El Excel — `040826_La Comer_Layout de datos RCA (OSA)_V5.xlsx` (35 MB)

| Hoja | Filas de datos | Nota |
|---|---:|---|
| CATALOGO | 26,586 | 179 duplicados en la llave (sku, tienda). Columna nueva `Nombre_Tienda`. |
| BOPS_OSA | 119,958 | **Columna nueva `venta_perdida_estimada`**, llena al 100%. |
| CEDIS_INVENTARIO | 23,349 | Se quedó en el Excel: cupo al quitar los SKU con existencia nula. |
| CEDIS_TRANSFERENCIAS | 30,000 | |
| SIMA_PEDIDOS_TIENDA | 0 | Sigue vacía, como se esperaba. |
| COMPRAS_PEDIDOS_PROV | 151,133 | |
| CITAS_PROV_CEDIS | 101,943 | |
| ~~TABLEAU_INV_TIENDA~~ | — | **Ya no existe la hoja.** |
| ~~TABLEAU_VENTAS~~ | — | **Ya no existe la hoja.** |

`osa_pct` viene en escala binaria 0/1, como ya estaba previsto: `_osa_pct()`
lo multiplica por 100. OSA general del periodo = **75.0%**. Días con
faltante (osa = 0) = **30,565**, con venta perdida de **$441,866.16**.

### 1.2 Los CSV

No son CSV de coma. Son exportes de Tableau: **UTF-16 LE con BOM, separados
por TAB**, y la cuarta columna **no tiene encabezado**.

```
Código de Barras <TAB> Tienda No <TAB> Day of Fecha <TAB> (vacío)
289              <TAB> 287       <TAB> February 16, 2026 <TAB> 0.00
```

| Archivo | Filas | SKU | Tienda | Rango |
|---|---:|---:|---|---|
| `TABLEAU_INV_TIENDA_1..6.csv` | 2,719,780 | 62,131 | 287 | 2026-02-16 → 2026-03-31 |
| `TABLEAU_VENTAS.csv` | 90,139 | 12,504 | **237** | 2025-01-02 → 2026-03-31 |

Los 6 archivos de inventario son cortes por fecha, sin traslape: 0 llaves
duplicadas al unirlos. Se leen como una sola tabla.

Diferencias contra lo que espera el spec hoy:

- **Encabezados distintos.** `Código de Barras` → `sku`, `Tienda No` →
  `tienda`, `Day of Fecha` / `Día en texto` → `fecha`. La 4ª columna viene
  **sin nombre**: hay que asignarla por posición.
- **Fechas como texto y en dos idiomas.** Inventario en inglés
  (`February 16, 2026`), ventas en español (`1 de Marzo de 2026`).
- **`TABLEAU_VENTAS` trae una fila `Grand Total / Total / Total`** que hay
  que descartar.
- **Columnas que ya no vienen.** Inventario perdió `hora_de_corte` y
  `existencia_minima_dia`. Ventas perdió `unidades_vendidas`,
  `importe_venta`, `venta_perdida_estimada` y `metodo_estimacion` — sólo
  queda una métrica sin nombre.

---

## 2. Tres hallazgos que hay que resolver con datos, no con código

Ninguno bloquea el diseño del back, pero los tres cambian lo que el
resultado puede afirmar. Van primero porque son más caros de arreglar que
el código.

### 2.1 `TABLEAU_VENTAS.csv` era de otra tienda — RESUELTO

Venía de la tienda 237 con todo lo demás en la 287. La Comer reextrajo el
2026-08-06: ahora son 254,986 filas de la **287**, periodo 2026-02-17 a
03-30. Correcto.

El cruce contra días con faltante sigue siendo bajo (53 de 30,565) pero eso
**es correcto por construcción**: si no hay producto no hay venta, y el
export no genera fila. Contra los días con el anaquel lleno cruza 14.5%. No
bloquea nada — la venta perdida vive en `BOPS_OSA`.

### 2.2 Sólo el 17% de los días con faltante tiene con qué clasificarse

Sobre los 30,565 días con OSA = 0:

| Fuente que necesita el árbol | Días que la tienen |
|---|---|
| Inventario en tienda (CSV) | 5,292 — **17.3%** |
| Fila en CATALOGO (vía de resurtido, CEDIS surtidor) | 5,201 — **17.0%** |
| Inventario en CEDIS | 17 — **0.1%** → 5,201 tras §2.2 bis |

No es un problema de formato de clave: probé normalizar ceros a la
izquierda y rellenar a 13 dígitos, y el cruce no se mueve ni un SKU.

**La causa es de alcance, no de datos.** De los 2,264 SKU con faltante en
`BOPS_OSA`, 1,768 no existen en el catálogo — y ninguno de esos tiene
inventario tampoco. Es un solo problema, no dos. El contraste que lo cierra:

| SKU con faltante | En COMPRAS_PEDIDOS_PROV |
|---|---|
| dentro del catálogo (496) | 484 — **98%** |
| fuera del catálogo (1,768) | 9 — **1%** |

El catálogo es 100% división Abarrotes (26,586 filas, una sola división), que
es el alcance declarado del proyecto. Catálogo, inventario de tienda y pedidos
a proveedor vienen filtrados a Abarrotes; **`BOPS_OSA` no**. Los códigos de
barras lo confirman: entre los de mayor venta perdida fuera de catálogo hay
4005900728807 (Beiersdorf/Nivea), 4548736108127 (Sony) y 3337875543248
(L'Oréal) — perfumería y electrónica.

Consecuencia para leer el resultado: el **17.2%** de cobertura mezcla dos
universos. Sobre los días cuyo SKU sí pertenece al alcance, la cobertura es
**5,258 de 5,292 = 99.4%**. Los 25,273 días restantes ($285,325.62) no son
una falla del modelo: son SKU de otras divisiones que no debían entrar.

**Decidido el 2026-08-06: se separan, no se descartan.** Implementado así:

- `FUERA_DE_CATALOGO` y la regla `R0_DentroDelCatalogo` en
  `orcmm_rca_engine.py`. Un SKU que el catálogo de la tienda no reconoce se
  marca con ese motivo en vez de caer como "falta inventario_tienda".
- `cobertura_modelo()` reporta **dos** cifras: sobre el alcance (mide al
  modelo) y sobre todo lo que entregó BOPS (mide la extracción).
- `dentro_del_alcance()` filtra el Pareto y el detalle por SKU. Sin eso, el
  Pareto salía con 64.8% de "Sin clasificar" tapando el Pareto real.
- La clasificación diaria conserva las 30,565 filas con su motivo, y la
  advertencia de lectura dice cuántos SKU llegaron de más. Nada se pierde.

Ojo con el 92.3% del README: se midió sobre el V2, que era **un solo SKU** de
ejemplo. Este V5 es el primer volumen real, así que no hay regresión — es la
primera medición honesta.

`CEDIS_INVENTARIO` **quedó resuelto sin pedir datos nuevos.** El desglose
mostró que las claves de CEDIS (280) y las fechas cuadran perfecto: lo que
faltaba eran los SKU. De los 496 SKU con faltante que están en el catálogo,
sólo 18 aparecían en el inventario de CEDIS, porque la extracción quitó los
que venían en cero — justo los que explican el faltante.

La Comer confirmó (2026-08-06) que **el reporte lista los SKU con existencia
y omite los que están en cero**: la ausencia de fila es un cero real, no un
dato desconocido. Queda implementado en `CEDIS_AUSENCIA_ES_CERO`
(`orcmm_pipeline.py`), y sólo aplica a los días que la extracción cubre para
ese CEDIS — un día no extraído sigue siendo dato ausente. Es la única
excepción a "vacío no es cero" en todo el modelo, así que se reporta como
advertencia en cada corrida y se apaga con un interruptor.

Efecto medido sobre el V5: los días con dato de CEDIS pasan de **17 a 5,201**;
sobre los días que llegan a esa rama, de **0.3% a 100%**.

### 2.3 La cuarta columna de los CSV no tiene nombre — RESUELTO

Inventario es `existencia_piezas`; ventas es **`importe_venta`**, confirmado
con La Comer el 2026-08-06. Queda declarado en `COLUMNA_SIN_NOMBRE`, y el
lector lo advierte en cada corrida hasta que el export salga con encabezado.

---

## 3. Plan de back

Cinco piezas. La 3.3 no es opcional: sin ella el análisis no termina.

### 3.1 Nuevo módulo `orcmm_fuentes_csv.py` — HECHO

Lector de CSV tolerante, aislado del resto:

- **Encoding**: detectar BOM (UTF-16 LE/BE, UTF-8), con fallback a cp1252.
- **Delimitador**: sniff entre tab, coma y punto y coma.
- **Encabezados**: tabla de alias por hoja, en el mismo estilo que el
  `ALIAS_ENCABEZADOS` que ya existe en el spec —
  `código de barras` → `sku`, `tienda no` → `tienda`,
  `day of fecha` / `día en texto` → `fecha`. Una columna sin encabezado se
  resuelve por posición contra el campo declarado en el spec.
- **Fechas**: ISO, más `February 16, 2026` y `1 de Marzo de 2026`. Una sola
  función, probada contra los dos idiomas.
- **Números**: separador de miles y coma decimal.
- **Filas de total**: descartar cuando la clave sea `Grand Total` / `Total`.

Devuelve la misma forma que `leer_hoja()` para que el pipeline no note la
diferencia entre una hoja del Excel y un CSV.

Medido sobre los 7 archivos reales: **2,719,780 filas de inventario en 7.5 s
con 22 MB de pico**, gracias al filtro por llave y a memorizar el parseo de
fecha (44 fechas distintas contra 2.7 millones de llamadas). El módulo corre
como inspector desde la línea de comandos:

```
python orcmm_fuentes_csv.py datos_reales_v5/*.csv \
       --contra "datos_reales_v5/...V5.xlsx"
```

y con `--contra` reporta qué porcentaje de los días con faltante quedaría
cubierto — es la verificación de 12 segundos para cualquier export nuevo.

### 3.2 `leer_fuentes()` pasa a recibir un paquete, no una ruta

```python
def leer_fuentes(paquete: PaqueteFuentes, umbral_osa: float = 100.0) -> Fuentes
```

`PaqueteFuentes` = el xlsx + un mapa `hoja -> [rutas de CSV]`. La asignación
de cada CSV a su hoja sale del **prefijo del nombre de archivo**:
`TABLEAU_INV_TIENDA_3.csv` → hoja `TABLEAU_INV_TIENDA`. Con ese criterio los
6 archivos de inventario se concatenan solos y agregar un séptimo no cuesta
nada.

Regla de precedencia: si una hoja existe en el Excel **y** llega por CSV, se
usa el Excel y se advierte. Los archivos que no correspondan a ninguna hoja
del spec se reportan como no reconocidos, no se ignoran en silencio.

### 3.3 Filtrar por llave al leer, en streaming — obligatorio

Dos problemas de escala, los dos nuevos porque el V2 tenía 55 filas.

**Memoria.** Hoy `leer_hoja()` construye la lista completa de filas y luego
un `dict` de dicts. 2.7 millones de filas de inventario así son varios GB.
La máquina de Fly tiene **512 MB**.

La salida es que el orden de lectura cambie: `BOPS_OSA` primero, de ahí el
conjunto de llaves `(sku, tienda, fecha)` con OSA bajo el umbral —30,565—, y
recorrer los CSV en streaming quedándose **sólo con esas llaves** y
guardando el escalar, no el dict. De 2.7 M de filas quedan ~5 K en memoria.
Por eso `leer_fuentes` necesita el `umbral_osa`, que hoy sólo conoce
`derivar_evidencias`.

Los conteos y rangos de fecha que alimentan la hoja "Cobertura y fuentes" se
acumulan durante el recorrido, para que sigan reportando el archivo completo
y no la parte que se guardó.

**Tiempo.** Peor que la memoria y no tiene que ver con el CSV.
`derivar_transito_vigente`, `derivar_envio_generado` y
`derivar_orden_proveedor` hacen un scan lineal de la lista entera de eventos
**por cada evidencia**:

```
30,565 evidencias × (30,000 transferencias × 2 + 151,133 pedidos) ≈ 6.4e9 iteraciones
```

Eso son horas, no segundos, y con un `_fecha()` reparseando la misma celda
millones de veces. Hay que indexar al leer:

- `transferencias_por[(sku, tienda_destino)]`
- `pedidos_prov_por[(sku, cedis_destino)]`
- `pedidos_tienda_por[(sku, tienda)]` — para cuando SIMA entregue

y normalizar fechas y claves **una vez**, al cargar, no en cada consulta.
Cada evidencia pasa a tocar sólo sus propios eventos. El resultado no cambia:
es el mismo filtro, sólo que sin recorrer lo que no aplica.

Mientras se toca esto: `derivar_evidencias` agrega una advertencia por cada
día sin catálogo. Con ~25 mil días así, son 25 mil strings en el JSON de la
API. Hay que agruparlas en una sola línea con el conteo y una muestra.

### 3.4 Spec, validador y corrector — HECHO

En `orcmm_layout_spec.py` — **hecho lo de la venta perdida y `hora_de_corte`**:

- ~~`BOPS_OSA` gana `venta_perdida_estimada`~~ (Decimal, obligatorio). Hecho.
- ~~`TABLEAU_INV_TIENDA`: `hora_de_corte` pasa a opcional~~ — ninguna regla la
  lee, se asume cierre del día. Hecho.
- ~~`TABLEAU_VENTAS`: al mudarse la venta perdida a `BOPS_OSA`, la hoja deja
  de ser crítica~~ — obligatorios sólo sku, tienda, fecha e `importe_venta`.
  Hecho.
- Falta: marcar cada hoja con su **origen** (`excel`, `csv` o `ambos`).
  `TABLEAU_INV_TIENDA` y `TABLEAU_VENTAS` pasan a `csv`.
- Falta: `CATALOGO` gana `Nombre_Tienda` como opcional, para que deje de
  salir como columna no reconocida.

En `orcmm_pipeline.py` — hecho: `derivar_evidencias` toma la venta perdida de
`BOPS_OSA` y cae a `TABLEAU_VENTAS` sólo si el día no la trae, para que los
layouts anteriores al V5 sigan corriendo. **Sube la cobertura sobre impacto
del 0% al 100%**: la venta perdida venía por la hoja de ventas, que en este
export no cruza ni un día, y `BOPS_OSA` la trae en las 119,958 filas.

En `orcmm_validar_layout.py`:

- Dejar de exigir como hoja del Excel lo que ahora es CSV, y en su lugar
  exigir que el CSV correspondiente venga en el paquete.
- Validar los CSV con las mismas reglas que las hojas: llave duplicada,
  obligatorios vacíos, fechas fuera de ventana, tipos.
- **Chequeo nuevo de cruce entre fuentes**, que es lo que hubiera cachado
  §2.1 y §2.2 antes de correr nada: qué % de los días con faltante tiene
  inventario de tienda, fila de catálogo e inventario de CEDIS, y qué tiendas
  aparecen en cada fuente. Sale como `faltan_datos`, no como error: el
  archivo está bien formado, lo que falla es la cobertura.

En `orcmm_corregir_layout.py`: no aplica a los CSV. Sólo hay que evitar que
intente reconstruir las dos hojas que ya no existen en el Excel.

Lo que costó tiempo y no se veía venir: `validar_archivo` cargaba el workbook
en modo normal y releía cada hoja hasta tres veces (una para encabezados, otra
para datos, otra para los cruces). **125 s.** Con `read_only=True` y una sola
pasada por hoja —de la que salen encabezados, datos y lo que necesitan los
cruces— baja a **40 s**, y a 78 s validando también los 2.7 millones de filas
de CSV.

`orcmm_corregir_layout` tarda **386 s** sobre este layout: reescribe un Excel
de 35 MB y para eso no puede usar `read_only`. Por eso el validador ahora
**marca cuáles errores sabe arreglar el corrector** (`Reporte.error(...,
corregible=True)`) y la API sólo lo invoca cuando alguno lo es. Sobre el V5 no
lo es ninguno, así que se ahorran esos seis minutos y medio.

También cambió un criterio: **"clave capturada como número" pasó de error a
advertencia** cuando el valor es un entero exacto. El pipeline normaliza toda
clave con `str()`, y para un entero eso reproduce el valor tal cual — el folio
26300925519 se lee igual venga como texto o como número. Sigue siendo error si
el valor es flotante o pasa de 2^53, que es cuando la lectura sí devuelve algo
distinto de lo capturado. Con eso el V5 pasó de 9 errores a 3, y los 3 son
datos genuinamente malos en `CITAS_PROV_CEDIS` que ninguna corrección
automática puede inventar.

### 3.5 API — subida de varios archivos — HECHO

`POST /api/validar` y `POST /api/analizar` pasan a aceptar:

1. un `.xlsx` (el layout), y
2. uno o varios `.csv`, **o** un `.zip` con todo dentro.

**El ZIP debería ser el camino recomendado.** Los 222 MB de CSV comprimen a
~11 MB medidos (ratio 20:1, son UTF-16 con relleno). Es la diferencia entre
una subida de minutos y una de segundos, y evita que el front tenga que
mandar 7 partes coordinadas.

Contrato nuevo:

```
POST /api/validar          xlsx + csv/zip  -> diagnóstico (síncrono, ~80 s)
POST /api/analizar         xlsx + csv/zip  -> 202 {id, estado: "en_proceso"}
GET  /api/analizar/{id}                    -> estado y, al terminar, el resumen
GET  /api/resultado/{id}                   -> el Excel
```

Todo llega por un solo campo repetido `archivos`, que acepta el .xlsx suelto
con sus CSV o un .zip con todo dentro. Medido: los 272 MB del paquete real
comprimen a **37.9 MB**.

- **Asíncrono**, con un `ThreadPoolExecutor` de **un solo worker**. No es sólo
  por el tiempo: dos análisis a la vez no caben en memoria (ver abajo).
- `corregir` y `forzar` como banderas. `forzar` analiza aun con errores que la
  corrección no arregla — sirve cuando lo roto es una hoja de la que depende
  sólo una parte del reporte, como las citas: afectan al scorecard del
  proveedor, no al Pareto. La validación completa viaja en la respuesta, así
  que se decide con los errores a la vista.
- Topes: 80 MB por archivo, 400 MB por paquete, y el mismo tope aplicado al
  contenido **expandido** del zip, que es lo que evita un zip bomb. De cada
  entrada del zip se toma sólo el nombre base, nunca la ruta.
- **Fly a 2 GB**, medido y no elegido al tanteo: el análisis deja el proceso
  en **843 MB de pico** — 400 MB tras leer las fuentes y el resto al construir
  el Excel, que openpyxl arma entero en memoria. Con los 512 MB anteriores la
  máquina moría a mitad.
- `read_only=True` en `leer_fuentes` y en el validador.

Probado de punta a punta contra la API con el paquete real en zip: **204 s**
(78 s de validación + 83 s de análisis + descompresión), mismos números que
el CLI, y el Excel de 4.3 MB se descarga. Los rechazos de entrada también:
sin layout, dos layouts, extensión no permitida, CSV que no corresponde a
ninguna hoja, zip corrupto.

---

### 3.6 Lo que apareció al correrlo — no estaba en el plan

- **`escribir_resultado` hacía `merge_cells` dentro de un loop de 38,087
  filas.** openpyxl compara cada rango fusionado contra todos los anteriores,
  así que es cuadrático y la escritura no terminaba. Quitado el merge (el
  texto se desborda sobre las columnas vacías y se lee igual): 32 s.
- **La advertencia de "falta el SKU en el catálogo" se emitía por día**, no
  por SKU. Con este layout eran ~25 mil líneas repetidas en el Excel y en el
  JSON de la API. Ahora sale una sola, agrupada y con ejemplos.

## 4. Orden sugerido

1. ~~§3.1 el lector de CSV~~ — hecho.
2. ~~§3.3 indexación de eventos y filtro por llave~~ — hecho.
3. ~~§3.2 `leer_fuentes` con paquete~~ — hecho.
4. ~~Corrida completa contra `datos_reales_v5/`~~ — hecho, 83 s.
5. ~~§3.4 spec y validador~~ — hecho, incluido el cruce entre fuentes.
6. ~~§3.5 API y Fly~~ — hecho y probado con el paquete real.
7. **Desplegar**: `flyctl deploy --ha=false`. Ojo: el cambio de VM a 2 GB va
   en el mismo deploy.
8. **Front** (documento aparte). El contrato cambió —subida multi-archivo y
   análisis asíncrono con poll— así que el front tiene que adaptarse sí o sí.
   A cambio, la respuesta ya trae las dos coberturas para encabezar.

---

## 5. Lo que hay que preguntarle a La Comer

- ~~¿Qué es la cuarta columna de `TABLEAU_VENTAS.csv`?~~ `importe_venta`.
- ~~`TABLEAU_VENTAS.csv` es de la tienda 237~~ — reextraído, ya es la 287.
- **`BOPS_OSA` trae SKU fuera de la división Abarrotes** (§2.2). ¿Se pide el
  export filtrado al alcance, o se filtra contra el catálogo del lado del
  modelo? Mientras no se decida, el 17.2% de cobertura de la portada mezcla
  dos universos y se lee como una falla del modelo que no es.
- **`CITAS_PROV_CEDIS` no cuadra con `COMPRAS_PEDIDOS_PROV`.** 61,671 citas
  traen un folio que no existe entre los pedidos, y 112,212 de 150,003
  pedidos (74.8%) no tienen cita — el modelo los dictamina como "el proveedor
  nunca agendó cita" (RC06). Con esa proporción, el scorecard de proveedores
  no se puede firmar: hay que confirmar si la extracción de citas vino
  completa. Se nota también en las tasas por arriba de 100% (BONAFONT
  confirmó 189.5% de lo pedido), que son de conciliación, no de desempeño.
- ~~`CEDIS_INVENTARIO` cruza con el 0.1%~~ — resuelto, ver §2.2. La ausencia
  de fila significa cero. Conviene ratificarlo por escrito: es el único punto
  del modelo donde vacío se lee como cero.
