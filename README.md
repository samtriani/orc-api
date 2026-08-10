# ORCMM — Clasificación de desabasto (OSA) por causa raíz · backend

Toma el Excel de captura que llenan los equipos de La Comer, revisa que el
layout venga bien y clasifica cada día de faltante en una de las seis causas
raíz de la matriz (RC01–RC06), con su responsable y su venta perdida.

Se puede usar de dos formas: desde la línea de comandos, o desde la API que
consume el front. Las dos corren exactamente el mismo motor.

> **El front vive en [samtriani/orc-gui](https://github.com/samtriani/orc-gui).**

> **¿Retomando en otra máquina?** Lee [`SEGUIMIENTO.md`](SEGUIMIENTO.md)
> primero — resume sesión a sesión qué se hizo con Claude Code y qué falta.

## Los archivos de datos no están en el repo

Las capturas traen SKU, tiendas, proveedores, folios de pedido y venta perdida
de un cliente, así que `*.xlsx` está en `.gitignore`. Para correr cualquier
ejemplo de este README hay que poner el archivo de captura en la raíz del
proyecto.

---

## Los módulos

| Archivo | Qué hace |
|---|---|
| `orcmm_layout_spec.py` | **Fuente única del layout.** Las 9 hojas, sus campos, tipos y ventanas. Si aquí se agrega un campo, todo lo demás se entera solo. |
| `orcmm_rca_engine.py` | El motor. Las 10 reglas de la matriz en orden de prioridad. Clasifica **un día** a la vez. |
| `orcmm_rca_periodo.py` | Agrega los veredictos diarios: por SKU-tienda, Pareto de causas y responsables, cobertura del modelo. |
| `orcmm_fuentes_csv.py` | Lee las fuentes que ya no caben en una hoja de Excel y llegan como CSV de Tableau. |
| `orcmm_pipeline.py` | De la captura al Excel de resultados: lee, deriva las banderas del día, clasifica y escribe. |
| `orcmm_validar_layout.py` | Revisa el paquete contra el spec **antes** de correrlo, y cruza las fuentes entre sí. No corrige nada. |
| `orcmm_corregir_layout.py` | Arregla lo que se puede arreglar solo. Nunca toca el original. |
| `orcmm_db.py` | Conexión a Postgres y upsert genérico — lo usan los tres scripts de abajo. |
| `orcmm_db_init.py` | Aplica `sql/schema.sql` (crea las tablas). |
| `orcmm_db_borrar.py` | Vacía las tablas de datos operativos, sin tocar el esquema ni los catálogos informativos. |
| `orcmm_etl_carga.py` | Carga el layout principal (xlsx + CSV) a Postgres. |
| `orcmm_etl_catalogos.py` | Carga sucursales y el catálogo de SKU por tienda (aparte del layout de captura). |
| `orcmm_fuentes_db.py` | Arma un `Fuentes` desde Postgres en vez de un archivo — lo usa `/api/analizar-tienda`. |
| `api/` | Backend FastAPI que expone todo lo anterior. |
| `web/` | Front en Angular. |

---

## Línea de comandos

Desde el layout V5, la captura son **varios archivos**: el `.xlsx` con la
mayoría de las hojas, más los CSV de las dos que ya no caben en una hoja de
Excel (`TABLEAU_INV_TIENDA` son 2.7 millones de filas, repartidas en seis
archivos). Los comandos reciben el Excel primero y los CSV después; se
reparten solos por el prefijo del nombre de archivo.

```bash
pip install -r requirements.txt

# 1. ¿el paquete cumple el layout? (también cruza las fuentes entre sí)
python orcmm_validar_layout.py "layout.xlsx" datos/*.csv

# 2. arreglar lo que se pueda (escribe "<archivo> corregido.xlsx", no toca el original)
python orcmm_corregir_layout.py "layout.xlsx"

# 3. clasificar
python orcmm_pipeline.py "layout.xlsx" datos/*.csv
```

Un layout anterior al V5, con todas las hojas dentro del Excel, se sigue
corriendo igual: sin los CSV.

El pipeline acepta `--umbral-osa` (por omisión 100: se analizan todos los días
que no estén al 100% de disponibilidad) y `-o` para el nombre de salida.

Para revisar un export de Tableau por su cuenta —qué encabezados detecta, qué
periodo cubre y qué tanto cruza contra los días con faltante— sin correr el
análisis completo:

```bash
python orcmm_fuentes_csv.py datos/*.csv --contra "layout.xlsx"
```

### Auditar un SKU

El resultado dictamina 30 mil días de golpe. Cuando uno se ve raro —o cuando un
SKU que debería aparecer no aparece— hay que ir a las ocho fuentes y cruzarlas.
`orcmm_expediente.py` hace ese cruce y deja el dato crudo de cada hoja al lado
del veredicto:

```bash
python orcmm_expediente.py 663985002478 "layout.xlsx" datos/*.csv \
    --tienda 287 --resultado "Resultado RCA.xlsx"
```

Sólo el layout es obligatorio; sin CSV y sin `--resultado` se omiten esas
secciones. Sirve sobre todo para el caso mudo: **un SKU que no produjo ni una
fila**. Casi siempre es que `BOPS_OSA` nunca lo reportó —el análisis arranca de
los días con faltante que entrega BOPS—, y eso no se ve en ningún lado porque
el SKU simplemente no está. El expediente lo dice con todas sus letras.

---

## Aplicación web

Dos procesos, cada uno en su terminal.

**Backend** — desde la raíz del proyecto:

```bash
pip install -r requirements.txt
python -m uvicorn api.main:app --reload --port 8000
```

Queda en `http://localhost:8000`; la documentación interactiva, en
`http://localhost:8000/docs`.

**Frontend** — está en su propio repo:

```bash
git clone https://github.com/samtriani/orc-gui.git
cd orc-gui
npm install
npm start
```

Queda en `http://localhost:4200`. En desarrollo `/api` va por proxy al puerto
8000, así que el navegador ve un solo origen y no hay que pelearse con CORS.
El backend además acepta CORS desde `localhost:4200` (ver `api/main.py`), por
si se sirve el front de otra forma.

### El flujo

1. Se sube el `.xlsx` de captura.
2. El back lo valida contra el spec. Si hay errores, además corre el corrector
   sobre una copia y vuelve a validar: es la única forma honesta de saber si el
   arreglo automático sirve para **ese** archivo.
3. Si el layout viene bien, analiza directo. Si trae errores corregibles, la
   pantalla ofrece **"Corregir y analizar"**, que aplica la corrección y sigue.
4. Devuelve el resumen en pantalla —cobertura, Pareto por causa y responsable,
   desempeño del proveedor, qué dato bloqueó la clasificación— y el Excel de
   resultados para descargar.

### Las tres severidades de la validación

No todo lo que está mal impide correr, y confundirlo deja la herramienta
parada en la puerta:

| | Qué es | ¿Bloquea? |
|---|---|---|
| **Errores de layout** | Columnas que no existen, obligatorios vacíos, claves deformadas al capturarse, hojas que no cuadran entre sí. El modelo leería mal. | Sí, salvo que se pase `forzar`. El validador marca cuáles sabe arreglar `orcmm_corregir_layout`. |
| **Datos incompletos** | El layout está bien, faltan renglones (hoja vacía, citas parciales, fuentes que no cruzan). | No. El motor ya sabe reportarlo como cobertura perdida nombrando el campo. |
| **Advertencias** | Vale la pena revisarlo antes de firmar el Pareto. | No. |

### Endpoints

Todo se sube por un solo campo repetido `archivos`: el `.xlsx` con sus CSV
sueltos, o un `.zip` con todo dentro. **El zip es lo recomendable** — los CSV
del layout real pesan 222 MB sueltos y menos de 40 MB comprimidos.

| Método | Ruta | Para qué |
|---|---|---|
| `POST` | `/api/validar` | Sólo valida. Dice si el paquete es corregible y qué cambiaría. |
| `POST` | `/api/analizar?corregir=&forzar=&umbral_osa=` | Encola el análisis desde un archivo subido. Responde `202` con el id. |
| `POST` | `/api/analizar-tienda` | Igual, pero leyendo Postgres: `{tienda, desde, hasta, umbral_osa}` en vez de un archivo. Ver [Persistencia en Postgres](#persistencia-en-postgres-neon). |
| `GET` | `/api/tiendas` | Tiendas con datos cargados en Postgres, para el selector del front. |
| `GET` | `/api/analizar/{id}` | Estado (`en_proceso`) y, al terminar, el resumen. Mismo endpoint para los dos tipos de análisis. |
| `GET` | `/api/resultado/{id}` | Descarga el Excel de resultados. |
| `GET` | `/api/salud` | Ping. |

El análisis es **asíncrono**: con volumen real tarda un par de minutos y un
request abierto tanto tiempo se lo lleva cualquier proxy de por medio. El
front encola y hace poll. Corre un solo análisis a la vez, porque dos no caben
en la memoria de la máquina.

`forzar` analiza aun con errores que la corrección automática no puede
arreglar. Sirve cuando lo roto es una hoja de la que depende sólo una parte
del reporte —una extracción de citas incompleta afecta al scorecard del
proveedor, no al Pareto— y el resultado se puede leer sabiendo eso. La
validación completa viaja en la respuesta: se decide con los errores a la
vista, no a ciegas.

El servidor no guarda nada permanente: cada análisis vive en una carpeta
temporal que se borra sola a la hora.

---

## Despliegue en Fly.io

```bash
fly auth login
fly launch --no-deploy --copy-config --name orc-api --region qro
fly deploy
```

Después, para que el front desplegado pueda hablarle:

```bash
fly secrets set ORCMM_ORIGENES="https://<dominio-del-front>"
```

Sin esa variable el backend sólo acepta al front de desarrollo
(`localhost:4200`), que es lo que se quiere en local.

### Por qué una sola máquina

`fly.toml` fija `min_machines_running = 1` y `auto_stop_machines = 'off'` a
propósito. El servicio guarda el Excel de resultados en un directorio temporal
y su índice en memoria del proceso: con dos máquinas, la petición de descarga
puede caer en la que **no** corrió el análisis y responder 404.

Mientras el estado no salga del proceso, esto no se escala en horizontal. Si
algún día hace falta, las opciones son un volumen compartido, un almacén de
objetos, o devolver el Excel en la misma respuesta del análisis.

`auto_stop_machines = 'off'` evita que la máquina se duerma entre el análisis
y la descarga y se lleve el archivo consigo.

---

## Persistencia en Postgres (Neon)

Las 9 hojas del layout, más dos catálogos informativos, viven también en
Postgres — así `/api/analizar-tienda` puede correr el mismo motor sin
volver a subir el Excel. El upload sigue funcionando igual; esto es
aditivo. `DATABASE_URL` va en un `.env` local (`.gitignore` lo excluye) y,
en Fly.io, como secreto:

```bash
fly secrets set DATABASE_URL="postgresql://usuario:password@host/db?sslmode=require"
```

### Aplicar el esquema

```bash
python orcmm_db_init.py
```

Corre `sql/schema.sql` (12 tablas: las 9 del layout, `sucursales`,
`catalogo_sku_tienda` y `etl_cargas` de bitácora). Usa `IF NOT EXISTS` en
todo, así que es seguro volver a correrlo.

### Cargar datos

```bash
# El layout principal (igual que la línea de comandos: xlsx + CSV sueltos)
python orcmm_etl_carga.py "layout.xlsx" datos/*.csv

# Los catálogos informativos (aparte, no son parte del layout de captura)
python orcmm_etl_catalogos.py \
    --sucursales "Listado_sucursales....xlsx" \
    --sku "Catalogo_Abarrotes_Coyoacan.xlsx" "Catalogo_Abarrotes_Centenario.xlsx" ...
```

Todo se carga por **UPSERT** sobre la llave natural de cada hoja — nunca
reemplazo total. Por eso volver a correrlo con una entrega corregida (o la
misma entrega dos veces) es seguro: no duplica filas.

### Borrar datos

```bash
python orcmm_db_borrar.py --si
```

Vacía las 9 tablas del layout + la bitácora, **sin tocar el esquema ni los
catálogos informativos** (`sucursales`/`catalogo_sku_tienda` tienen su
propio ciclo de vida — no se recargan cada vez que llega una versión nueva
de los datos operativos). Para incluirlos también: `--con-catalogos`. Sin
`--si` no borra nada, sólo avisa qué haría.

---

## Estado actual del modelo

Dos interruptores viven en `orcmm_rca_engine.py` porque son **acuerdos de
negocio**, no detalles de implementación. Cambiarlos cambia a quién le cae la
venta perdida en el Pareto.

### `EVALUAR_PEDIDO_TIENDA = False` — temporal

SIMA todavía no entrega los pedidos de tienda. Con la prioridad 3 apagada, el
árbol pasa de largo de la pregunta "¿la tienda pidió?" y puede llegar hasta
CEDIS y proveedor.

- Sin el interruptor: **0 %** de cobertura, todo RC99.
- Con el interruptor, sobre el layout V5 real: **99.6 %** de los días que
  entran al alcance (el 92.3 % que decía antes esta línea se midió sobre un
  archivo de ejemplo de un solo SKU, no sobre volumen real).

El costo es que **RC03 "Pedido No Generado" se vuelve inalcanzable** y sus días
se reparten entre RC04, RC05 y RC06. Sirve para leer la rama de abasto; no para
repartir responsabilidades en firme. El aviso va en rojo en tres hojas del Excel
y en la pantalla. **Cuando llegue SIMA: poner `True` y volver a correr.** Si
llega y se olvida, el pipeline avisa que está ignorando datos que ya existen.

### Pendientes de ratificar con La Comer

- `RESPONSABLE_SIN_CITA = Responsable.PROVEEDOR` — se asume que el proveedor
  solicita la cita. Si la agenda Compras, cambiarlo.
- `CLASIFICAR_CITA_PENDIENTE = True` — un faltante anterior a la cita del
  proveedor se dictamina RC05 / Compras-Abasto. El responsable es el correcto,
  pero la etiqueta RC05 ("Pedido Proveedor No Generado") queda forzada, porque
  el pedido sí existe. En `False` se reporta como hueco de la matriz.

### Huecos conocidos de la matriz

- Entrega completa del proveedor con CEDIS en cero: ninguna de las 10 reglas lo
  cubre. Cae como RC99 nombrando el hueco.
- `CITAS_PROV_CEDIS`: 23 de 33 pedidos no traen cita. El modelo lo lee como "el
  proveedor nunca agendó" (RC06). **Falta que Compras confirme** que la
  extracción trae todas las citas del periodo; si vino parcial, es un hueco de
  captura convertido en acusación.
