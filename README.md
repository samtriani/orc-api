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
| `orcmm_pipeline.py` | Del Excel de captura al Excel de resultados: lee, deriva las banderas del día, clasifica y escribe. |
| `orcmm_validar_layout.py` | Revisa el archivo contra el spec **antes** de correrlo. No corrige nada. |
| `orcmm_corregir_layout.py` | Arregla lo que se puede arreglar solo. Nunca toca el original. |
| `api/` | Backend FastAPI que expone todo lo anterior. |
| `web/` | Front en Angular. |

---

## Línea de comandos

```bash
pip install -r requirements.txt

# 1. ¿el archivo cumple el layout?
python orcmm_validar_layout.py "040826_La Comer_Layout de datos RCA (OSA)_V2_Con Datos.xlsx"

# 2. arreglar lo que se pueda (escribe "<archivo> corregido.xlsx", no toca el original)
python orcmm_corregir_layout.py "040826_La Comer_Layout de datos RCA (OSA)_V2_Con Datos.xlsx"

# 3. clasificar
python orcmm_pipeline.py "040826_La Comer_Layout de datos RCA (OSA)_V2_Con Datos corregido.xlsx"
```

El pipeline acepta `--umbral-osa` (por omisión 100: se analizan todos los días
que no estén al 100% de disponibilidad) y `-o` para el nombre de salida.

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
| **Errores de layout** | Columnas que no existen, claves capturadas como número, tipos que no cuadran. El modelo leería mal. | Sí. Casi siempre los arregla `orcmm_corregir_layout`. |
| **Datos incompletos** | El layout está bien, faltan renglones (hoja vacía, citas parciales). | No. El motor ya sabe reportarlo como cobertura perdida nombrando el campo. |
| **Advertencias** | Vale la pena revisarlo antes de firmar el Pareto. | No. |

### Endpoints

| Método | Ruta | Para qué |
|---|---|---|
| `POST` | `/api/validar` | Sólo valida. Dice si el archivo es corregible y qué cambiaría. |
| `POST` | `/api/analizar?corregir=&umbral_osa=` | Valida, corrige si se pidió, clasifica y devuelve el resumen. |
| `GET` | `/api/resultado/{id}` | Descarga el Excel de resultados. |
| `GET` | `/api/salud` | Ping. |

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

## Estado actual del modelo

Dos interruptores viven en `orcmm_rca_engine.py` porque son **acuerdos de
negocio**, no detalles de implementación. Cambiarlos cambia a quién le cae la
venta perdida en el Pareto.

### `EVALUAR_PEDIDO_TIENDA = False` — temporal

SIMA todavía no entrega los pedidos de tienda. Con la prioridad 3 apagada, el
árbol pasa de largo de la pregunta "¿la tienda pidió?" y puede llegar hasta
CEDIS y proveedor.

- Sin el interruptor: **0 %** de cobertura, todo RC99.
- Con el interruptor: **92.3 %**.

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
