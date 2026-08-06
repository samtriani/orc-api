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

## Sesión 2026-08-06

**Se hizo:**

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
