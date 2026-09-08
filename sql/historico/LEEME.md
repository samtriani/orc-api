# Histórico — NO ejecutar

Estos archivos son la historia de cómo se fue construyendo la base durante el
desarrollo: el esquema original de agosto de 2026 y las siete migraciones que
se le aplicaron encima, en orden.

**Para instalar la base se usa `../orcmm_ddl_completo.sql`**, que ya los
incluye a todos consolidados y es el que se le entrega al cliente.

Se conservan por una sola razón: la base que corre hoy se construyó
aplicándolos en este orden, así que si alguna vez hay que explicar por qué una
columna es como es, la respuesta está aquí con su fecha y su motivo.

| Archivo | Qué agregó |
|---|---|
| `schema.sql` | Esquema original: las 9 tablas del layout, `sucursales`, `catalogo_sku_tienda` y `etl_cargas`. |
| `migracion_v8.sql` | Columnas del layout V8: proveedor en catálogo, banderas de alerta en BOPS, tienda destino en compras, mínimo intradía en inventario. |
| `migracion_sima_origen.sql` | `sima_pedidos_tienda.tienda` pasa a llamarse `origen`: dejó de ser siempre una tienda (el `300` es el proceso central). |
| `migracion_catalogo_comercial.sql` | Categoría, subcategoría y marca, para poder filtrar el tablero por la jerarquía completa. |
| `migracion_runs.sql` | Tabla `runs`: cada análisis ejecutado, con la versión del motor que lo produjo. |
| `migracion_run_dias.sql` | Tabla `run_dias`: el veredicto diario con su evidencia. |
| `migracion_busqueda_skus.sql` | Extensiones `unaccent` y `pg_trgm` más el índice de búsqueda de producto por nombre. |
| `migracion_decil.sql` | `catalogo.decil`, para filtrar el análisis por rotación del SKU. |

## Una diferencia que vale la pena conocer

El archivo consolidado **no es una copia literal** de estos: al armarlo se
comparó contra la base en producción y apareció una diferencia. En
`schema.sql`, `compras_pedidos_prov.cedis_destino` está como `NOT NULL`, pero
en la base real es nullable — se relajó en algún momento sin dejar migración.

Medido sobre la información real, 4,967 de 852,994 pedidos (0.6%) llegan sin
CEDIS destino. El archivo consolidado la declara nullable, que es lo correcto:
con `NOT NULL` esas filas se rechazarían al cargar.
