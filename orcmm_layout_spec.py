"""
ORCMM - Especificación única del layout de captura.

Este módulo es la ÚNICA fuente de verdad del formato del Excel. Lo importan:
  - generar_layout_datos.py   (escribe el archivo vacío para los equipos)
  - generar_ejemplo_lleno.py  (escribe el mismo archivo con datos de ejemplo)
  - orcmm_pipeline.py         (lo lee y clasifica)

Si aquí se agrega o quita un campo, los tres se mantienen sincronizados solos.

Estructura de cada hoja:
    fila 1  título
    fila 2  fuente / owner / ventana
    fila 3  ENCABEZADOS  <- el pipeline lee de aquí
    fila 4  tipo esperado (* = obligatorio)
    fila 5  descripción con ejemplo
    fila 6+ DATOS        <- el pipeline lee de aquí
"""

from __future__ import annotations

FILA_TITULO = 1
FILA_META = 2
FILA_ENCABEZADO = 3
FILA_TIPO = 4
FILA_DESCRIPCION = 5
FILA_DATOS = 6

TXT, ENT, DEC, FEC, LST, BOL, HOR = "Texto", "Entero", "Decimal", "Fecha", "Lista", "Sí/No", "Hora"

# De dónde llega cada hoja. Desde el layout V5, dos de ellas pasan del millón
# de filas que aguanta una hoja de Excel y se entregan como CSV aparte. El
# validador usa esto para saber qué exigir dentro del .xlsx y qué exigir como
# archivo suelto; sin el marcador, un layout V5 correcto se reporta como roto.
EXCEL, CSV, AMBOS = "excel", "csv", "ambos"

# Formas de los datos
ESTATICO = "Estático — una fila por SKU y tienda"
DIARIO = "Foto diaria — una fila por SKU, tienda y día"
EVENTO = "Evento — una fila por folio, con sus fechas"


HOJAS = {

    # -----------------------------------------------------------------------
    "CATALOGO": {
        "titulo": "Catálogo de SKUs por tienda",
        "forma": ESTATICO,
        "equipo": "Catálogo / Comercial",
        "owner": "Rebeca",
        "ventana": "Foto vigente. Si un SKU cambió de vía durante el periodo, una fila por vigencia.",
        "para_que": "Define si el SKU se surte vía CEDIS o directo del proveedor. Sin esto no se puede "
                    "elegir la rama del árbol (prioridad 4).",
        "campos": [
            ("sku", TXT, True, "Clave del artículo, como TEXTO para no perder ceros a la izquierda.", "7501059236776"),
            ("descripcion", TXT, False, "Descripción comercial.", "PAPEL HIGIENICO MAXIRESIST PETALO 18 PZA"),
            ("tienda", TXT, True, "Clave de tienda.", "287"),
            ("cedis_surtidor", TXT, True, "Clave del CEDIS que surte a esa tienda. Liga el inventario de CEDIS con la tienda.", "280"),
            ("division", TXT, True, "División del SKU.", "Abarrotes"),
            ("via_resurtido", LST, True, "Vía 1 (por CEDIS) / Vía 2 / DSD (directo del proveedor a tienda).", "Vía 1"),
            ("piezas_por_caja", ENT, True, "Unidad de empaque vigente. Convierte las cajas del reporte NS a piezas.", "12"),
            ("tipo_resurtido", LST, False,
             "Manual / Automático. Afina el responsable de RC03 'Pedido No Generado' "
             "(prioridad 3) cuando SIMA confirma que no hubo pedido de tienda.", "Automático"),
            ("rol_frecuencia", TXT, False, "Rol por frecuencia (base F9).", "A"),
            ("linea_vs_io", LST, False, "Línea / In&Out. No se usa para clasificar; sirve para segmentar el Pareto.", "Línea"),
            ("estatus_activo", BOL, True, "SKU vigente en catálogo para esa tienda.", "Sí"),
            ("Nombre_Tienda", TXT, False, "Nombre comercial de la tienda. Sólo para leer el "
                                          "reporte; ninguna regla lo usa.", "LA COMER COYOACAN"),
        ],
    },

    # -----------------------------------------------------------------------
    "TABLEAU_INV_TIENDA": {
        "titulo": "Inventario en tienda — diario",
        "forma": DIARIO,
        "origen": CSV,
        "equipo": "Tableau",
        "owner": "Andrés",
        "ventana": "2026-03-01 a 2026-03-31. Una fila por SKU, tienda y día, incluyendo los días en cero.",
        "para_que": "Primera pregunta del árbol: ¿había producto en la tienda? (prioridad 1).",
        "campos": [
            ("sku", TXT, True, "", "7501059236776"),
            ("tienda", TXT, True, "", "287"),
            ("fecha", FEC, True, "AAAA-MM-DD.", "2026-03-09"),
            ("existencia_piezas", ENT, True, "Existencia del día en piezas. CERO es un dato válido y necesario.", "0"),
            ("hora_de_corte", HOR, False,
             "Hora a la que se toma la foto de inventario. Ninguna regla la lee: "
             "se asume el cierre del día (23:59), que es como sale el reporte. "
             "Documenta el supuesto, no cambia ningún resultado.", "23:59"),
            ("existencia_minima_dia", ENT, False, "Mínimo alcanzado durante el día. Si existe, el motor lo prefiere sobre la foto.", "0"),
        ],
    },

    # -----------------------------------------------------------------------
    "BOPS_OSA": {
        "titulo": "OSA diario en anaquel",
        "forma": DIARIO,
        "equipo": "BOPS",
        "owner": "Por asignar",
        "ventana": "2026-03-01 a 2026-03-31.",
        "para_que": "Define qué días entran al análisis: sólo se clasifican los días con OSA menor a 100%. "
                    "Desde el layout V5 trae también la venta perdida, que es lo que pondera el Pareto.",
        "campos": [
            ("sku", TXT, True, "", "7501059236776"),
            ("tienda", TXT, True, "", "287"),
            ("fecha", FEC, True, "AAAA-MM-DD.", "2026-03-09"),
            ("osa_pct", DEC, True, "Disponibilidad en anaquel del día, de 0 a 100.", "12.0"),
            ("venta_perdida_estimada", DEC, True,
             "Venta perdida del día en pesos. Vive aquí y no en TABLEAU_VENTAS porque "
             "es BOPS quien sabe cuántas horas estuvo vacío el anaquel.", "1408.00"),
            ("numero_rupturas", ENT, False, "Cuántas veces se agotó el anaquel ese día.", "2"),
            ("minutos_sin_producto", ENT, False, "Duración total del hueco en minutos.", "845"),
        ],
    },

    # -----------------------------------------------------------------------
    "TABLEAU_VENTAS": {
        "titulo": "Ventas y venta perdida — diario",
        "forma": DIARIO,
        "origen": CSV,
        "equipo": "Tableau",
        "owner": "Andrés",
        "ventana": "2026-03-01 a 2026-03-31.",
        "para_que": "Contexto de venta real del día. Desde el layout V5 ya NO es la fuente de la venta "
                    "perdida: ésa vive en BOPS_OSA. El motor no depende de esta hoja para clasificar "
                    "ni para ponderar el Pareto.",
        "campos": [
            ("sku", TXT, True, "", "7501059236776"),
            ("tienda", TXT, True, "", "287"),
            ("fecha", FEC, True, "AAAA-MM-DD.", "2026-03-09"),
            ("importe_venta", DEC, True, "Venta real del día en pesos.", "0.00"),
            ("unidades_vendidas", ENT, False, "", "0"),
            ("venta_perdida_estimada", DEC, False,
             "Se dejó opcional al mudarse a BOPS_OSA. Si viene, el motor la usa sólo "
             "cuando BOPS_OSA no trae el dato de ese día.", "1408.00"),
            ("metodo_estimacion", TXT, False, "Cómo se calculó la venta perdida.", "Promedio 4 semanas"),
        ],
    },

    # -----------------------------------------------------------------------
    "CEDIS_INVENTARIO": {
        "titulo": "Inventario en CEDIS — diario",
        "forma": DIARIO,
        "equipo": "CEDIS",
        "owner": "Andrés",
        "ventana": "2026-03-01 a 2026-03-31.",
        "para_que": "Separa 'CEDIS no tenía' de 'CEDIS tenía y no lo mandó' (prioridades 5 a 8). "
                    "ÚNICA hoja donde la ausencia de fila SÍ significa cero: el reporte de CEDIS "
                    "lista los SKU con existencia y omite los que están en cero (confirmado con "
                    "La Comer, 2026-08-06). Aplica sólo a los días que la extracción cubre; un día "
                    "que no se extrajo sigue siendo un dato ausente.",
        "campos": [
            ("sku", TXT, True, "", "7501059236776"),
            ("cedis", TXT, True, "Debe coincidir con cedis_surtidor de la hoja CATALOGO.", "280"),
            ("fecha", FEC, True, "AAAA-MM-DD.", "2026-03-11"),
            ("existencia_piezas", ENT, True, "Existencia del día en CEDIS, en piezas.", "192"),
            ("piezas_reservadas", ENT, False, "Piezas ya comprometidas a otras tiendas. Distingue 'no surtió' de 'no tenía disponible'.", "0"),
        ],
    },

    # -----------------------------------------------------------------------
    "CEDIS_TRANSFERENCIAS": {
        "titulo": "Transferencias CEDIS a tienda",
        "forma": EVENTO,
        "equipo": "CEDIS / Logística",
        "owner": "Andrés",
        "ventana": "2026-02-01 a 2026-03-31. Arranca en febrero a propósito: un tránsito de febrero explica faltantes del 1 de marzo.",
        "para_que": "¿Salió mercancía de CEDIS? ¿Iba en camino ese día? (prioridades 2, 5 y 6).",
        "campos": [
            ("folio", TXT, True, "Identificador de la transferencia.", "TR-884213"),
            ("sku", TXT, True, "", "7501059236776"),
            ("cedis_origen", TXT, True, "", "280"),
            ("tienda_destino", TXT, True, "", "287"),
            ("fecha_generacion", FEC, True, "Cuándo se generó la transferencia en sistema.", "2026-03-14"),
            ("fecha_salida_cedis", FEC, True, "Cuándo salió físicamente del CEDIS. Abre la ventana de tránsito.", "2026-03-15"),
            ("fecha_recepcion_tienda", FEC, True, "Cuándo se recibió en tienda. VACÍO si no se había recibido al corte.", "2026-03-16"),
            ("cantidad_enviada_piezas", ENT, True, "", "192"),
            ("cantidad_recibida_piezas", ENT, False, "Permite detectar merma en tránsito.", "192"),
            ("estatus", TXT, False, "Estatus del sistema tal cual, sin recodificar.", "Recibido"),
        ],
    },

    # -----------------------------------------------------------------------
    "SIMA_PEDIDOS_TIENDA": {
        "titulo": "Pedidos de tienda a CEDIS",
        "forma": EVENTO,
        "equipo": "SIMA",
        "owner": "Gabriel Medina",
        "ventana": "2026-02-01 a 2026-03-31. Arranca en febrero a propósito: un pedido de febrero sigue vigente en marzo.",
        "para_que": "¿La tienda pidió el producto? (prioridad 3). Es el cuello de botella del árbol: sin esto, "
                    "nada de las prioridades 3 a 10 se puede clasificar.",
        "campos": [
            ("folio", TXT, True, "", "PD-1120945"),
            ("sku", TXT, True, "", "7501059236776"),
            ("tienda", TXT, True, "", "287"),
            ("fecha_pedido", FEC, True, "Cuándo se generó el pedido.", "2026-03-03"),
            ("fecha_requerida", FEC, False, "Fecha comprometida de surtido.", "2026-03-06"),
            ("cantidad_pedida_piezas", ENT, True, "", "192"),
            ("cantidad_surtida_piezas", ENT, True, "Cero si no se surtió.", "192"),
            ("fecha_surtido", FEC, True, "Cierra la vigencia del pedido. VACÍO si sigue abierto.", "2026-03-16"),
            ("estatus", TXT, False, "Estatus del sistema tal cual.", "Surtido"),
        ],
    },

    # -----------------------------------------------------------------------
    "COMPRAS_PEDIDOS_PROV": {
        "titulo": "Pedidos a proveedor con entrega a CEDIS",
        "forma": EVENTO,
        "equipo": "Compras / Abasto",
        "owner": "Andrés",
        "ventana": "2026-02-01 a 2026-03-31. Arranca en febrero a propósito.",
        "para_que": "¿Se le pidió al proveedor y entregó completo? (prioridades 7 y 8).",
        "campos": [
            ("folio", TXT, True, "", "OC-556120"),
            ("sku", TXT, True, "", "7501059236776"),
            ("proveedor_id", TXT, True, "", "PRV-0442"),
            ("proveedor_nombre", TXT, False, "", "Kimberly Clark"),
            ("cedis_destino", TXT, True, "", "280"),
            ("fecha_pedido", FEC, True, "", "2026-03-02"),
            ("fecha_cita", FEC, False, "Cita de entrega agendada.", "2026-03-05"),
            ("fecha_recibo", FEC, True, "Recepción real en CEDIS. VACÍO si no se recibió.", "2026-03-11"),
            ("cajas_pedidas", ENT, True, "EN CAJAS, no en piezas. Se convierte con piezas_por_caja del catálogo.", "40"),
            ("cajas_entregadas", ENT, True, "EN CAJAS. Cero si el proveedor no entregó.", "16"),
            ("estatus", TXT, False, "Estatus del sistema tal cual.", "Surtido parcial"),
        ],
    },

    # -----------------------------------------------------------------------
    "CITAS_PROV_CEDIS": {
        "titulo": "Citas de entrega del proveedor en CEDIS",
        "forma": EVENTO,
        "equipo": "Compras / Abasto",
        "owner": "Andrés",
        "ventana": "2026-02-01 a 2026-03-31. Arranca en febrero a propósito, igual que los pedidos a proveedor.",
        "para_que": "Refina la prioridad 8. El pedido dice qué se le pidió al proveedor; la cita dice a qué "
                    "se comprometió y si lo cumplió. Separa tres fallas distintas que hoy se ven iguales: "
                    "nunca agendó cita, confirmó menos cajas de las pedidas, o no cumplió su propia cita. "
                    "Además evita culpar al proveedor por un pedido cuya cita todavía no vence.",
        "campos": [
            ("folio", TXT, True, "Folio del PEDIDO a proveedor. Debe existir en COMPRAS_PEDIDOS_PROV: es la liga entre ambas hojas.", "26300915590"),
            ("folio_cita", TXT, True, "Número de cita. Único por cita; un pedido puede tener varias.", "922966"),
            ("sku", TXT, True, "", "7501059236776"),
            ("proveedor_id", TXT, True, "", "PRV-0442"),
            ("proveedor_nombre", TXT, False, "", "Kimberly Clark"),
            ("cedis_destino", TXT, True, "", "280"),
            ("fecha_pedido", FEC, True, "La misma del pedido en COMPRAS_PEDIDOS_PROV.", "2026-03-02"),
            ("fecha_cita", FEC, True, "Fecha agendada de entrega. Es la fecha del compromiso del proveedor.", "2026-03-05"),
            ("cajas_confirmadas_cita", ENT, True, "EN CAJAS. A cuánto se comprometió el proveedor AL AGENDAR. Si es igual a cajas_pedidas siempre, avisar: el dato se está copiando del pedido.", "40"),
            ("cajas_entregadas", ENT, True, "EN CAJAS recibidas EN ESA CITA. Cero si no se presentó.", "16"),
            ("estatus_cita", TXT, False, "Estatus del sistema tal cual (Confirmada / Cancelada / Reprogramada / No se presentó).", "Confirmada"),
            ("fecha_entrega_real", FEC, False, "Sólo si la mercancía llegó en fecha distinta a la cita. Vacío si llegó el día de la cita.", "2026-03-06"),
        ],
    },
}


# ---------------------------------------------------------------------------
# Encabezados tolerados
#
# Los equipos entregan a veces la columna con el nombre "de negocio" en vez del
# nombre técnico. En lugar de romper la lectura, se traduce aquí y el validador
# lo reporta como renombre pendiente. La llave se compara en minúsculas y con
# los espacios colapsados.
# ---------------------------------------------------------------------------

ALIAS_ENCABEZADOS = {
    "CITAS_PROV_CEDIS": {
        "folio (cambiar a pedido)": "folio",
        "folio pedido": "folio",
        "pedido": "folio",
        "# de cita": "folio_cita",
        "no. de cita": "folio_cita",
        "numero_cita": "folio_cita",
        "cita": "folio_cita",
        "estatus cita": "estatus_cita",
    },
}


def origen_de(hoja: str) -> str:
    """De dónde se espera esa hoja. Por omisión, del Excel."""
    return HOJAS.get(hoja, {}).get("origen", EXCEL)


def normalizar_encabezado(hoja: str, texto):
    """Nombre canónico del campo a partir del encabezado tal como vino."""
    if texto is None:
        return None
    limpio = " ".join(str(texto).split())
    if not limpio:
        return None
    return ALIAS_ENCABEZADOS.get(hoja, {}).get(limpio.lower(), limpio)


REGLAS_DE_LLENADO = [
    ("Pegar desde la fila 6",
     "Las filas 1 a 5 de cada hoja son la guía: título, fuente, nombre del campo, tipo esperado y "
     "descripción. No hay que borrarlas ni moverlas; el programa lee los encabezados de la fila 3."),
    ("Vacío NO es cero",
     "Una celda vacía significa SIN DATO. Un cero significa 'se confirmó que no había'. Llevan a "
     "resultados opuestos: cero inventario abre la evaluación de CEDIS y proveedor; sin dato deja el "
     "día sin clasificar. No rellenar huecos con 0."),
    ("Los campos con asterisco son obligatorios",
     "Sin ellos el modelo no puede evaluar la regla correspondiente y ese día queda sin clasificar."),
    ("Registros tal como salen del sistema",
     "No calcular columnas Sí/No como '¿había tránsito?'. Esas se derivan después, con una definición "
     "única y auditable. Entregar los folios y sus fechas."),
    ("Formatos",
     "SKU y tienda como TEXTO, para no perder ceros a la izquierda ni pasar a notación científica. "
     "Todas las fechas en AAAA-MM-DD. Estatus del sistema sin recodificar."),
    ("Las hojas de eventos arrancan en febrero",
     "CEDIS_TRANSFERENCIAS, SIMA_PEDIDOS_TIENDA, COMPRAS_PEDIDOS_PROV y CITAS_PROV_CEDIS se extraen desde "
     "el 1 de febrero aunque el análisis sea de marzo: un faltante del 1 de marzo se explica con un pedido "
     "y un tránsito generados en febrero."),
    ("CITAS_PROV_CEDIS tiene que venir COMPLETA",
     "El modelo lee 'pedido sin fila en CITAS' como 'el proveedor nunca agendó cita', y eso dictamina "
     "contra el proveedor. Si la extracción trae sólo una parte de las citas del periodo, hay que decirlo: "
     "un hueco de captura se vuelve una acusación."),
]
