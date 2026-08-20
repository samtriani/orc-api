"""
ORCMM - OSA Root Cause Management Model
Motor de clasificación de desabasto (Root Cause Analysis)

Implementa la matriz de causa raíz validada con La Comer:
  Archivo: "040826_La Comer _ Clasificación desabasto (OSA)_V1.xlsx"
  Hoja:    "Matriz Causa raiz"

Taxonomía oficial: 6 causas raíz (RC01-RC06) + RC99 "Sin clasificar".
Se evalúa a nivel SKU + TIENDA + FECHA, en el orden de prioridad de la
matriz (1 a 10); la primera condición que se cumple dictamina.

Principio de diseño: la matriz es DETERMINISTA. No se emite un dictamen
con evidencia incompleta. Si una variable requerida por la regla en turno
no tiene dato, el caso se marca RC99 indicando exactamente qué falta,
en lugar de asumir un valor y arrastrar el error al Pareto.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from enum import Enum
from typing import Dict, List, Optional, Union


# ---------------------------------------------------------------------------
# 1. TAXONOMÍA OFICIAL (hoja "Matriz Causa raiz", filas 17-23)
# ---------------------------------------------------------------------------

class CausaRaiz(str, Enum):
    RC01 = "Ejecución en Tienda"
    RC02 = "Transporte / Tránsito"
    # Neutra a propósito: la causa cubre los dos casos y el responsable los
    # separa —automático a Compras, manual a Tienda—, así que el nombre no
    # puede comprometerse con uno. Que fue el sistema lo dice la subcausa,
    # que sólo aparece cuando el catálogo lo respalda. "De tienda" y no "a
    # tienda": el pedido va de la sucursal al CEDIS.
    RC03 = "Pedido de tienda no generado"
    RC04 = "CEDIS No Surtió"
    RC05 = "Pedido Proveedor No Generado"
    RC06 = "Incumplimiento Proveedor"
    RC99 = "Sin clasificar"


class Responsable(str, Enum):
    TIENDA = "Tienda"
    LOGISTICA = "Logística"
    TIENDA_ABASTO = "Tienda / Abasto"
    CEDIS = "CEDIS"
    COMPRAS_ABASTO = "Compras / Abasto"
    PROVEEDOR = "Proveedor"
    PENDIENTE = "Pendiente"


class SubcausaProveedor(str, Enum):
    """Refinamiento de la prioridad 8 con la hoja CITAS_PROV_CEDIS.

    No es una causa raíz nueva: la taxonomía sigue siendo RC01-RC06. Es el
    detalle de QUÉ hizo el proveedor, que es lo que cambia la conversación
    con él. 'No entregó' puede significar tres cosas muy distintas:

      - nunca agendó la cita              -> el pedido nunca entró a su calendario
      - agendó por menos de lo pedido     -> recortó el compromiso desde el inicio
      - agendó y no cumplió su propia cita-> incumplió lo que él mismo confirmó

    La primera se negocia con el área de citas; la segunda es un problema de
    capacidad o de asignación del proveedor; la tercera es incumplimiento puro.
    """
    SIN_CITA = "Nunca agendó cita para el pedido"
    CROSSDOCK_PEDIDO_TARDE = ("Entrega completa en crossdock; el pedido se "
                              "generó tarde para cubrir el consumo")
    CITA_PENDIENTE = "Cita agendada para después del día del faltante"
    RECORTE_EN_CITA = "Confirmó en cita menos cajas de las pedidas"
    NO_SE_PRESENTO = "Cita vencida sin entrega"
    ENTREGA_PARCIAL = "Entregó menos de lo que confirmó en la cita"


class SubcausaEjecucionTienda(str, Enum):
    """Refinamiento de la prioridad 1 (RC01) con las banderas de alerta de BOPS.

    NO cambia la causa ni el responsable: sigue siendo RC01 / Tienda. Sólo
    afina el PORQUÉ del hueco de anaquel habiendo inventario en trastienda,
    que es lo que cambia la conversación con el gerente de tienda:

      - alerta enviada y NO atendida -> el sistema avisó y la tienda no surtió
      - alerta enviada y atendida    -> surtió y aun así el hueco persistió
                                        (surtido insuficiente o rezago de medición)
      - sin alerta                    -> nunca se le avisó a la tienda del hueco
                                        (posible falla del sistema de alertas)

    Fuente: BOPS_OSA.alerta_enviada / alerta_ejecutada (layout V8). Sin las
    banderas (None) no se refina: se respeta 'vacío no es cero'.
    """
    ALERTA_NO_ATENDIDA = "Alerta de BOPS enviada y no atendida en tienda"
    ALERTA_ATENDIDA_PERSISTE = "Alerta atendida pero el hueco de anaquel persistió"
    SIN_ALERTA = "Sin alerta de BOPS: la tienda no fue notificada del hueco"


# Un día cuyo SKU no está en el catálogo de la tienda no es un día que el
# modelo no supo explicar: es un día que no le tocaba explicar. En el layout
# V5, BOPS_OSA entrega SKU de divisiones fuera del alcance (perfumería,
# electrónica) mientras que catálogo, inventario y pedidos vienen filtrados a
# Abarrotes. Sin separarlos, esos días entran al denominador de la cobertura
# como si fueran fuentes faltantes y hunden el resultado.
#
# Se marcan con este campo faltante para que la cobertura pueda reportar dos
# cifras: sobre el alcance y sobre todo lo que llegó. No se descartan — que
# BOPS mande SKU de más es justamente lo que hay que corregir en el origen.
FUERA_DE_CATALOGO = "sku_fuera_del_catalogo_de_la_tienda"

# Código propio para los días fuera de alcance, separado de RC99.
#
# No son lo mismo y mezclarlos miente: RC99 es "el modelo no supo explicarlo"
# —una deuda del análisis— y esto es "no le tocaba explicarlo", porque BOPS
# entregó un SKU que el catálogo de la tienda no reconoce. En la corrida de
# Coyoacán eran 145,645 días contra 30,245 de RC99 real: el reporte diario
# enseñaba 175,890 "Sin clasificar" y el 83% no era culpa del modelo.
#
# No entra al Pareto ni a la cobertura —eso ya lo filtraba dentro_del_alcance—
# pero sí aparece en la clasificación diaria, que es donde se auditan.
RC00_FUERA_DE_ALCANCE = "RC00"
CAUSA_FUERA_DE_ALCANCE = "Fuera de alcance · división no analizada"

# ===========================================================================
# EXCLUIR DEL ALCANCE LOS SKU QUE SIMA NO TRAE   (2026-08-20)
#
# APAGADO. Se construyó cuando parecía que SIMA venía incompleta y los SKU sin
# pedido llenaban el Pareto de RC99. La Comer confirmó que la extracción está
# completa: si un SKU del catálogo no aparece en SIMA es porque la tienda NO
# lo pidió en el periodo. Eso es evidencia, no un hueco — el día se clasifica
# como RC03 y ya no hay nada que excluir. Ver derivar_pedido_tienda.
#
# Se deja el interruptor porque el mecanismo sigue siendo correcto para el
# caso que sí lo justifica: una entrega parcial. En True, los SKU sin ningún
# pedido salen del alcance —no entran al Pareto, ni al waterfall, ni a la
# cobertura, ni al denominador del OSA— y el resultado reporta cuántos se
# dejaron fuera, para que el front lo diga.
#
# Ojo si se vuelve a prender: sube la cobertura a costa de achicar el
# universo, y el Pareto queda calculado sobre los SKU que sí tienen pedido,
# que suelen ser los de más rotación. Es sesgo de selección, y sin el letrero
# a la vista el resultado miente por omisión.
EXCLUIR_SKU_SIN_SIMA = False

SIN_DATO_SIMA = "sku_sin_pedidos_en_sima"
CAUSA_SIN_DATO_SIMA = "Fuera de alcance · sin datos de SIMA"


class TipoResurtido(str, Enum):
    """Fuente: CATALOGO.tipo_resurtido. Refina el responsable de la prioridad 3
    (RC03 'Pedido de tienda no generado') cuando SIMA confirma que no hubo pedido de
    tienda: ver RESPONSABLE_PEDIDO_NO_GENERADO más abajo.
    """
    MANUAL = "Manual"
    AUTOMATICO = "Automático"


class SubcausaPedidoTienda(str, Enum):
    """Refinamiento de la prioridad 3 con CATALOGO.tipo_resurtido.

    'No se generó pedido' es una causa distinta según quién debía generarlo:
      - resurtido automático -> falló el algoritmo de resurtido (o su
        parametrización), no una persona en tienda.
      - resurtido manual     -> nadie en tienda lo generó a mano.
    Sin el dato de catálogo no se puede distinguir cuál de las dos aplica.
    """
    AUTOMATICO_NO_GENERO = "Resurtido automático: el sistema no generó el pedido"
    MANUAL_NO_GENERO = "Resurtido manual: no se generó el pedido a mano"


# ===========================================================================
# PRIORIDAD 3 / SIMA_PEDIDOS_TIENDA   (apagada 2026-08-05, PRENDIDA 2026-08-18)
#
# SIMA ya entrega los pedidos de tienda, así que la prioridad 3 vuelve a
# evaluarse y RC03 'Pedido de tienda no generado' es alcanzable otra vez.
#
# Lo que estuvo pasando mientras estuvo apagada, y que ahora se corrige: la
# regla 3 se saltaba y el árbol pasaba de largo del inventario en tienda a la
# vía de resurtido, como si la pregunta '¿la tienda pidió?' no existiera. Los
# días que le tocaban a RC03 se repartieron entre RC04, RC05 y RC06 — o sea,
# se le estuvo cobrando a CEDIS y al proveedor un faltante que en realidad se
# explicaba porque la tienda nunca pidió.
#
# CONSECUENCIA AL PRENDERLA: el Pareto se mueve. Los días que hoy salen como
# RC04/RC05/RC06 y que tengan un pedido de tienda ausente van a pasar a RC03 /
# Tienda. No es un dato nuevo: es el que faltaba para no culpar al de junto.
#
# OJO con la cobertura de la hoja: si SIMA llega recortada a unos días, un día
# sin pedido NO significa que la tienda no pidió, sino que no se sabe — y la
# regla lo leería como culpa de la tienda. Antes de publicar un Pareto con
# esto, confirmar que la extracción cubre TODO el periodo analizado.
# ===========================================================================

EVALUAR_PEDIDO_TIENDA = True

NOTA_SIN_SIMA = ("Prioridad 3 omitida: sin datos de SIMA no se sabe si la tienda pidió")


# ---------------------------------------------------------------------------
# Refinamiento de RC01 con las banderas de alerta de BOPS (layout V8).
#
# No cambia la causa ni el responsable —sigue siendo RC01 / Tienda—, sólo
# agrega la subcausa que explica el porqué del hueco. Se deja como constante,
# no enterrada en un if, porque es un acuerdo de negocio:
#   True  -> se afina la subcausa de RC01 con alerta_enviada/alerta_ejecutada
#   False -> RC01 no se refina (comportamiento previo al V8)
# PENDIENTE DE RATIFICAR CON LA COMER: si 'sin alerta' debe seguir siendo
# responsabilidad de la tienda o abrir una causa de sistema de alertas. Por
# ahora se mantiene RC01 / Tienda en los tres casos y sólo se distingue el
# detalle, tal como se acordó.
REFINAR_RC01_CON_ALERTA = True


# ---------------------------------------------------------------------------
# Decisiones de modelado del bloque de citas — PENDIENTES DE RATIFICAR
#
# Se dejan como constantes, no enterradas en un if, porque son acuerdos de
# negocio y no detalles de implementación. Cambiar el valor cambia a quién le
# cae la venta perdida en el Pareto.
# ---------------------------------------------------------------------------

# ¿De quién es un pedido que nunca llegó a tener cita? En el proceso de La
# Comer el proveedor solicita la cita, así que por omisión es suyo. Si en
# realidad la agenda Compras, cambiar a Responsable.COMPRAS_ABASTO.
RESPONSABLE_SIN_CITA = Responsable.PROVEEDOR

# Faltante en tienda mientras el pedido a proveedor sigue en tiempo (su cita
# es posterior al día del faltante). El proveedor NO está incumpliendo: el
# hueco es de cobertura, el pedido se generó tarde para el consumo.
#   True  -> se dictamina RC05 / Compras-Abasto con subcausa CITA_PENDIENTE
#   False -> se reporta como hueco de la matriz y el día queda sin clasificar
CLASIFICAR_CITA_PENDIENTE = True

# Entrega completa del proveedor y aun así CEDIS amaneció en cero.
#
# Era el último hueco de la matriz. La explicación la dio Compras: en TODOS
# esos casos el proveedor entregó completo y a tiempo, pero el mismo día que
# entregó al CEDIS éste hizo crossdock — en Vía 2 el CEDIS no resguarda
# inventario, así que amanecer en cero es lo normal, no una falla.
#
# No es culpa del proveedor ni del CEDIS: el pedido se generó tarde para
# cubrir el consumo. Con más antelación no se habría llegado al desabasto.
# Mismo razonamiento y mismo responsable que CLASIFICAR_CITA_PENDIENTE.
#
# SÓLO aplica a Vía 2. En Vía 1 el CEDIS sí debe resguardar, así que entrega
# completa con CEDIS en cero es otra pregunta —qué pasó con lo que recibió— y
# se deja como hueco de la matriz en vez de dictaminarla por analogía.
CLASIFICAR_ENTREGA_COMPLETA_CEDIS_CERO = True

# ¿A quién le cae un pedido de tienda que nunca se generó (RC03, prioridad 3)?
# Depende de quién debía generarlo, y eso lo dice CATALOGO.tipo_resurtido:
#   Automático -> el algoritmo de resurtido debía generarlo solo. Si no lo
#                 hizo, es una falla de sistema/parametrización, no de la
#                 tienda -> Responsable.COMPRAS_ABASTO.
#   Manual     -> alguien en tienda tenía que generarlo a mano y no lo hizo
#                 -> Responsable.TIENDA.
#   Sin dato   -> CATALOGO no trae tipo_resurtido para ese SKU-tienda; no se
#                 puede saber cuál de las dos rutas aplica y se mantiene el
#                 responsable conjunto histórico -> Responsable.TIENDA_ABASTO.
# Depende de EVALUAR_PEDIDO_TIENDA (arriba): mientras SIMA no entrega, la
# prioridad 3 no dictamina y este mapa no se consulta. PENDIENTE DE
# RATIFICAR CON LA COMER, igual que RESPONSABLE_SIN_CITA.
RESPONSABLE_PEDIDO_NO_GENERADO = {
    TipoResurtido.AUTOMATICO: Responsable.COMPRAS_ABASTO,
    TipoResurtido.MANUAL: Responsable.TIENDA,
}

SUBCAUSA_PEDIDO_NO_GENERADO = {
    TipoResurtido.AUTOMATICO: SubcausaPedidoTienda.AUTOMATICO_NO_GENERO,
    TipoResurtido.MANUAL: SubcausaPedidoTienda.MANUAL_NO_GENERO,
}


class ViaResurtido(str, Enum):
    """Bifurcación de la matriz (prioridad 4).

    Confirmado con La Comer (2026-08-05): Vía 1 y Vía 2 son ambas resurtido
    vía CEDIS y comparten la misma rama causal (reglas 5-8) — la diferencia
    es operativa, no de diagnóstico:
      Vía 1 = CEDIS guarda inventario para surtir más rápido a tiendas.
      Vía 2 = CEDIS opera como cross-dock: recibe del proveedor y redistribuye
              a tienda sin resguardar inventario.
    En ambas, las preguntas que importan son las mismas: ¿CEDIS tenía/recibió
    producto disponible?, ¿lo envió a tienda?, si no tenía, ¿el proveedor
    entregó completo?
    """
    VIA_1 = "Vía 1"       # resurtido vía CEDIS con inventario en resguardo
    VIA_2 = "Vía 2"       # resurtido vía CEDIS cross-dock (misma rama que Vía 1)
    DSD = "DSD"           # entrega directa del proveedor a tienda


# ---------------------------------------------------------------------------
# 2. MODELO CANÓNICO DE EVIDENCIA (SKU + TIENDA + FECHA)
# ---------------------------------------------------------------------------

@dataclass
class EvidenciaSKUTienda:
    """Una observación diaria. None significa SIN DATO, nunca cero.

    La distinción es crítica: 'no sé si había inventario' y 'confirmé que
    había cero' llevan a ramas opuestas del árbol.
    """

    sku: str
    tienda: str
    fecha: date

    # --- Métricas de negocio (no discriminan causa, se arrastran al output)
    osa: Optional[float] = None
    venta_perdida: Optional[float] = None

    # --- Prioridad 0: ¿el SKU pertenece al catálogo de la tienda?  Fuente: CATALOGO
    # None = no se pudo comprobar (catálogo vacío); False = está fuera del alcance.
    en_catalogo: Optional[bool] = None
    # ¿SIMA trae algún pedido de este SKU? None = no se evaluó
    # (SIMA vacía o exclusión apagada). Ver EXCLUIR_SKU_SIN_SIMA.
    sku_en_sima: Optional[bool] = None

    # --- Prioridad 1: ¿había producto en tienda?      Fuente: Inventario tienda / BOPS
    inventario_tienda: Optional[int] = None

    # --- Prioridad 1, refinamiento: banderas de alerta de BOPS (layout V8).
    # 1 / 0 / None. Sólo afinan la SUBCAUSA de RC01; no cambian causa ni
    # responsable. None = sin dato (vacío no es cero): RC01 no se refina.
    alerta_enviada: Optional[bool] = None
    alerta_ejecutada: Optional[bool] = None

    # --- Prioridad 2: ¿había tránsito vigente?        Fuente: Tránsitos
    transito_vigente: Optional[bool] = None

    # --- Prioridad 3: ¿la tienda/sistema generó pedido?   Fuente: SIMA
    pedido_tienda_generado: Optional[bool] = None

    # --- Prioridad 3, refinamiento: manual o automático   Fuente: CATALOGO
    tipo_resurtido: Optional[TipoResurtido] = None

    # --- Prioridad 4: bifurcación de la vía           Fuente: CATALOGO SECO VIA1
    via_resurtido: Optional[ViaResurtido] = None

    # --- Prioridades 5-6 (Vía 1): CEDIS               Fuente: Inv. CEDIS + Transferencias
    inventario_cedis: Optional[int] = None
    envio_cedis_generado: Optional[bool] = None

    # --- Prioridades 7-8 (Vía 1): proveedor a CEDIS   Fuente: NS / Recibos / Citas
    pedido_proveedor_generado: Optional[bool] = None
    proveedor_cajas_pedidas: Optional[int] = None
    proveedor_cajas_entregadas: Optional[int] = None

    # --- Prioridad 8, refinamiento              Fuente: CITAS_PROV_CEDIS
    # Todo None = no hay hoja de citas; la regla 8 opera como antes.
    proveedor_cita_agendada: Optional[bool] = None
    proveedor_cita_vencida: Optional[bool] = None      # ¿la cita ya pasó al día D?
    proveedor_cajas_confirmadas_cita: Optional[int] = None
    proveedor_fecha_cita: Optional[date] = None
    proveedor_folio_pedido: Optional[str] = None
    proveedor_folio_cita: Optional[str] = None

    # --- Prioridades 9-10 (DSD): entrega directa      Fuente: Pedido DSD / Recibo tienda
    dsd_entrego_tienda: Optional[bool] = None


# ---------------------------------------------------------------------------
# 3. RESULTADO DE UNA REGLA
# ---------------------------------------------------------------------------

@dataclass
class Dictamen:
    """Una regla de la matriz se cumplió."""
    prioridad: int
    root_cause_id: str
    causa: CausaRaiz
    responsable: Responsable
    fuente: str
    evidencia: List[str] = field(default_factory=list)
    subcausa: Optional[Union[SubcausaProveedor, SubcausaPedidoTienda,
                             SubcausaEjecucionTienda]] = None


@dataclass
class Indeterminado:
    """La regla en turno aplica al caso, pero le falta el dato para decidir."""
    prioridad: int
    campos_faltantes: List[str]
    evidencia: List[str] = field(default_factory=list)


# None = la regla no aplica a este caso, continúa la cadena
Evaluacion = Union[Dictamen, Indeterminado, None]


# ---------------------------------------------------------------------------
# Refinamiento de RC01 con alerta de BOPS — usado por prioridad 1 y prioridad 10
# (ambas dictaminan RC01 / Tienda "Ejecución en Tienda"). No cambia causa ni
# responsable: sólo devuelve la subcausa y anota la evidencia.
# ---------------------------------------------------------------------------

def subcausa_por_alerta(ev: EvidenciaSKUTienda, evidencia: List[str]):
    """Subcausa de RC01 según las banderas de alerta de BOPS.

    Devuelve None si el refinamiento está apagado o si no hay dato de alerta
    (vacío no es cero). Cuando la alerta se envió pero no se sabe si se
    ejecutó, tampoco se refina más allá de 'enviada', para no inventar.
    """
    if not REFINAR_RC01_CON_ALERTA or ev.alerta_enviada is None:
        return None

    if not ev.alerta_enviada:
        evidencia.append("BOPS no envió alerta de este hueco de anaquel")
        return SubcausaEjecucionTienda.SIN_ALERTA

    # Alerta enviada: distinguir por ejecución. Sin dato de ejecución no se
    # refina más (vacío no es cero).
    if ev.alerta_ejecutada is None:
        return None

    if ev.alerta_ejecutada:
        evidencia.append("Alerta de BOPS enviada y atendida; el hueco persistió")
        return SubcausaEjecucionTienda.ALERTA_ATENDIDA_PERSISTE

    evidencia.append("Alerta de BOPS enviada y no atendida en tienda")
    return SubcausaEjecucionTienda.ALERTA_NO_ATENDIDA


# ---------------------------------------------------------------------------
# 4. REGLAS — orden y contenido calcados de la hoja "Matriz Causa raiz"
# ---------------------------------------------------------------------------

class Regla:
    prioridad: int = 0

    def evalua(self, ev: EvidenciaSKUTienda, ctx: List[str]) -> Evaluacion:
        """ctx acumula la evidencia confirmada por las reglas anteriores."""
        raise NotImplementedError


class R0_DentroDelCatalogo(Regla):
    """Prioridad 0 — el SKU tiene que estar en el catálogo de esa tienda.

    No es una regla de la matriz: es el filtro de alcance.

    BOPS entrega TODAS las divisiones de la tienda y el catálogo cubre sólo
    Abarrotes (confirmado con La Comer). Un SKU de otra división no se puede
    clasificar con este árbol —sin catálogo no hay vía de resurtido ni CEDIS
    surtidor que consultar— y sobre todo no debería contarse contra la
    cobertura: no es que el modelo no supiera, es que no le tocaba. En
    Coyoacán son 17,567 SKU y 145,645 días. Ver FUERA_DE_CATALOGO.
    """
    prioridad = 0

    def evalua(self, ev, ctx):
        if ev.en_catalogo is False:
            return Indeterminado(
                self.prioridad, [FUERA_DE_CATALOGO],
                ["El SKU no está en el catálogo de la tienda: queda fuera del "
                 "alcance del análisis, no es un dato faltante"])
        # Mismo trato para los que SIMA no trae: sin un solo pedido no se puede
        # contestar la prioridad 3, y el día no se puede clasificar sin
        # inventarle un culpable. Se separa con su propio marcador porque el
        # motivo es distinto —aquí el dato SÍ debía existir— y el reporte
        # tiene que poder decirlo. Ver EXCLUIR_SKU_SIN_SIMA.
        if EXCLUIR_SKU_SIN_SIMA and ev.sku_en_sima is False:
            return Indeterminado(
                self.prioridad, [SIN_DATO_SIMA],
                ["SIMA no trae ningún pedido de este SKU en el periodo: queda "
                 "fuera del alcance por falta de dato, no por un fallo del "
                 "modelo"])
        return None


class R1_InventarioEnTienda(Regla):
    """Prioridad 1 — Inventario tienda > 0 → Ejecución en tienda.

    Desde el layout V8 se refina con las banderas de alerta de BOPS
    (alerta_enviada / alerta_ejecutada). El refinamiento NO cambia la causa ni
    el responsable —sigue siendo RC01 / Tienda—: sólo distingue la SUBCAUSA,
    para saber si la tienda ignoró la alerta, la atendió sin efecto, o nunca
    fue notificada. Ver subcausa_por_alerta y REFINAR_RC01_CON_ALERTA.
    """
    prioridad = 1

    def evalua(self, ev, ctx):
        if ev.inventario_tienda is None:
            return Indeterminado(self.prioridad, ["inventario_tienda"])

        if ev.inventario_tienda > 0:
            evidencia = [f"Inventario en tienda = {ev.inventario_tienda} (> 0)"]
            subcausa = subcausa_por_alerta(ev, evidencia)
            return Dictamen(
                self.prioridad, "RC01", CausaRaiz.RC01, Responsable.TIENDA,
                "Inventario tienda / BOPS",
                evidencia,
                subcausa=subcausa,
            )

        ctx.append("Inventario en tienda = 0")
        return None


class R2_TransitoVigente(Regla):
    """Prioridad 2 — Inventario = 0 y existe tránsito vigente → Transporte."""
    prioridad = 2

    def evalua(self, ev, ctx):
        if ev.transito_vigente is None:
            return Indeterminado(self.prioridad, ["transito_vigente"], list(ctx))

        if ev.transito_vigente:
            return Dictamen(
                self.prioridad, "RC02", CausaRaiz.RC02, Responsable.LOGISTICA,
                "Tránsitos",
                ctx + ["Tránsito vigente hacia tienda"],
            )

        ctx.append("Sin tránsito vigente")
        return None


class R3_PedidoTiendaNoGenerado(Regla):
    """Prioridad 3 — Inventario = 0, sin tránsito y no existe pedido.

    Apagada temporalmente con EVALUAR_PEDIDO_TIENDA mientras SIMA no entrega.
    Cuando dictamina, CATALOGO.tipo_resurtido afina a quién le cae: ver
    RESPONSABLE_PEDIDO_NO_GENERADO.
    """
    prioridad = 3

    def evalua(self, ev, ctx):
        if not EVALUAR_PEDIDO_TIENDA:
            ctx.append(NOTA_SIN_SIMA)
            return None

        if ev.pedido_tienda_generado is None:
            return Indeterminado(self.prioridad, ["pedido_tienda_generado"], list(ctx))

        if not ev.pedido_tienda_generado:
            responsable = RESPONSABLE_PEDIDO_NO_GENERADO.get(
                ev.tipo_resurtido, Responsable.TIENDA_ABASTO)
            subcausa = SUBCAUSA_PEDIDO_NO_GENERADO.get(ev.tipo_resurtido)

            detalle = "No existe pedido de tienda"
            fuente = "SIMA"
            if ev.tipo_resurtido is not None:
                detalle += f" (resurtido {ev.tipo_resurtido.value})"
                fuente += " / CATALOGO"
            else:
                detalle += " (CATALOGO no trae tipo_resurtido para este SKU-tienda: no se pudo afinar el responsable)"

            return Dictamen(
                self.prioridad, "RC03", CausaRaiz.RC03, responsable,
                fuente,
                ctx + [detalle],
                subcausa=subcausa,
            )

        ctx.append("Pedido de tienda generado")
        return None


class R4_BifurcacionVia(Regla):
    """Prioridad 4 — Existe pedido: bifurcar según vía de resurtido.

    No dictamina causa; enruta hacia la rama CEDIS (Vía 1 y Vía 2, que
    comparten reglas 5-8) o hacia DSD.
    """
    prioridad = 4

    def evalua(self, ev, ctx):
        if ev.via_resurtido is None:
            return Indeterminado(self.prioridad, ["via_resurtido"], list(ctx))

        ctx.append(f"Vía de resurtido: {ev.via_resurtido.value}")
        return None


class R5_R6_RamaCedis(Regla):
    """Prioridades 5 y 6 — Vía 1 o Vía 2 con inventario en CEDIS.

      5: Inv. CEDIS > 0 + Sin envío   → CEDIS No Surtió
      6: Inv. CEDIS > 0 + Existe envío → Transporte / Tránsito
    """
    prioridad = 5

    def evalua(self, ev, ctx):
        if ev.via_resurtido not in (ViaResurtido.VIA_1, ViaResurtido.VIA_2):
            return None

        if ev.inventario_cedis is None:
            return Indeterminado(self.prioridad, ["inventario_cedis"], list(ctx))

        if ev.inventario_cedis <= 0:
            ctx.append("Inventario en CEDIS = 0")
            return None

        base = ctx + [f"Inventario en CEDIS = {ev.inventario_cedis} (> 0)"]

        if ev.envio_cedis_generado is None:
            return Indeterminado(self.prioridad, ["envio_cedis_generado"], base)

        if not ev.envio_cedis_generado:
            return Dictamen(
                5, "RC04", CausaRaiz.RC04, Responsable.CEDIS,
                "Inventario CEDIS + Transferencias",
                base + ["CEDIS no generó envío a tienda"],
            )

        return Dictamen(
            6, "RC02", CausaRaiz.RC02, Responsable.LOGISTICA,
            "Transferencias",
            base + ["Envío generado desde CEDIS, producto no llegó a anaquel"],
        )


class R7_R8_RamaProveedorCedis(Regla):
    """Prioridades 7 y 8 — Vía 1 o Vía 2 con CEDIS en cero.

      7: Sin pedido a proveedor                → Pedido Proveedor No Generado
      8: Con pedido + entrega incompleta       → Incumplimiento Proveedor

    La prioridad 8 se resuelve con la hoja de citas cuando existe. La cita es
    el compromiso del proveedor, y sin ella la regla original acusa de más:
    una orden abierta lleva cero cajas recibidas al día D, así que un pedido
    generado ayer con cita para dentro de una semana se contaba hoy como
    incumplimiento. Con la cita se sabe si el proveedor ya estaba en falta.

    Si la hoja de citas no viene, la evidencia llega en None y la regla opera
    exactamente como antes (pedidas vs entregadas).
    """
    prioridad = 7

    def evalua(self, ev, ctx):
        if ev.via_resurtido not in (ViaResurtido.VIA_1, ViaResurtido.VIA_2):
            return None
        if ev.inventario_cedis is None or ev.inventario_cedis > 0:
            return None

        if ev.pedido_proveedor_generado is None:
            return Indeterminado(self.prioridad, ["pedido_proveedor_generado"], list(ctx))

        if not ev.pedido_proveedor_generado:
            return Dictamen(
                7, "RC05", CausaRaiz.RC05, Responsable.COMPRAS_ABASTO,
                "NS / Pedidos proveedor",
                ctx + ["No existe pedido a proveedor"],
            )

        etiqueta = f"Pedido a proveedor {ev.proveedor_folio_pedido} vigente" \
            if ev.proveedor_folio_pedido else "Pedido a proveedor existente"
        base = ctx + [etiqueta]

        if ev.proveedor_cita_agendada is not None:
            return self._con_citas(ev, base)
        return self._sin_citas(ev, base)

    # -- prioridad 8 con la hoja de citas ----------------------------------

    def _con_citas(self, ev, base) -> Evaluacion:
        pedidas = ev.proveedor_cajas_pedidas

        if not ev.proveedor_cita_agendada:
            return Dictamen(
                8, "RC06", CausaRaiz.RC06, RESPONSABLE_SIN_CITA,
                "Citas proveedor a CEDIS",
                base + ["El pedido no tiene cita agendada en CEDIS"],
                subcausa=SubcausaProveedor.SIN_CITA,
            )

        cita = f" (cita {ev.proveedor_folio_cita})" if ev.proveedor_folio_cita else ""
        fecha = ev.proveedor_fecha_cita.isoformat() if ev.proveedor_fecha_cita else "sin fecha"

        if not ev.proveedor_cita_vencida:
            detalle = (f"Cita agendada para el {fecha}{cita}, posterior al día del "
                       f"faltante: el proveedor todavía está en tiempo")
            if not CLASIFICAR_CITA_PENDIENTE:
                return Indeterminado(
                    8, ["regla_matriz_para_cita_de_proveedor_no_vencida"], base + [detalle])
            return Dictamen(
                8, "RC05", CausaRaiz.RC05, Responsable.COMPRAS_ABASTO,
                "Citas proveedor a CEDIS",
                base + [detalle + ". El pedido se generó tarde para cubrir el consumo"],
                subcausa=SubcausaProveedor.CITA_PENDIENTE,
            )

        confirmadas = ev.proveedor_cajas_confirmadas_cita
        entregadas = ev.proveedor_cajas_entregadas

        faltantes = [c for c, v in
                     (("proveedor_cajas_confirmadas_cita", confirmadas),
                      ("proveedor_cajas_entregadas", entregadas))
                     if v is None]
        if faltantes:
            return Indeterminado(8, faltantes, base)

        base = base + [f"Cita del {fecha}{cita} vencida: confirmó {confirmadas} cajas, "
                       f"entregó {entregadas}"]

        if entregadas < confirmadas:
            sub = (SubcausaProveedor.NO_SE_PRESENTO if entregadas == 0
                   else SubcausaProveedor.ENTREGA_PARCIAL)
            cumplimiento = round(entregadas / confirmadas * 100, 1) if confirmadas else 0.0
            return Dictamen(
                8, "RC06", CausaRaiz.RC06, Responsable.PROVEEDOR,
                "Citas proveedor a CEDIS",
                base + [f"{sub.value} (cumplimiento de cita {cumplimiento}%)"],
                subcausa=sub,
            )

        if pedidas is not None and confirmadas < pedidas:
            return Dictamen(
                8, "RC06", CausaRaiz.RC06, Responsable.PROVEEDOR,
                "Citas proveedor a CEDIS",
                base + [f"Cumplió su cita, pero la agendó por {confirmadas} de las "
                        f"{pedidas} cajas pedidas"],
                subcausa=SubcausaProveedor.RECORTE_EN_CITA,
            )

        # Mismo hueco de la matriz que en la ruta sin citas: el proveedor
        # entregó todo lo pedido y aun así CEDIS amaneció en cero.
        evidencia = base + ["Entrega completa contra la cita, pero inventario CEDIS = 0"]
        dictamen = self._entrega_completa_cedis_cero(ev, evidencia)
        if dictamen is not None:
            return dictamen
        return Indeterminado(
            8, ["regla_matriz_para_entrega_completa_con_cedis_en_cero"], evidencia,
        )


    def _entrega_completa_cedis_cero(self, ev, evidencia):
        """Dictamen del hueco de entrega completa con CEDIS en cero.

        Devuelve None cuando no aplica —está apagado, o el SKU no es Vía 2— y
        entonces el caso sigue saliendo como hueco de la matriz, que es lo
        honesto: no se dictamina por analogía algo que no se verificó.
        """
        if not CLASIFICAR_ENTREGA_COMPLETA_CEDIS_CERO:
            return None
        if ev.via_resurtido is not ViaResurtido.VIA_2:
            return None
        return Dictamen(
            8, "RC05", CausaRaiz.RC05, Responsable.COMPRAS_ABASTO,
            "Citas proveedor a CEDIS",
            evidencia + ["Vía 2 (crossdock): el CEDIS no resguarda inventario, "
                         "así que el cero es normal. El pedido se generó tarde "
                         "para cubrir el consumo"],
            subcausa=SubcausaProveedor.CROSSDOCK_PEDIDO_TARDE,
        )

    # -- prioridad 8 original, sin hoja de citas ---------------------------

    def _sin_citas(self, ev, base) -> Evaluacion:
        if ev.proveedor_cajas_pedidas is None or ev.proveedor_cajas_entregadas is None:
            faltantes = [c for c, v in
                         (("proveedor_cajas_pedidas", ev.proveedor_cajas_pedidas),
                          ("proveedor_cajas_entregadas", ev.proveedor_cajas_entregadas))
                         if v is None]
            return Indeterminado(self.prioridad, faltantes, base)

        pedidas = ev.proveedor_cajas_pedidas
        entregadas = ev.proveedor_cajas_entregadas

        if entregadas < pedidas:
            fill_rate = round(entregadas / pedidas * 100, 1) if pedidas else 0.0
            return Dictamen(
                8, "RC06", CausaRaiz.RC06, Responsable.PROVEEDOR,
                "Recibos / Citas",
                base + [f"Entrega incompleta: {entregadas}/{pedidas} cajas "
                        f"(fill rate {fill_rate}%)"],
            )

        # Brecha de la matriz: CEDIS en cero pese a entrega completa del
        # proveedor. Ninguna de las 10 reglas cubre este caso.
        evidencia = base + [f"Entrega completa: {entregadas}/{pedidas} cajas, "
                            f"pero inventario CEDIS = 0"]
        dictamen = self._entrega_completa_cedis_cero(ev, evidencia)
        if dictamen is not None:
            return dictamen
        return Indeterminado(
            self.prioridad, ["regla_matriz_para_entrega_completa_con_cedis_en_cero"],
            evidencia,
        )


class R9_R10_RamaDSD(Regla):
    """Prioridades 9 y 10 — Entrega directa del proveedor a tienda.

      9:  No entregó a tienda → Incumplimiento Proveedor
      10: Sí entregó a tienda → Ejecución en Tienda

    La prioridad 10 dictamina RC01 igual que la prioridad 1, así que también se
    refina con las banderas de alerta de BOPS (misma subcausa, mismo criterio).
    """
    prioridad = 9

    def evalua(self, ev, ctx):
        if ev.via_resurtido is not ViaResurtido.DSD:
            return None

        if ev.dsd_entrego_tienda is None:
            return Indeterminado(self.prioridad, ["dsd_entrego_tienda"], list(ctx))

        if not ev.dsd_entrego_tienda:
            return Dictamen(
                9, "RC06", CausaRaiz.RC06, Responsable.PROVEEDOR,
                "Pedido DSD",
                ctx + ["Proveedor no entregó en tienda"],
            )

        evidencia = ctx + ["Proveedor entregó en tienda, producto no llegó a anaquel"]
        subcausa = subcausa_por_alerta(ev, evidencia)
        return Dictamen(
            10, "RC01", CausaRaiz.RC01, Responsable.TIENDA,
            "Recibo tienda",
            evidencia,
            subcausa=subcausa,
        )


# ---------------------------------------------------------------------------
# 5. MOTOR
# ---------------------------------------------------------------------------

class MotorRCA:
    """Recorre la matriz en orden de prioridad. La primera regla que se
    cumple dictamina; si a la regla en turno le falta evidencia, el caso
    se cierra como RC99 con el detalle de qué dato hace falta."""

    def __init__(self):
        self.reglas: List[Regla] = [
            R0_DentroDelCatalogo(),
            R1_InventarioEnTienda(),
            R2_TransitoVigente(),
            R3_PedidoTiendaNoGenerado(),
            R4_BifurcacionVia(),
            R5_R6_RamaCedis(),
            R7_R8_RamaProveedorCedis(),
            R9_R10_RamaDSD(),
        ]

    def diagnosticar(self, ev: EvidenciaSKUTienda) -> dict:
        ctx: List[str] = []

        for regla in self.reglas:
            resultado = regla.evalua(ev, ctx)

            if isinstance(resultado, Dictamen):
                return self._salida(ev, resultado)

            if isinstance(resultado, Indeterminado):
                return self._salida_sin_clasificar(ev, resultado)

        # Ninguna rama aplicó (prioridad 99 de la matriz)
        return self._salida_sin_clasificar(
            ev, Indeterminado(99, ["ninguna_condicion_de_la_matriz_aplica"], ctx)
        )

    # -- serialización -----------------------------------------------------

    @staticmethod
    def _salida(ev: EvidenciaSKUTienda, d: Dictamen) -> dict:
        return {
            "sku": ev.sku,
            "tienda": ev.tienda,
            "fecha": ev.fecha.isoformat(),
            "osa": ev.osa,
            "venta_perdida": ev.venta_perdida,
            "clasificado": True,
            "root_cause_id": d.root_cause_id,
            "causa_raiz": d.causa.value,
            "responsable": d.responsable.value,
            "subcausa": d.subcausa.value if d.subcausa else None,
            "prioridad_regla": d.prioridad,
            "fuente": d.fuente,
            "evidencia": d.evidencia,
            "datos_faltantes": [],
        }

    @staticmethod
    def _salida_sin_clasificar(ev: EvidenciaSKUTienda, i: Indeterminado) -> dict:
        return {
            "sku": ev.sku,
            "tienda": ev.tienda,
            "fecha": ev.fecha.isoformat(),
            "osa": ev.osa,
            "venta_perdida": ev.venta_perdida,
            "clasificado": False,
            # Un día fuera de catálogo no es "no supimos": es "no nos tocaba".
            # Lleva código propio para que el reporte diario no los sume al
            # Sin clasificar. Ver RC00_FUERA_DE_ALCANCE.
            "root_cause_id": (RC00_FUERA_DE_ALCANCE
                              if (FUERA_DE_CATALOGO in i.campos_faltantes
                                  or SIN_DATO_SIMA in i.campos_faltantes) else "RC99"),
            # Los dos salen del alcance, pero por motivos distintos y el
            # renglón tiene que decir cuál: uno es "no le tocaba al modelo",
            # el otro es "falta el dato que sí debía llegar".
            "causa_raiz": (CAUSA_FUERA_DE_ALCANCE
                           if FUERA_DE_CATALOGO in i.campos_faltantes
                           else CAUSA_SIN_DATO_SIMA
                           if SIN_DATO_SIMA in i.campos_faltantes
                           else CausaRaiz.RC99.value),
            "responsable": Responsable.PENDIENTE.value,
            "subcausa": None,
            "prioridad_regla": i.prioridad,
            "fuente": None,
            "evidencia": i.evidencia,
            "datos_faltantes": i.campos_faltantes,
        }


# ---------------------------------------------------------------------------
# 6. EJEMPLOS — los tres deep-dives de La Comer Coyoacán (TRAX 3, Abr'26)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import json

    casos: Dict[str, EvidenciaSKUTienda] = {
        # Lámina 11 — Papel Higiénico Maxiresist: fill rate proveedor 0%
        "Papel Higiénico Maxiresist": EvidenciaSKUTienda(
            sku="7501059236776", tienda="287", fecha=date(2026, 4, 15),
            osa=24.1, venta_perdida=3556,
            inventario_tienda=0,
            transito_vigente=False,
            pedido_tienda_generado=True,
            via_resurtido=ViaResurtido.VIA_1,
            inventario_cedis=0,
            pedido_proveedor_generado=True,
            proveedor_cajas_pedidas=40,
            proveedor_cajas_entregadas=0,
        ),

        # Lámina 12 — Café Soluble Jacobs: proveedor surtió 100% a CEDIS,
        # pero no hubo asignación a la tienda 287
        "Café Soluble Jacobs": EvidenciaSKUTienda(
            sku="7501234567890", tienda="287", fecha=date(2026, 4, 10),
            osa=60.0, venta_perdida=800,
            inventario_tienda=0,
            transito_vigente=False,
            pedido_tienda_generado=True,
            via_resurtido=ViaResurtido.VIA_1,
            inventario_cedis=663,
            envio_cedis_generado=False,
        ),

        # Lámina 13 — Suavitel: 168 cajas recibidas el 20, enviadas el 29
        "Suavizante Suavitel Acqua": EvidenciaSKUTienda(
            sku="7501111222333", tienda="287", fecha=date(2026, 4, 24),
            osa=0.0, venta_perdida=1240,
            inventario_tienda=0,
            transito_vigente=False,
            pedido_tienda_generado=True,
            via_resurtido=ViaResurtido.VIA_1,
            inventario_cedis=144,
            envio_cedis_generado=False,
        ),

        # --- Prioridad 8 con la hoja de citas ------------------------------

        # El pedido nunca entró al calendario de CEDIS
        "Proveedor sin cita agendada": EvidenciaSKUTienda(
            sku="7509546080680", tienda="287", fecha=date(2026, 3, 10),
            osa=0.0, venta_perdida=320,
            inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=True,
            via_resurtido=ViaResurtido.VIA_2, inventario_cedis=0,
            pedido_proveedor_generado=True,
            proveedor_cajas_pedidas=28, proveedor_cajas_entregadas=0,
            proveedor_cita_agendada=False,
            proveedor_folio_pedido="26300960730",
        ),

        # Confirmó 25 cajas para el 19, llegó el 19 con las manos vacías
        "Proveedor no se presentó a su cita": EvidenciaSKUTienda(
            sku="7509546080680", tienda="287", fecha=date(2026, 3, 20),
            osa=0.0, venta_perdida=280,
            inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=True,
            via_resurtido=ViaResurtido.VIA_2, inventario_cedis=0,
            pedido_proveedor_generado=True,
            proveedor_cajas_pedidas=25, proveedor_cajas_entregadas=0,
            proveedor_cita_agendada=True, proveedor_cita_vencida=True,
            proveedor_cajas_confirmadas_cita=25,
            proveedor_fecha_cita=date(2026, 3, 19),
            proveedor_folio_pedido="26300964685", proveedor_folio_cita="967844",
        ),

        # Confirmó 67 y entregó 31: cumplió a medias su propio compromiso
        "Proveedor entregó parcial en la cita": EvidenciaSKUTienda(
            sku="7509546080680", tienda="287", fecha=date(2026, 2, 21),
            osa=0.0, venta_perdida=150,
            inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=True,
            via_resurtido=ViaResurtido.VIA_2, inventario_cedis=0,
            pedido_proveedor_generado=True,
            proveedor_cajas_pedidas=67, proveedor_cajas_entregadas=31,
            proveedor_cita_agendada=True, proveedor_cita_vencida=True,
            proveedor_cajas_confirmadas_cita=67,
            proveedor_fecha_cita=date(2026, 2, 20),
            proveedor_folio_pedido="26300915063", proveedor_folio_cita="959829",
        ),

        # El faltante ocurre ANTES de la cita: no es incumplimiento del proveedor
        "Faltante con la cita todavía por vencer": EvidenciaSKUTienda(
            sku="7509546080680", tienda="287", fecha=date(2026, 3, 16),
            osa=0.0, venta_perdida=95,
            inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=True,
            via_resurtido=ViaResurtido.VIA_2, inventario_cedis=0,
            pedido_proveedor_generado=True,
            proveedor_cajas_pedidas=43, proveedor_cajas_entregadas=0,
            proveedor_cita_agendada=True, proveedor_cita_vencida=False,
            proveedor_cajas_confirmadas_cita=43,
            proveedor_fecha_cita=date(2026, 3, 26),
            proveedor_folio_pedido="26300971835", proveedor_folio_cita="969303",
        ),

        # Caso con evidencia incompleta: no se dictamina, se reporta el hueco
        "SKU sin dato de tránsito": EvidenciaSKUTienda(
            sku="7500000000000", tienda="317", fecha=date(2026, 4, 18),
            osa=45.0, venta_perdida=210,
            inventario_tienda=0,
        ),
    }

    motor = MotorRCA()
    for nombre, evidencia in casos.items():
        print(f"\n--- {nombre} ---")
        print(json.dumps(motor.diagnosticar(evidencia), indent=2, ensure_ascii=False))
