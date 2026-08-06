"""
ORCMM - Pipeline: del Excel de captura al resultado de causa raíz.

    python orcmm_pipeline.py "040826_La Comer_Layout de datos RCA (OSA)_V2_Con Datos.xlsx"

Qué hace, en orden:

  1. LEE     las 9 hojas del Excel que llenaron los equipos
  2. DERIVA  las banderas de cada día (¿había tránsito vigente? ¿pedido abierto?
             ¿la cita del proveedor ya venció?) a partir de los eventos con sus fechas
  3. CLASIFICA cada día de faltante con la matriz RC01-RC06 (orcmm_rca_engine)
  4. ESCRIBE un Excel de resultados con la clasificación diaria, el Pareto, el
             scorecard del proveedor y el reporte de cobertura

Principio: si una fuente viene vacía o le falta el día, la bandera queda en
None y el día se marca Sin clasificar nombrando el campo que falta. El
pipeline nunca rellena huecos con ceros.
"""

from __future__ import annotations

import argparse
import sys
import textwrap
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from orcmm_layout_spec import (FILA_DATOS, FILA_ENCABEZADO, HOJAS,
                               normalizar_encabezado)
from orcmm_rca_engine import (EVALUAR_PEDIDO_TIENDA, EvidenciaSKUTienda,
                              TipoResurtido, ViaResurtido)
from orcmm_rca_periodo import (clasificar, cobertura_modelo, diagnosticar_periodo,
                               pareto_periodo, resumen_por_causa,
                               resumen_por_responsable, resumen_por_subcausa)

AMARILLO, GRIS, AZUL, AMBAR, VERDE, ROJO = "FFE600", "F2F2F2", "DDEBF7", "FFEB9C", "C6EFCE", "FFC7CE"
COLOR_CAUSA = {"RC01": AZUL, "RC02": AMBAR, "RC03": ROJO,
               "RC04": AMBAR, "RC05": ROJO, "RC06": ROJO, "RC99": GRIS}

NEGRITA = Font(bold=True)
TITULO = Font(bold=True, size=14)
CHICA = Font(size=9, italic=True, color="595959")
BORDE = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
WRAP = Alignment(wrap_text=True, vertical="top")


# ===========================================================================
# 1. LECTURA
# ===========================================================================

def _fecha(v) -> Optional[date]:
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return date.fromisoformat(str(v).strip()[:10])


def _texto(v) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def _entero(v) -> Optional[int]:
    if v is None or v == "":
        return None
    return int(float(v))


def _decimal(v) -> Optional[float]:
    if v is None or v == "":
        return None
    return float(v)


def _osa_pct(v) -> Optional[float]:
    """BOPS reporta OSA binario por diseño de su sistema: 1 = producto visible
    en anaquel ese día, 0 = no visible. No es un % granular, aunque el layout
    lo declara como Decimal 0-100 (confirmado con el equipo, 2026-08-05). Se
    normaliza aquí a la escala 0-100 para que el umbral y el reporte 'OSA %'
    no necesiten un caso especial."""
    v = _decimal(v)
    if v is None:
        return None
    return v * 100.0


def leer_hoja(wb, nombre: str, advertencias: List[str]) -> List[dict]:
    """Lee una hoja de captura: encabezados en la fila 3, datos desde la 6."""
    if nombre not in wb.sheetnames:
        advertencias.append(f"La hoja {nombre} no existe en el archivo.")
        return []

    ws = wb[nombre]
    # Se traduce el encabezado al nombre canónico del spec: los equipos a veces
    # entregan la columna con el nombre de negocio ("# de Cita" por folio_cita).
    encabezados = [normalizar_encabezado(nombre, c.value) for c in ws[FILA_ENCABEZADO]]

    presentes = [h for h in encabezados if h]
    faltantes = [c[0] for c in HOJAS[nombre]["campos"]
                 if c[2] and c[0] not in presentes]
    if faltantes:
        advertencias.append(f"{nombre}: faltan columnas obligatorias {', '.join(faltantes)}")

    filas = []
    for fila in ws.iter_rows(min_row=FILA_DATOS, values_only=True):
        if all(v is None or v == "" for v in fila):
            continue
        filas.append({h: v for h, v in zip(encabezados, fila) if h})
    return filas


@dataclass
class Fuentes:
    catalogo: Dict[Tuple[str, str], dict] = field(default_factory=dict)
    inv_tienda: Dict[Tuple[str, str, date], dict] = field(default_factory=dict)
    osa: Dict[Tuple[str, str, date], dict] = field(default_factory=dict)
    ventas: Dict[Tuple[str, str, date], dict] = field(default_factory=dict)
    inv_cedis: Dict[Tuple[str, str, date], dict] = field(default_factory=dict)
    transferencias: List[dict] = field(default_factory=list)
    pedidos_tienda: List[dict] = field(default_factory=list)
    pedidos_prov: List[dict] = field(default_factory=list)
    citas_prov: List[dict] = field(default_factory=list)
    # citas indexadas por (folio del pedido, sku): un pedido puede traer varias
    citas_por_pedido: Dict[Tuple[str, str], List[dict]] = field(default_factory=dict)

    conteo: Dict[str, int] = field(default_factory=dict)
    rango: Dict[str, Tuple[Optional[date], Optional[date]]] = field(default_factory=dict)
    advertencias: List[str] = field(default_factory=list)

    def vacia(self, hoja: str) -> bool:
        return self.conteo.get(hoja, 0) == 0


def leer_fuentes(ruta: Path) -> Fuentes:
    wb = openpyxl.load_workbook(ruta, data_only=True)
    fu = Fuentes()
    adv = fu.advertencias

    def registrar(hoja: str, filas: List[dict], campos_fecha: List[str],
                  nota_si_vacia: Optional[str] = None):
        fu.conteo[hoja] = len(filas)
        fechas = [_fecha(f.get(c)) for f in filas for c in campos_fecha]
        fechas = [x for x in fechas if x]
        fu.rango[hoja] = (min(fechas), max(fechas)) if fechas else (None, None)
        if not filas:
            adv.append(f"{hoja}: hoja VACÍA — " + (
                nota_si_vacia or "todas las reglas que dependen de ella "
                                 "quedarán sin clasificar."))

    filas = leer_hoja(wb, "CATALOGO", adv)
    for f in filas:
        fu.catalogo[(_texto(f["sku"]), _texto(f["tienda"]))] = f
    registrar("CATALOGO", filas, [])

    for hoja, destino, llave3 in [
        ("TABLEAU_INV_TIENDA", fu.inv_tienda, "tienda"),
        ("BOPS_OSA", fu.osa, "tienda"),
        ("TABLEAU_VENTAS", fu.ventas, "tienda"),
        ("CEDIS_INVENTARIO", fu.inv_cedis, "cedis"),
    ]:
        filas = leer_hoja(wb, hoja, adv)
        for f in filas:
            d = _fecha(f["fecha"])
            if d:
                destino[(_texto(f["sku"]), _texto(f[llave3]), d)] = f
        registrar(hoja, filas, ["fecha"])

    fu.transferencias = leer_hoja(wb, "CEDIS_TRANSFERENCIAS", adv)
    registrar("CEDIS_TRANSFERENCIAS", fu.transferencias,
              ["fecha_generacion", "fecha_salida_cedis", "fecha_recepcion_tienda"])

    fu.pedidos_tienda = leer_hoja(wb, "SIMA_PEDIDOS_TIENDA", adv)
    registrar("SIMA_PEDIDOS_TIENDA", fu.pedidos_tienda, ["fecha_pedido", "fecha_surtido"],
              "esperado por ahora: SIMA todavía la está procesando. La prioridad 3 está "
              "apagada a propósito para que el árbol pueda seguir hasta CEDIS y proveedor.")

    fu.pedidos_prov = leer_hoja(wb, "COMPRAS_PEDIDOS_PROV", adv)
    registrar("COMPRAS_PEDIDOS_PROV", fu.pedidos_prov, ["fecha_pedido", "fecha_recibo"])

    fu.citas_prov = leer_hoja(wb, "CITAS_PROV_CEDIS", adv)
    registrar("CITAS_PROV_CEDIS", fu.citas_prov, ["fecha_pedido", "fecha_cita"],
              "la prioridad 8 vuelve a operar sólo con cajas pedidas contra cajas "
              "entregadas. No se pierde cobertura, se pierde el detalle: no se puede "
              "separar 'no agendó cita' de 'no cumplió su cita', y un pedido cuya cita "
              "aún no vence se cuenta como incumplimiento del proveedor.")

    for c in fu.citas_prov:
        fu.citas_por_pedido.setdefault(
            (_texto(c.get("folio")), _texto(c.get("sku"))), []).append(c)

    _revisar_cobertura_de_citas(fu)

    aviso = aviso_prioridad_3(fu)
    if aviso:
        fu.advertencias.insert(0, aviso)

    return fu


def aviso_prioridad_3(fu: Fuentes) -> Optional[str]:
    """El texto que tiene que acompañar a cualquier resultado sin prioridad 3.

    Va en la consola y en tres hojas del Excel. Un Pareto en el que RC03 no
    puede aparecer se lee mal si no dice por qué: parece que la tienda nunca
    falla, cuando lo que pasa es que no se le preguntó.
    """
    if EVALUAR_PEDIDO_TIENDA:
        return None

    aviso = ("ANÁLISIS PARCIAL — la prioridad 3 está apagada mientras SIMA entrega los "
             "pedidos de tienda. Ningún día puede salir como RC03 'Pedido No Generado', y "
             "los días que le corresponderían se están repartiendo entre RC04 (CEDIS), "
             "RC05 y RC06 (proveedor). Sirve para leer la rama de abasto; NO para repartir "
             "responsabilidades en firme. Se reactiva con EVALUAR_PEDIDO_TIENDA = True.")

    if not fu.vacia("SIMA_PEDIDOS_TIENDA"):
        n = fu.conteo["SIMA_PEDIDOS_TIENDA"]
        aviso += (f" OJO: SIMA_PEDIDOS_TIENDA ya trae {n} "
                  f"{'registro' if n == 1 else 'registros'} y se están IGNORANDO. "
                  f"Ya se puede prender el interruptor.")
    return aviso


def _revisar_cobertura_de_citas(fu: Fuentes) -> None:
    """La ausencia de cita dictamina contra el proveedor, así que conviene
    saber de antemano qué tan completa viene la hoja.

    Un pedido sin fila en CITAS se lee como 'nunca agendó cita'. Si la hoja
    llegó parcial, eso deja de ser un dictamen y se vuelve un artefacto de la
    extracción: hay que verlo antes de leer el Pareto, no después.
    """
    if fu.vacia("CITAS_PROV_CEDIS") or fu.vacia("COMPRAS_PEDIDOS_PROV"):
        return

    pedidos = {(_texto(o.get("folio")), _texto(o.get("sku"))) for o in fu.pedidos_prov}
    con_cita = set(fu.citas_por_pedido)

    huerfanas = con_cita - pedidos
    if huerfanas:
        fu.advertencias.append(
            f"CITAS_PROV_CEDIS: {len(huerfanas)} citas con folio que no existe en "
            f"COMPRAS_PEDIDOS_PROV (ej: {sorted(f for f, _ in huerfanas)[:3]}). "
            f"No se pueden ligar a un pedido y quedan fuera del análisis.")

    sin_cita = pedidos - con_cita
    if sin_cita:
        pct = round(len(sin_cita) / len(pedidos) * 100, 1)
        fu.advertencias.append(
            f"CITAS_PROV_CEDIS: {len(sin_cita)} de {len(pedidos)} pedidos a proveedor "
            f"({pct}%) no tienen cita. El modelo los dictamina como 'el proveedor nunca "
            f"agendó cita' (RC06). CONFIRMAR con Compras que la hoja trae TODAS las citas "
            f"del periodo antes de usar este Pareto: si la extracción vino parcial, la "
            f"causa es un hueco de captura, no del proveedor.")


# ===========================================================================
# 2. DERIVACIÓN — de eventos con fechas a banderas del día D
# ===========================================================================

VIAS = {v.value.lower(): v for v in ViaResurtido}
TIPOS_RESURTIDO = {t.value.lower(): t for t in TipoResurtido}


def derivar_transito_vigente(fu: Fuentes, sku, tienda, D) -> Optional[bool]:
    """Existe transferencia que ya salió de CEDIS y aún no se recibe en tienda."""
    if fu.vacia("CEDIS_TRANSFERENCIAS"):
        return None
    for t in fu.transferencias:
        if _texto(t["sku"]) != sku or _texto(t["tienda_destino"]) != tienda:
            continue
        salida = _fecha(t.get("fecha_salida_cedis"))
        recepcion = _fecha(t.get("fecha_recepcion_tienda"))
        if salida and salida <= D and (recepcion is None or D < recepcion):
            return True
    return False


def derivar_envio_generado(fu: Fuentes, sku, tienda, D) -> Optional[bool]:
    """Existe transferencia generada y todavía no recibida en tienda."""
    if fu.vacia("CEDIS_TRANSFERENCIAS"):
        return None
    for t in fu.transferencias:
        if _texto(t["sku"]) != sku or _texto(t["tienda_destino"]) != tienda:
            continue
        generacion = _fecha(t.get("fecha_generacion"))
        recepcion = _fecha(t.get("fecha_recepcion_tienda"))
        if generacion and generacion <= D and (recepcion is None or recepcion > D):
            return True
    return False


def derivar_pedido_tienda(fu: Fuentes, sku, tienda, D) -> Optional[bool]:
    """Existe pedido de tienda abierto al día D."""
    if fu.vacia("SIMA_PEDIDOS_TIENDA"):
        return None
    for p in fu.pedidos_tienda:
        if _texto(p["sku"]) != sku or _texto(p["tienda"]) != tienda:
            continue
        pedido = _fecha(p.get("fecha_pedido"))
        surtido = _fecha(p.get("fecha_surtido"))
        if pedido and pedido <= D and (surtido is None or surtido > D):
            return True
    return False


@dataclass
class OrdenProveedor:
    """Estado de la orden a proveedor que explica el faltante del día D."""
    existe: Optional[bool] = None
    cajas_pedidas: Optional[int] = None
    cajas_entregadas: Optional[int] = None
    folio: Optional[str] = None
    # Bloque de citas. Todo en None = no hay hoja CITAS y la regla 8 opera
    # como antes, comparando cajas pedidas contra entregadas.
    cita_agendada: Optional[bool] = None
    cita_vencida: Optional[bool] = None
    cajas_confirmadas_cita: Optional[int] = None
    fecha_cita: Optional[date] = None
    folio_cita: Optional[str] = None


def _compromiso(o: dict) -> date:
    """Fecha en la que se esperaba la mercancía. Ordena las órdenes abiertas."""
    return _fecha(o.get("fecha_cita")) or _fecha(o.get("fecha_recibo")) \
        or _fecha(o.get("fecha_pedido")) or date.max


def derivar_orden_proveedor(fu: Fuentes, sku, cedis, D) -> OrdenProveedor:
    """Orden a proveedor vigente al día D, con el estado de su cita.

    Con 33 órdenes traslapadas para un mismo SKU, cualquier día tiene varias
    vigentes. Se elige la de compromiso más próximo: es la entrega que debía
    haber tapado el faltante de hoy, y por lo tanto la que lo explica. El
    criterio es determinista, no depende del orden de las filas del Excel.
    """
    if fu.vacia("COMPRAS_PEDIDOS_PROV"):
        return OrdenProveedor()

    vigentes = []
    for o in fu.pedidos_prov:
        if _texto(o["sku"]) != sku or _texto(o.get("cedis_destino")) != cedis:
            continue
        pedido = _fecha(o.get("fecha_pedido"))
        recibo = _fecha(o.get("fecha_recibo"))
        if pedido and pedido <= D and (recibo is None or recibo > D):
            vigentes.append(o)

    if not vigentes:
        return OrdenProveedor(existe=False)

    o = min(vigentes, key=lambda x: (_compromiso(x), _texto(x.get("folio")) or ""))
    folio = _texto(o.get("folio"))
    orden = OrdenProveedor(
        existe=True,
        cajas_pedidas=_entero(o.get("cajas_pedidas")),
        cajas_entregadas=0,   # la orden sigue abierta al día D: aún no hay recibo
        folio=folio,
    )

    if fu.vacia("CITAS_PROV_CEDIS"):
        return orden

    return _aplicar_citas(fu, orden, folio, sku, D)


def _aplicar_citas(fu: Fuentes, orden: OrdenProveedor, folio, sku, D) -> OrdenProveedor:
    """Estado de las citas del pedido al día D.

    Se separan las citas ya vencidas de las que aún no llegan. Sólo las
    vencidas juzgan al proveedor: de ellas salen las cajas que confirmó y las
    que realmente entregó. Una cita futura dice lo contrario, que todavía
    está en tiempo.
    """
    citas = fu.citas_por_pedido.get((folio, sku), [])
    if not citas:
        orden.cita_agendada = False
        return orden

    orden.cita_agendada = True
    vencidas = [c for c in citas if (_fecha(c.get("fecha_cita")) or date.max) <= D]

    if vencidas:
        ultima = max(vencidas, key=lambda c: _fecha(c.get("fecha_cita")) or date.min)
        confirmadas = [_entero(c.get("cajas_confirmadas_cita")) for c in vencidas]
        entregadas = [_entero(c.get("cajas_entregadas")) for c in vencidas]
        orden.cita_vencida = True
        orden.cajas_confirmadas_cita = None if None in confirmadas else sum(confirmadas)
        orden.cajas_entregadas = None if None in entregadas else sum(entregadas)
        orden.fecha_cita = _fecha(ultima.get("fecha_cita"))
        orden.folio_cita = _texto(ultima.get("folio_cita"))
        return orden

    proxima = min(citas, key=lambda c: _fecha(c.get("fecha_cita")) or date.max)
    orden.cita_vencida = False
    orden.fecha_cita = _fecha(proxima.get("fecha_cita"))
    orden.folio_cita = _texto(proxima.get("folio_cita"))
    return orden


def derivar_evidencias(fu: Fuentes, umbral_osa: float) -> List[EvidenciaSKUTienda]:
    """Una evidencia por cada día con faltante (OSA por debajo del umbral)."""
    evidencias = []

    for (sku, tienda, D), fila_osa in sorted(fu.osa.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2])):
        osa = _osa_pct(fila_osa.get("osa_pct"))
        if osa is None or osa >= umbral_osa:
            continue   # día sin faltante: no entra al análisis

        cat = fu.catalogo.get((sku, tienda))
        if cat is None:
            fu.advertencias.append(f"CATALOGO: falta el SKU {sku} en la tienda {tienda}.")
            via, cedis, tipo_resurtido = None, None, None
        else:
            via = VIAS.get(str(_texto(cat.get("via_resurtido")) or "").lower())
            cedis = _texto(cat.get("cedis_surtidor"))
            tipo_resurtido = TIPOS_RESURTIDO.get(str(_texto(cat.get("tipo_resurtido")) or "").lower())

        inv_t = fu.inv_tienda.get((sku, tienda, D))
        existencia = None
        if inv_t is not None:
            existencia = _entero(inv_t.get("existencia_minima_dia"))
            if existencia is None:
                existencia = _entero(inv_t.get("existencia_piezas"))

        inv_c = fu.inv_cedis.get((sku, cedis, D)) if cedis else None
        existencia_cedis = None
        if inv_c is not None:
            existencia_cedis = _entero(inv_c.get("existencia_piezas"))
            reservadas = _entero(inv_c.get("piezas_reservadas")) or 0
            existencia_cedis = max(0, existencia_cedis - reservadas)

        venta = fu.ventas.get((sku, tienda, D))
        orden = derivar_orden_proveedor(fu, sku, cedis, D) if cedis else OrdenProveedor()

        evidencias.append(EvidenciaSKUTienda(
            sku=sku, tienda=tienda, fecha=D,
            osa=osa,
            venta_perdida=_decimal(venta.get("venta_perdida_estimada")) if venta else None,
            inventario_tienda=existencia,
            transito_vigente=derivar_transito_vigente(fu, sku, tienda, D),
            pedido_tienda_generado=derivar_pedido_tienda(fu, sku, tienda, D),
            tipo_resurtido=tipo_resurtido,
            via_resurtido=via,
            inventario_cedis=existencia_cedis,
            envio_cedis_generado=derivar_envio_generado(fu, sku, tienda, D),
            pedido_proveedor_generado=orden.existe,
            proveedor_cajas_pedidas=orden.cajas_pedidas,
            proveedor_cajas_entregadas=orden.cajas_entregadas,
            proveedor_cita_agendada=orden.cita_agendada,
            proveedor_cita_vencida=orden.cita_vencida,
            proveedor_cajas_confirmadas_cita=orden.cajas_confirmadas_cita,
            proveedor_fecha_cita=orden.fecha_cita,
            proveedor_folio_pedido=orden.folio,
            proveedor_folio_cita=orden.folio_cita,
        ))

    return evidencias


# ===========================================================================
# 3. SALIDA
# ===========================================================================

def si_no(v: Optional[bool]) -> str:
    return "" if v is None else ("Sí" if v else "No")


def _celda_cita(ev: EvidenciaSKUTienda) -> str:
    """Resume en una celda si el pedido tenía cita y si ya había vencido."""
    if ev.proveedor_cita_agendada is None:
        return ""
    if not ev.proveedor_cita_agendada:
        return "Sin cita"
    fecha = ev.proveedor_fecha_cita.isoformat() if ev.proveedor_fecha_cita else "s/f"
    return f"{fecha} {'vencida' if ev.proveedor_cita_vencida else 'por vencer'}"


def _banner(ws, fila: int, col_ini: int, col_fin: int, texto: str, alto: int = 46):
    """Aviso a todo lo ancho, para que no se pueda leer la tabla sin verlo."""
    c = ws.cell(row=fila, column=col_ini, value=texto)
    c.font = Font(bold=True, size=9, color="9C0006")
    c.fill = PatternFill("solid", fgColor=ROJO)
    c.alignment = WRAP
    c.border = BORDE
    ws.merge_cells(start_row=fila, start_column=col_ini, end_row=fila, end_column=col_fin)
    ws.row_dimensions[fila].height = alto


def _encabezado(ws, fila, valores):
    for j, v in enumerate(valores, 1):
        c = ws.cell(row=fila, column=j, value=v)
        c.font = NEGRITA
        c.fill = PatternFill("solid", fgColor=AMARILLO)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BORDE


@dataclass
class DesempenoProveedor:
    """Cumplimiento de un proveedor contra CEDIS en el periodo.

    Las tres tasas de la cita se miden sobre el mismo universo, los pedidos que
    SÍ tienen cita, para que se puedan leer en cadena:
      confirmado   = de lo que se le pidió, ¿a cuánto se comprometió al agendar?
      cumplimiento = de lo que confirmó, ¿cuánto entregó?
      efectivo     = de lo que se le pidió, ¿cuánto llegó?  (el que duele)

    Los pedidos sin cita se reportan aparte y no se meten al denominador: hoy
    no se sabe si el proveedor nunca agendó o si la extracción vino corta, y
    mezclarlos hundiría la tasa por una razón que todavía no está confirmada.
    El surtido del pedido (columna aparte) sí cubre los 33 pedidos, tomado de
    COMPRAS_PEDIDOS_PROV.
    """
    proveedor_id: str
    nombre: str
    pedidos: int = 0
    cajas_pedidas: int = 0
    cajas_surtidas_pedido: int = 0     # cierre de la orden, según COMPRAS
    pedidos_con_cita: int = 0
    cajas_pedidas_con_cita: int = 0
    citas: int = 0
    citas_incumplidas: int = 0
    cajas_confirmadas: int = 0
    cajas_entregadas: int = 0

    @property
    def pedidos_sin_cita(self) -> int:
        return self.pedidos - self.pedidos_con_cita

    @staticmethod
    def _tasa(parte: int, todo: int) -> Optional[float]:
        return round(parte / todo, 4) if todo else None

    @property
    def pct_surtido_pedido(self) -> Optional[float]:
        return self._tasa(self.cajas_surtidas_pedido, self.cajas_pedidas)

    @property
    def pct_confirmado(self) -> Optional[float]:
        return self._tasa(self.cajas_confirmadas, self.cajas_pedidas_con_cita)

    @property
    def pct_cumplimiento(self) -> Optional[float]:
        return self._tasa(self.cajas_entregadas, self.cajas_confirmadas)

    @property
    def pct_efectivo(self) -> Optional[float]:
        return self._tasa(self.cajas_entregadas, self.cajas_pedidas_con_cita)


def desempeno_proveedores(fu: Fuentes) -> List[DesempenoProveedor]:
    """Scorecard por proveedor, construido sobre las hojas de origen.

    No depende de la clasificación: mide al proveedor sobre TODOS sus pedidos
    del periodo, no sólo sobre los días en que hubo faltante en tienda. Un
    proveedor puede incumplir sin que se note en anaquel, y eso también hay
    que verlo antes de que sí se note.
    """
    prov: Dict[str, DesempenoProveedor] = {}

    for o in fu.pedidos_prov:
        pid = _texto(o.get("proveedor_id")) or "(sin id)"
        d = prov.setdefault(pid, DesempenoProveedor(
            pid, _texto(o.get("proveedor_nombre")) or ""))
        pedidas = _entero(o.get("cajas_pedidas")) or 0
        d.pedidos += 1
        d.cajas_pedidas += pedidas
        d.cajas_surtidas_pedido += _entero(o.get("cajas_entregadas")) or 0

        citas = fu.citas_por_pedido.get((_texto(o.get("folio")), _texto(o.get("sku"))), [])
        if not citas:
            continue
        d.pedidos_con_cita += 1
        d.cajas_pedidas_con_cita += pedidas
        for c in citas:
            confirmadas = _entero(c.get("cajas_confirmadas_cita")) or 0
            entregadas = _entero(c.get("cajas_entregadas")) or 0
            d.citas += 1
            d.cajas_confirmadas += confirmadas
            d.cajas_entregadas += entregadas
            if entregadas < confirmadas:
                d.citas_incumplidas += 1

    return sorted(prov.values(), key=lambda d: d.cajas_pedidas, reverse=True)


def citas_incumplidas(fu: Fuentes) -> List[dict]:
    """Citas vencidas en las que el proveedor entregó menos de lo confirmado."""
    fallas = []
    for c in fu.citas_prov:
        confirmadas = _entero(c.get("cajas_confirmadas_cita"))
        entregadas = _entero(c.get("cajas_entregadas"))
        if confirmadas is None or entregadas is None or entregadas >= confirmadas:
            continue
        fallas.append({
            "folio": _texto(c.get("folio")),
            "folio_cita": _texto(c.get("folio_cita")),
            "sku": _texto(c.get("sku")),
            "proveedor": _texto(c.get("proveedor_nombre")) or _texto(c.get("proveedor_id")),
            "fecha_cita": _fecha(c.get("fecha_cita")),
            "confirmadas": confirmadas,
            "entregadas": entregadas,
            "faltantes": confirmadas - entregadas,
            "estatus": _texto(c.get("estatus_cita")) or "",
        })
    fallas.sort(key=lambda f: f["faltantes"], reverse=True)
    return fallas


def discrepancias_pedido_cita(fu: Fuentes) -> List[dict]:
    """Dónde el pedido y la cita cuentan historias distintas del mismo folio.

    Las dos hojas traen fecha_cita y cajas entregadas. Cuando no coinciden hay
    que decidir cuál manda antes de presentar cualquier número al proveedor:
    aquí manda la cita para lo que pasó en la cita, y el pedido para el cierre
    de la orden. Esta lista es para que Compras concilie el origen del dato.
    """
    salida = []
    for o in fu.pedidos_prov:
        folio, sku = _texto(o.get("folio")), _texto(o.get("sku"))
        citas = fu.citas_por_pedido.get((folio, sku), [])
        if not citas:
            continue

        ent_pedido = _entero(o.get("cajas_entregadas"))
        ent_cita = sum(_entero(c.get("cajas_entregadas")) or 0 for c in citas)
        cita_pedido = _fecha(o.get("fecha_cita"))
        fechas_cita = sorted(f for f in (_fecha(c.get("fecha_cita")) for c in citas) if f)

        motivos = []
        if ent_pedido is not None and ent_pedido != ent_cita:
            motivos.append(f"cajas entregadas: pedido dice {ent_pedido}, citas suman {ent_cita}")
        if cita_pedido and fechas_cita and cita_pedido not in fechas_cita:
            motivos.append(f"fecha de cita: pedido dice {cita_pedido.isoformat()}, "
                           f"cita dice {fechas_cita[0].isoformat()}")
        if motivos:
            salida.append({"folio": folio, "sku": sku, "motivos": " · ".join(motivos)})
    return salida


def escribir_hoja_proveedor(wb, fu: Fuentes, diagnosticos: List[dict]) -> None:
    ws = wb.create_sheet("Proveedor y citas")
    ws.sheet_view.showGridLines = False
    for j, w in enumerate([3, 30, 9, 11, 13, 8, 10, 13, 11, 12, 13, 13, 12, 13], 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    ws["B2"] = "Cumplimiento del proveedor en CEDIS"
    ws["B2"].font = TITULO
    ws["B3"] = ("El pedido dice qué se le pidió; la cita dice a qué se comprometió y qué entregó. "
                "Las tres tasas de cita se miden sobre los pedidos que tienen cita: confirmado = "
                "cajas de cita sobre cajas pedidas, cumplimiento = entregado sobre confirmado, "
                "efectivo = entregado sobre pedido. Los pedidos sin cita van aparte.")
    ws["B3"].font = CHICA

    if fu.vacia("CITAS_PROV_CEDIS"):
        ws["B5"] = ("CITAS_PROV_CEDIS viene vacía. Sin ella no se puede separar 'no agendó cita' "
                    "de 'no cumplió su cita', y la prioridad 8 vuelve a operar sólo con "
                    "cajas pedidas contra cajas entregadas.")
        ws["B5"].fill = PatternFill("solid", fgColor=AMBAR)
        ws["B5"].alignment = WRAP
        ws.merge_cells(start_row=5, start_column=2, end_row=5, end_column=14)
        return

    f = 5
    _encabezado(ws, f, ["", "Proveedor", "Pedidos", "Cajas ped.", "% surtido pedido",
                        "Citas", "Sin cita", "Cajas ped. c/cita", "Cajas conf.",
                        "Cajas entreg.", "% confirmado", "% cumplim.", "% efectivo",
                        "Citas falladas"])
    ws.row_dimensions[f].height = 44
    f += 1

    for d in desempeno_proveedores(fu):
        etiqueta = f"{d.nombre} ({d.proveedor_id})" if d.nombre else d.proveedor_id
        valores = [etiqueta, d.pedidos, d.cajas_pedidas, d.pct_surtido_pedido,
                   d.citas, d.pedidos_sin_cita, d.cajas_pedidas_con_cita,
                   d.cajas_confirmadas, d.cajas_entregadas,
                   d.pct_confirmado, d.pct_cumplimiento, d.pct_efectivo,
                   d.citas_incumplidas]
        for j, v in enumerate(valores, 2):
            c = ws.cell(row=f, column=j, value=v)
            c.border = BORDE
        for col in (5, 11, 12, 13):
            ws.cell(row=f, column=col).number_format = "0.0%"
        for col, tasa in ((5, d.pct_surtido_pedido), (13, d.pct_efectivo)):
            if tasa is not None:
                ws.cell(row=f, column=col).fill = PatternFill(
                    "solid", fgColor=VERDE if tasa >= 0.95 else
                    (AMBAR if tasa >= 0.8 else ROJO))
        if d.pedidos_sin_cita:
            ws.cell(row=f, column=7).fill = PatternFill("solid", fgColor=AMBAR)
        f += 1

    # --- Impacto de cada subcausa (pedido de tienda y proveedor) -----------
    subcausas = resumen_por_subcausa(diagnosticos)

    f += 2
    ws.cell(row=f, column=2, value="Qué detalle de la causa costó más").font = TITULO
    f += 1
    ws.cell(row=f, column=2, value=(
        "Sólo los días con faltante que la matriz alcanzó a clasificar en la prioridad 3 "
        "(pedido de tienda, según CATALOGO.tipo_resurtido) u 8 (proveedor). Si está vacío, "
        "el árbol se detuvo antes de llegar a alguna de las dos.")).font = CHICA
    f += 2

    _encabezado(ws, f, ["", "Detalle / subcausa", "Días", "Venta perdida"] + [""] * 10)
    f += 1
    if subcausas:
        for fila in subcausas:
            ws.cell(row=f, column=2, value=fila["subcausa"]).alignment = WRAP
            ws.cell(row=f, column=3, value=fila["dias"])
            ws.cell(row=f, column=4, value=fila["venta_perdida"]).number_format = '$#,##0.00'
            for j in range(2, 5):
                ws.cell(row=f, column=j).border = BORDE
            f += 1
    else:
        c = ws.cell(row=f, column=2, value="Ningún día llegó a la prioridad 3 u 8 con dictamen de subcausa.")
        c.fill = PatternFill("solid", fgColor=GRIS)
        f += 1

    # --- Citas falladas, una por una --------------------------------------
    fallas = citas_incumplidas(fu)
    f += 2
    ws.cell(row=f, column=2, value="Citas en las que el proveedor entregó de menos").font = TITULO
    f += 2
    _encabezado(ws, f, ["", "Proveedor", "Pedido", "Cita", "Fecha cita", "SKU",
                        "Confirmó", "Entregó", "Faltaron", "Estatus del sistema"] + [""] * 4)
    f += 1
    for x in fallas:
        valores = [x["proveedor"], x["folio"], x["folio_cita"],
                   x["fecha_cita"].isoformat() if x["fecha_cita"] else "",
                   x["sku"], x["confirmadas"], x["entregadas"], x["faltantes"], x["estatus"]]
        for j, v in enumerate(valores, 2):
            c = ws.cell(row=f, column=j, value=v)
            c.border = BORDE
        ws.cell(row=f, column=10).fill = PatternFill("solid", fgColor=ROJO)
        f += 1
    if not fallas:
        ws.cell(row=f, column=2, value="Ninguna: el proveedor entregó todo lo que confirmó.")\
            .fill = PatternFill("solid", fgColor=VERDE)
        f += 1

    # --- Conciliación entre las dos hojas ---------------------------------
    disc = discrepancias_pedido_cita(fu)
    if disc:
        f += 2
        ws.cell(row=f, column=2, value="Pedido y cita no cuadran").font = TITULO
        f += 1
        ws.cell(row=f, column=2, value=(
            "El mismo folio trae dos versiones del mismo dato. Conciliar con Compras antes "
            "de presentarle números al proveedor.")).font = CHICA
        f += 2
        _encabezado(ws, f, ["", "Pedido", "SKU", "Qué no cuadra"] + [""] * 10)
        f += 1
        for d in disc:
            ws.cell(row=f, column=2, value=d["folio"]).border = BORDE
            ws.cell(row=f, column=3, value=d["sku"]).border = BORDE
            c = ws.cell(row=f, column=4, value=d["motivos"])
            c.alignment = WRAP
            c.fill = PatternFill("solid", fgColor=AMBAR)
            ws.merge_cells(start_row=f, start_column=4, end_row=f, end_column=14)
            f += 1


def escribir_resultado(ruta: Path, fu: Fuentes, evidencias: List[EvidenciaSKUTienda],
                       diagnosticos: List[dict]):
    wb = openpyxl.Workbook()
    aviso = aviso_prioridad_3(fu)

    # --- Clasificación diaria --------------------------------------------
    ws = wb.active
    ws.title = "Clasificacion diaria"
    ws.sheet_view.showGridLines = False
    for j, w in enumerate([16, 9, 12, 8, 13, 11, 10, 11, 12, 8, 11, 11, 10, 10, 12, 11, 10,
                           8, 26, 20, 30, 10, 26, 60], 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    ws["A1"] = "Clasificación de desabasto por causa raíz — resultado por día"
    ws["A1"].font = TITULO
    ws["A2"] = ("Las columnas de evidencia se derivaron de las hojas de eventos; la causa la dictaminó "
                "la matriz RC01-RC06. Un día Sin clasificar nombra el dato que faltó.")
    ws["A2"].font = CHICA
    if aviso:
        _banner(ws, 3, 1, 24, aviso, alto=32)

    _encabezado(ws, 4, [
        "sku", "tienda", "fecha", "OSA %", "venta perdida",
        "inv tienda", "tránsito", "pedido tda", "tipo resurtido", "vía", "inv CEDIS",
        "envío CEDIS", "ped prov", "cajas ped", "cita prov", "cajas conf", "cajas ent",
        "RC", "causa raíz", "responsable", "detalle / subcausa",
        "prioridad", "fuente", "evidencia / dato faltante",
    ])
    ws.row_dimensions[4].height = 32

    for i, (ev, dg) in enumerate(zip(evidencias, diagnosticos)):
        r = 5 + i
        detalle = dg["evidencia"][-1] if dg["evidencia"] else ""
        if dg["datos_faltantes"]:
            detalle = "FALTA EL DATO: " + ", ".join(dg["datos_faltantes"])

        valores = [ev.sku, ev.tienda, ev.fecha.isoformat(), ev.osa, ev.venta_perdida,
                   ev.inventario_tienda, si_no(ev.transito_vigente),
                   si_no(ev.pedido_tienda_generado),
                   ev.tipo_resurtido.value if ev.tipo_resurtido else "",
                   ev.via_resurtido.value if ev.via_resurtido else "",
                   ev.inventario_cedis, si_no(ev.envio_cedis_generado),
                   si_no(ev.pedido_proveedor_generado),
                   ev.proveedor_cajas_pedidas, _celda_cita(ev),
                   ev.proveedor_cajas_confirmadas_cita, ev.proveedor_cajas_entregadas,
                   dg["root_cause_id"], dg["causa_raiz"], dg["responsable"],
                   dg.get("subcausa") or "",
                   dg["prioridad_regla"], dg["fuente"], detalle]

        for j, v in enumerate(valores, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.border = BORDE
        rc = ws.cell(row=r, column=18)
        rc.font = NEGRITA
        rc.fill = PatternFill("solid", fgColor=COLOR_CAUSA.get(dg["root_cause_id"], GRIS))
        rc.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "D5"
    ws.auto_filter.ref = f"A4:X{4 + len(evidencias)}"

    # --- Pareto -----------------------------------------------------------
    ws = wb.create_sheet("Pareto")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", [4, 40, 12, 18, 14, 26]):
        ws.column_dimensions[col].width = w

    diags = diagnosticar_periodo(diagnosticos)
    par = pareto_periodo(diags)

    ws["B2"] = "Pareto del periodo"
    ws["B2"].font = TITULO
    ws["B3"] = (f"{par['sku_tienda_analizados']} combinaciones SKU-tienda · "
                f"{par['dias_con_faltante']} días con faltante · "
                f"venta perdida ${par['venta_perdida_total']:,.2f}")
    ws["B3"].font = CHICA

    f = 5
    if aviso:
        _banner(ws, f, 2, 6, aviso, alto=60)
        f += 2
    _encabezado(ws, f, ["", "Causa raíz", "Días", "Venta perdida", "% del total", "Responsable"])
    f += 1
    for fila in resumen_por_causa(diagnosticos):
        ws.cell(row=f, column=2, value=f"{fila['root_cause_id']} · {fila['causa']}")
        ws.cell(row=f, column=2).fill = PatternFill(
            "solid", fgColor=COLOR_CAUSA.get(fila["root_cause_id"], GRIS))
        ws.cell(row=f, column=3, value=fila["dias"])
        ws.cell(row=f, column=4, value=fila["venta_perdida"])
        ws.cell(row=f, column=5, value=fila["pct"] / 100).number_format = "0.0%"
        ws.cell(row=f, column=6, value=fila["responsable"])
        for j in range(2, 7):
            ws.cell(row=f, column=j).border = BORDE
        f += 1

    f += 2
    _encabezado(ws, f, ["", "Responsable", "Días", "Venta perdida", "% del total", ""])
    f += 1
    for fila in resumen_por_responsable(diagnosticos):
        ws.cell(row=f, column=2, value=fila["responsable"]).font = NEGRITA
        ws.cell(row=f, column=3, value=fila["dias"])
        ws.cell(row=f, column=4, value=fila["venta_perdida"])
        ws.cell(row=f, column=5, value=fila["pct"] / 100).number_format = "0.0%"
        for j in range(2, 7):
            ws.cell(row=f, column=j).border = BORDE
        f += 1

    # --- Por SKU-Tienda ---------------------------------------------------
    # Es la lista accionable: qué combinación está costando más y de quién es.
    ws = wb.create_sheet("Por SKU-Tienda")
    ws.sheet_view.showGridLines = False
    for j, w in enumerate([16, 9, 12, 12, 11, 15, 11, 8, 26, 20, 60], 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    ws["A1"] = "Desglose por SKU y tienda — ordenado por impacto"
    ws["A1"].font = TITULO
    ws["A2"] = ("La causa dominante es la que más venta perdida acumuló en el periodo. "
                "El desglose muestra cómo se repartió el impacto entre causas.")
    ws["A2"].font = CHICA

    _encabezado(ws, 4, ["sku", "tienda", "días faltante", "días clasif.", "cobertura",
                        "venta perdida", "OSA prom.", "RC", "causa dominante",
                        "responsable", "desglose de causas"])
    ws.row_dimensions[4].height = 32

    for i, dd in enumerate(diags):
        r = 5 + i
        desglose = " · ".join(
            f"{c}: ${v:,.0f}" for c, v in
            sorted(dd.desglose_causas.items(), key=lambda kv: kv[1], reverse=True))

        valores = [dd.sku, dd.tienda, dd.dias_con_faltante, dd.dias_clasificados,
                   dd.cobertura_pct / 100, dd.venta_perdida_total, dd.osa_promedio,
                   dd.root_cause_id_dominante, dd.causa_dominante,
                   dd.responsable_dominante, desglose]
        for j, v in enumerate(valores, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.border = BORDE
        ws.cell(row=r, column=5).number_format = "0.0%"
        ws.cell(row=r, column=6).number_format = '$#,##0.00'
        rc = ws.cell(row=r, column=8)
        rc.font = NEGRITA
        rc.fill = PatternFill("solid", fgColor=COLOR_CAUSA.get(dd.root_cause_id_dominante, GRIS))
        rc.alignment = Alignment(horizontal="center")
        if dd.cobertura_pct < 100:
            ws.cell(row=r, column=5).fill = PatternFill("solid", fgColor=AMBAR)

    ws.freeze_panes = "C5"
    ws.auto_filter.ref = f"A4:K{4 + len(diags)}"

    # --- Proveedor y citas ------------------------------------------------
    escribir_hoja_proveedor(wb, fu, diagnosticos)

    # --- Cobertura y fuentes ---------------------------------------------
    ws = wb.create_sheet("Cobertura y fuentes")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", [4, 34, 14, 16, 16, 54]):
        ws.column_dimensions[col].width = w

    cob = cobertura_modelo(diagnosticos)

    ws["B2"] = "Cobertura del modelo"
    ws["B2"].font = TITULO
    f = 4
    if aviso:
        _banner(ws, f, 2, 6, aviso, alto=60)
        f += 2
    for k, v in [("Días con faltante analizados", cob["casos_totales"]),
                 ("Días clasificados", cob["casos_clasificados"]),
                 ("Cobertura sobre días", f"{cob['cobertura_casos_pct']}%"),
                 ("Venta perdida total", f"${cob['venta_perdida_total']:,.2f}"),
                 ("Venta perdida clasificada", f"${cob['venta_perdida_clasificada']:,.2f}"),
                 ("Cobertura sobre impacto", f"{cob['cobertura_venta_perdida_pct']}%")]:
        ws.cell(row=f, column=2, value=k).font = NEGRITA
        ws.cell(row=f, column=2).fill = PatternFill("solid", fgColor=GRIS)
        ws.cell(row=f, column=4, value=v)
        for j in range(2, 5):
            ws.cell(row=f, column=j).border = BORDE
        f += 1

    if cob["campos_que_bloquean"]:
        f += 2
        ws.cell(row=f, column=2, value="Qué dato bloqueó la clasificación").font = TITULO
        f += 2
        _encabezado(ws, f, ["", "Campo faltante", "Días", "Venta perdida sin explicar",
                            "", "A quién pedírselo"])
        f += 1
        vp_por_campo = cob["venta_perdida_por_campo_faltante"]
        de_quien = {
            "inventario_tienda": "Tableau — inventario en tienda",
            "transito_vigente": "CEDIS / Logística — transferencias",
            "envio_cedis_generado": "CEDIS / Logística — transferencias",
            "pedido_tienda_generado": "SIMA — pedidos de tienda",
            "via_resurtido": "Catálogo — vía de resurtido del SKU",
            "inventario_cedis": "CEDIS — inventario diario",
            "pedido_proveedor_generado": "Compras — pedidos a proveedor",
            "proveedor_cajas_pedidas": "Compras — pedidos a proveedor",
            "proveedor_cajas_entregadas": "Compras — citas de proveedor",
            "proveedor_cajas_confirmadas_cita": "Compras — citas de proveedor",
            "regla_matriz_para_cita_de_proveedor_no_vencida":
                "Decisión de negocio — la matriz no cubre el faltante previo a la cita",
            "regla_matriz_para_entrega_completa_con_cedis_en_cero":
                "Decisión de negocio — la matriz no cubre CEDIS en cero con entrega completa",
        }
        for campo, n in cob["campos_que_bloquean"].items():
            ws.cell(row=f, column=2, value=campo).font = Font(name="Consolas", size=10)
            ws.cell(row=f, column=3, value=n)
            ws.cell(row=f, column=4, value=vp_por_campo.get(campo, 0)).number_format = '$#,##0.00'
            ws.cell(row=f, column=6, value=de_quien.get(campo, "Revisar matriz")).alignment = WRAP
            for j in range(2, 7):
                ws.cell(row=f, column=j).border = BORDE
                ws.cell(row=f, column=j).fill = PatternFill("solid", fgColor=AMBAR)
            f += 1

    f += 2
    ws.cell(row=f, column=2, value="Fuentes leídas").font = TITULO
    f += 2
    _encabezado(ws, f, ["", "Hoja", "Filas", "Primera fecha", "Última fecha", "Equipo"])
    f += 1
    for hoja in HOJAS:
        n = fu.conteo.get(hoja, 0)
        d0, d1 = fu.rango.get(hoja, (None, None))
        ws.cell(row=f, column=2, value=hoja).font = Font(name="Consolas", size=10)
        ws.cell(row=f, column=3, value=n)
        ws.cell(row=f, column=4, value=d0.isoformat() if d0 else "")
        ws.cell(row=f, column=5, value=d1.isoformat() if d1 else "")
        ws.cell(row=f, column=6, value=HOJAS[hoja]["equipo"])
        for j in range(2, 7):
            ws.cell(row=f, column=j).border = BORDE
            if n == 0:
                ws.cell(row=f, column=j).fill = PatternFill("solid", fgColor=ROJO)
        f += 1

    if fu.advertencias:
        f += 2
        ws.cell(row=f, column=2, value="Advertencias de lectura").font = TITULO
        f += 2
        for a in fu.advertencias:
            c = ws.cell(row=f, column=2, value=a)
            c.alignment = WRAP
            c.fill = PatternFill("solid", fgColor=AMBAR)
            ws.merge_cells(start_row=f, start_column=2, end_row=f, end_column=6)
            f += 1

    wb.save(ruta)
    return cob, par


# ===========================================================================
# 4. CLI
# ===========================================================================

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Lee el Excel de captura y clasifica el desabasto por causa raíz.")
    ap.add_argument("archivo",
                    help="Excel de captura llenado por los equipos.")
    ap.add_argument("-o", "--salida", default=None,
                    help="Excel de resultados. Por omisión, 'Resultado RCA - <archivo>'.")
    ap.add_argument("--umbral-osa", type=float, default=100.0,
                    help="Se analizan los días con OSA por debajo de este valor (default 100).")
    args = ap.parse_args()

    entrada = Path(args.archivo)
    if not entrada.exists():
        print(f"No se encontró el archivo: {entrada}", file=sys.stderr)
        return 1

    salida = Path(args.salida) if args.salida else \
        entrada.with_name(f"Resultado RCA - {entrada.stem}.xlsx")

    print(f"\nLeyendo  {entrada.name}")
    fu = leer_fuentes(entrada)
    for hoja in HOJAS:
        n = fu.conteo.get(hoja, 0)
        d0, d1 = fu.rango.get(hoja, (None, None))
        rango = f"{d0} a {d1}" if d0 else "sin fechas"
        marca = "  " if n else "!!"
        print(f"  {marca} {hoja:<24} {n:>4} filas   {rango}")

    if fu.advertencias:
        print("\nAdvertencias:")
        for a in fu.advertencias:
            print(f"   - {a}")

    evidencias = derivar_evidencias(fu, args.umbral_osa)
    if not evidencias:
        print("\nNo hay días con faltante que analizar. Revisa la hoja BOPS_OSA.")
        return 1

    diagnosticos = clasificar(evidencias)   # una sola pasada del motor

    cob, par = escribir_resultado(salida, fu, evidencias, diagnosticos)

    print(f"\nAnalizados {cob['casos_totales']} días con faltante · "
          f"clasificados {cob['casos_clasificados']} ({cob['cobertura_casos_pct']}%)")
    print(f"Venta perdida ${cob['venta_perdida_total']:,.2f} · "
          f"clasificada ${cob['venta_perdida_clasificada']:,.2f} "
          f"({cob['cobertura_venta_perdida_pct']}%)")

    print("\nPareto por causa raíz")
    for causa, pct in par["pareto_por_causa"].items():
        print(f"  {pct:>5.1f}%   {causa}")

    print("\nPareto por responsable")
    for resp, pct in par["pareto_por_responsable"].items():
        print(f"  {pct:>5.1f}%   {resp}")

    if cob["campos_que_bloquean"]:
        print("\nDatos que bloquearon la clasificación")
        vp_campo = cob["venta_perdida_por_campo_faltante"]
        for campo, n in cob["campos_que_bloquean"].items():
            print(f"  {n:>3} días   ${vp_campo.get(campo, 0):>12,.2f}   falta {campo}")

    prov = desempeno_proveedores(fu)
    if prov and not fu.vacia("CITAS_PROV_CEDIS"):
        print("\nCumplimiento del proveedor en CEDIS")
        for d in prov:
            def p(v):
                return f"{v*100:5.1f}%" if v is not None else "   n/d"
            print(f"  {(d.nombre or d.proveedor_id)[:34]:<34} "
                  f"confirmado {p(d.pct_confirmado)}  "
                  f"cumplimiento {p(d.pct_cumplimiento)}  "
                  f"efectivo {p(d.pct_efectivo)}  "
                  f"({d.pedidos_sin_cita} pedidos sin cita, "
                  f"{d.citas_incumplidas} citas falladas)")

    print("\nTop SKU-Tienda por impacto")
    diags_st = diagnosticar_periodo(diagnosticos)
    for dd in diags_st[:10]:
        print(f"  ${dd.venta_perdida_total:>12,.2f}   SKU {dd.sku} / tienda {dd.tienda}   "
              f"[{dd.root_cause_id_dominante}] {dd.causa_dominante}")

    aviso = aviso_prioridad_3(fu)
    if aviso:
        print(f"\n{'!' * 78}")
        for linea in textwrap.wrap(aviso, 76):
            print(f"! {linea}")
        print("!" * 78)

    print(f"\nResultado escrito en: {salida.name}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
