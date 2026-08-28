"""Prueba de humo del motor RCA, sin tocar Postgres.

No pretende cubrir la matriz entera: fija los veredictos que costaron trabajo
averiguar esta semana, para que un cambio futuro no los mueva en silencio.
Cada caso trae de dónde salió.

Corre en el CI y tarda menos de un segundo.
"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from orcmm_rca_engine import (CAUSAS_DE_PEDIDO, FUSIONAR_PEDIDOS, PROPAGAR_RC06,
                              EvidenciaSKUTienda, MotorRCA, TipoResurtido,
                              ViaResurtido)
from orcmm_rca_periodo import clasificar

MOTOR = MotorRCA()
FALLOS = []
CASOS = [0]      # se cuenta solo: el numero fijo se desfasaba al agregar casos


def caso(nombre, esperado_rc, esperado_resp=None, **campos):
    campos.setdefault("en_catalogo", True)
    ev = EvidenciaSKUTienda(sku="X", tienda="287", fecha=date(2026, 3, 20),
                            osa=0.0, venta_perdida=100.0, **campos)
    CASOS[0] += 1
    d = MOTOR.diagnosticar(ev)
    # causa_base y no root_cause_id: estos casos fijan las REGLAS del árbol,
    # no cómo se presentan. Ver FUSIONAR_PEDIDOS — con la fusión puesta,
    # root_cause_id dice "RC08" para seis de estos casos, y compararlo aquí
    # los reprobaría a todos sin que ninguna regla hubiera cambiado.
    rc = d.get("causa_base", d.get("root_cause_id"))
    resp = d.get("responsable")
    mal = rc != esperado_rc or (esperado_resp and resp != esperado_resp)
    print(f"  {'MAL ' if mal else 'ok  '} {nombre:<52} {rc} · {resp}")
    if mal:
        FALLOS.append(f"{nombre}: esperaba {esperado_rc}/{esperado_resp}, dio {rc}/{resp}")


print("Prioridad 0-2")
caso("SKU fuera del catálogo -> fuera de alcance", "RC00",
     en_catalogo=False)
caso("Inventario en tienda > 0 -> ejecución", "RC01", "Tienda",
     inventario_tienda=12)
caso("Inventario 0 y tránsito vigente -> transporte", "RC02", "Logística",
     inventario_tienda=0, transito_vigente=True)
caso("Sin dato de inventario -> se detiene", "RC99",
     inventario_tienda=None)

print("\nPrioridad 3 — el pedido de tienda")
# El responsable depende de tipo_resurtido: automático es de Compras. Es el
# empate que fallaba por la tilde de "Automático" (ver clave_catalogo).
caso("No pidió y el resurtido es automático -> Compras", "RC03", "Compras / Abasto",
     inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=False,
     tipo_resurtido=TipoResurtido.AUTOMATICO)
caso("No pidió y el resurtido es manual -> Tienda", "RC03", "Tienda",
     inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=False,
     tipo_resurtido=TipoResurtido.MANUAL)

print("\nRama CEDIS")
comun = dict(inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=True,
             via_resurtido=ViaResurtido.VIA_1)
caso("CEDIS tenía y no envió -> CEDIS", "RC04", "CEDIS",
     **comun, inventario_cedis=50, envio_cedis_generado=False)
# El sub-paso que suele faltar en las versiones escritas del árbol: tener
# inventario en CEDIS no basta para culpar a CEDIS.
caso("CEDIS tenía y SÍ envió -> transporte", "RC02", "Logística",
     **comun, inventario_cedis=50, envio_cedis_generado=True)
# La costura entre RC05 y RC07: NO hay pedido vigente contra SI hay pero en
# plazo. Es la separacion que pidio La Comer el 2026-08-22, y es justo donde
# alguien podria volver a juntarlas sin darse cuenta.
caso("Sin pedido a proveedor -> RC05 no generado", "RC05", "Compras / Abasto",
     **comun, inventario_cedis=0, pedido_proveedor_generado=False)
caso("Con pedido y cita aun por vencer -> RC07 tardio", "RC07", "Compras / Abasto",
     **comun, inventario_cedis=0, pedido_proveedor_generado=True,
     proveedor_cajas_pedidas=10, proveedor_cita_agendada=True,
     proveedor_cita_vencida=False)
# Este caso fija SIN_CITA_VA_A, que es un acuerdo de negocio y no una regla:
# estuvo en "compras" del 22 al 28 de agosto y volvio a "proveedor" cuando se
# midio que el 73% de los pedidos de marzo siguen sin cita DENTRO de la
# ventana que las citas cubren. Si alguien lo mueve, este caso lo dice.
caso("Con pedido y sin cita -> incumplimiento del proveedor", "RC06", "Proveedor",
     **comun, inventario_cedis=0, pedido_proveedor_generado=True,
     proveedor_cajas_pedidas=10, proveedor_cita_agendada=False)

print("\nRama DSD — estaba muerta hasta el 2026-08-21")
dsd = dict(inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=True,
           via_resurtido=ViaResurtido.DSD, pedido_dsd_generado=True)
caso("El proveedor no entregó -> incumplimiento", "RC06", "Proveedor",
     **dsd, dsd_entrego_tienda=False)
# La 10 sorprende: si el proveedor SÍ dejó el producto y aun así el anaquel
# estuvo vacío, vuelve a ser ejecución en tienda.
caso("El proveedor sí entregó -> ejecución en tienda", "RC01", "Tienda",
     **dsd, dsd_entrego_tienda=True)
caso("Nadie le pidió al proveedor -> RC05", "RC05", "Compras / Abasto",
     inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=True,
     via_resurtido=ViaResurtido.DSD, pedido_dsd_generado=False)


print()
print("La fusión de las causas de pedido")


def caso_fusion(nombre, base, **campos):
    """Verifica la capa de presentación, no la regla.

    Lo que puede romperse aquí es que una causa de pedido escape de la bolsa
    —se vería suelta en el Pareto— o que una que no lo es caiga dentro.
    """
    campos.setdefault("en_catalogo", True)
    ev = EvidenciaSKUTienda(sku="X", tienda="287", fecha=date(2026, 3, 20),
                            osa=0.0, venta_perdida=100.0, **campos)
    CASOS[0] += 1
    d = MOTOR.diagnosticar(ev)
    esperado = "RC08" if (FUSIONAR_PEDIDOS and base in CAUSAS_DE_PEDIDO) else base
    rc, sub = d.get("root_cause_id"), d.get("subcausa")
    # Fusionar sin subcausa dejaría el día dentro de la bolsa sin nada que
    # dijera qué pasó: la subcausa es lo único que sobrevive a la fusión.
    mal = rc != esperado or (rc == "RC08" and not sub)
    print(f"  {'MAL ' if mal else 'ok  '} {nombre:<52} {rc} · {sub or '(sin subcausa)'}")
    if mal:
        FALLOS.append(f"{nombre}: esperaba {esperado} con subcausa, dio {rc}/{sub}")


caso_fusion("RC03 entra a Pedidos y conserva subcausa", "RC03",
            inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=False,
            tipo_resurtido=TipoResurtido.AUTOMATICO)
# Sin tipo_resurtido no había subcausa: es el caso que obligó a inventarle una
# antes de poder fusionar.
caso_fusion("RC03 sin tipo de resurtido tampoco se queda mudo", "RC03",
            inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=False)
caso_fusion("RC05 entra a Pedidos y conserva subcausa", "RC05",
            **comun, inventario_cedis=0, pedido_proveedor_generado=False)
caso_fusion("RC07 entra a Pedidos y conserva subcausa", "RC07",
            **comun, inventario_cedis=0, pedido_proveedor_generado=True,
            proveedor_cajas_pedidas=10, proveedor_cita_agendada=True,
            proveedor_cita_vencida=False)
caso_fusion("RC05 de la rama DSD también entra", "RC05",
            inventario_tienda=0, transito_vigente=False, pedido_tienda_generado=True,
            via_resurtido=ViaResurtido.DSD, pedido_dsd_generado=False)
# Los controles: ninguna de estas es causa de pedido y no deben caer dentro.
caso_fusion("RC01 se queda fuera de la bolsa", "RC01", inventario_tienda=12)
caso_fusion("RC04 se queda fuera de la bolsa", "RC04",
            **comun, inventario_cedis=50, envio_cedis_generado=False)


# ---------------------------------------------------------------------------
# La propagacion temporal de RC06 (ver PROPAGAR_RC06 en el motor)
#
# Estos casos NO pasan por diagnosticar(): la propagacion depende de la serie
# completa del SKU-tienda, asi que van por clasificar(), que es por donde
# entra todo el pipeline. Cada caso fija una frontera distinta, porque las
# fronteras son lo unico dificil de esta regla.
# ---------------------------------------------------------------------------

def dia(d, folio, sku="X", tienda="287", osa=0.0, inv=0, **kw):
    """Un dia de Via 1 con CEDIS en cero: la rama que juzga al proveedor.

    Sin folio no hay pedido vigente, y la matriz sola dictamina RC05 (la
    bolsa "Pedidos"). Con folio y cero cajas entregadas dictamina RC06. Esa
    es la pinza: los dias "sin folio" son los que la propagacion debe poder
    convertir, y por eso son los que se usan de relleno.
    """
    return EvidenciaSKUTienda(
        sku=sku, tienda=tienda, fecha=date(2026, 3, d), osa=osa,
        venta_perdida=100.0, en_catalogo=True, inventario_tienda=inv,
        transito_vigente=False, pedido_tienda_generado=True,
        via_resurtido=ViaResurtido.VIA_1, inventario_cedis=0,
        envio_cedis_generado=False,
        pedido_proveedor_generado=folio is not None,
        proveedor_cajas_pedidas=40 if folio else None,
        proveedor_cajas_entregadas=0 if folio else None,
        proveedor_folio_pedido=folio, **kw)


def dia_dsd(d, entrego, osa=0.0):
    return EvidenciaSKUTienda(
        sku="X", tienda="287", fecha=date(2026, 3, d), osa=osa,
        venta_perdida=100.0, en_catalogo=True, inventario_tienda=0,
        transito_vigente=False, pedido_tienda_generado=True,
        via_resurtido=ViaResurtido.DSD, pedido_dsd_generado=True,
        dsd_entrego_tienda=entrego)


def caso_prop(nombre, evidencias, esperado):
    """esperado: la cadena de causas presentadas, con * en las propagadas."""
    CASOS[0] += 1
    dgs = clasificar(evidencias)
    real = " ".join(d["root_cause_id"] + ("*" if d.get("propagacion_rc06") else "")
                    for d in dgs)
    mal = real != esperado
    print(f"  {'MAL ' if mal else 'ok  '} {nombre:<52} {real}")
    if mal:
        FALLOS.append(f"{nombre}: esperaba '{esperado}', dio '{real}'")


if PROPAGAR_RC06:
    print("\nPropagacion temporal de RC06 — * marca el dia heredado")

    caso_prop("Hereda mientras no haya nueva vigencia",
              [dia(10, "A"), dia(11, None), dia(12, None)],
              "RC06 RC06* RC06*")
    # Un pedido nuevo es la oportunidad de reevaluar: manda lo que diga la
    # matriz ese dia, aunque venga de un incumplimiento.
    caso_prop("Nueva vigencia con otra causa corta la cadena",
              [dia(10, "A"), dia(11, None),
               dia(12, "B", proveedor_cita_agendada=True,
                   proveedor_cita_vencida=False)],
              "RC06 RC06* RC08")
    caso_prop("Nueva vigencia que vuelve a fallar reinicia el periodo",
              [dia(10, "A"), dia(11, "B"), dia(12, None)],
              "RC06 RC06 RC06*")
    caso_prop("No cruza entre SKU",
              [dia(10, "A", sku="X"), dia(11, None, sku="Y")],
              "RC06 RC08")
    caso_prop("No cruza entre tiendas",
              [dia(10, "A", tienda="287"), dia(11, None, tienda="280")],
              "RC06 RC08")
    caso_prop("Un OSA por encima de cero no se sobrescribe",
              [dia(10, "A"), dia(11, None, osa=50.0)],
              "RC06 RC08")
    # Vacio no es cero, aqui tambien: sin OSA no se puede afirmar que el
    # hueco siga abierto.
    caso_prop("Sin dato de OSA no se propaga",
              [dia(10, "A"), dia(11, None, osa=None)],
              "RC06 RC08")
    caso_prop("Tambien hereda en la rama DSD",
              [dia_dsd(10, False), dia_dsd(11, None)],
              "RC06 RC06*")
    # derivar_evidencias solo emite dias CON faltante, asi que un OSA
    # recuperado no llega como fila: llega como fecha ausente.
    caso_prop("Un hueco de fechas corta la cadena",
              [dia(10, "A"), dia(11, None), dia(14, None)],
              "RC06 RC06* RC08")
    # Las tres condiciones de cierre: el dia trae producto, asi que el
    # incumplimiento ya no es lo que aprieta.
    caso_prop("Que el proveedor SI entregue en tienda cierra el periodo",
              [dia_dsd(10, False), dia_dsd(11, True), dia_dsd(12, None)],
              "RC06 RC01 RC99")
    caso_prop("Producto en tienda cierra el periodo",
              [dia(10, "A"), dia(11, None, inv=5), dia(12, None)],
              "RC06 RC01 RC08")
    # Que el folio desaparezca no es una oportunidad nueva: es la ausencia
    # de una. Es la mitad del valor de la regla.
    caso_prop("Quedarse sin folio no es una nueva oportunidad",
              [dia(10, "A"), dia(11, None), dia(12, None)],
              "RC06 RC06* RC06*")

    # El orden de salida es contrato: hay dos lugares que vuelven a parear
    # evidencias y dictamenes con zip().
    CASOS[0] += 1
    evs = [dia(12, None), dia(10, "A"), dia(11, None)]
    dgs = clasificar(evs)
    pareado = all(d["fecha"] == e.fecha.isoformat() for e, d in zip(evs, dgs))
    print(f"  {'ok  ' if pareado else 'MAL '} "
          f"{'La salida conserva el orden de entrada':<52} "
          f"{[d['fecha'][-2:] for d in dgs]}")
    if not pareado:
        FALLOS.append("La salida de clasificar() se reordeno")

    # Nada se pisa en silencio: el dia heredado conserva que decia antes.
    CASOS[0] += 1
    d = clasificar([dia(10, "A"), dia(11, None)])[1]
    traza = {"causa_base_original": "RC05", "root_cause_id_original": "RC08",
             "tipo_clasificacion": "propagada", "propagacion_rc06": True,
             "fecha_origen_propagacion_rc06": "2026-03-10",
             "folio_origen_propagacion_rc06": "A", "prioridad_regla_original": 7}
    faltan = {k: (v, d.get(k)) for k, v in traza.items() if d.get(k) != v}
    print(f"  {'MAL ' if faltan else 'ok  '} "
          f"{'El dia heredado conserva su dictamen original':<52} "
          f"{d.get('root_cause_id_original')} -> {d.get('root_cause_id')}")
    if faltan:
        FALLOS.append(f"Trazabilidad incompleta: {faltan}")


if FALLOS:
    print("\nFALLÓ:")
    for f in FALLOS:
        print("   -", f)
    sys.exit(1)
print()
print(f"OK - {CASOS[0]} veredictos del motor sin cambios")
