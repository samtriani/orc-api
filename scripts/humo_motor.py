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

from orcmm_rca_engine import EvidenciaSKUTienda, MotorRCA, TipoResurtido, ViaResurtido

MOTOR = MotorRCA()
FALLOS = []
CASOS = [0]      # se cuenta solo: el numero fijo se desfasaba al agregar casos


def caso(nombre, esperado_rc, esperado_resp=None, **campos):
    campos.setdefault("en_catalogo", True)
    ev = EvidenciaSKUTienda(sku="X", tienda="287", fecha=date(2026, 3, 20),
                            osa=0.0, venta_perdida=100.0, **campos)
    CASOS[0] += 1
    d = MOTOR.diagnosticar(ev)
    rc = d.get("root_cause_id")
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
caso("Con pedido y sin cita -> RC07 tardio", "RC07", "Compras / Abasto",
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

if FALLOS:
    print("\nFALLÓ:")
    for f in FALLOS:
        print("   -", f)
    sys.exit(1)
print()
print(f"OK - {CASOS[0]} veredictos del motor sin cambios")
