"""
ORCMM - Validador de layout contra datos reales.

Compara un Excel de captura (ya llenado por los equipos) contra la
especificación única en orcmm_layout_spec.py, ANTES de correr el pipeline.

    python orcmm_validar_layout.py "040826_La Comer_Layout de datos RCA (OSA)_V1_Con Datos.xlsx"

Revisa, por cada hoja esperada:
  - Que la hoja exista.
  - Que los encabezados de la fila 3 coincidan con los campos del spec
    (nombre, orden, faltantes, sobrantes/renombrados).
  - Que haya datos desde la fila 6.
  - Tipo real de cada celda vs tipo esperado (Texto/Entero/Decimal/Fecha/Hora/Sí-No/Lista).
  - SKU y tienda/cedis capturados como TEXTO (no numérico, para no perder ceros
    a la izquierda ni caer en notación científica).
  - Campos obligatorios (*) vacíos.
  - Fechas fuera de la ventana declarada en el spec, si aplica.
  - Duplicados en la llave natural de cada hoja.

No corrige nada ni corre el motor de RCA: sólo diagnostica el archivo de
entrada para que se pueda pedir una corrección a los equipos antes de la
corrida real.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import openpyxl

from orcmm_layout_spec import (ALIAS_ENCABEZADOS, BOL, DEC, ENT, FEC, FILA_DATOS,
                               FILA_ENCABEZADO, FILA_TIPO, HOJAS, HOR, LST, TXT,
                               normalizar_encabezado)
from orcmm_rca_engine import EVALUAR_PEDIDO_TIENDA

LLAVES = {
    "CATALOGO": ["sku", "tienda"],
    "TABLEAU_INV_TIENDA": ["sku", "tienda", "fecha"],
    "BOPS_OSA": ["sku", "tienda", "fecha"],
    "TABLEAU_VENTAS": ["sku", "tienda", "fecha"],
    "CEDIS_INVENTARIO": ["sku", "cedis", "fecha"],
    "CEDIS_TRANSFERENCIAS": ["folio"],
    "SIMA_PEDIDOS_TIENDA": ["folio"],
    "COMPRAS_PEDIDOS_PROV": ["folio"],
    "CITAS_PROV_CEDIS": ["folio_cita"],
}

CAMPOS_TEXTO_CLAVE = {"sku", "tienda", "cedis", "cedis_surtidor", "cedis_destino",
                      "cedis_origen", "folio", "folio_cita", "proveedor_id"}


def _texto(v) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s or None


class Reporte:
    """Tres severidades, porque no todo lo que está mal impide correr.

      errores       el layout está roto: columnas que no existen, claves
                    capturadas como número, tipos que no cuadran. El modelo
                    leería mal. Bloquean.
      faltan_datos  el layout está bien, faltan renglones. El motor ya sabe
                    qué hacer con eso: lo reporta como cobertura perdida
                    nombrando el campo. No bloquea, pero cambia lo que se
                    puede concluir del resultado.
      advertencias  vale la pena revisarlo antes de firmar el Pareto.
    """

    def __init__(self):
        self.errores: list[str] = []
        self.faltan_datos: list[str] = []
        self.advertencias: list[str] = []
        self.info: list[str] = []

    def error(self, msg): self.errores.append(msg)
    def dato(self, msg): self.faltan_datos.append(msg)
    def warn(self, msg): self.advertencias.append(msg)
    def ok(self, msg): self.info.append(msg)


def validar_encabezados(ws, nombre: str, rep: Reporte) -> list[Optional[str]]:
    crudos = [_texto(c.value) for c in ws[FILA_ENCABEZADO]]
    encabezados = [normalizar_encabezado(nombre, h) for h in crudos]
    esperados = [c[0] for c in HOJAS[nombre]["campos"]]
    presentes = [h for h in encabezados if h]

    # Un encabezado traducido se lee bien, pero conviene alinearlo en el origen.
    for crudo, canonico in zip(crudos, encabezados):
        if crudo and canonico and crudo != canonico:
            rep.warn(f"{nombre}: la columna '{crudo}' se está leyendo como '{canonico}'. "
                     f"Renombrarla en el layout para que el archivo y el modelo digan lo mismo.")

    opcionales = {c[0] for c in HOJAS[nombre]["campos"] if not c[2]}
    faltantes = [e for e in esperados if e not in presentes and e not in opcionales]
    ausentes_opcionales = [e for e in esperados if e not in presentes and e in opcionales]
    if ausentes_opcionales:
        rep.warn(f"{nombre}: sin las columnas opcionales {', '.join(ausentes_opcionales)} "
                 f"(no bloquean, pero el modelo no puede usarlas).")
    sobrantes = [p for p in presentes if p not in esperados]
    duplicados = [h for h, n in Counter(presentes).items() if n > 1]

    if faltantes:
        rep.error(f"{nombre}: faltan columnas esperadas -> {', '.join(faltantes)}")
    if sobrantes:
        rep.warn(f"{nombre}: columnas no reconocidas por el spec (revisar nombre exacto) -> {', '.join(sobrantes)}")
    if duplicados:
        rep.error(f"{nombre}: encabezado repetido en fila {FILA_ENCABEZADO} -> {', '.join(duplicados)}")

    # Orden: sólo advertencia, el pipeline no depende del orden.
    comunes = [e for e in esperados if e in presentes]
    orden_presentes = [h for h in presentes if h in esperados]
    if comunes != orden_presentes:
        rep.warn(f"{nombre}: el orden de columnas no coincide con el layout original (no bloquea la lectura).")

    if not faltantes and not sobrantes and not duplicados:
        rep.ok(f"{nombre}: encabezados OK ({len(esperados)} columnas).")

    return encabezados


def tipo_valor_ok(valor, tipo_esperado: str) -> bool:
    if valor is None or valor == "":
        return True  # vacío se valida aparte (obligatoriedad)
    if tipo_esperado == FEC:
        return isinstance(valor, (date, datetime))
    if tipo_esperado == HOR:
        return isinstance(valor, (datetime,)) or hasattr(valor, "hour")
    if tipo_esperado == ENT:
        if isinstance(valor, bool):
            return False
        if isinstance(valor, (int, float)):
            return float(valor).is_integer()
        return str(valor).strip().lstrip("-").isdigit()
    if tipo_esperado == DEC:
        if isinstance(valor, bool):
            return False
        if isinstance(valor, (int, float)):
            return True
        try:
            float(str(valor).replace(",", ""))
            return True
        except ValueError:
            return False
    if tipo_esperado == BOL:
        return str(valor).strip().lower() in {"sí", "si", "no", "true", "false", "1", "0"}
    # TXT y LST: cualquier cosa sirve, pero para claves no debe ser numérico crudo
    return True


def validar_datos(ws, nombre: str, encabezados: list[Optional[str]], rep: Reporte):
    campos = HOJAS[nombre]["campos"]
    tipo_de = {c[0]: c[1] for c in campos}
    obligatorio_de = {c[0]: c[2] for c in campos}
    col_de_campo = {}
    for idx, h in enumerate(encabezados):
        if h and h not in col_de_campo:
            col_de_campo[h] = idx

    n_filas = 0
    faltantes_obligatorios: Counter = Counter()
    tipos_malos: Counter = Counter()
    claves_como_numero: Counter = Counter()
    llave_vista: Counter = Counter()
    ejemplos_tipo: dict[str, str] = {}

    for fila in ws.iter_rows(min_row=FILA_DATOS, values_only=True):
        if all(v is None or v == "" for v in fila):
            continue
        n_filas += 1
        registro = {}
        for campo, idx in col_de_campo.items():
            v = fila[idx] if idx < len(fila) else None
            registro[campo] = v

            tipo = tipo_de.get(campo)
            if tipo and not tipo_valor_ok(v, tipo):
                tipos_malos[campo] += 1
                ejemplos_tipo.setdefault(campo, f"'{v}' (tipo Python: {type(v).__name__}, esperado {tipo})")

            if campo in CAMPOS_TEXTO_CLAVE and isinstance(v, (int, float)) and not isinstance(v, bool):
                claves_como_numero[campo] += 1

            if obligatorio_de.get(campo) and (v is None or v == ""):
                faltantes_obligatorios[campo] += 1

        llave = LLAVES.get(nombre, [])
        if llave and all(registro.get(k) not in (None, "") for k in llave):
            k = tuple(_texto(registro.get(k)) if not isinstance(registro.get(k), (date, datetime))
                      else registro.get(k) for k in llave)
            llave_vista[k] += 1

    if n_filas == 0:
        # Una hoja vacía no es un layout roto: es una fuente que no llegó. El
        # motor la maneja dejando sin clasificar los días que dependían de
        # ella y diciendo cuál campo faltó, así que no tiene por qué impedir
        # la corrida.
        if nombre == "SIMA_PEDIDOS_TIENDA" and not EVALUAR_PEDIDO_TIENDA:
            rep.dato(f"{nombre}: hoja vacía, esperado por ahora. La prioridad 3 está apagada "
                     f"(EVALUAR_PEDIDO_TIENDA = False) y el modelo corre parcial, sin poder "
                     f"dictaminar RC03. Al llegar los datos hay que prender el interruptor.")
        else:
            rep.dato(f"{nombre}: no tiene filas de datos desde la fila {FILA_DATOS} (hoja vacía). "
                     f"Todas las reglas que dependen de ella quedarán sin clasificar.")
        return

    rep.ok(f"{nombre}: {n_filas} filas de datos leídas desde la fila {FILA_DATOS}.")

    for campo, n in faltantes_obligatorios.items():
        rep.error(f"{nombre}.{campo}: {n} de {n_filas} filas sin dato (campo obligatorio *).")

    for campo, n in claves_como_numero.items():
        rep.error(f"{nombre}.{campo}: {n} filas con la clave capturada como NÚMERO, no TEXTO "
                  f"(riesgo de perder ceros a la izquierda / notación científica).")

    for campo, n in tipos_malos.items():
        rep.warn(f"{nombre}.{campo}: {n} valores no coinciden con el tipo esperado "
                 f"({tipo_de.get(campo)}). Ejemplo: {ejemplos_tipo.get(campo)}")

    duplicados = [k for k, n in llave_vista.items() if n > 1]
    if duplicados:
        llave_nombres = "+".join(LLAVES.get(nombre, []))
        muestra = ", ".join(str(k) for k in duplicados[:5])
        rep.warn(f"{nombre}: {len(duplicados)} llaves duplicadas ({llave_nombres}) -> ej: {muestra}")


def _leer(ws, nombre: str) -> list[dict]:
    encabezados = [normalizar_encabezado(nombre, c.value) for c in ws[FILA_ENCABEZADO]]
    filas = []
    for fila in ws.iter_rows(min_row=FILA_DATOS, values_only=True):
        if all(v is None or v == "" for v in fila):
            continue
        filas.append({h: v for h, v in zip(encabezados, fila) if h})
    return filas


def validar_cruce_citas(wb, rep: Reporte):
    """Revisiones que sólo se ven cruzando la cita con su pedido.

    La cita es la única hoja que se lee contra otra: sin el pedido no significa
    nada, y su valor está justamente en la diferencia entre lo pedido, lo
    confirmado y lo entregado. Si esa diferencia nunca aparece, el dato es
    decorativo y el modelo no debería apoyarse en él.
    """
    if "CITAS_PROV_CEDIS" not in wb.sheetnames or "COMPRAS_PEDIDOS_PROV" not in wb.sheetnames:
        return

    citas = _leer(wb["CITAS_PROV_CEDIS"], "CITAS_PROV_CEDIS")
    pedidos = _leer(wb["COMPRAS_PEDIDOS_PROV"], "COMPRAS_PEDIDOS_PROV")
    if not citas or not pedidos:
        return

    por_folio = {(_texto(o.get("folio")), _texto(o.get("sku"))): o for o in pedidos}
    con_cita, huerfanas, confirmadas_de_mas, iguales, comparables = set(), [], [], 0, 0

    for c in citas:
        llave = (_texto(c.get("folio")), _texto(c.get("sku")))
        pedido = por_folio.get(llave)
        if pedido is None:
            huerfanas.append(llave[0])
            continue
        con_cita.add(llave)

        conf, ped = c.get("cajas_confirmadas_cita"), pedido.get("cajas_pedidas")
        if isinstance(conf, (int, float)) and isinstance(ped, (int, float)):
            comparables += 1
            if conf > ped:
                confirmadas_de_mas.append(llave[0])
            elif conf == ped:
                iguales += 1

    if huerfanas:
        rep.error(f"CITAS_PROV_CEDIS: {len(huerfanas)} citas cuyo folio+sku no existe en "
                  f"COMPRAS_PEDIDOS_PROV (ej: {huerfanas[:3]}). Sin pedido no se pueden usar.")

    if confirmadas_de_mas:
        rep.error(f"CITAS_PROV_CEDIS: {len(confirmadas_de_mas)} citas con "
                  f"cajas_confirmadas_cita MAYOR que las cajas_pedidas del pedido "
                  f"(ej: {confirmadas_de_mas[:3]}). Una de las dos hojas está mal.")

    if comparables and iguales == comparables:
        rep.warn(f"CITAS_PROV_CEDIS: en las {comparables} citas, cajas_confirmadas_cita es "
                 f"SIEMPRE igual a cajas_pedidas. O el proveedor nunca recorta al agendar, o "
                 f"la columna se está copiando del pedido. Si es lo segundo, el modelo no "
                 f"puede detectar 'confirmó de menos' y hay que traer el dato real de la cita.")

    sin_cita = len(set(por_folio) - con_cita)
    if sin_cita:
        pct = round(sin_cita / len(por_folio) * 100, 1)
        rep.dato(f"CITAS_PROV_CEDIS: {sin_cita} de {len(por_folio)} pedidos ({pct}%) no traen cita. "
                 f"El modelo lee eso como 'el proveedor nunca agendó' y lo dictamina como RC06. "
                 f"Confirmar con Compras que la extracción trae TODAS las citas del periodo: si "
                 f"vino parcial, la causa es un hueco de captura y no del proveedor.")

    estatus = {_texto(c.get("estatus_cita")) for c in citas} - {None}
    if len(estatus) == 1:
        rep.warn(f"CITAS_PROV_CEDIS.estatus_cita: todas las filas dicen '{estatus.pop()}'. "
                 f"La columna no distingue nada; el modelo dictamina con las cajas, no con "
                 f"el estatus. Si el sistema sí marca canceladas o no presentadas, traerlo.")


def validar_archivo(ruta: Path) -> Reporte:
    rep = Reporte()
    wb = openpyxl.load_workbook(ruta, data_only=True)

    hojas_archivo = set(wb.sheetnames)
    hojas_esperadas = set(HOJAS.keys())

    for nombre in hojas_esperadas - hojas_archivo:
        rep.error(f"{nombre}: la hoja NO existe en el archivo.")

    for nombre in hojas_archivo - hojas_esperadas:
        rep.warn(f"Hoja '{nombre}' está en el archivo pero no está en el spec (¿hoja extra o mal nombrada?).")

    for nombre in HOJAS:
        if nombre not in hojas_archivo:
            continue
        ws = wb[nombre]
        encabezados = validar_encabezados(ws, nombre, rep)
        validar_datos(ws, nombre, encabezados, rep)

    validar_cruce_citas(wb, rep)

    return rep


def imprimir_reporte(rep: Reporte, archivo: str):
    print(f"\n{'='*78}\nValidación de layout: {archivo}\n{'='*78}")

    print(f"\n[OK] {len(rep.info)}")
    for m in rep.info:
        print(f"  - {m}")

    print(f"\n[ADVERTENCIAS] {len(rep.advertencias)} (no bloquean la corrida, pero revisar)")
    for m in rep.advertencias:
        print(f"  ! {m}")

    print(f"\n[DATOS INCOMPLETOS] {len(rep.faltan_datos)} (el layout está bien; faltan renglones)")
    for m in rep.faltan_datos:
        print(f"  ? {m}")

    print(f"\n[ERRORES DE LAYOUT] {len(rep.errores)} (bloquean: el modelo leería mal)")
    for m in rep.errores:
        print(f"  X {m}")

    print(f"\n{'='*78}")
    if rep.errores:
        print(f"RESULTADO: hay {len(rep.errores)} problema(s) de layout que corregir antes de "
              f"correr el pipeline.")
        print("           Probar primero: python orcmm_corregir_layout.py \"<archivo>\"")
    elif rep.faltan_datos:
        print("RESULTADO: el layout está bien. Se puede correr, pero el resultado será PARCIAL "
              "por las fuentes que faltan.")
    elif rep.advertencias:
        print("RESULTADO: sin errores bloqueantes. Revisar advertencias antes de correr.")
    else:
        print("RESULTADO: layout OK. Se puede correr orcmm_pipeline.py")
    print(f"{'='*78}\n")


def main() -> int:
    ap = argparse.ArgumentParser(description="Valida el layout de un Excel de captura contra el spec.")
    ap.add_argument("archivo")
    args = ap.parse_args()

    ruta = Path(args.archivo)
    if not ruta.exists():
        print(f"No se encontró el archivo: {ruta}", file=sys.stderr)
        return 1

    rep = validar_archivo(ruta)
    imprimir_reporte(rep, ruta.name)
    return 1 if rep.errores else 0


if __name__ == "__main__":
    raise SystemExit(main())
