"""
ORCMM - API del analizador de desabasto.

    uvicorn api.main:app --reload --port 8000     (desde la raíz del proyecto)

Un solo flujo, que el front encadena:

    POST /api/validar          ¿el paquete está bien? ¿se puede arreglar solo?
    POST /api/analizar         encola el análisis y devuelve un id
    GET  /api/analizar/{id}    estado y, cuando termina, el resumen
    GET  /api/resultado/{id}   descarga el Excel de resultados

Desde el layout V5 el paquete son VARIOS archivos: el .xlsx del layout más los
CSV de las hojas que ya no caben en una hoja de Excel (TABLEAU_INV_TIENDA son
2.7 millones de filas). Se pueden subir sueltos o dentro de un .zip — el zip
es lo recomendable: esos CSV pesan 222 MB sueltos y ~11 MB comprimidos.

El análisis es asíncrono porque con volumen real tarda alrededor de minuto y
medio, y un request abierto tanto tiempo se lo lleva cualquier proxy de por
medio. Corre en un solo hilo de trabajo: dos análisis a la vez no caben en la
memoria de la máquina.

El servidor no guarda nada permanente: cada análisis vive en su carpeta
temporal y se borra sola pasado un rato.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from psycopg2.extras import RealDictCursor
from pydantic import BaseModel, Field
from starlette.concurrency import run_in_threadpool

# En Fly.io DATABASE_URL llega por `fly secrets set` (variable de entorno
# real); en desarrollo local no existe ese secreto, así que se carga del
# .env si está — no hace nada si el archivo no existe.
load_dotenv()

from api.servicio import analizar, diagnosticar_layout               # noqa: E402
from orcmm_db import conectar                                        # noqa: E402
from orcmm_expediente_db import expediente_sku                       # noqa: E402
from orcmm_fuentes_csv import hoja_de_archivo                        # noqa: E402
from orcmm_fuentes_db import leer_fuentes_db                         # noqa: E402

# El CSV de inventario más grande de la entrega real pesa 48 MB; el layout
# lleno, 35 MB. El tope por archivo deja margen sin volverse una invitación.
MAX_ARCHIVO = 80 * 1024 * 1024
# Los 6 CSV más el Excel suman ~260 MB sueltos. El mismo tope aplica al
# contenido expandido de un zip, que es lo que evita un zip bomb.
MAX_PAQUETE = 400 * 1024 * 1024
VIDA_TRABAJO_S = 60 * 60          # una hora
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

EXTENSIONES = {".xlsx", ".csv", ".zip"}

# De dónde se acepta al navegador. En local, el front de desarrollo; en
# despliegue, la variable de entorno, para no tener que redesplegar el back
# cada vez que el front cambia de dominio:
#   fly secrets set ORCMM_ORIGENES="https://orc-gui.vercel.app"
ORIGENES_LOCALES = "http://localhost:4200,http://127.0.0.1:4200"
ORIGENES = [o.strip() for o in os.getenv("ORCMM_ORIGENES", ORIGENES_LOCALES).split(",")
            if o.strip()]

app = FastAPI(title="ORCMM — Clasificación de desabasto por causa raíz",
              version="3.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=ORIGENES,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

# Un solo trabajador: el análisis carga el Excel entero y varios índices en
# memoria, y dos en paralelo tumbarían la máquina. Los demás esperan turno.
EJECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="orcmm")


# ---------------------------------------------------------------------------
# Trabajos en curso
# ---------------------------------------------------------------------------

@dataclass
class Trabajo:
    carpeta: Path
    archivo: str = ""
    estado: str = "en_proceso"
    resultado: Optional[dict] = None
    error: Optional[str] = None
    salida: Optional[Path] = None
    nombre_descarga: Optional[str] = None
    creado: float = field(default_factory=time.time)


TRABAJOS: Dict[str, Trabajo] = {}
CANDADO = threading.Lock()


def _limpiar_viejos() -> None:
    corte = time.time() - VIDA_TRABAJO_S
    with CANDADO:
        viejos = [k for k, t in TRABAJOS.items()
                  if t.creado < corte and t.estado != "en_proceso"]
        for id_ in viejos:
            shutil.rmtree(TRABAJOS.pop(id_).carpeta, ignore_errors=True)


# ---------------------------------------------------------------------------
# Recepción del paquete
# ---------------------------------------------------------------------------

@dataclass
class Paquete:
    xlsx: Path
    csvs: List[Path] = field(default_factory=list)
    nombre_original: str = "captura.xlsx"


def _nombre_seguro(nombre: str) -> str:
    """Sólo el nombre del archivo, nunca una ruta.

    Vale para lo que sube el usuario y para lo que viene dentro de un zip:
    una entrada llamada '../../etc/algo' es la forma clásica de escribir
    fuera de la carpeta de trabajo.
    """
    return Path(str(nombre or "").replace("\\", "/")).name


async def _volcar(archivo: UploadFile, destino: Path, presupuesto: int) -> int:
    """Escribe el archivo subido en disco sin cargarlo entero en memoria."""
    tamano = 0
    with destino.open("wb") as f:
        while trozo := await archivo.read(1024 * 1024):
            tamano += len(trozo)
            if tamano > MAX_ARCHIVO or tamano > presupuesto:
                raise HTTPException(
                    413, f"'{_nombre_seguro(archivo.filename)}' pasa del tamaño "
                         f"permitido ({MAX_ARCHIVO // 1024 // 1024} MB por archivo, "
                         f"{MAX_PAQUETE // 1024 // 1024} MB en total).")
            f.write(trozo)
    return tamano


def _expandir_zip(ruta: Path, carpeta: Path, presupuesto: int) -> List[Path]:
    """Saca del zip lo que corresponde a una fuente del layout.

    Se ignora todo lo demás —carpetas, archivos de sistema, cualquier otra
    extensión— y se lleva la cuenta del tamaño ya descomprimido: un zip chico
    puede expandirse a gigabytes.
    """
    salidas: List[Path] = []
    with zipfile.ZipFile(ruta) as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            nombre = _nombre_seguro(info.filename)
            if not nombre or nombre.startswith(".") or nombre.startswith("~$"):
                continue
            ext = Path(nombre).suffix.lower()
            if ext not in (".xlsx", ".csv"):
                continue

            presupuesto -= info.file_size
            if info.file_size > MAX_ARCHIVO or presupuesto < 0:
                raise HTTPException(
                    413, f"El contenido del zip pasa de "
                         f"{MAX_PAQUETE // 1024 // 1024} MB descomprimido.")

            destino = carpeta / nombre
            with z.open(info) as origen, destino.open("wb") as f:
                shutil.copyfileobj(origen, f, 1024 * 1024)
            salidas.append(destino)
    return salidas


async def _recibir(archivos: List[UploadFile]) -> tuple[str, Paquete]:
    """Guarda el paquete subido en una carpeta propia y regresa su id.

    Se acepta el layout suelto con sus CSV, o todo dentro de un zip. Los
    nombres originales sólo se usan para etiquetar y para saber a qué hoja
    corresponde cada CSV; en disco se escriben siempre con un nombre saneado.
    """
    if not archivos:
        raise HTTPException(400, "No llegó ningún archivo.")

    _limpiar_viejos()
    id_ = uuid.uuid4().hex
    carpeta = Path(tempfile.mkdtemp(prefix=f"orcmm_{id_}_"))

    try:
        restante = MAX_PAQUETE
        recibidos: List[Path] = []

        for archivo in archivos:
            nombre = _nombre_seguro(archivo.filename)
            ext = Path(nombre).suffix.lower()
            if ext not in EXTENSIONES:
                raise HTTPException(
                    400, f"'{nombre}': sólo se aceptan .xlsx (el layout), .csv "
                         f"(las fuentes grandes) o un .zip con todo dentro.")

            destino = carpeta / f"{len(recibidos):02d}{ext}"
            restante -= await _volcar(archivo, destino, restante)

            if ext == ".zip":
                try:
                    salidas = _expandir_zip(destino, carpeta, restante)
                except zipfile.BadZipFile:
                    raise HTTPException(400, f"'{nombre}' no se pudo abrir como zip.")
                destino.unlink(missing_ok=True)
                if not salidas:
                    raise HTTPException(
                        400, f"'{nombre}' no trae ningún .xlsx ni .csv adentro.")
                recibidos.extend(salidas)
            else:
                # El nombre importa: de él sale a qué hoja pertenece cada CSV.
                final = carpeta / nombre
                if final != destino:
                    destino.replace(final)
                recibidos.append(final)

        xlsxs = [r for r in recibidos if r.suffix.lower() == ".xlsx"]
        if not xlsxs:
            raise HTTPException(400, "Falta el .xlsx del layout de captura.")
        if len(xlsxs) > 1:
            raise HTTPException(
                400, f"Llegaron {len(xlsxs)} archivos .xlsx y sólo puede haber un "
                     f"layout: {', '.join(x.name for x in xlsxs)}.")

        csvs = [r for r in recibidos if r.suffix.lower() == ".csv"]
        desconocidos = [c.name for c in csvs if hoja_de_archivo(c.name) is None]
        if desconocidos:
            raise HTTPException(
                400, f"Estos CSV no corresponden a ninguna hoja del layout: "
                     f"{', '.join(desconocidos)}. El nombre del archivo tiene que "
                     f"empezar con el nombre de la hoja (por ejemplo "
                     f"TABLEAU_INV_TIENDA_1.csv).")

        if xlsxs[0].stat().st_size == 0:
            raise HTTPException(400, "El layout llegó vacío.")

        paquete = Paquete(xlsx=xlsxs[0], csvs=csvs, nombre_original=xlsxs[0].name)
        with CANDADO:
            TRABAJOS[id_] = Trabajo(carpeta=carpeta, archivo=paquete.nombre_original)
        return id_, paquete

    except Exception:
        shutil.rmtree(carpeta, ignore_errors=True)
        raise


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.get("/api/salud")
def salud() -> dict:
    return {"estado": "ok", "trabajos_en_memoria": len(TRABAJOS),
            "origenes_permitidos": ORIGENES}


@app.post("/api/validar")
async def validar(archivos: List[UploadFile] = File(...)) -> dict:
    """Revisa el paquete sin analizar. Dice además si se puede corregir solo."""
    id_, paquete = await _recibir(archivos)
    trabajo = TRABAJOS[id_]
    try:
        diagnostico = diagnosticar_layout(paquete.xlsx, trabajo.carpeta, paquete.csvs)
    except Exception as e:
        shutil.rmtree(TRABAJOS.pop(id_).carpeta, ignore_errors=True)
        raise HTTPException(422, f"No se pudo leer el paquete: {e}")

    trabajo.estado = "ok"
    return {"id": id_, "archivo": paquete.nombre_original, **diagnostico}


def _correr(id_: str, paquete: Paquete, corregir: bool, forzar: bool,
            umbral_osa: float) -> None:
    """El análisis en sí, ya fuera del request. Nunca levanta: deja el motivo
    en el trabajo para que /api/analizar/{id} lo pueda contar."""
    trabajo = TRABAJOS.get(id_)
    if trabajo is None:                       # lo limpiaron mientras esperaba turno
        return
    try:
        diagnostico = diagnosticar_layout(paquete.xlsx, trabajo.carpeta, paquete.csvs)

        a_analizar = paquete.xlsx
        correccion = None
        if not diagnostico["valido"]:
            if corregir and diagnostico["ruta_corregida"]:
                a_analizar = Path(diagnostico["ruta_corregida"])
                correccion = {
                    "cambios": diagnostico["cambios_propuestos"],
                    "errores_que_siguen": diagnostico["errores_tras_correccion"],
                }
                # Los errores que la corrección no resuelve (hoja vacía, citas
                # incompletas) no bloquean: son datos que faltan, no layout
                # roto, y el modelo ya sabe reportarlos como cobertura perdida.
            elif not forzar:
                trabajo.estado = "bloqueado"
                trabajo.resultado = {"archivo": paquete.nombre_original, **diagnostico}
                return

        base = Path(paquete.nombre_original).stem
        trabajo.nombre_descarga = f"Resultado RCA - {base}.xlsx"
        salida = trabajo.carpeta / "resultado.xlsx"

        resumen = analizar(a_analizar, salida, umbral_osa, paquete.csvs)

        if not resumen["hay_resultados"]:
            trabajo.estado = "sin_datos"
            trabajo.resultado = {"archivo": paquete.nombre_original,
                                 "validacion": diagnostico["validacion"],
                                 "correccion": correccion, **resumen}
            return

        trabajo.salida = salida
        trabajo.estado = "ok"
        trabajo.resultado = {
            "archivo": paquete.nombre_original,
            "validacion": diagnostico["validacion"],
            "correccion": correccion,
            "nombre_salida": trabajo.nombre_descarga,
            **resumen,
        }
    except Exception as e:
        trabajo.estado = "error"
        trabajo.error = f"El análisis falló: {e}"


# ---------------------------------------------------------------------------
# Análisis directo desde Postgres — misma cola/Trabajo que el de archivo,
# pero sin diagnosticar_layout/corregir/forzar: no hay archivo que validar,
# el ETL (orcmm_etl_carga.py) ya validó y dejó constancia al cargar.
# ---------------------------------------------------------------------------

class SolicitudTienda(BaseModel):
    tienda: str
    desde: date
    hasta: date
    umbral_osa: float = Field(100.0, gt=0, le=100)


def _correr_desde_bd(id_: str, tienda: str, desde: date, hasta: date,
                      umbral_osa: float) -> None:
    """Mismo patrón que _correr: nunca levanta, deja el motivo en trabajo.error."""
    trabajo = TRABAJOS.get(id_)
    if trabajo is None:
        return
    try:
        fu = leer_fuentes_db(tienda, desde, hasta, umbral_osa)
        trabajo.nombre_descarga = f"Resultado RCA - tienda {tienda}.xlsx"
        salida = trabajo.carpeta / "resultado.xlsx"
        resumen = analizar(None, salida, umbral_osa, fu=fu)

        # El front espera 'validacion'/'correccion' siempre presentes (mismas
        # claves que _correr ya manda) — aquí van vacías porque no hay
        # archivo que validar.
        validacion_vacia = {"errores": [], "faltan_datos": [], "advertencias": [], "ok": []}

        if not resumen["hay_resultados"]:
            trabajo.estado = "sin_datos"
            trabajo.resultado = {"archivo": trabajo.archivo, "validacion": validacion_vacia,
                                 "correccion": None, **resumen}
            return

        trabajo.salida = salida
        trabajo.estado = "ok"
        trabajo.resultado = {
            "archivo": trabajo.archivo,
            "validacion": validacion_vacia,
            "correccion": None,
            "nombre_salida": trabajo.nombre_descarga,
            **resumen,
        }
    except Exception as e:
        trabajo.estado = "error"
        trabajo.error = f"El análisis desde base de datos falló: {e}"


@app.post("/api/analizar-tienda", status_code=202)
async def analizar_tienda(cuerpo: SolicitudTienda) -> dict:
    """Igual que /api/analizar, pero lee de Postgres para una tienda y un
    periodo en vez de un archivo subido. La tienda es obligatoria y única:
    el análisis siempre corre sobre exactamente la tienda pedida."""
    if cuerpo.hasta < cuerpo.desde:
        raise HTTPException(400, "'hasta' no puede ser anterior a 'desde'.")

    _limpiar_viejos()
    id_ = uuid.uuid4().hex
    carpeta = Path(tempfile.mkdtemp(prefix=f"orcmm_bd_{id_}_"))
    etiqueta = f"Tienda {cuerpo.tienda} · {cuerpo.desde} a {cuerpo.hasta}"
    with CANDADO:
        TRABAJOS[id_] = Trabajo(carpeta=carpeta, archivo=etiqueta)

    EJECUTOR.submit(_correr_desde_bd, id_, cuerpo.tienda, cuerpo.desde, cuerpo.hasta,
                     cuerpo.umbral_osa)
    return {"id": id_, "archivo": etiqueta, "estado": "en_proceso",
            "seguir_en": f"/api/analizar/{id_}"}


def _listar_tiendas() -> list:
    conn = conectar()
    try:
        with conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT c.tienda, s.nombre, s.formato,
                       MIN(b.fecha) AS fecha_min, MAX(b.fecha) AS fecha_max
                FROM (SELECT DISTINCT tienda FROM catalogo) c
                JOIN bops_osa b ON b.tienda = c.tienda
                LEFT JOIN sucursales s ON s.tienda = c.tienda
                GROUP BY c.tienda, s.nombre, s.formato
                ORDER BY c.tienda
            """)
            filas = cur.fetchall()
    finally:
        conn.close()
    return [{**f, "fecha_min": f["fecha_min"].isoformat() if f["fecha_min"] else None,
             "fecha_max": f["fecha_max"].isoformat() if f["fecha_max"] else None}
            for f in filas]


@app.get("/api/tiendas")
async def tiendas() -> dict:
    """Tiendas con datos cargados en Postgres, para el selector del front.
    Sólo las que ya tienen BOPS_OSA (si no, elegirlas sólo llevaría a "sin
    datos"); nombre/formato son informativos, de la tabla `sucursales`."""
    try:
        filas = await run_in_threadpool(_listar_tiendas)
    except Exception as e:
        raise HTTPException(503, f"No se pudo leer la base de datos: {e}")
    return {"tiendas": filas}


@app.get("/api/expediente")
async def expediente(tienda: str, sku: str, desde: date, hasta: date,
                      umbral_osa: float = Query(100.0, gt=0, le=100)) -> dict:
    """Detalle diario de un SKU en una tienda: inventario, venta, pedidos y
    la causa raíz de cada día con faltante. A diferencia de /api/analizar*,
    no hace falta cola: un solo SKU es una consulta rápida."""
    try:
        return await run_in_threadpool(expediente_sku, tienda, sku, desde, hasta, umbral_osa)
    except Exception as e:
        raise HTTPException(503, f"No se pudo leer el expediente: {e}")


@app.post("/api/analizar", status_code=202)
async def analizar_archivo(
    archivos: List[UploadFile] = File(...),
    corregir: bool = Query(False, description="Aplicar la corrección automática de layout."),
    forzar: bool = Query(False, description="Analizar aunque queden errores que la "
                                            "corrección automática no puede arreglar."),
    umbral_osa: float = Query(100.0, gt=0, le=100,
                              description="Se analizan los días con OSA por debajo de este valor."),
) -> dict:
    """Encola el análisis y devuelve el id con el que se sigue.

    No analiza dentro del request: con volumen real la corrida tarda minutos y
    el navegador (o cualquier proxy de por medio) cortaría antes. El front hace
    poll a /api/analizar/{id}.

    Con `forzar` se analiza aun con errores pendientes. Sirve cuando lo que
    está roto es una hoja de la que sólo depende una parte del reporte —una
    extracción de citas incompleta afecta al scorecard del proveedor, no al
    Pareto— y el resultado se puede leer sabiendo eso. La validación viaja
    entera en la respuesta: se decide con los errores a la vista, no a ciegas.
    """
    id_, paquete = await _recibir(archivos)
    EJECUTOR.submit(_correr, id_, paquete, corregir, forzar, umbral_osa)
    return {"id": id_, "archivo": paquete.nombre_original, "estado": "en_proceso",
            "seguir_en": f"/api/analizar/{id_}"}


@app.get("/api/analizar/{id_}")
def estado_analisis(id_: str) -> dict:
    """Estado del análisis y, cuando termina, el mismo resumen de siempre."""
    trabajo = TRABAJOS.get(id_)
    if trabajo is None:
        raise HTTPException(404, "Ese análisis ya no está disponible. Hay que volver "
                                 "a subir el paquete.")

    if trabajo.estado == "en_proceso":
        return {"id": id_, "archivo": trabajo.archivo, "estado": "en_proceso",
                "segundos": round(time.time() - trabajo.creado, 1)}

    if trabajo.estado == "error":
        raise HTTPException(422, trabajo.error or "El análisis falló.")

    return {"id": id_, "estado": trabajo.estado, **(trabajo.resultado or {})}


@app.get("/api/resultado/{id_}")
def descargar(id_: str) -> FileResponse:
    """Entrega el Excel de resultados del análisis."""
    trabajo = TRABAJOS.get(id_)
    if trabajo is None or trabajo.salida is None or not trabajo.salida.exists():
        raise HTTPException(404, "Ese resultado ya no está disponible. Hay que volver "
                                 "a subir el paquete.")
    return FileResponse(trabajo.salida, media_type=XLSX,
                        filename=trabajo.nombre_descarga or "resultado.xlsx")
