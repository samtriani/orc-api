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
from concurrent.futures import Future, ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.responses import FileResponse
import orcmm_runs                                              # noqa: E402
from psycopg2.extras import RealDictCursor
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool

# En Fly.io DATABASE_URL llega por `fly secrets set` (variable de entorno
# real); en desarrollo local no existe ese secreto, así que se carga del
# .env si está — no hace nada si el archivo no existe.
load_dotenv()

from api.servicio import (analizar, diagnosticar_layout,             # noqa: E402
                          escribir_excel, evidencia_y_dictamen)
from orcmm_pipeline import universo_osa                              # noqa: E402
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

# El resumen de una corrida real pesa 5.74 MB sin comprimir y 481 KB
# comprimido: doce veces menos, medido sobre Coyoacán marzo. Iba en claro
# hasta ahora, y ese tamaño es justo el que nos costó el spinner infinito
# cuando la respuesta tardaba más que el intervalo del poll.
#
# minimum_size deja pasar sin tocar las respuestas chicas —el estado del
# análisis son 111 bytes— donde comprimir costaría más CPU que el ahorro.
app.add_middleware(GZipMiddleware, minimum_size=2048)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ORIGENES,
    allow_methods=["GET", "POST", "DELETE"],
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
    # Cuándo el ejecutor le dio turno. Mientras sea None el trabajo está en
    # cola, no corriendo: sin esta marca el contador de la pantalla sumaba la
    # espera al tiempo de análisis y decía "llevo 193 s trabajando" cuando la
    # verdad era "llevo 193 s formado".
    iniciado: Optional[float] = None
    # En qué fase va. Sin esto el poll dice "corriendo" durante minutos y
    # cuando algo falla no se sabe ni dónde murió.
    etapa: Optional[str] = None
    futuro: Optional[Future] = None
    cancelado: bool = False

    def transcurrido(self) -> float:
        """Segundos EN LA FASE actual: desde que le dieron turno si ya
        corre, o desde que se creó si sigue en cola. Es el mismo cálculo que
        hacía el poll en dos lados; aquí vive una vez."""
        return round(time.time() - (self.iniciado or self.creado), 1)


TRABAJOS: Dict[str, Trabajo] = {}
CANDADO = threading.Lock()

# Estados en los que el trabajo ya no va a cambiar solo.
TERMINALES = ("ok", "bloqueado", "sin_datos", "error", "cancelado")


def _limpiar_viejos() -> None:
    corte = time.time() - VIDA_TRABAJO_S
    with CANDADO:
        viejos = [k for k, t in TRABAJOS.items()
                  if t.creado < corte and t.estado != "en_proceso"]
        for id_ in viejos:
            shutil.rmtree(TRABAJOS.pop(id_).carpeta, ignore_errors=True)


def _activo() -> Optional[str]:
    """El id del análisis en vuelo, si hay uno.

    El ejecutor tiene un solo trabajador a propósito, así que encolar un
    segundo análisis no lo hace ir más rápido: lo pone a esperar detrás del
    primero mientras la pantalla cuenta segundos como si estuviera trabajando.
    Es mejor decir que ya hay uno corriendo y ofrecer seguirlo.
    """
    with CANDADO:
        for id_, t in TRABAJOS.items():
            if t.estado == "en_proceso" and not t.cancelado:
                return id_
    return None


def _rechazar_si_hay_activo() -> None:
    id_ = _activo()
    if id_ is None:
        return
    t = TRABAJOS[id_]
    raise HTTPException(409, {
        "mensaje": "Ya hay un análisis en curso. El servidor corre uno a la vez.",
        "id_activo": id_,
        "archivo": t.archivo,
        "segundos": t.transcurrido(),
    })


def _tomar_turno(id_: str) -> bool:
    """Marca el arranque real del trabajo. False si ya lo cancelaron mientras
    esperaba en la cola, en cuyo caso no hay que analizar nada."""
    trabajo = TRABAJOS.get(id_)
    if trabajo is None or trabajo.cancelado:
        return False
    trabajo.iniciado = time.time()
    return True


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
    if not _tomar_turno(id_):                 # lo cancelaron mientras hacia cola
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
    if not _tomar_turno(id_):                 # lo cancelaron mientras hacia cola
        return
    try:
        def marcar(etapa: str) -> None:
            trabajo.etapa = etapa

        fu = leer_fuentes_db(tienda, desde, hasta, umbral_osa, avisar=marcar)
        trabajo.nombre_descarga = f"Resultado RCA - tienda {tienda}.xlsx"
        salida = trabajo.carpeta / "resultado.xlsx"
        # SIN el Excel: la pantalla se pinta con esto y el archivo se escribe
        # después. Medido en Coyoacán, el Excel son ~287 s de los ~342 que
        # tardaba la corrida completa. Ver escribir_excel.
        resumen = analizar(None, salida, umbral_osa, fu=fu, avisar=marcar,
                           con_excel=False)

        # El front espera 'validacion'/'correccion' siempre presentes (mismas
        # claves que _correr ya manda) — aquí van vacías porque no hay
        # archivo que validar.
        validacion_vacia = {"errores": [], "faltan_datos": [], "advertencias": [], "ok": []}

        if not resumen["hay_resultados"]:
            trabajo.estado = "sin_datos"
            trabajo.resultado = {"archivo": trabajo.archivo, "validacion": validacion_vacia,
                                 "correccion": None, **resumen}
            return

        trabajo.estado = "ok"
        trabajo.resultado = {
            "archivo": trabajo.archivo,
            "validacion": validacion_vacia,
            "correccion": None,
            "nombre_salida": trabajo.nombre_descarga,
            **resumen,
        }

        # Al histórico. Va DESPUÉS de dejar el trabajo en 'ok': si guardar
        # falla, la corrida que costó minutos sigue estando disponible en
        # esta pantalla; lo único que se pierde es poder volver a consultarla.
        aviso = orcmm_runs.guardar(
            id_, tienda, desde, hasta, resumen, umbral_osa,
            segundos=trabajo.transcurrido(), origen="bd",
            # Sale de los días SANOS, que no llegan al detalle diario: sin
            # esto el Excel regenerado saldría sin el OSA por SKU.
            universo=universo_osa(fu, umbral_osa))
        if aviso:
            trabajo.resultado.setdefault("advertencias", []).append(aviso)

        # Y AHORA el Excel, con el resultado ya servido. La pantalla lleva
        # minutos pintada mientras esto corre; `trabajo.salida` se pone hasta
        # el final para que la descarga no entregue un archivo a medio
        # escribir. Si truena, el análisis sigue siendo válido: lo único que
        # se pierde es el botón de descarga, y se dice.
        try:
            # Una sola pasada de clasificación (~1 s sobre 44 mil días) para
            # las dos cosas: guardar el detalle y escribir el Excel.
            trabajo.etapa = "guardando el detalle diario"
            evidencias, diagnosticos = evidencia_y_dictamen(fu, umbral_osa)
            aviso_dias = orcmm_runs.guardar_dias(id_, evidencias, diagnosticos)
            if aviso_dias:
                trabajo.resultado.setdefault("advertencias", []).append(aviso_dias)

            trabajo.etapa = "generando el Excel de resultados"
            escribir_excel(fu, salida, umbral_osa,
                           evidencias=evidencias, diagnosticos=diagnosticos)
            trabajo.salida = salida
        except Exception as e:
            trabajo.resultado.setdefault("advertencias", []).append(
                f"El análisis terminó bien, pero no se pudo generar el Excel ({e}). "
                f"Las cifras de esta pantalla son válidas; sólo no hay archivo que "
                f"descargar.")
        finally:
            trabajo.etapa = None
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
    _rechazar_si_hay_activo()
    id_ = uuid.uuid4().hex
    carpeta = Path(tempfile.mkdtemp(prefix=f"orcmm_bd_{id_}_"))
    etiqueta = f"Tienda {cuerpo.tienda} · {cuerpo.desde} a {cuerpo.hasta}"
    with CANDADO:
        TRABAJOS[id_] = Trabajo(carpeta=carpeta, archivo=etiqueta)

    TRABAJOS[id_].futuro = EJECUTOR.submit(
        _correr_desde_bd, id_, cuerpo.tienda, cuerpo.desde, cuerpo.hasta, cuerpo.umbral_osa)
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


# ---------------------------------------------------------------------------
# Histórico de corridas
#
# Una corrida completa tarda ~5.7 minutos. Guardarla y volver a leerla es la
# diferencia entre esperar seis minutos y pintar la pantalla de inmediato.
# ---------------------------------------------------------------------------

@app.get("/api/runs")
async def runs(limite: int = Query(50, ge=1, le=200),
               tienda: Optional[str] = None) -> dict:
    """Las corridas ya hechas, lo más reciente primero.

    Sin el resumen: son ~5 MB cada uno y aquí sólo se pinta un renglón por
    corrida. El detalle se pide aparte.
    """
    try:
        filas = await run_in_threadpool(orcmm_runs.listar, limite, tienda)
    except Exception as e:
        raise HTTPException(503, f"No se pudo leer el histórico de corridas: {e}")
    return {"runs": filas}


@app.get("/api/runs/{id_}")
async def run(id_: str) -> dict:
    """El resumen guardado de una corrida, listo para pintar.

    Es el mismo cuerpo que /api/analizar/{id}/resumen, más un bloque
    `guardado` con la procedencia: cuándo se corrió y con qué versión del
    motor. La pantalla tiene que poder decir que esto salió del histórico y
    no hacerlo pasar por un análisis recién hecho — sobre todo porque las
    reglas cambian, y un resultado de hace un mes puede no coincidir con lo
    que daría el motor de hoy.
    """
    try:
        guardado = await run_in_threadpool(orcmm_runs.leer, id_)
    except Exception as e:
        raise HTTPException(503, f"No se pudo leer la corrida: {e}")
    if guardado is None:
        raise HTTPException(404, "Esa corrida no está en el histórico.")
    return guardado


@app.get("/api/runs/{id_}/excel")
async def excel_de_run(id_: str) -> FileResponse:
    """El Excel de una corrida guardada, generado al pedirlo.

    No vuelve a leer las fuentes ni a clasificar: se reconstruye desde
    `run_dias` y el resumen. Se ahorran los ~56 s de análisis y queda sólo el
    costo de escribir el archivo.

    Se escribe en un temporal por petición y se borra al terminar de
    enviarlo: son 16 MB, y guardarlos por corrida serían gigas en disco por
    algo que se pide de vez en cuando.
    """
    carpeta = Path(tempfile.mkdtemp(prefix="orcmm-xlsx-"))
    salida = carpeta / "resultado.xlsx"
    try:
        hubo = await run_in_threadpool(orcmm_runs.regenerar_excel, id_, salida)
    except Exception as e:
        shutil.rmtree(carpeta, ignore_errors=True)
        raise HTTPException(503, f"No se pudo generar el Excel de esa corrida: {e}")
    if not hubo:
        shutil.rmtree(carpeta, ignore_errors=True)
        raise HTTPException(404, "Esa corrida no tiene detalle guardado. Sólo se puede "
                                 "regenerar el Excel de las corridas hechas después de "
                                 "que se empezó a guardar el detalle diario.")
    return FileResponse(
        salida, media_type=XLSX, filename=f"Resultado RCA - {id_}.xlsx",
        background=BackgroundTask(shutil.rmtree, carpeta, ignore_errors=True),
    )


@app.delete("/api/runs/{id_}")
async def borrar_run(id_: str) -> dict:
    try:
        fue = await run_in_threadpool(orcmm_runs.borrar, id_)
    except Exception as e:
        raise HTTPException(503, f"No se pudo borrar la corrida: {e}")
    if not fue:
        raise HTTPException(404, "Esa corrida no está en el histórico.")
    return {"borrado": id_}


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
    _rechazar_si_hay_activo()
    id_, paquete = await _recibir(archivos)
    TRABAJOS[id_].futuro = EJECUTOR.submit(_correr, id_, paquete, corregir, forzar, umbral_osa)
    return {"id": id_, "archivo": paquete.nombre_original, "estado": "en_proceso",
            "seguir_en": f"/api/analizar/{id_}"}


@app.get("/api/analizar/{id_}")
def estado_analisis(id_: str) -> dict:
    """SÓLO el estado. Unos bytes, no el resumen.

    Antes esto devolvía el resumen completo en cada vuelta del poll, y el
    resumen del análisis real pesa ~1.5 MB: 496 filas de SKU-tienda, 829 de
    proveedor y ~16,500 citas falladas. Preguntar "¿ya acabaste?" cada 3
    segundos arrastraba ese megabyte y medio cada vez.

    Peor: la respuesta tardaba 3.7 s directo y 17-30 s pasando por el proxy de
    Vercel — siempre más que el intervalo del poll—, así que el front la
    cancelaba en cada vuelta y la pantalla se quedaba girando para siempre con
    el análisis ya terminado del otro lado.

    El resumen se pide UNA vez, al final, en /api/analizar/{id}/resumen.
    """
    trabajo = TRABAJOS.get(id_)
    if trabajo is None:
        raise HTTPException(404, "Ese análisis ya no está disponible. Hay que volver "
                                 "a subir el paquete.")

    if trabajo.estado == "error":
        raise HTTPException(422, trabajo.error or "El análisis falló.")

    cuerpo = {"id": id_, "archivo": trabajo.archivo, "estado": trabajo.estado}
    if trabajo.estado == "en_proceso":
        # 'en_cola' y 'corriendo' son cosas distintas y hay que decirlo: el
        # contador de un trabajo formado no mide trabajo, mide espera.
        en_cola = trabajo.iniciado is None
        cuerpo["fase"] = "en_cola" if en_cola else "corriendo"
        if trabajo.etapa and not en_cola:
            cuerpo["etapa"] = trabajo.etapa
        cuerpo["segundos"] = trabajo.transcurrido()
        if en_cola:
            cuerpo["delante"] = sum(
                1 for t in TRABAJOS.values()
                if t.estado == "en_proceso" and not t.cancelado and t.creado < trabajo.creado)
    return cuerpo


@app.delete("/api/analizar/{id_}")
def cancelar_analisis(id_: str) -> dict:
    """Cancela un análisis.

    Si todavía hacía cola, `Future.cancel()` lo saca antes de que arranque y el
    turno queda libre de inmediato. Si ya está corriendo no se puede matar el
    hilo sin arriesgar la máquina, así que se marca cancelado —el front deja de
    esperarlo y su resultado se descarta— pero el trabajador sigue ocupado
    hasta que esa corrida termine. Se dice tal cual en la respuesta para no
    prometer algo que no pasa.
    """
    trabajo = TRABAJOS.get(id_)
    if trabajo is None:
        raise HTTPException(404, "Ese análisis ya no está disponible.")

    if trabajo.estado in TERMINALES:
        return {"id": id_, "estado": trabajo.estado, "libero_el_turno": True,
                "detalle": "El análisis ya había terminado."}

    trabajo.cancelado = True
    trabajo.estado = "cancelado"
    salio_de_la_cola = bool(trabajo.futuro and trabajo.futuro.cancel())

    return {
        "id": id_,
        "estado": "cancelado",
        "libero_el_turno": salio_de_la_cola,
        "detalle": ("Se sacó de la cola antes de arrancar." if salio_de_la_cola else
                    "Ya estaba corriendo: se descarta su resultado, pero el "
                    "trabajador sigue ocupado hasta que esa corrida termine."),
    }


@app.get("/api/analizar/{id_}/resumen")
def resumen_analisis(id_: str) -> dict:
    """El resumen completo, una sola vez, cuando el análisis ya terminó.

    Sirve a los tres estados terminales: 'ok' trae el análisis, 'bloqueado'
    trae los errores de layout y 'sin_datos' el motivo. El front necesita el
    cuerpo en los tres casos, no sólo en el bueno.
    """
    trabajo = TRABAJOS.get(id_)
    if trabajo is None:
        raise HTTPException(404, "Ese análisis ya no está disponible. Hay que volver "
                                 "a subir el paquete.")

    if trabajo.estado == "en_proceso":
        raise HTTPException(409, "El análisis todavía está corriendo.")

    if trabajo.estado == "error":
        raise HTTPException(422, trabajo.error or "El análisis falló.")

    return {"id": id_, "estado": trabajo.estado, **(trabajo.resultado or {})}


@app.get("/api/resultado/{id_}")
def descargar(id_: str) -> FileResponse:
    """Entrega el Excel de resultados del análisis."""
    trabajo = TRABAJOS.get(id_)
    if trabajo is None:
        raise HTTPException(404, "Ese resultado ya no está disponible. Hay que volver "
                                 "a subir el paquete.")
    # El Excel se escribe DESPUÉS de servir el resultado, así que hay una
    # ventana de varios minutos en la que el análisis ya está en pantalla y el
    # archivo todavía no existe. Es un estado legítimo, no un error: 409 y se
    # dice, en vez de un 404 que suena a "se perdió".
    if trabajo.salida is None or not trabajo.salida.exists():
        if trabajo.estado == "ok":
            raise HTTPException(409, "El Excel todavía se está generando. Son unos "
                                     "minutos: las cifras de la pantalla ya son las "
                                     "definitivas, el archivo va detrás.")
        raise HTTPException(404, "Ese resultado ya no está disponible. Hay que volver "
                                 "a subir el paquete.")
    return FileResponse(trabajo.salida, media_type=XLSX,
                        filename=trabajo.nombre_descarga or "resultado.xlsx")
