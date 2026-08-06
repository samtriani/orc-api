"""
ORCMM - API del analizador de desabasto.

    uvicorn api.main:app --reload --port 8000     (desde la raíz del proyecto)

Un solo flujo, en tres pasos que el front puede encadenar:

    POST /api/validar    ¿el layout está bien? ¿se puede arreglar solo?
    POST /api/analizar   valida, opcionalmente corrige, clasifica y resume
    GET  /api/resultado/{id}   descarga el Excel de resultados

El servidor no guarda nada permanente: cada análisis vive en su carpeta
temporal y se borra sola pasado un rato.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Optional

from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from api.servicio import analizar, diagnosticar_layout

MAX_BYTES = 25 * 1024 * 1024      # el layout lleno pesa ~40 KB; 25 MB es holgado
VIDA_TRABAJO_S = 60 * 60          # una hora
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# De dónde se acepta al navegador. En local, el front de desarrollo; en
# despliegue, la variable de entorno, para no tener que redesplegar el back
# cada vez que el front cambia de dominio:
#   fly secrets set ORCMM_ORIGENES="https://orc-gui.vercel.app"
ORIGENES_LOCALES = "http://localhost:4200,http://127.0.0.1:4200"
ORIGENES = [o.strip() for o in os.getenv("ORCMM_ORIGENES", ORIGENES_LOCALES).split(",")
            if o.strip()]

app = FastAPI(title="ORCMM — Clasificación de desabasto por causa raíz",
              version="2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=ORIGENES,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Trabajos en curso
# ---------------------------------------------------------------------------

@dataclass
class Trabajo:
    carpeta: Path
    salida: Optional[Path] = None
    nombre_descarga: Optional[str] = None
    creado: float = field(default_factory=time.time)


TRABAJOS: Dict[str, Trabajo] = {}


def _limpiar_viejos() -> None:
    corte = time.time() - VIDA_TRABAJO_S
    for id_ in [k for k, t in TRABAJOS.items() if t.creado < corte]:
        shutil.rmtree(TRABAJOS.pop(id_).carpeta, ignore_errors=True)


async def _recibir(archivo: UploadFile) -> tuple[str, Path]:
    """Guarda el archivo subido en una carpeta propia y regresa su id.

    El nombre original sólo se usa para etiquetar la descarga; en disco se
    escribe con un nombre que fija el servidor, para que nada de lo que venga
    en el request decida una ruta.
    """
    if not (archivo.filename or "").lower().endswith(".xlsx"):
        raise HTTPException(400, "El archivo tiene que ser un .xlsx de captura ORCMM.")

    _limpiar_viejos()
    id_ = uuid.uuid4().hex
    carpeta = Path(tempfile.mkdtemp(prefix=f"orcmm_{id_}_"))
    destino = carpeta / "captura.xlsx"

    tamano = 0
    with destino.open("wb") as f:
        while trozo := await archivo.read(1024 * 1024):
            tamano += len(trozo)
            if tamano > MAX_BYTES:
                shutil.rmtree(carpeta, ignore_errors=True)
                raise HTTPException(413, f"El archivo pasa de {MAX_BYTES // 1024 // 1024} MB.")
            f.write(trozo)

    if tamano == 0:
        shutil.rmtree(carpeta, ignore_errors=True)
        raise HTTPException(400, "El archivo llegó vacío.")

    TRABAJOS[id_] = Trabajo(carpeta=carpeta)
    return id_, destino


def _nombre_original(archivo: UploadFile) -> str:
    return Path(archivo.filename or "captura.xlsx").name


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.get("/api/salud")
def salud() -> dict:
    return {"estado": "ok", "trabajos_en_memoria": len(TRABAJOS),
            "origenes_permitidos": ORIGENES}


@app.post("/api/validar")
async def validar(archivo: UploadFile = File(...)) -> dict:
    """Revisa el layout sin analizar. Dice además si se puede corregir solo."""
    id_, ruta = await _recibir(archivo)
    try:
        diagnostico = diagnosticar_layout(ruta, TRABAJOS[id_].carpeta)
    except Exception as e:
        shutil.rmtree(TRABAJOS.pop(id_).carpeta, ignore_errors=True)
        raise HTTPException(422, f"No se pudo leer el archivo como Excel: {e}")

    return {"id": id_, "archivo": _nombre_original(archivo), **diagnostico}


@app.post("/api/analizar")
async def analizar_archivo(
    archivo: UploadFile = File(...),
    corregir: bool = Query(False, description="Aplicar la corrección automática de layout."),
    umbral_osa: float = Query(100.0, gt=0, le=100,
                              description="Se analizan los días con OSA por debajo de este valor."),
) -> dict:
    """Valida, opcionalmente corrige, clasifica y devuelve el resumen.

    Si el layout trae errores y no se pidió corregir, no analiza: regresa el
    reporte para que el front ofrezca el arreglo. Analizar sobre un archivo con
    claves mal capturadas produce un Pareto que se ve bien y está mal.
    """
    id_, ruta = await _recibir(archivo)
    trabajo = TRABAJOS[id_]

    try:
        diagnostico = diagnosticar_layout(ruta, trabajo.carpeta)
    except Exception as e:
        shutil.rmtree(TRABAJOS.pop(id_).carpeta, ignore_errors=True)
        raise HTTPException(422, f"No se pudo leer el archivo como Excel: {e}")

    a_analizar = ruta
    correccion = None

    if not diagnostico["valido"]:
        if not corregir:
            return {"id": id_, "archivo": _nombre_original(archivo),
                    "estado": "bloqueado", **diagnostico}

        if not diagnostico["ruta_corregida"]:
            return {"id": id_, "archivo": _nombre_original(archivo),
                    "estado": "bloqueado", **diagnostico}

        a_analizar = Path(diagnostico["ruta_corregida"])
        correccion = {
            "cambios": diagnostico["cambios_propuestos"],
            "errores_que_siguen": diagnostico["errores_tras_correccion"],
        }
        # Los errores que la corrección no resuelve (hoja vacía, citas
        # incompletas) no bloquean: son datos que faltan, no layout roto, y el
        # modelo ya sabe reportarlos como cobertura perdida.

    base = Path(_nombre_original(archivo)).stem
    trabajo.salida = trabajo.carpeta / "resultado.xlsx"
    trabajo.nombre_descarga = f"Resultado RCA - {base}.xlsx"

    try:
        resumen = analizar(a_analizar, trabajo.salida, umbral_osa)
    except Exception as e:
        raise HTTPException(422, f"El análisis falló: {e}")

    if not resumen["hay_resultados"]:
        trabajo.salida = None
        return {"id": id_, "archivo": _nombre_original(archivo),
                "estado": "sin_datos", "validacion": diagnostico["validacion"],
                "correccion": correccion, **resumen}

    return {
        "id": id_,
        "archivo": _nombre_original(archivo),
        "estado": "ok",
        "validacion": diagnostico["validacion"],
        "correccion": correccion,
        "nombre_salida": trabajo.nombre_descarga,
        **resumen,
    }


@app.get("/api/resultado/{id_}")
def descargar(id_: str) -> FileResponse:
    """Entrega el Excel de resultados del análisis."""
    trabajo = TRABAJOS.get(id_)
    if trabajo is None or trabajo.salida is None or not trabajo.salida.exists():
        raise HTTPException(404, "Ese resultado ya no está disponible. Hay que volver "
                                 "a subir el archivo.")
    return FileResponse(trabajo.salida, media_type=XLSX,
                        filename=trabajo.nombre_descarga or "resultado.xlsx")
