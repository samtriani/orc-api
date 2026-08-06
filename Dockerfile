# ORCMM — backend
#
# Imagen chica a propósito: el servicio no hace más que leer un Excel de 40 KB,
# clasificarlo y escribir otro. No necesita compilador ni nada nativo.

FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# Las dependencias van antes que el código para que un cambio en las reglas
# no invalide la capa de instalación.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY orcmm_*.py ./
COPY api/ ./api/

# El proceso no corre como root: si algún día se sube un .xlsx malicioso, que
# tenga los menos permisos posibles.
RUN useradd --create-home --uid 1000 orcmm && chown -R orcmm:orcmm /app
USER orcmm

EXPOSE 8080

CMD ["python", "-m", "uvicorn", "api.main:app", "--host", "0.0.0.0", "--port", "8080"]
