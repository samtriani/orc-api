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

# El commit con el que se construyó la imagen. Dentro del contenedor no hay
# repo git que preguntar, así que si esto no llega, cada corrida queda sellada
# como "desconocida" —o peor, con un sha viejo si alguien lo puso a mano una
# vez y nadie lo volvió a tocar—. Y ese sello es lo que permite explicar por
# qué dos corridas del mismo periodo no cuadran: las reglas cambian.
#
# Lo inyecta scripts/deploy.sh; no desplegar a mano sin él.
ARG ORCMM_VERSION=desconocida
ENV ORCMM_VERSION=$ORCMM_VERSION

# El proceso no corre como root: si algún día se sube un .xlsx malicioso, que
# tenga los menos permisos posibles.
RUN useradd --create-home --uid 1000 orcmm && chown -R orcmm:orcmm /app
USER orcmm

EXPOSE 8080

CMD ["python", "-m", "uvicorn", "api.main:app", "--host", "0.0.0.0", "--port", "8080"]
