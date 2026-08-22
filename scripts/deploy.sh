#!/usr/bin/env bash
#
# Despliegue de orc-api a Fly.
#
# Existe por una sola razón: sellar la imagen con el commit que la produjo.
# Dentro del contenedor no hay repo git, así que `orcmm_runs.version_motor()`
# depende de ORCMM_VERSION; sin él cada corrida queda marcada "desconocida", y
# ese sello es lo que permite explicar por qué dos corridas del mismo periodo
# no cuadran — entre el 18 y el 22 de agosto las reglas cambiaron cinco veces.
#
# Un `flyctl deploy` a pelo compila igual y no avisa de nada. Por eso el
# despliegue pasa por aquí.
set -euo pipefail

cd "$(dirname "$0")/.."

if [[ -n "$(git status --porcelain)" ]]; then
    echo "Hay cambios sin commitear. La imagen quedaría sellada con un commit"
    echo "que no es el que se está desplegando."
    exit 1
fi

SHA="$(git rev-parse --short HEAD)"
echo "Desplegando $SHA"

# --ha=false: UNA sola máquina. El índice de trabajos vive en memoria del
# proceso, así que con dos la descarga puede caer en la que no corrió el
# análisis. Ver el comentario de fly.toml.
flyctl deploy --ha=false --build-arg "ORCMM_VERSION=$SHA" "$@"

echo
echo "Verificando..."
curl -fsS "https://orc-api.fly.dev/api/salud" && echo
