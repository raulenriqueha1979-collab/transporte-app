#!/usr/bin/env bash
#
# update.sh — Actualiza la app de transporte en PythonAnywhere (o local).
#
# Qué hace:
#   1. git pull de la última versión.
#   2. Activa el entorno virtual e instala dependencias.
#   3. Respalda la base SQLite con marca de tiempo.
#   4. Ejecuta la migración (crea tablas nuevas y agrega columnas nuevas sin borrar datos).
#   5. Toca el archivo WSGI para que PythonAnywhere recargue de inmediato.
#
# Uso:
#   bash update.sh
#
# Variables opcionales (exportar antes de ejecutar si tu ruta difiere):
#   PROJECT_DIR   Carpeta del proyecto (por defecto: la carpeta de este script).
#   VENV_DIR      Carpeta del virtualenv (por defecto: $PROJECT_DIR/venv).
#   WSGI_FILE     Ruta del archivo WSGI de PythonAnywhere.
#                 Ej: /var/www/tuusuario_pythonanywhere_com_wsgi.py
#   DB_PATH       Ruta de la base SQLite (por defecto: $PROJECT_DIR/database.db;
#                 en PythonAnywhere suele ser /home/MIGAB2026/database.db).
#
set -euo pipefail

# --- Rutas ------------------------------------------------------------------
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="${PROJECT_DIR:-$SCRIPT_DIR}"
VENV_DIR="${VENV_DIR:-$PROJECT_DIR/venv}"

cd "$PROJECT_DIR"
echo "==> Proyecto: $PROJECT_DIR"

# --- 1. Traer últimos cambios ----------------------------------------------
echo "==> git pull"
git pull --ff-only

# --- 2. Entorno virtual + dependencias -------------------------------------
if [ ! -d "$VENV_DIR" ]; then
    echo "==> Creando entorno virtual en $VENV_DIR"
    python3 -m venv "$VENV_DIR"
fi
# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"
echo "==> Instalando dependencias"
pip install --upgrade pip >/dev/null
pip install -r requirements.txt

# --- 3. Backup de la base de datos -----------------------------------------
# Si /home/MIGAB2026 existe (PythonAnywhere), la app usa esa carpeta.
if [ -d "/home/MIGAB2026" ]; then
    DEFAULT_DB="/home/MIGAB2026/database.db"
else
    DEFAULT_DB="$PROJECT_DIR/database.db"
fi
DB_PATH="${DB_PATH:-$DEFAULT_DB}"
if [ -f "$DB_PATH" ]; then
    BACKUP="${DB_PATH}.bak.$(date +%Y%m%d_%H%M%S)"
    cp "$DB_PATH" "$BACKUP"
    echo "==> Backup de la base: $BACKUP"
else
    echo "==> No hay base previa en $DB_PATH (se creará nueva)."
fi

# --- 4. Migración (tablas y columnas nuevas, sin borrar datos) -------------
echo "==> Migrando base de datos"
python migrate.py

# --- 5. Recargar la app en PythonAnywhere ----------------------------------
if [ -n "${WSGI_FILE:-}" ]; then
    if [ -f "$WSGI_FILE" ]; then
        touch "$WSGI_FILE"
        echo "==> WSGI tocado, la web se recargará: $WSGI_FILE"
    else
        echo "!! WSGI_FILE no encontrado: $WSGI_FILE (revisa la ruta)."
    fi
else
    echo "!! WSGI_FILE no definido. En PythonAnywhere expórtalo, por ejemplo:"
    echo "     export WSGI_FILE=/var/www/TUUSUARIO_pythonanywhere_com_wsgi.py"
    echo "   o pulsa 'Reload' en la pestaña Web del panel de PythonAnywhere."
fi

echo "==> Actualización completada."
