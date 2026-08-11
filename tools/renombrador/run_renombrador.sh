#!/bin/bash
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV="$SCRIPT_DIR/.venv"
VENV_PY="$VENV/bin/python"

# Verificar Python
if ! command -v python3 &> /dev/null; then
    echo "No se encontro Python3."
    echo "Instala con: sudo apt install python3 python3-venv"
    exit 1
fi

# Crear entorno virtual si no existe
if [ ! -f "$VENV_PY" ]; then
    echo "Creando entorno virtual..."
    python3 -m venv "$VENV"
    if [ $? -ne 0 ]; then
        echo "Error creando entorno virtual."
        exit 1
    fi
    echo "Entorno virtual creado."
fi

# Instalar dependencias
echo "Verificando dependencias..."
"$VENV_PY" -m pip install -r "$SCRIPT_DIR/requirements.txt" --quiet
if [ $? -ne 0 ]; then
    echo "Error instalando dependencias."
    exit 1
fi

# Ejecutar app
echo "Iniciando Renombrador de Fotos..."
"$VENV_PY" "$SCRIPT_DIR/renombrar_fotos_gui.py" &
