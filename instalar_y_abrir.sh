#!/bin/bash
clear
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_DIR="$SCRIPT_DIR/.venv"
PY="$VENV_DIR/bin/python3"
PIP="$VENV_DIR/bin/pip"

echo " ============================================================"
echo "    MetaTag v8.9 - Instalador y Lanzador Linux"
echo " ============================================================"
echo ""

# Auto-organizar: si no existe src/, mover .py a src/
if [ ! -d "$SCRIPT_DIR/src" ]; then
    if [ -f "$SCRIPT_DIR/metatag_v8.py" ]; then
        mkdir -p "$SCRIPT_DIR/src"
        echo "  [INFO] Organizando archivos en carpeta src/..."
        mv "$SCRIPT_DIR"/metatag_v8.py       "$SCRIPT_DIR/src/"
        mv "$SCRIPT_DIR"/metatag_graficas.py  "$SCRIPT_DIR/src/"
        mv "$SCRIPT_DIR"/metatag_widgets.py   "$SCRIPT_DIR/src/"
        mv "$SCRIPT_DIR"/metatag_writer.py    "$SCRIPT_DIR/src/"
        mv "$SCRIPT_DIR"/Visor.py             "$SCRIPT_DIR/src/"
        mv "$SCRIPT_DIR"/editor_casillas_backup.py "$SCRIPT_DIR/src/" 2>/dev/null
        echo "  [OK] Archivos movidos a src/."
    fi
fi

# Crear carpeta data si no existe
if [ ! -d "$SCRIPT_DIR/data" ]; then
    mkdir -p "$SCRIPT_DIR/data"
    echo "  [OK] Carpeta data creada."
fi

# Verificar Python
if ! command -v python3 &> /dev/null; then
    echo "  [CRITICO] Python3 no esta instalado."
    echo "  Instalar con: sudo apt install python3 python3-venv python3-pip"
    exit 1
fi

# Crear venv si no existe
if [ ! -f "$PY" ]; then
    echo "  [INFO] Creando entorno virtual..."
    python3 -m venv "$VENV_DIR"
    if [ $? -ne 0 ]; then
        echo "  [ERROR] No se pudo crear el venv. Instalar: sudo apt install python3-venv"
        exit 1
    fi
    echo "  [OK] Entorno virtual creado."
fi

# Instalar dependencias
LIBRERIAS="pillow piexif reportlab pandas openpyxl numpy matplotlib"

echo "  Verificando librerias necesarias..."
for LIB in $LIBRERIAS; do
    $PY -c "import $LIB" 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "  [INFO] Instalando $LIB..."
        $PIP install "$LIB" --quiet
        if [ $? -ne 0 ]; then
            echo "  [ERROR] No se pudo instalar $LIB."
            exit 1
        fi
        echo "  [OK] $LIB instalada."
    else
        echo "  [OK] $LIB ya disponible."
    fi
done

echo ""
echo "  Todo listo. Abriendo MetaTag..."
echo ""

# Ejecutar MetaTag脱离 de la terminal y cerrarla
LOG_FILE="$SCRIPT_DIR/data/metatag_debug.log"
setsid "$PY" "$SCRIPT_DIR/src/metatag_v8.py" > "$LOG_FILE" 2>&1 &
disown
sleep 0.3
exit 0
