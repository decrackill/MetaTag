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

# Ejecutar MetaTag en primer plano (la terminal se cierra cuando cierres la app)
"$PY" "$SCRIPT_DIR/metatag_v8.py"
EXIT_CODE=$?

if [ $EXIT_CODE -ne 0 ]; then
    echo ""
    echo "  [ERROR] MetaTag terminó con error (código $EXIT_CODE)."
    echo ""
fi
read -p "  Presiona Enter para cerrar..."
