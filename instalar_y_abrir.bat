@echo off
setlocal enabledelayedexpansion
title MetaTag v8.9 - Lanzador Profesional

set PROGRAMA=src\metatag_v8.py
set VERSION=v8.9

echo.
echo  ============================================================
echo     MetaTag %VERSION% - Escritor de Metadatos Arqueologicos
echo  ============================================================
echo.

:: ── Auto-organizar: si no existe src/, mover .py a src/ ──
if not exist "%~dp0src" (
    if exist "%~dp0metatag_v8.py" (
        mkdir "%~dp0src"
        echo  [INFO] Organizando archivos en carpeta src/...
        move "%~dp0metatag_v8.py"      "%~dp0src\"
        move "%~dp0metatag_graficas.py" "%~dp0src\"
        move "%~dp0metatag_widgets.py"  "%~dp0src\"
        move "%~dp0metatag_writer.py"   "%~dp0src\"
        move "%~dp0Visor.py"            "%~dp0src\"
        move "%~dp0editor_casillas_backup.py" "%~dp0src\" 2>nul
        move "%~dp0renombrar_fotos_gui.py"    "%~dp0src\image_sync.py" 2>nul
        echo  [OK] Archivos movidos a src/.
    )
)

:: ── Crear carpeta data si no existe ──
if not exist "%~dp0data" (
    mkdir "%~dp0data"
    echo  [OK] Carpeta data creada.
)

:: ── Buscar Python: py launcher, python, pythonw ──
set PYTHON_CMD=
py --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON_CMD=py
) else (
    python --version >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        set PYTHON_CMD=python
    ) else (
        echo  [CRITICO] Python no esta instalado o no se encuentra en el PATH.
        echo.
        echo  INSTRUCCIONES PARA EL ARQUEOLOGO:
        echo  1. Ve a https://www.python.org/downloads/
        echo  2. Descarga e instala la ultima version.
        echo  3. IMPORTANTE: Marca la casilla "Add Python to PATH" al instalar.
        echo.
        pause
        exit /b 1
    )
)

if not exist "%~dp0%PROGRAMA%" (
    echo  [ERROR] No se encontro el archivo %PROGRAMA%.
    echo  Asegurate de que este .bat este en la misma carpeta que el codigo.
    pause
    exit /b 1
)

echo  Verificando librerias necesarias...
set LIBRERIAS=pandas openpyxl pillow piexif matplotlib numpy reportlab customtkinter

for %%L in (%LIBRERIAS%) do (
    set IMPORT_NAME=%%L
    if /I "%%L"=="pillow" set IMPORT_NAME=PIL
    %PYTHON_CMD% -c "import !IMPORT_NAME!" >nul 2>&1
    if !errorlevel! neq 0 (
        echo  [INFO] Instalando %%L...
        %PYTHON_CMD% -m pip install %%L --quiet
        if !errorlevel! neq 0 (
            echo  [ERROR] No se pudo instalar %%L. Revisa tu conexion a internet.
            pause
            exit /b 1
        )
        echo  [OK] %%L instalada con exito.
    ) else (
        echo  [OK] %%L ya se encuentra disponible.
    )
)

echo.
echo  Todo listo. Abriendo MetaTag %VERSION%...
echo  Esta ventana se cerrara sola en un instante.
echo.

:: ── Lanzar con pythonw (sin consola) o python como fallback ──
where pythonw >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    start "" pythonw "%~dp0%PROGRAMA%"
) else (
    start "" %PYTHON_CMD% "%~dp0%PROGRAMA%"
)

exit /b 0
