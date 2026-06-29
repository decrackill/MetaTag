@echo off
setlocal enabledelayedexpansion
title MetaTag v8.6 - Lanzador Profesional

:: Configuración de nombres (Asegúrate de que el nombre coincida con tu archivo de código)
set PROGRAMA=MetaTag_v8.py
set VERSION=v8.6

echo.
echo  ============================================================
echo     MetaTag %VERSION% - Escritor de Metadatos Arqueologicos
echo  ============================================================
echo.

:: 1. Verificar Python (Sin redirecciones molestas)
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
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

:: 2. Verificar archivo principal
if not exist "%~dp0%PROGRAMA%" (
    echo  [ERROR] No se encontro el archivo %PROGRAMA%.
    echo  Asegurate de que este .bat este en la misma carpeta que el codigo.
    pause
    exit /b 1
)

:: 3. Verificar y actualizar dependencias (Una por una con control de error)
echo  Verificando librerias necesarias...
set LIBRERIAS=pandas openpyxl pillow piexif matplotlib

for %%L in (%LIBRERIAS%) do (
    python -c "import %%L" >nul 2>&1
    if !errorlevel! neq 0 (
        echo  [INFO] Instalando %%L...
        python -m pip install %%L --quiet
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

:: 4. Ejecucion del programa en segundo plano (Magia para ocultar consola)
start "" pythonw "%~dp0%PROGRAMA%"

:: 5. Cerrar la terminal inmediatamente
exit