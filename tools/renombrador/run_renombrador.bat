@echo off
setlocal enabledelayedexpansion

set "SCRIPT=%~dp0renombrar_fotos_gui.py"
set "VENV=%~dp0.venv"
set "VENV_PY=%VENV%\Scripts\python.exe"

REM Verificar si existe Python
where py >nul 2>&1
if not errorlevel 1 (
    set "SYSTEM_PY=py"
) else (
    where python >nul 2>&1
    if not errorlevel 1 (
        set "SYSTEM_PY=python"
    ) else (
        echo No se encontro Python instalado.
        echo Descarga Python desde: https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

REM Crear entorno virtual si no existe
if not exist "%VENV_PY%" (
    echo Creando entorno virtual...
    %SYSTEM_PY% -m venv "%VENV%"
    if errorlevel 1 (
        echo Error creando entorno virtual.
        pause
        exit /b 1
    )
    echo Entorno virtual creado.
)

REM Instalar dependencias
echo Verificando dependencias...
"%VENV_PY%" -m pip install -r "%~dp0requirements.txt" --quiet
if errorlevel 1 (
    echo Error instalando dependencias.
    pause
    exit /b 1
)

REM Ejecutar app
echo Iniciando Renombrador de Fotos...
start "" "%VENV_PY%" "%SCRIPT%"
exit /b 0
