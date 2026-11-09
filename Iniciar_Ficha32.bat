@echo off
setlocal enabledelayedexpansion
title Ficha 32 - Telemedicina (MINSA / DITEL)
cd /d "%~dp0"

echo =====================================================================
echo   Ficha N. 32 - Indicador de Telemedicina (MINSA / DITEL)
echo   Iniciando interfaz grafica de escritorio...
echo =====================================================================
echo.

rem Buscar Python disponible en el sistema
set PYTHON_BIN=
where pythonw >nul 2>nul && set PYTHON_BIN=pythonw
if not defined PYTHON_BIN (
    where python >nul 2>nul && set PYTHON_BIN=python
)
if not defined PYTHON_BIN (
    where py >nul 2>nul && set PYTHON_BIN=py
)

if not defined PYTHON_BIN (
    echo [ERROR] No se encontro Python instalado en el sistema.
    echo Por favor instale Python 3.10 o superior descargandolo de:
    echo https://www.python.org/downloads/
    echo.
    echo Asegurese de marcar la casilla "Add Python to PATH" durante la instalacion.
    echo.
    pause
    exit /b 1
)

rem Verificar dependencias requeridas
echo Verificando librerias necesarias...
python -c "import tkinter, pandas, xlsxwriter, docx, requests, matplotlib" >nul 2>nul
if errorlevel 1 (
    echo.
    echo [AVISO] Se detectaron dependencias faltantes.
    echo Instalando librerias automaticamente (esto solo toma un momento la primera vez)...
    pip install pandas xlsxwriter python-docx requests matplotlib openpyxl
    if errorlevel 1 (
        echo [ERROR] No se pudieron instalar las dependencias con pip.
        pause
        exit /b 1
    )
    echo Dependencias instaladas con exito.
    echo.
)

rem Iniciar la aplicacion de escritorio
echo Abriendo aplicacion Ficha 32...
start %PYTHON_BIN% app_ficha32.py

exit /b 0
