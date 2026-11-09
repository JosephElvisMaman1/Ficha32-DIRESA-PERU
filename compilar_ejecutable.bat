@echo off
setlocal
title Compilador de Instalador Standalone - Ficha 32
cd /d "%~dp0"

echo =====================================================================
echo   Ficha N. 32 - Indicador de Telemedicina (MINSA / DITEL)
echo   Generador de Aplicacion de Escritorio Standalone (.EXE)
echo =====================================================================
echo.

rem 1. Verificar PyInstaller
echo [1/3] Verificando instalacion de PyInstaller...
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller no detectado. Instalando PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo [ERROR] No se pudo instalar PyInstaller. Verifique su conexion.
        pause
        exit /b 1
    )
)
echo PyInstaller listo.
echo.

rem 2. Compilar aplicacion
echo [2/3] Compilando aplicacion Ficha 32 a ejecutable nativo de Windows...
echo (Este proceso toma 1-2 minutos la primera vez)...
pyinstaller --clean Ficha32_Telemedicina.spec

if errorlevel 1 (
    echo.
    echo [ERROR] Hubo un error durante la compilacion.
    pause
    exit /b 1
)

echo.
echo [3/3] Compilacion completada con exito!
echo.
echo =====================================================================
echo   El ejecutable portable ha sido creado en:
echo   dist\Ficha32_Telemedicina.exe
echo.
echo   Puede copiar ese archivo o crear un acceso directo en su Escritorio.
echo   Funciona de forma autonoma y 100%% offline.
echo =====================================================================
echo.
pause
