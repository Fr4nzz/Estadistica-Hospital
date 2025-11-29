@echo off
REM ============================================
REM Build script para crear el ejecutable
REM ============================================
REM 
REM Este script compila EstadisticaHospital.py en un
REM ejecutable portable (.exe) usando PyInstaller.
REM
REM El ejecutable resultante requiere que Chrome esté
REM instalado en la computadora destino.
REM
REM Tamaño aproximado del .exe: 50-80 MB
REM ============================================

echo.
echo ============================================
echo  Compilando Estadistica Hospital v3.0
echo ============================================
echo.

REM Verificar Python
python --version >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python no esta instalado o no esta en PATH
    echo         Descargue Python desde: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Paso 1: Instalar dependencias
echo [1/3] Instalando dependencias de Python...
pip install pandas openpyxl playwright pyinstaller --quiet --upgrade

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Fallo la instalacion de dependencias
    pause
    exit /b 1
)

REM Paso 2: Instalar Playwright (solo el driver, no los navegadores)
echo [2/3] Configurando Playwright...
playwright install-deps >nul 2>nul

REM Paso 3: Compilar con PyInstaller
echo [3/3] Compilando ejecutable (esto puede tomar unos minutos)...

pyinstaller --onefile ^
    --windowed ^
    --name "EstadisticaHospital" ^
    --add-data "config.ini;." ^
    --add-data "config_examenes.json;." ^
    --hidden-import "playwright.sync_api" ^
    --hidden-import "playwright._impl" ^
    --collect-submodules "playwright" ^
    EstadisticaHospital.py

echo.
echo ============================================
if exist "dist\EstadisticaHospital.exe" (
    echo  [OK] Compilacion exitosa!
    echo.
    echo  Archivo: dist\EstadisticaHospital.exe
    echo  Tamano:  
    for %%A in ("dist\EstadisticaHospital.exe") do echo           %%~zA bytes
    echo.
    echo  IMPORTANTE: 
    echo  - El .exe requiere Chrome instalado en el destino
    echo  - Copie config.ini y config_examenes.json junto al .exe
    echo  - Cree la carpeta ExcelsDescargados junto al .exe
) else (
    echo  [ERROR] La compilacion fallo
    echo  Revise los mensajes de error arriba
)
echo ============================================
echo.
pause
