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
REM Tamaño aproximado del .exe: 80-120 MB
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

REM Paso 2: Limpiar builds anteriores
echo [2/3] Limpiando builds anteriores...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec

REM Paso 3: Compilar con PyInstaller (excluyendo paquetes innecesarios)
echo [3/3] Compilando ejecutable (esto puede tomar 2-5 minutos)...
echo      Excluyendo librerias innecesarias para reducir tamano...

pyinstaller --onefile ^
    --windowed ^
    --name "EstadisticaHospital" ^
    --exclude-module torch ^
    --exclude-module torchvision ^
    --exclude-module torchaudio ^
    --exclude-module tensorflow ^
    --exclude-module keras ^
    --exclude-module scipy ^
    --exclude-module matplotlib ^
    --exclude-module IPython ^
    --exclude-module jupyter ^
    --exclude-module notebook ^
    --exclude-module PIL ^
    --exclude-module cv2 ^
    --exclude-module sklearn ^
    --exclude-module sympy ^
    --exclude-module bokeh ^
    --exclude-module plotly ^
    --exclude-module seaborn ^
    --exclude-module networkx ^
    --exclude-module pygments ^
    --exclude-module jedi ^
    --exclude-module parso ^
    --exclude-module zmq ^
    --exclude-module tornado ^
    --exclude-module nbformat ^
    --exclude-module nbconvert ^
    --exclude-module jsonschema ^
    --exclude-module lark ^
    --exclude-module triton ^
    --exclude-module fsspec ^
    --exclude-module win32com ^
    --exclude-module pythoncom ^
    --exclude-module pywintypes ^
    EstadisticaHospital.py

echo.
echo ============================================
if exist "dist\EstadisticaHospital.exe" (
    echo  [OK] Compilacion exitosa!
    echo.
    echo  Archivo: dist\EstadisticaHospital.exe
    for %%A in ("dist\EstadisticaHospital.exe") do (
        set size=%%~zA
        setlocal enabledelayedexpansion
        set /a sizeMB=!size! / 1048576
        echo  Tamano:  !sizeMB! MB
        endlocal
    )
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
