@echo off
echo ========================================
echo   Building Estadistica Hospital v3.6
echo ========================================
echo.

REM Install dependencies
echo Installing dependencies...
pip install pyinstaller pandas openpyxl playwright tkcalendar babel --quiet

REM Build the executable
echo.
echo Building executable...
pyinstaller --noconfirm --onefile --windowed ^
    --name "EstadisticaHospital" ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=playwright ^
    --hidden-import=playwright.sync_api ^
    --hidden-import=tkcalendar ^
    --hidden-import=babel.numbers ^
    --collect-all tkcalendar ^
    --collect-all babel ^
    EstadisticaHospital.py

echo.
echo ========================================
echo   Build complete!
echo   Executable: dist\EstadisticaHospital.exe
echo.
echo   NOTE: The app creates config files
echo   automatically on first run.
echo ========================================
pause
