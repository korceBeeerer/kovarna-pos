@echo off
echo ====================================
echo   KOVARENSKA POKLADNA - START
echo ====================================
echo.

cd backend

echo Kontroluji Python...
python --version
if errorlevel 1 (
    echo CHYBA: Python neni nainstalovan!
    echo Nainstalujte Python z https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo Kontroluji zavislosti...
pip show flask >nul 2>&1
if errorlevel 1 (
    echo Instaluji zavislosti...
    pip install -r requirements.txt
)

echo.
echo Spoustim server...
echo.
python server.py

pause
