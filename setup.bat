@echo off
echo ===============================================
echo   Access Table Exporter - Setup
echo ===============================================
echo.

echo Pruefe Python-Installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python ist nicht installiert!
    echo Bitte installieren Sie Python 3.10 oder hoher von:
    echo https://www.python.org/downloads/
    echo.
    echo Stellen Sie sicher, dass Sie "Add Python to PATH" ankreuzen!
    pause
    exit /b 1
)

echo Python gefunden:
python --version

echo.
echo Erstelle virtuelle Umgebung...
if exist .venv (
    echo Virtuelle Umgebung existiert bereits. Losche alte Version...
    rmdir /s /q .venv
)
python -m venv .venv

echo.
echo Aktiviere virtuelle Umgebung...
call .venv\Scripts\activate.bat

echo.
echo Installiere erforderliche Pakete...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo ===============================================
echo   Setup erfolgreich abgeschlossen!
echo ===============================================
echo.
echo Zum Starten des Exporters verwenden Sie:
echo   start_exporter.bat
echo.
pause
