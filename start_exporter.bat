@echo off
echo ===============================================
echo   Access Table Exporter
echo ===============================================
echo.

if not exist .venv (
    echo ERROR: Virtuelle Umgebung nicht gefunden!
    echo Bitte fueren Sie zuerst setup.bat aus.
    pause
    exit /b 1
)

echo Aktiviere virtuelle Umgebung...
call .venv\Scripts\activate.bat

echo Starte Access Table Exporter...
echo.
python access_table_exporter.py

echo.
echo Exporter beendet.
pause
