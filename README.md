# Access Table Exporter

Ein professionelles Python-Tool zum Export von Microsoft Access-Tabellen in verschiedene Formate mit Fortschrittsanzeige und vollständigem Logging.

## Features

✅ **Unterstützte Formate:**
- CSV (Comma-Separated Values)
- XLSX (Excel-Dateien)
- JSON (JavaScript Object Notation)
- PDF (Portable Document Format)

✅ **Benutzerfreundlichkeit:**
- GUI-Dialoge für Datei- und Ordnerauswahl
- Tqdm-Fortschrittsbalken in der Konsole
- Interaktive Tabellenauswahl
- Automatische Ordnerstruktur-Erstellung

✅ **Professionelle Features:**
- Rotating File Logging
- Standard-Verzeichnisse (input, export, logs)
- Chunked Processing für große Tabellen
- Umfassende Fehlerbehandlung
- Detaillierte Konsolenausgaben

## Systemanforderungen

- **Betriebssystem:** Windows 10/11
- **Python:** Version 3.10 oder höher
- **Microsoft Access Database Engine:** 
  - Entweder Microsoft Office mit Access installiert
  - Oder Microsoft Access Database Engine Redistributable

## Installation

### 1. Repository herunterladen
```bash
git clone <repository-url>
cd access_converter
```

### 2. Automatische Installation
Führen Sie das Setup-Skript aus:
```cmd
setup.bat
```

Das Skript:
- Prüft die Python-Installation
- Erstellt eine virtuelle Umgebung
- Installiert alle erforderlichen Pakete

### 3. Manuelle Installation (falls erforderlich)
```cmd
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Verwendung

### Schnellstart
```cmd
start_exporter.bat
```

### Manueller Start
```cmd
.venv\Scripts\activate
python access_table_exporter.py
```

### Arbeitsablauf
1. **Datei auswählen:** GUI-Dialog öffnet sich zur .accdb/.mdb-Dateiauswahl
2. **Zielordner wählen:** Exportverzeichnis auswählen (Standard: ./export)
3. **Tabellen auswählen:** Interaktive Auswahl der zu exportierenden Tabellen
4. **Format wählen:** CSV, XLSX, JSON oder PDF
5. **Export:** Fortschrittsbalken zeigt den Export-Status
6. **Abschluss:** Dateien werden im gewählten Verzeichnis gespeichert

## Verzeichnisstruktur

```
access_converter/
├── access_table_exporter.py    # Hauptanwendung
├── requirements.txt            # Python-Abhängigkeiten
├── setup.bat                  # Installationsskript
├── start_exporter.bat         # Start-Skript
├── README.md                  # Diese Datei
├── .venv/                     # Virtuelle Python-Umgebung
├── input/                     # Standard-Eingabeverzeichnis
├── export/                    # Standard-Exportverzeichnis
└── logs/                      # Log-Dateien
    ├── access_exporter.log
    └── access_exporter.log.1  # Rotierte Logs
```

## Konfiguration

### Standard-Verzeichnisse
- **Input:** `./input` - Für Access-Dateien
- **Export:** `./export` - Für exportierte Dateien  
- **Logs:** `./logs` - Für Log-Dateien

### Logging
- **Konsole:** INFO-Level und höher
- **Datei:** DEBUG-Level mit detaillierten Informationen
- **Rotation:** Automatisch bei 10MB, maximal 5 Dateien

### Chunked Processing
Große Tabellen werden automatisch in Blöcken von 1000 Zeilen verarbeitet für optimale Performance.

## Fehlerbehebung

### ODBC-Fehler
```
Microsoft Access Driver (*.mdb, *.accdb) not found
```
**Lösung:** Installieren Sie den Microsoft Access Database Engine:
- Für 64-bit Python: AccessDatabaseEngine_X64.exe
- Für 32-bit Python: AccessDatabaseEngine.exe

### Import-Fehler
```
ModuleNotFoundError: No module named 'xyz'
```
**Lösung:** 
```cmd
.venv\Scripts\activate
pip install -r requirements.txt
```

### Datei-Zugriffsfehler
```
PermissionError: [Errno 13] Permission denied
```
**Lösung:** 
- Schließen Sie die Access-Datei in Microsoft Access
- Prüfen Sie Schreibrechte im Zielverzeichnis
- Führen Sie als Administrator aus (falls erforderlich)

## Logging-Ausgabe

### Konsole
```
[2024-01-XX 10:30:15] INFO - Access Table Exporter gestartet
[2024-01-XX 10:30:16] INFO - Datei ausgewählt: beispiel.accdb
[2024-01-XX 10:30:17] INFO - Verfügbare Tabellen: 5
[2024-01-XX 10:30:20] INFO - Beginne Export von 3 Tabellen...
100%|████████████| 3/3 [00:05<00:00,  1.67s/table]
[2024-01-XX 10:30:25] INFO - Export erfolgreich abgeschlossen!
```

### Log-Datei (DEBUG)
```
2024-01-XX 10:30:15,123 - DEBUG - setup_dirs() - Erstelle Verzeichnis: input
2024-01-XX 10:30:15,124 - DEBUG - setup_dirs() - Erstelle Verzeichnis: export  
2024-01-XX 10:30:15,125 - DEBUG - setup_dirs() - Erstelle Verzeichnis: logs
2024-01-XX 10:30:16,200 - INFO - choose_file() - Datei ausgewählt: C:\Beispiel\data.accdb
2024-01-XX 10:30:17,150 - DEBUG - Verbindung zur Datenbank hergestellt
2024-01-XX 10:30:17,155 - INFO - get_table_list() - Gefundene Tabellen: ['Kunden', 'Artikel', 'Bestellungen']
```

## Unterstützte Access-Versionen

- Microsoft Access 97-2003 (.mdb)
- Microsoft Access 2007+ (.accdb)
- Alle ODBC-kompatiblen Access-Formate

## Performance

- **Kleine Tabellen** (< 1000 Zeilen): Sofortiger Export
- **Mittlere Tabellen** (1000-100.000 Zeilen): Chunked Processing
- **Große Tabellen** (> 100.000 Zeilen): Optimierte Batch-Verarbeitung

## Lizenz

Dieses Projekt steht unter der MIT-Lizenz. Siehe LICENSE-Datei für Details.

## Support

Bei Problemen oder Fragen:
1. Prüfen Sie die Fehlerbehebung in dieser README
2. Schauen Sie in die Log-Dateien unter `./logs/`
3. Erstellen Sie ein Issue im Repository

MIT Licence GCNG Software 2025