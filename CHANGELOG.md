# Changelog

Alle wichtigen Änderungen an diesem Projekt werden in dieser Datei dokumentiert.

Das Format basiert auf [Keep a Changelog](https://keepachangelog.com/de/1.0.0/),
und dieses Projekt folgt der [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-08-07

### Hinzugefügt
- **Erste stabile Version** des Access Table Exporters
- **Multi-Format-Export**: CSV, XLSX, JSON, PDF-Unterstützung
- **GUI-Integration**: Tkinter-Dialoge für Datei- und Ordnerauswahl
- **Fortschrittsanzeige**: Tqdm-Progressbars für visuelles Feedback
- **Professionelles Logging**: Rotating file handler mit DEBUG/INFO-Levels
- **Standard-Verzeichnisse**: Automatische input/export/logs-Struktur
- **Chunked Processing**: Optimierte Verarbeitung großer Tabellen
- **Interaktive Auswahl**: Flexible Tabellenauswahl (Nummern, Namen, Bereiche)
- **Detaillierte Zusammenfassungen**: Export-Statistiken und Dateimetadaten
- **Umfassende Fehlerbehandlung**: Robuste ODBC-Verbindungen
- **Automatische Installation**: setup.bat für einfache Einrichtung
- **Cross-Platform-Support**: Windows-optimiert mit ODBC-Treibern

### Technische Details
- **Python 3.10+** Kompatibilität
- **ODBC-basiert** für maximale Access-Kompatibilität (.mdb/.accdb)
- **Pandas-Integration** für effiziente Datenverarbeitung
- **Virtuelle Umgebung** für isolierte Abhängigkeiten
- **Professionelle Code-Struktur** mit Klassen und Modulen

### Performance
- Erfolgreich getestet mit **67.655 Datensätzen** in 6.41 Sekunden
- Unterstützt Tabellen mit **34.550+ Zeilen**
- Durchschnittliche Verarbeitungszeit: **0.64s pro Tabelle**

### Dokumentation
- Vollständige **README.md** mit Installationsanleitung
- **Fehlerbehebung** und häufige Probleme
- **Beispiele** und Verwendungshinweise
- **MIT-Lizenz** für freie Nutzung
