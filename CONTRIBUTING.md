# Beitrag zum Access Table Exporter

Vielen Dank für Ihr Interesse, zum Access Table Exporter beizutragen! 🎉

## Wie Sie beitragen können

### 🐛 Bugs melden
- Verwenden Sie die [GitHub Issues](../../issues)
- Beschreiben Sie das Problem detailliert
- Fügen Sie Log-Dateien aus `./logs/` hinzu
- Geben Sie Ihre Python- und Access-Version an

### 💡 Features vorschlagen
- Öffnen Sie ein [Feature Request Issue](../../issues)
- Beschreiben Sie den gewünschten Use Case
- Erklären Sie, warum das Feature nützlich wäre

### 🔧 Code beitragen

#### Entwicklungsumgebung einrichten
```cmd
git clone https://github.com/IhrUsername/access_converter.git
cd access_converter
setup.bat
```

#### Code-Style
- Verwenden Sie **deutsche Kommentare** für Konsistenz
- Folgen Sie PEP 8 Python-Konventionen
- Fügen Sie Docstrings für neue Funktionen hinzu
- Testen Sie mit verschiedenen Access-Dateien

#### Pull Request Prozess
1. Forken Sie das Repository
2. Erstellen Sie einen Feature-Branch (`git checkout -b feature/neue-funktion`)
3. Committen Sie Ihre Änderungen (`git commit -m 'Füge neue Funktion hinzu'`)
4. Pushen Sie zum Branch (`git push origin feature/neue-funktion`)
5. Öffnen Sie einen Pull Request

### 📋 Checkliste für Pull Requests
- [ ] Code folgt dem bestehenden Style
- [ ] Neue Features sind dokumentiert
- [ ] Tests wurden durchgeführt
- [ ] CHANGELOG.md wurde aktualisiert
- [ ] README.md wurde bei Bedarf angepasst

## Entwicklung

### Testen
```cmd
# Virtuelle Umgebung aktivieren
.venv\Scripts\activate

# Verschiedene Access-Dateien testen
python access_table_exporter.py
```

### Neue Exportformate hinzufügen
Erweitern Sie die `export_table()` Methode in `AccessTableExporter`:

```python
def export_table(self, table_name, format_type, output_path):
    # Bestehende Formate: CSV, XLSX, JSON, PDF
    # Neues Format hier hinzufügen
    if format_type.upper() == 'NEUES_FORMAT':
        # Implementierung hier
        pass
```

## Community

### Verhaltenskodex
- Seien Sie respektvoll und konstruktiv
- Helfen Sie anderen bei Problemen
- Teilen Sie Ihr Wissen und Ihre Erfahrungen

### Support
- 📧 Issues für Bugs und Features
- 💬 Diskussionen für allgemeine Fragen
- 📚 Wiki für erweiterte Dokumentation

Vielen Dank für Ihren Beitrag! 🚀
