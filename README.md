# Document Slicer

Ein Python-Tool zur automatischen Verarbeitung und Aufteilung großer Dokumente. Das Tool überwacht ein Verzeichnis auf neue Dokumente und verarbeitet diese basierend auf Größe und Zeichenanzahl.

## Features

- Überwachung eines Verzeichnisses auf neue Dateien
- Unterstützung für DOCX, PPTX und PDF Dateien
- Automatische Konvertierung von DOCX zu PDF
- Intelligentes Slicing von PDFs basierend auf:
  - Dateigröße (> 15 MB)
  - Zeichenanzahl (> 60.000 Zeichen)
- Automatisches Löschen der Original-PDFs nach dem Slicing
- Detailliertes Logging aller Operationen

## Installation

1. Repository klonen:
```bash
git clone https://github.com/mariusschermann/slicer.git
cd slicer
```

2. Virtuelle Umgebung erstellen und aktivieren:
```bash
# Unter macOS/Linux:
python -m venv .venv
source .venv/bin/activate

# Unter Windows:
python -m venv .venv
.venv\Scripts\activate
```

3. Abhängigkeiten installieren:
```bash
pip install -r requirements.txt
```

## Konfiguration

Die wichtigsten Einstellungen können in der `document_processor.py` angepasst werden:

```python
SIZE_THRESHOLD = 15 * 1024 * 1024  # 15 MB in bytes
CHAR_THRESHOLD = 60000
WATCH_DIRECTORY = "/path/to/your/directory"  # Hier Ihren Pfad eintragen
```

## Verwendung

1. Starten Sie das Script:
```bash
python document_processor.py
```

2. Das Script wird:
   - Alle vorhandenen Dateien im überwachten Verzeichnis verarbeiten
   - Neue Dateien automatisch verarbeiten, sobald sie hinzugefügt werden
   - Die Verarbeitung fortsetzen, bis Sie das Script mit Ctrl+C beenden

## Ausgabe

- Verarbeitete PDFs werden in Teile aufgeteilt und mit Seitenzahlen benannt (z.B. `dokument_1-5.pdf`)
- Alle Operationen werden in der Konsole protokolliert
- Fehler und Warnungen werden entsprechend gekennzeichnet

## Anforderungen

- Python 3.6 oder höher
- Microsoft Word (für DOCX zu PDF Konvertierung)
- Die in `requirements.txt` aufgelisteten Python-Pakete

## Fehlerbehebung

### Häufige Probleme

1. **PDF-Konvertierung funktioniert nicht**
   - Stellen Sie sicher, dass Microsoft Word installiert ist
   - Überprüfen Sie die Berechtigungen des überwachten Verzeichnisses

2. **Script startet nicht**
   - Überprüfen Sie, ob die virtuelle Umgebung aktiviert ist
   - Stellen Sie sicher, dass alle Abhängigkeiten installiert sind

3. **Keine Dateien werden verarbeitet**
   - Überprüfen Sie den `WATCH_DIRECTORY` Pfad
   - Stellen Sie sicher, dass die Dateien die Schwellenwerte überschreiten

## Lizenz

MIT License - siehe [LICENSE](LICENSE) Datei für Details.

## Beitragen

Beiträge sind willkommen! Bitte:
1. Forken Sie das Repository
2. Erstellen Sie einen Feature-Branch (`git checkout -b feature/AmazingFeature`)
3. Committen Sie Ihre Änderungen (`git commit -m 'Add some AmazingFeature'`)
4. Pushen Sie den Branch (`git push origin feature/AmazingFeature`)
5. Erstellen Sie einen Pull Request

## Support

Bei Fragen oder Problemen:
1. Öffnen Sie ein Issue im GitHub Repository
2. Beschreiben Sie das Problem detailliert
3. Fügen Sie ggf. Logs oder Screenshots hinzu 