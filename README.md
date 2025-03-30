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
git clone https://github.com/yourusername/document-slicer.git
cd document-slicer
```

2. Virtuelle Umgebung erstellen und aktivieren:
```bash
python -m venv .venv
source .venv/bin/activate  # Unter Windows: .venv\Scripts\activate
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
WATCH_DIRECTORY = "/path/to/your/directory"
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

## Lizenz

MIT License - siehe [LICENSE](LICENSE) Datei für Details.

## Beitragen

Beiträge sind willkommen! Bitte:
1. Forken Sie das Repository
2. Erstellen Sie einen Feature-Branch
3. Committen Sie Ihre Änderungen
4. Pushen Sie den Branch
5. Erstellen Sie einen Pull Request 