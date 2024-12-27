# Dienstplan Splitter

Ein Tool zum Aufteilen von TopSecure Dienstplänen, die in einer einzelnen großen PDF liegen in einzelne PDF-Dateien pro Person.

## Installation

### Voraussetzungen
- Python 3.11 oder höher
- pip (Python Package Manager)

### Abhängigkeiten installieren
```bash
pip install numpy==1.26.4 PyPDF2==3.0.1 pdfplumber==0.10.3 pyinstaller
```

## App erstellen

### Für macOS
1. Terminal öffnen
2. In das Projektverzeichnis wechseln
3. Folgenden Befehl ausführen:
```bash
pyinstaller --onefile --windowed --icon=icon.png dienstplan_splitter_gui.py
```
4. Die fertige App befindet sich im `dist`-Ordner
5. Die App kann in den Programme-Ordner verschoben werden

### Für Windows
1. Kommandozeile (cmd) öffnen
2. In das Projektverzeichnis wechseln
3. Folgenden Befehl ausführen:
```bash
pyinstaller --onefile --windowed --icon=icon.png dienstplan_splitter_gui.py
```
4. Die fertige EXE-Datei befindet sich im `dist`-Ordner

## Verwendung
1. App starten
2. PDF-Datei auswählen
3. Zielverzeichnis auswählen
4. "Dienstplan aufteilen" klicken
5. Die einzelnen PDF-Dateien werden im gewählten Verzeichnis erstellt 
