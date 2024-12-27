# Dienstplan Splitter für Top Secure

Dieses Programm teilt einen PDF-Dienstplan in einzelne PDFs pro Person auf.

Die Konvention ist 
Nachname_Vorname_YYYY_MM für Pläne, die nicht für den aktuellen Monat sind, und 
Nachname_Vorname_YYYY_MM_DD für Pläne, die für den aktuellen Monat sind. So können Planänderungen im laufenden Monat nachvollzogen werden.  




## Features

- Benutzerfreundliche grafische Oberfläche
- Einfache Dateiauswahl per Klick
- Fortschrittsanzeige während der Verarbeitung
- Erkennt automatisch ein- und zweiseitige Dienstpläne
- Extrahiert Namen und Datum aus dem PDF
- Erstellt separate PDFs für jede Person
- Benennt Dateien automatisch 

## Installation

### Option 1: Ausführbare Datei (Empfohlen)
1. Laden Sie die ausführbare Datei für Ihr Betriebssystem herunter:
   - Windows: `DienstplanSplitter.exe`
   - Mac: `DienstplanSplitter.app`
2. Doppelklicken Sie auf die heruntergeladene Datei

### Option 2: Python-Installation
Falls Sie das Programm aus dem Quellcode ausführen möchten:
1. Installieren Sie Python 3.x
2. Installieren Sie die benötigten Pakete:
   ```
   pip install -r requirements.txt
   ```
3. Starten Sie das Programm:
   ```
   python dienstplan_splitter_gui.py
   ```

## Verwendung

1. Starten Sie das Programm per Doppelklick
2. Klicken Sie auf "PDF auswählen" und wählen Sie Ihre PDF-Datei aus
3. Klicken Sie auf "Dienstpläne aufteilen"
4. Die aufgeteilten PDFs werden im Ordner `split_schedules` gespeichert

## Executable erstellen

Um eine ausführbare Datei zu erstellen:

1. Installieren Sie die Entwicklungsabhängigkeiten:
   ```
   pip install -r requirements.txt
   ```

2. Führen Sie auto-py-to-exe aus:
   ```
   auto-py-to-exe
   ```

3. In auto-py-to-exe:
   - Wählen Sie `dienstplan_splitter_gui.py` als Script Location
   - Wählen Sie "One File" und "Window Based"
   - Klicken Sie auf "Convert" 
