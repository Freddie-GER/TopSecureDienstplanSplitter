# Dienstplan Splitter (Virtual Printer Edition)

Ein Tool zum Aufteilen von TopSecure Dienstplänen durch einen virtuellen Drucker.
(TopSecure ist ein Planungsprogramm der BYTE AG und deprecated)

## Features
- Erscheint als virtueller Drucker im System
- Teilt Dienstpläne automatisch beim Drucken auf
- Keine manuelle PDF-Verarbeitung notwendig
- Extrahiert Namen und Datum aus dem PDF
- Erstellt separate PDFs für jede Person
- Benennt Dateien automatisch
  - Die Konvention ist
    - Nachname_Vorname_YYYY_MM für Pläne, die nicht für den aktuellen Monat sind, und
    - Nachname_Vorname_YYYY_MM_DD für Pläne, die für den aktuellen Monat sind. So können Planänderungen im laufenden Monat nachvollzogen werden.  

## Installation

### Voraussetzungen
- Windows 10/11
- Administrative Rechte für die Druckerinstallation
- Python 3.11 oder höher (für Entwicklung)

### Für Benutzer
1. Installer herunterladen
2. Als Administrator ausführen
3. Virtuellen Drucker "Dienstplan Splitter" im System auswählen
4. Zielverzeichnis für die aufgeteilten PDFs konfigurieren

### Für Entwickler
```bash
pip install -r requirements.txt
```

## Verwendung
1. In TopSecure den Dienstplan wie gewohnt zum Drucken auswählen
2. Als Drucker "Dienstplan Splitter" wählen
3. Drucken klicken
4. Die einzelnen PDF-Dateien werden automatisch im konfigurierten Verzeichnis erstellt

## Entwicklung
Dieses Projekt ist ein Fork des originalen [Dienstplan Splitter](https://github.com/Freddie-GER/TopSecureDienstplanSplitter), 
erweitert um die Funktionalität eines virtuellen Druckers.

### Technische Details
- Verwendet PDF Creator SDK für die virtuelle Druckerfunktionalität
- Basiert auf der bewährten Parsing-Logik des originalen Dienstplan Splitters
- Implementiert als Windows-Druckerservice

