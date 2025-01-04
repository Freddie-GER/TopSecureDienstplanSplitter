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

### Installation des virtuellen Druckers
1. Installer von der [Releases-Seite](https://github.com/Freddie-GER/TopSecureDienstplanSplitter/releases) herunterladen
   - Die Datei heißt `dienstplan_splitter_printer_setup.exe`
2. Installer als Administrator ausführen
   - Rechtsklick auf die .exe → "Als Administrator ausführen"
3. Der virtuelle Drucker "Dienstplan Splitter" wird automatisch installiert
4. Das Ausgabeverzeichnis wird standardmäßig unter "Dokumente/Dienstplan Splitter" erstellt

### Überprüfen der Installation
1. Öffnen Sie die Windows Einstellungen
2. Gehen Sie zu "Drucker & Scanner"
3. Der "Dienstplan Splitter" sollte in der Liste erscheinen

## Verwendung
1. In TopSecure den Dienstplan wie gewohnt zum Drucken auswählen
2. Als Drucker "Dienstplan Splitter" wählen
3. Drucken klicken
4. Die einzelnen PDF-Dateien werden automatisch im konfigurierten Verzeichnis erstellt

### Ausgabeverzeichnis
Die aufgeteilten PDFs finden Sie unter:
- `C:\Users\IHR_BENUTZERNAME\Documents\Dienstplan Splitter`

### Deinstallation
1. Installer erneut als Administrator ausführen
2. `uninstall` auswählen
3. Der virtuelle Drucker wird vollständig entfernt

## Entwicklung
Dieses Projekt ist ein Fork des originalen [Dienstplan Splitter](https://github.com/Freddie-GER/TopSecureDienstplanSplitter), 
erweitert um die Funktionalität eines virtuellen Druckers.

### Technische Details
- Verwendet Windows Printer API für die virtuelle Druckerfunktionalität
- Basiert auf der bewährten Parsing-Logik des originalen Dienstplan Splitters
- Implementiert als Windows-Druckerservice

