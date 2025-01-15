# DienstplanSplitter

Ein Tool zum Aufteilen von TopSecure Dienstplänen, die in einer einzelnen großen PDF liegen in einzelne PDF-Dateien pro Person.
(TopSecure ist ein Planungsprogramm der BYTE AG und deprecated)

## Installation & Ausführung
1. EXE-Datei herunterladen
2. **Wichtiger Hinweis zu Windows SmartScreen:**
   - Beim ersten Start erscheint möglicherweise eine Windows SmartScreen Warnung
   - Dies ist normal, da die Software nicht über ein teures Code-Signing-Zertifikat verfügt
   - Um die Software trotzdem auszuführen:
     1. Auf "Weitere Informationen" klicken
     2. Dann auf "Trotzdem ausführen" klicken
   - Die Warnung erscheint nur beim ersten Start
   - Der Quellcode ist öffentlich auf GitHub verfügbar und kann geprüft werden

## Features
- Benutzerfreundliche grafische Oberfläche
- Einfache Dateiauswahl per Klick
- Fortschrittsanzeige während der Verarbeitung
- Erkennt automatisch ein- und zweiseitige Dienstpläne
- Extrahiert Namen und Datum aus dem PDF
- Erstellt separate PDFs für jede Person
- Benennt Dateien automatisch
  - Die Konvention ist
    - Nachname_Vorname_YYYY_MM für Pläne, die nicht für den aktuellen Monat sind, und
    - Nachname_Vorname_YYYY_MM_DD für Pläne, die für den aktuellen Monat sind. So können Planänderungen im laufenden Monat nachvollzogen werden.  

## Systemanforderungen
- Windows 10 oder neuer
- Keine zusätzliche Software erforderlich

## Lizenz
Dieses Projekt steht unter der GNU General Public License v3.0 mit zusätzlichen Einschränkungen:
- Keine kommerzielle Nutzung ohne ausdrückliche Genehmigung
- Namensnennung erforderlich
- Speziell entwickelt für TopSecure Dienstpläne

Vollständige Lizenzdetails finden Sie in der [LICENSE](LICENSE) Datei.

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

## Herausgeber
R2 Brainworks B.V.
Amsterdam, Niederlande