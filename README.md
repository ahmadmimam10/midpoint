# MidPoint

Offline-Datenbanken in Excel: Tabellen und Eigenschaften wie in Notion zusammenstellen, Daten komfortabel pflegen, in Tabelle, Kalender, Board und Projektleiste bearbeiten, mit Dashboards auswerten, mit **Erfassungsdateien** Daten von vielen Stellen einsammeln und mit **Liste übernehmen** vorhandene Excel- und CSV-Listen sowie 2ask-Befragungen einlesen.

- läuft vollständig in Excel (Windows, Microsoft 365 / Excel 2016+)
- keine Cloud, keine Add-ins, keine fremden Bibliotheken
- erzeugte Dateien (Exporte, Erfassungsdateien) sind normale Excel-Arbeitsmappen ohne Makros

## Vorstellung

https://github.com/user-attachments/assets/bf8ff6df-4880-4c00-a739-dec8a059ede6

**Präsentation als PDF (36 Folien):** [doku/MidPoint_Vorstellung.pdf](doku/MidPoint_Vorstellung.pdf): typische Probleme mit Excel-Listen, die Lösung mit MidPoint, ein Praxisbeispiel von vorne bis hinten (Grünflächenamt: Spielplätze) und die Fakten.

**Lernreihe:** 17 kurze Videos bauen Schritt für Schritt eine eigene Datei, von der leeren Masterdatei über Board, Kalender und Zeitleiste bis zum Dashboard und zu Erfassungsdateien, dazu Baupläne und Startdateien zum Mitbauen. Im Lieferpaket unter `Lernen/Lernportal.html`.

## Aufbau des Repositorys

| Ordner | Inhalt |
|---|---|
| `src/modules`, `src/classes`, `src/forms` | gesamter VBA-Quellcode (Text, UTF-8) |
| `src/res` | eingebettete Bilder (MidPoint-Logo) |
| `build/` | PowerShell-Skripte: Build, Selbsttest, Leistungstest, Lint, Paket, Doku-Export, Sicherheitsprüfung, Testlisten und Testlauf für „Liste übernehmen“ |
| `doku/` | Liesmich, Hilfe, Kurzanleitung, Technische Dokumentation, Testmatrix, Versionshinweise, Vorstellung (PDF), Konzept „Liste übernehmen“ |
| `testdaten/importlisten` | Beschreibung der Testlisten und echte 2ask-Exporte; die übrigen Testlisten erzeugt `build\testlisten.ps1` |
| `Logos/` | Original-Logodateien |

Die Excel-Dateien selbst (`dist/`, Lieferpaket `paket/`) liegen nicht im Repository: Sie entstehen aus den Quellen neu.

## Bauen

Voraussetzung (nur auf dem Entwicklungsrechner): Excel Desktop mit der Option „Zugriff auf das VBA-Projektobjektmodell vertrauen“.

```powershell
& .\build\build.ps1 -Target master    # leere Datei  -> dist\MidPoint_Master.xlsm
& .\build\build.ps1 -Target demo      # mit Beispieldaten -> dist\MidPoint_Demo.xlsm
& .\build\selftest.ps1 -NoBuild       # automatische Prüfungen
& .\build\package.ps1                 # Lieferpaket (ZIP) in paket\
```

## Versionen

Jede Version ist in `doku/Versionshinweise.md` beschrieben. Das Lieferpaket (ZIP mit Masterdatei, Demodatei, Lernreihe, Dokumentation und Prüfberichten) liegt jeweils als Release bei GitHub.

## Lizenz

MIT, siehe [LICENSE](LICENSE). MidPoint ist ein eigenständiges Projekt und kein Produkt von Notion oder Microsoft.
