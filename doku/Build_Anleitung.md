# Build-Anleitung (nur Entwicklung)

Die ausgelieferten Dateien `MidPoint_Master.xlsm` und `MidPoint_Demo.xlsm` werden reproduzierbar aus den Textquellen in `src` erzeugt. **Anwenderrechner brauchen nichts davon**: Sie öffnen nur die fertige XLSM.

## Voraussetzungen (Entwicklungsrechner)
- Windows 10/11, Microsoft Excel 365 Desktop (32 oder 64 Bit), Windows PowerShell 5.1.
- In Excel: *Datei › Optionen › Trust Center › Einstellungen › Makroeinstellungen › „Zugriff auf das VBA-Projektobjektmodell vertrauen“*, nur für den Build, nicht für den Betrieb.
- Während des Builds keine andere Excel-Sitzung mit ungespeicherten Änderungen offen lassen.

## Befehle (im Ordner `build`)
```
powershell -ExecutionPolicy Bypass -File build.ps1 -Target master   # leere Masterdatei
powershell -ExecutionPolicy Bypass -File build.ps1 -Target demo     # Demodatei mit fiktiven Daten
powershell -ExecutionPolicy Bypass -File build.ps1 -Target stress   # Lasttestdatei (20.000 DS, 40.000 Zuordnungen)
powershell -ExecutionPolicy Bypass -File build.ps1 -Update "C:\Pfad\Fachdatei.xlsm"   # Programmcode einer Fachdatei erneuern, Daten bleiben
```
Ergebnis liegt in `dist`. `-Update` erzeugt eine Kopie `<Name>_aktualisiert.xlsm`; die Originaldatei bleibt unverändert. Beim ersten Öffnen der Kopie ergänzt die Migration fehlende Spalten/Blätter (z. B. Format 2 → 3). Geprüft mit einer Demodatei 1.1.0.

## Prüfskripte
- `lint.ps1`: statische Prüfung der Quellen (reservierte Namen, Doppeldeklarationen, Zeichensatz Windows-1252, verschmolzene Zeilen, Deklarationen nach Prozeduren).
- `selftest.ps1`: baut die Demodatei, öffnet eine Kopie in `pruefung\selbsttest` und führt die automatisierten Abnahmeprüfungen aus (inkl. Sicherung/Wiederherstellung und Speichern/Schließen/Wiederöffnen). Ergebnis: `pruefung\selbsttest\Selbsttest_Ergebnis.txt`.
- `perf.ps1`: baut die Lasttestdatei und misst die Laufzeiten. Ergebnis: `pruefung\leistung\Leistung_Ergebnis.txt`.
- `exportdocs.ps1`: schreibt Hilfe und Technische Dokumentation aus der Masterdatei nach `doku\Hilfe.md` und `doku\Technische_Dokumentation.md`.
- `reload.ps1`, `runmacro.ps1`, `devopen.ps1`, `snap.ps1`: Entwicklungshilfen (Module in eine offene Datei neu laden, Makro ausführen, Datei mit Makros öffnen, Bildschirmaufnahme).

## Ablauf des Builds
1. Neue Arbeitsmappe in einer eigenen Excel-Instanz.
2. Module und Klassen aus `src\modules` und `src\classes` importieren, UserForms anlegen und Code aus `src\forms\*.frm.vba` einfügen, `src\ThisWorkbook.cls` übernehmen.
3. `mBuild.Build_Init`: Hinweisblatt, Oberflächen- und Konfigurationsblätter, Vorlagen, Hilfe, Technische Dokumentation.
4. Bei Demo/Lasttest: `mBuild.Build_Demo` bzw. `mBuild.Build_Stress`.
5. `mBuild.Build_CodeCatalog` (Quellcode in die Datei), `mBuild.Build_Finalize`, als XLSM speichern.

Die Datei enthält danach keine Verweise auf die Entwicklungsumgebung.
