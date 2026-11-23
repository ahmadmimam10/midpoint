# Leistungsbericht für MidPoint 1.7.0

## Testumgebung
| | |
|---|---|
| Rechner | AMD Ryzen 5 2600X (6 Kerne), 32 GB RAM, SSD |
| Betriebssystem | Microsoft Windows 11 Home |
| Excel | Microsoft 365, Version 16.0 (Build 20326), **64 Bit** |
| Messmethode | VBA-Timer in echter Excel-Laufzeit (`mSelfTest.Measure` über `build\perf.ps1`; „Liste übernehmen“ mit `build\listimport_test.ps1`, Teilschritte mit `mListImport.LiProfile`). Die Datei wird vorher frisch geöffnet; gemessen werden der erste Aufruf („kalt“) und die Wiederholung („warm“). |
| Stand | November 2026: Version 1.7.0, Lasttestdatei neu gebaut, zwei Messläufe. Vergleich mit 1.5.0 (Oktober 2026). |

## Lastprofil (Projektauftrag Abschnitt 14)
Lasttestdatei `MidPoint_Lasttest.xlsm` (8,3 MB; in 1.5.0 waren es 7,2 MB, der Zuwachs ist Programmcode), reproduzierbar erzeugt mit `build.ps1 -Target stress`:
- **10 Fachtabellen** mit zusammen **20.000 Datensätzen**
- davon eine Tabelle mit **10.000 Datensätzen und 50 Feldern** (Text, Zahl, Geld, Datum, Auswahl, Status, Mehrfachauswahl, Ja/Nein, Prozent, Langtext, Verknüpfung, Rollup)
- **40.000 Beziehungszuordnungen** (n:m) und 10.000 Rollup-Berechnungen darüber
- Tabellen-, Board- und Kalenderansicht sowie ein Dashboard mit 6 Standardelementen

## Messwerte 1.7.0 (Sekunden)
| Vorgang | Ziel | 1.7.0 (Lauf 1 / Lauf 2) | 1.5.0 | Ergebnis |
|---|---|---|---|---|
| Einzelnen Datensatz öffnen (Maske mit 50 Feldern) | höchstens 2 | **0,41 / 0,38** | 0,35 | erfüllt |
| Einzelnen Datensatz speichern (Prüfung, Transaktion, Protokoll, Neuberechnung) | höchstens 2 | **0,02 / 0,02** | 0,02 | erfüllt |
| Einfache Suche in 10.000 Einträgen (erste Suche, 50 Spalten, baut Suchindex) | höchstens 3 | **2,74 / 2,67** | 2,63 | erfüllt |
| Weitere Suche / Suchbegriff verfeinern | höchstens 3 | **0,02 / 0,02** | 0,02 | erfüllt |
| Filter „Status offen“ in 10.000 Einträgen (7.500 Treffer) | höchstens 3 | **1,62 / 1,59** | 1,39 | erfüllt |
| Sortierung über 10.000 Einträge | kein Zielwert | 0,20 / 0,19 | 0,19 | |
| Dashboard mit 6 Standardelementen über 10.000 Einträge aktualisieren | höchstens 10 | **9,42 / 9,28** | 8,48 | erfüllt |
| Dashboard erneut aufrufen, Daten unverändert | kein Zielwert | **0,05 / 0,05** | 0,05 | Seite wird nur eingeblendet |
| Tabellenansicht öffnen (Seite 1 von 10.000) | kein Zielwert | 2,38 / 2,38 | 2,35 | |
| Tabellenansicht erneut aufrufen, unverändert | kein Zielwert | **0,04 / 0,04** | 0,05 | Seite wird nur eingeblendet |
| Board-Ansicht öffnen | kein Zielwert | 2,54 / 2,77 | 2,42 | |
| Kalender-Ansicht öffnen | kein Zielwert | 2,25 / 2,21 | 2,38 | |
| Alle Berechnungen neu (10.000 Rollups über 40.000 Zuordnungen) | kein Zielwert | 3,66 / 2,88 | 2,79 | |
| Startseite | kein Zielwert | 0,38 / 0,38 | 0,30 | |
| Aufbau der kompletten Lasttestdatei (Build) | kein Zielwert | 131 bis 150 s | 128 s | |

Rohdaten: `03_Pruefung/Leistung_Ergebnis.txt` (Lauf 2). Zwischen zwei Läufen derselben Datei schwanken einzelne Werte um bis zu 0,8 s (Rollups 3,66 und 2,88 s), meist um 0,1 bis 0,3 s.

**Vergleich mit 1.5.0:** Alle Zielwerte sind erfüllt. Das Dashboard braucht rund 0,9 s mehr (9,3 statt 8,5 s, Ziel 10 s), der Filter 0,2 s und die Startseite 0,08 s mehr. In 1.7.0 wurde an diesen Stellen nichts geändert; dazwischen liegen 1.6.0 und 1.6.1 (Navigation, Zoom-Takt, Seitenspeicher), deren Laufzeiten nicht gesondert gemessen wurden. Dabei gefunden und behoben: Ein erster Lauf mit 1.7.0 zeigte Filter 3,15 s und Zeitverlauf 4,70 s. Ursache war eine Änderung an der Schlüsselklasse `cMap` (erst einfügen, dann ersetzen), die beim Aktualisieren vorhandener Schlüssel jedes Mal eine Fehlerbehandlung auslöste. Sie gilt jetzt nur noch dort, wo die Schlüssel fast immer neu sind (Werte einer Importzeile).

## Liste übernehmen (neu in 1.7.0)
Gemessen mit `build\listimport_test.ps1` an allen Testlisten, Übernahme jeweils in eine neue Tabelle der leeren Masterdatei.

| Liste | Umfang | Lesen und auswerten | Prüfen | Übernehmen (mit Sicherung) |
|---|---|---|---|---|
| 13 Verwaltungslisten (14 Tabellen, darunter Formular, Monate als Spalten, 4 Blätter zusammen, CSV) | 1 bis 90 Einträge, 2 bis 14 Spalten | 0,0 bis 0,5 s | höchstens 0,1 s | 2,3 bis 3,3 s |
| 2ask-Nachbau, kodiert mit Codeplan | 24 Fragebögen, 43 Felder | 0,6 s | 0,2 s | 10,1 s |
| echte 2ask-Exporte, kodiert mit Codeplan | 3 Fragebögen, 43 Felder | 0,5 s | 0,0 s | 13,3 und 17,4 s |
| 2ask unkodiert als sonstige Liste | 24 Zeilen, 75 Spalten | 0,7 s | 0,3 s | 36,3 s |
| Große Liste | 5.000 Zeilen, 12 Spalten | 9,1 s | 9,8 s | 15,4 s |

Die Übernahme kostet vor allem die automatische Sicherung vorher (rund 2 s) und das Anlegen der Felder mit Spalte, Ansichten und Auswahllisten. Jedes weitere Feld wird dabei etwas teurer: 10 Felder brauchen zusammen rund 0,5 s, 43 Felder rund 8 s, 75 Felder rund 34 s. Die Einträge selbst werden in einem Zug geschrieben (5.000 Einträge in 0,8 s).

## Optimierung in 1.7.0 (gemessen, nach Befund)
Große Liste mit 5.000 Zeilen und 12 Spalten, Messung je Teilschritt mit `mListImport.LiProfile`:

| Schritt | vorher | nachher | Maßnahme |
|---|---|---|---|
| Lesen und auswerten | 27,2 s | 9,1 s | Formeln, Zellformate und verbundene Zellen werden nur in Spalten bzw. Zeilen nachgeschlagen, in denen es sie gibt (ein fehlender Schlüssel kostet in VBA eine Fehlerbehandlung); die Prüfung auf Kreuzspalten bricht ab, sobald sie entschieden ist; der Zellzugriff merkt sich Spalte und Blatt; der Zähler in `cMap.AddNum` schlägt einmal statt dreimal nach |
| Prüfen | 14,6 s | 9,8 s | Feldbeschreibung einmal je Spalte statt für jede Zelle neu; Zeilenwerte ohne Fehlerbehandlung eintragen (`cMap.SetNew`) |
| Übernehmen | 31,5 s | 15,4 s | Auswahlwerte über eine vorbereitete Tabelle Name zu Option statt Suche über alle Optionen je Zelle; Zielfeld und Feldart je Spalte vorbereitet |

Prüfen und Übernehmen gehören zum allgemeinen Import und machen auch „Daten importieren“ schneller.

## Seitenwechsel in der Demodatei (Millisekunden, je Klick in der Navigation, Messung 1.3.0)
| Seite | 1.2.0 (Seite + Navigation) | 1.3.0 erster Besuch | 1.3.0 unverändert erneut |
|---|---|---|---|
| Start | 234 + 195 | 180 bis 220 | **63 bis 70** |
| Tabelle „Projekte“ | 266 bis 320 + 188 | 266 bis 390 | (Wechsel zwischen zwei Tabellen: neu) |
| Tabelle „Aufgaben“ | 508 bis 539 + 203 bis 234 | 375 bis 477 | (Wechsel zwischen zwei Tabellen: neu) |
| Dashboard | 86 bis 102 + 180 | 773 (erste Berechnung) | **70 bis 78** |
| Einrichten | 227 bis 242 + 195 | 266 bis 313 | **78** |
| Technik | 78 bis 86 + 188 | 125 | **63 bis 78** |
| Hilfe, Erfassung (immer neu) | 94 bis 164 + 180 bis 227 | 125 bis 180 | immer neu |

## Änderungen in 1.7.0 mit Laufzeitwirkung
- „Liste übernehmen“ läuft nur, wenn es benutzt wird; an Seiten, Ansichten und Dashboards ändert sich nichts.
- Schnellerer Import großer Listen (siehe oben), auch bei „Daten importieren“.

## Änderungen in 1.5.0 mit Laufzeitwirkung
- Tabellenansicht: ein zusätzlicher Schreibvorgang (Leerzeichen in der Nachbarspalte gegen überlaufenden Text), im Messrauschen.
- Neue Dashboard-Darstellungen und Kennzahlen: nur bei Verwendung; Median sortiert die Werte je Gruppe. Im Lasttest nicht verwendet.

## Änderungen in 1.4.0 mit Laufzeitwirkung
- **Kalender:** Die Spurvergabe je Woche für mehrtägige Einträge kostet im Lasttest rund 0,4 s mehr beim ersten Aufbau der Monatsansicht. Bei üblichen Datenmengen ist das nicht spürbar.
- **Projektleiste:**
  - Tagesgenaue Leisten werden als Formen gezeichnet statt als Zellfärbung.
  - Je Zeile entsteht eine Form; mit Fortschritt bzw. Namen ein bis zwei weitere.
  - Je Zeitraum werden höchstens 300 Zeilen gezeigt.
  - Für die Projektleiste gibt es keine eigene Lastmessung.
- **Dashboard „Anordnen“:** Die Prüfung nach dem Loslassen einer Karte vergleicht nur die Lage der Karten und läuft nur im Anordnen-Modus.

## Optimierungen in 1.3.0 (gemessen, nach Befund)
1. **Navigation: von rund 190 ms auf 16 bis 60 ms je Seitenwechsel.** Vorher wurden bei jedem Wechsel alle Einträge (rund 40 Steuerelemente) gelöscht und neu angelegt, und das Hintergrundbild wurde komplett mit sechsfacher Abtastung gezeichnet. Jetzt wird nur die Markierung umgesetzt; mehrfach abgetastet werden nur die Streifen mit Rundungen. Neu aufgebaut wird nur bei geänderten Einträgen oder Farben, ohne Zwischenzeichnen.
2. **Wiederverwendung unveränderter Seiten:** 60 bis 80 ms statt 200 bis 450 ms (Demo), 0,04 bis 0,05 s statt 1,9 bis 9,5 s (Lasttest).
3. **Farbwechsel ohne Stottern:** Die Navigation wird nicht mehr geschlossen und neu geöffnet.
4. **Formen:** Eine leere Form kostet 0,2 ms, eine Schaltfläche aber 3 bis 6 ms durch Text- und Größenberechnung (AutoSize 1,7 ms). Breiten werden je Beschriftung einmal gemessen; Startseiten-Karten brauchen 3 statt 5 Formen.
5. **Tabellenansicht:**
   - Optionsfarben werden je Farbe gemeinsam gesetzt statt Zelle für Zelle.
   - Spaltenköpfe werden in einem Zug geschrieben.
   - Die Markierungsspalte liegt im Datenblock.
   - Bezeichnungen verknüpfter Einträge werden zwischengespeichert.
6. **Excel-Formelspalten:** Alle Einträge werden in einer Berechnung erledigt: Bereich auf dem Hilfsblatt füllen, nur diesen Bereich rechnen, Ergebnis in einem Zug lesen.

## Verbleibende Grenzen (offen benannt)
- Das Dashboard mit 6 Elementen über 10.000 Einträge braucht beim ersten Aufruf bzw. nach Datenänderungen 9,3 s und liegt damit knapp unter dem Zielwert von 10 s. Danach ist der Wechsel sofort.
- Die erste Suche nach dem Öffnen bzw. nach einer Datenänderung baut den Suchindex auf (rund 3 s); folgende Suchen sind sofort.
- Der Wechsel zwischen zwei verschiedenen Tabellen baut die Seite neu auf, weil alle Tabellen ein Blatt teilen.
- „Liste übernehmen“: Sehr breite Listen brauchen beim Übernehmen spürbar länger, weil jedes weitere Feld etwas teurer angelegt wird (75 Spalten rund 36 s). Eine Liste mit 5.000 Zeilen braucht vom Lesen bis zur fertigen Tabelle rund 35 s.
- Gemessen auf einem Rechner mit SSD und 32 GB RAM. Auf schwächerer Hardware oder auf Netzlaufwerken sind längere Zeiten zu erwarten.
- 32-Bit-Excel wurde nicht gemessen; der Code ist 32/64-Bit-fähig.
