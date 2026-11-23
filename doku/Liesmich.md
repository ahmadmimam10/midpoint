# MidPoint 1.8.2: Lieferpaket

Offline-Datenbanken in Excel: Tabellen und Eigenschaften wie in Notion zusammenstellen, Daten komfortabel pflegen, in Tabelle, Kalender, Board und Projektleiste bearbeiten, auswerten, mit Erfassungsdateien Daten von anderen einsammeln und vorhandene Listen übernehmen. Ohne Cloud, ohne Add-ins, ohne Programmierung.

Neu in 1.8.2 (Wartungsversion):
- **Modul 4 der Lernreihe neu, jetzt in vier Videos:** 4.1 Das Board, 4.2 Kalender und Zeitleiste, 4.3 Das Dashboard, 4.4 Daten einsammeln. Ruhiger erklärt und mit mehr Inhalt: eine vierte Tabelle „Arbeiten“ mit Kalender und Zeitleiste, ein Dashboard mit Kacheln, Balken, Ring und Säulen. Die Lernreihe hat jetzt 17 Videos.
- **Korrekturen:** Erfassungsdateien blieben hängen, wenn ein Drucker wie „OneNote (Desktop)“ als Standarddrucker eingestellt war; in der Farbstil-Auswahl sind alle Beschreibungen gleich groß; die Terminliste des Kalenders zeigt auch mehrtägige Einträge aus dem Vormonat; Mehrfachlisten in Dialogen zeigen auch den letzten Eintrag ganz.
- **Open Source** unter der MIT-Lizenz (`LICENSE.txt`).

Neu in 1.8.1 (Wartungsversion):
- **Neue Lernreihe:** 16 kurze Videos in fünf Modulen und einem Extra. Von der leeren Masterdatei an entsteht Schritt für Schritt eine Datei für die Spielplätze eines erfundenen Grünflächenamts: erste Tabelle, Einträge, Verknüpfungen, Board, Kalender, Dashboard und Erfassungsdateien. Dazu die Übung „Ihr Fuhrpark“, eine Ideenkiste und ein Video für IT und Datenschutz. Alles im `Lernen/Lernportal.html`, ohne Internet.
- **Mitbauen:** je Modul ein Bauplan zum Abhaken und eine Startdatei, die alte Baumliste, die fertige Datei und die Fuhrpark-Übung mit Lösung (`Lernen/Mitbauen`).
- **Korrekturen:** Jahreszahlen ohne Tausenderpunkt auch bei Feldern wie „Pflanzjahr“ aus „Liste übernehmen“; das Quellcode-Fenster nennt nur noch Monat und Jahr des Programmstands; die Beschreibung der Vorlage „Spielplätze & Grünflächen“ ist nicht mehr abgeschnitten.
- **Paket neu gegliedert:** die Masterdatei ganz oben, dann `Vorstellung` (Demodatei, Video „Warum MidPoint?“), `Lernen`, `Dokumentation` und `IT und Datenschutz`.

Neu in 1.8.0:
- **Diagramm kopieren:** Jedes Dashboard-Element lässt sich über „…“ kopieren und in PowerPoint oder Word mit Strg+V einfügen, als scharfes Bild mit Titel, Farben und Schrift wie im Dashboard.
- **Liste übernehmen, nachgeschärft** an echten Behördenlisten und gewachsenen Verwaltungslisten: Hinweiszeilen unter dem Kopf (Stichtag, Übersetzung), bessere Feldnamen, „-“ und „k. A.“ als leer, „ca. 1.200 €“ und „500,-“ als Betrag, „dito“ als Wert der Zeile darüber, Blätter mit leicht abweichenden Spalten, keine unechten Mehrfachauswahlen mehr, unkodierte 2ask-Exporte vollständig, deutlich schneller bei vielen Auswahlwerten.

Neu in 1.7.1 (Wartungsversion):
- **Liste übernehmen:** Wird ein Feld auf „Mehrfachauswahl“ umgestellt, ordnet eine Häkchenliste jeden Wert beliebig vielen Werten zu. Zeilenumbrüche in einer Zelle trennen Werte. Die Meldung nach der Übernahme ist jetzt immer sichtbar (bisher blieb sie nach langen 2ask-Übernahmen verborgen).
- **Design „Mainz“:** Vorlage unter Einrichten > Dokumentstil und Farbstil für Dashboards angelehnt an das Erscheinungsbild der Stadt Mainz (Gelb, Hellgrün, Calibri). Kein offizielles Produkt der Stadt.

Neu in 1.7.0:
- **Liste übernehmen** (auf Start und unter Einrichten): Ein Assistent macht aus einer vorhandenen Excel- oder CSV-Liste eine MidPoint-Tabelle, auch wenn sie nicht aus MidPoint stammt. MidPoint findet die Tabellen in der Datei, erkennt Kopf-, Gruppen- und Summenzeilen und die Art jeder Spalte und fragt nach, wo es nicht sicher ist (z. B. Kreuzspalten als Mehrfachauswahl, Monate als Spalten, Abkürzungen wie „erl.“).
- **Befragungen aus 2ask:** Der kodierte Excel-Export mit Codeplan wird genau übernommen: Antwortcodes als Optionen in der Reihenfolge des Fragebogens, Mehrfachauswahl, Matrixfragen, Skalen, Rangfolge und Eingabefelder mit passendem Typ.
- **Schnellerer Import großer Listen**, auch bei „Daten importieren“: 5.000 Zeilen werden in 9,8 statt 14,6 s geprüft und in 15,4 statt 31,5 s übernommen.
- **Sicherheitsprüfung vor jedem Release:** Ergebnis auf der Seite Technik und in `IT und Datenschutz/Pruefung/Sicherheitspruefung.md`.
- **Zum Ausprobieren:** `IT und Datenschutz/Pruefung/Testlisten` enthält 13 typische Verwaltungslisten und 2ask-Exporte mit Beschreibung, was jeweils erkannt werden soll.

Neu in 1.6.1 (Wartungsversion):
- **Excel-Zoom:** Vergrößern und Verkleinern (Schieberegler unten rechts oder Strg + Mausrad) funktioniert auf allen Seiten. Die Navigation behält ihre Größe und liegt immer neben dem Inhalt. Der Zoom gilt für alle Seiten.
- **Waagerechtes Scrollen** in breiten Tabellen ist ruhig: kein Zurückspringen nach links, kein Stottern.
- **Bildschirmskalierung 125 bis 175 %:** saubere Kurven in der Navigation, alle Einträge erreichbar, Logo und Titel mittig.
- **Seitenkopf:** Logo, Status und Untertitel überlagern nie Titel oder „Speichern“.
- **Fehlanzeige per Klick:** Meldet eine Stelle „nichts zu melden“ (z. B. per Mail), genügt in den Rückmeldungen ein Klick auf „offen“.
- **Schnellere Seitenwechsel:** Unveränderte Seiten erscheinen sofort (z. B. das Dashboard in 70 statt 770 ms).

Neu in 1.6.0:
- **Neuer Name: MidPoint**, mit Logo in der Navigation. Erzeugte Dateien bleiben neutral, ohne Logo.
- **Erfassungsdateien einfacher:**
  - Ablaufbild und „So funktioniert's“.
  - Assistent in vier Schritten.
  - Eine Datei für alle; Absender wählen ihren Namen aus einer Liste.
  - Übersicht „Rückmeldungen“: wer hat geantwortet, wer fehlt noch.
- **Fehler behoben:**
  - Dieselbe Datei von mehreren Stellen wird jetzt vollständig eingelesen.
  - Eine Erfassung, die nach dem Erzeugen nicht gespeichert wurde, lässt sich wiederherstellen.
  - Spaltenköpfe und Schaltflächen in den Meldungsfenstern sitzen richtig.
- **Bildschirme mit Skalierung (z. B. 175 %):** Die Navigation passt sich an. Kein Scrollen ins Leere mehr, und nichts rutscht unter die Navigation.

Neu in 1.5.0:
- **Dokumentstil:** 10 Vorlagen (Farben + Schrift) oder eigene Auswahl für die ganze Datei.
- **„Welches Diagramm passt?“:** Bildübersicht aller Darstellungen; neu ist das **Sankey-Diagramm**.
- **Listen:** Schalter „Zeilenumbruch“; alle Werte linksbündig.
- **Erfassungsdateien:** bestehende und neue Einträge auf einem Blatt klar getrennt; das Layout passt sich an wenige oder viele Spalten an.
- **Mehrfachauswahl mit Höchstzahl** (z. B. 6 Möglichkeiten, höchstens 3). Sie gilt in Formular, Import und Erfassungsdateien; dort nimmt Excel keinen Haken über der Grenze an.
- **Farbstile mit Farbmustern:** 10 gedeckte Stile (kräftig und pastell), jeder mit seinen acht Farben. Für Dashboards, Erfassungsdateien und Export.
- **Erfassungsdateien:**
  - mehrere Tabellen direkt beim Anlegen wählen (je ein Blatt)
  - beim Prüfen klar getrennter Bereich „Neue Einträge“ mit vorbelegtem Empfänger
  - Ankreuzspalten auch beim Prüfen
  - Farbstil und Schriftart wählbar (u. a. Garamond)
- **Excel-Export:**
  - mehrere Ansichten oder Tabellen als Blätter einer Datei
  - schlicht oder gestaltet, mit Farbe und Schrift
- **Dashboards:**
  - neue Darstellungen: Netz, Kreis, Fläche, Trichter, Baumdiagramm, Pareto, Rangliste
  - neue Kennzahlen: Median, Spannweite, verschiedene Werte
- **Tabellen und Formulare:**
  - lange Texte in der letzten Spalte enden am Rand
  - Zahlen in Formularen linksbündig
  - Dialoge zeigen Angaben erst, wenn sie gebraucht werden

Neu in 1.4.0:
- Tabellen lassen sich über „Ansicht > Tabelle löschen…“ (sowie im Eigenschaftsdialog und unter Einrichten) löschen, mit Prüfung und Sicherung.
- Fehler behoben: Fehler 91 beim Anlegen einer Tabelle, nachdem die einzige Tabelle gelöscht wurde; feine weiße Linien in der Navigation bei hellen Farben, „Erste Tabelle anlegen“ in der Navigation, Dialog „Neue Tabelle“ unter der Taskleiste, „Typ ändern“ an der Hauptspalte (Fehler 9).
- Zahlen einheitlich ohne hängendes Komma und überall rechtsbündig.
- Kalender: Heute als Rahmen; mehrtägige Einträge als saubere Leisten in eigenen Spuren.
- Projektleiste: tagesgenau, optional mit Namen auf den Leisten.
- Verknüpfung und Rollup werden geführt eingerichtet, mit Beispiel und Klartext.
- Spaltenauswahl ohne versehentliches Ein-/Ausblenden.
- „Zurück“ in mehrstufigen Dialogen.
- Dashboard-Karten rasten sofort nach dem Loslassen ein.

Neu in 1.3.0: deutlich schnellere Seitenwechsel (unveränderte Seiten werden nur eingeblendet), ruhigere Navigation ohne Flackern, helle Pastellfarben mit dunkler Schrift, Formelspalten als echte Excel-Formeln mit Spalten per Klick, Dashboard-Anordnen im festen 4-Spalten-Raster. 

Neu in 1.2.0: alles in einer Oberfläche (kein Umbaumodus), Schalter „Automatisch speichern“, schlankere Formulare, Dashboard-Stile/-Paletten mit Vorschau und Anordnen per Ziehen, Ankreuzspalten für Mehrfachauswahl, vorbelegte Prüfdateien je Empfänger mit Aktualisierung über die ID, mehrere Blätter je Erfassungsdatei. Details: `Dokumentation/Versionshinweise.md`.

**Bestehende Dateien 1.1.0 aktualisieren:** `powershell -ExecutionPolicy Bypass -File "IT und Datenschutz\Quellcode\build\build.ps1" -Update "Pfad\zur\Datei.xlsm"` erzeugt `…_aktualisiert.xlsm` mit dem neuen Programmcode und allen Daten; beim ersten Öffnen wird das Dateiformat automatisch ergänzt. (Benötigt auf dem Entwicklungsrechner „Zugriff auf das VBA-Projektobjektmodell vertrauen“.)

## Inhalt
| Ordner oder Datei | Inhalt |
|---|---|
| `MidPoint_Master.xlsm` | leere Masterdatei: für jedes Thema eine Kopie anlegen und umbenennen |
| `LICENSE.txt` | Lizenz (MIT) |
| `Vorstellung` | `MidPoint_Demo.xlsm` mit fiktiven Beispieldaten zum Zeigen und Ausprobieren, dazu das Video „Warum MidPoint?“ (1:43) |
| `Lernen` | `Lernportal.html` mit 17 Lernvideos (Ordner `Lernvideos`), dazu `Mitbauen`: Baupläne, Startdateien je Modul, alte Baumliste, fertige Datei, Fuhrpark-Übung mit Lösung |
| `Dokumentation` | Kurzanleitung „In zehn Schritten“, Hilfe, Versionshinweise |
| `IT und Datenschutz` | Video „Für IT und Datenschutz“, Technische Dokumentation, `Pruefung` (Testmatrix, Sicherheitsprüfung, Leistungsbericht, Rohdaten, Testlauf und Testlisten „Liste übernehmen“, Lasttestdatei, echte Bildschirmaufnahmen), `Quellcode` (VBA-Quellcode `src`, Build- und Prüfskripte `build`, Build-Anleitung) |

## Schnellstart
1. `Lernen/Lernportal.html` öffnen: Video 1.1 zeigt in drei Minuten, worum es geht.
2. `MidPoint_Master.xlsm` kopieren und umbenennen, die Kopie öffnen, „Inhalt aktivieren“ (bzw. Freigabe durch die IT).
3. Oben links den Titel festlegen, „Aus Vorlage“ oder „Neue Tabelle“, dann über „+“ im Tabellenkopf Eigenschaften hinzufügen. Schritt für Schritt zeigt es die Lernreihe, mit den Bauplänen in `Lernen/Mitbauen`.

Die Demodatei (`Vorstellung/MidPoint_Demo.xlsm`) zeigt alle Funktionen mit fiktiven Daten. Ausführliche Hilfe: in der Datei unter „Hilfe & Sicherung“.

Datenschutz: Alle Daten bleiben in der Datei bzw. in bewusst erzeugten Exporten, Sicherungen und Erfassungsdateien.
