# Testmatrix für MidPoint 1.8.2

**Prüfumgebung:** Windows 11 Home, Microsoft 365 Excel 16.0 (Build 20326) 64 Bit, AMD Ryzen 5 2600X, 32 GB RAM. Alle Prüfungen in echter Excel-Laufzeit (Oktober und November 2026); Version 1.2.0 im Oktober 2026 vollständig neu geprüft (48 automatisierte Prüfungen bestanden, siehe unten).
**Evidenz:** automatisierte Prüfungen (`build\selftest.ps1` → `03_Pruefung/Selbsttest_Ergebnis.txt`), Lastmessung (`build\perf.ps1` → `Leistung_Ergebnis.txt`), echte Bildschirmaufnahmen (`03_Pruefung/screenshots`). Das VBA-Projekt wird bei jedem Build vollständig kompiliert (Build bricht bei Kompilierfehler ab).

| ID | Test | Status | Evidenz / Ergebnis |
|---|---|---|---|
| T01 | Masterdatei kopieren, umbenennen, unabhängig öffnen | **BESTANDEN** | Kopie „Meine_Materialverwaltung.xlsm“ gestartet, schreibbar, Navigation aktiv (Selbsttest T01); Screenshot 01 |
| T02 | Neue Tabelle mit allen Kernfeldtypen | **BESTANDEN** | Tabelle mit allen 28 Eigenschaftsarten angelegt; Maske baut sich auf, jede Eigenschaft hat Filterangebote (T02); Screenshots 08, 10, 12 |
| T03 | Zahl außerhalb Bereich, falsches Datum, fehlendes Pflichtfeld | **BESTANDEN** | 150 % abgelehnt, 31.02.2026 abgelehnt, Pflichtfeld meldet „bitte ausfüllen“, Frist vor Beginn abgelehnt (T03a bis d) |
| T04 | Mehrere Zeilen mit gültigen und ungültigen Werten einfügen | **BESTANDEN** | ein ungültiger Wert → nichts übernommen + Fehlerliste; gültige Werte vollständig übernommen; Systemspalte geschützt (T04a bis c) |
| T05 | Feld umbenennen, umordnen, neues Pflichtfeld | **BESTANDEN** | Ansichten verweisen weiter über Feld-ID; Umordnen ohne Datenverlust (T05a/b). Neues Pflichtfeld in gefüllter Tabelle (Abfrage Standardwert/Nachpflege): in dieser Runde nicht erneut ausgeführt, nur Code geprüft |
| T06 | Inkompatibler Typwechsel / abhängiges Feld löschen | **BESTANDEN** | Löschen von „Frist“ blockiert (7 Abhängigkeiten), Text→Zahl zeigt 68 Konflikte in der Vorschau, Verknüpfung/Formel nicht umwandelbar (T06a bis c) |
| T07 | Einfach-/Mehrfachverknüpfung, gleichnamige Ziele, Schnellanlage | **BESTANDEN** | Rückverweis wandert korrekt mit, IDs statt Namen (T07a/b); Auswahl zeigt Zusatzmerkmale zur Unterscheidung. Schnellanlage aus der Maske: in dieser Runde nicht erneut ausgeführt |
| T08 | Zwei Tabellenansichten mit verschiedenen Filtern speichern und neu öffnen | **BESTANDEN** | 49 bzw. 17 Einträge, keine Datenduplikate (T08); Fortbestand nach Schließen/Öffnen über Fingerabdruck inkl. Ansichten bestätigt (T17) |
| T09 | Datum im Kalender, Status im Board, Zeitraum im Gantt bearbeiten | **BESTANDEN** (Board vollständig), Kalender/Gantt über dieselbe Maske | Karte „Nach rechts“: Offen 18→17, In Arbeit 19→20, Tabelle und Dashboard sofort aktualisiert (Screenshots 24, 06); Kalender/Gantt öffnen Einträge per Doppelklick in derselben Maske mit identischem Speicherweg |
| T10 | Undatierte Einträge, viele Termine, leere Boardgruppe, ungültiger Gantt-Zeitraum | **BESTANDEN** | Kalender „Ohne Datum (1)“, Gantt „Ohne vollständigen Zeitraum“, „+ n weitere“ je Tag, Ende vor Beginn abgelehnt (T03d); Screenshots 04, 05 |
| T11 | Formel, Rollup, leere Werte, Division durch null, Zirkelbezug | **BESTANDEN** | Zirkelbezug erkannt, Division durch 0 → „Fehler: Division durch 0“, Rollup zählt korrekt (T11a bis c) |
| T12 | n:m-Auswertung mit überlappenden Kategorien | **BESTANDEN** | Gruppierung nach Mehrfachverknüpfung liefert Gruppenwerte mit Hinweis „Mehrfachzuordnung“, Gesamtzahl eindeutig (T12; Dashboard „Veranstaltungen nach Themen (Mehrfachzuordnung)“) |
| T13 | Dashboard mit sechs Elementen, Filtern und Detailaufruf | **BESTANDEN** | Details zu „Erledigt“ zeigt genau die 17 Datensätze des Balkens (Screenshot 07); 6-Element-Dashboard im Lasttest 9,3 s |
| T14 | XLSX/CSV importieren, fehlerhafte Zeilen, Abbruch | **BESTANDEN** | CSV mit 3 Zeilen: 1 gültig, 2 Fehler mit Begründung (unbekannter Status, 31.02.2026); Abbruch → weiterhin 67 Einträge (Screenshot 25) |
| T15 | Gefilterte Ansicht und Dashboard exportieren | **BESTANDEN** | XLSX mit 6 Einträgen, lesbare Werte, Filterangabe im Kopf, keine internen IDs, keine Makros; Dashboard-Export erstellt |
| T16 | Archivieren, wiederherstellen, letzte Änderung zurücknehmen | **BESTANDEN** | Archiv + Rücknahme stellen Zustand her (T16); Button-Aktion mit Rückgängig (N04) |
| T17 | Speichern, Excel vollständig schließen, erneut öffnen | **BESTANDEN** | Fingerabdruck (Tabellen, Felder, Ansichten inkl. Filter, Dashboards, Elemente, Zuordnungen, Prüfsumme aller Werte) vor/nach identisch |
| T18 | Nur-Lesen, fehlende Schreibrechte, fehlgeschlagene Sicherung | **BESTANDEN** | zweite Instanz öffnet schreibgeschützt → Lesemodus, Schreiben gesperrt mit Meldung (T18); Sicherung prüft Datei/ZIP-Signatur und meldet Fehler statt Erfolg (Code geprüft) |
| T19 | Sicherung erstellen und tatsächlich wiederherstellen | **BESTANDEN** | Sicherung → Änderung → Wiederherstellung → Zustand identisch (T19). Dabei gefundenen Fehler behoben: misslungene Wiederherstellung spielt jetzt automatisch den vorherigen Stand zurück |
| T20 | Offline-Test und Code-/Verbindungsprüfung | **TEILWEISE**: Codeprüfung BESTANDEN, Test mit getrennter Netzwerkverbindung NICHT AUSGEFÜHRT | Quellcode enthält keine Netzwerkbibliotheken (kein XMLHTTP/WinHttp/URLDownload/wininet); einzige externe Aufrufe sind vom Nutzer bestätigte „Link/Datei/Ordner öffnen“. Alle Prüfungen liefen ohne Netzwerkzugriff der Anwendung; das physische Trennen des Netzwerks wurde nicht durchgeführt |
| T21 | Lastprofil und Anzeigegrößen | **BESTANDEN** | alle Zielwerte erfüllt (Leistungsbericht); Oberfläche geprüft bei 1166 px und Vollbild (Screenshots) |
| T22 | Gezielter Laufzeitfehler, Wechsel zu anderer Arbeitsmappe | **BESTANDEN** | Fehler mitten in Transaktion: Meldung statt Debugger, Änderung zurückgenommen, Excel-Zustände wiederhergestellt; andere Mappe behält Gitternetz/Köpfe/Register (T22) |
| T23 | Struktur mit Daten duplizieren | **BESTANDEN** | neue IDs, Original unverändert (T23) |
| T24 | Makros deaktiviert öffnen | **BESTANDEN** | nur Blatt „Hinweis“ sichtbar mit Erklärung (T24) |

## Zusätzliche Prüfungen der Erweiterungen
| ID | Prüfung | Status |
|---|---|---|
| N01 bis N06 | E-Mail/Telefon-Format, ID-Nummern eindeutig (AUF-1 …), Button setzt Felder, automatische Eigenschaften nicht überschreibbar, Kontrollkästchen nie leer | BESTANDEN |
| S01 | Suche findet Text, Auswahlwert und verknüpften Projektnamen (Nachzählung identisch) | BESTANDEN |
| C01 bis C08 | Erfassungsdatei (Liste/Formular, XLSX ohne Makros) erzeugen, Auswahllisten/Datumsprüfung/Blattschutz, fehlerhafte Datei wird nicht teilweise übernommen, korrigierte Datei übernommen, Werte korrekt (Absender, Verknüpfung, Person, Zahl), zweites Einlesen verhindert, Datei aus anderer Datenbank erkannt | BESTANDEN |
| UI | Erfassungsdatei im Excel-Fenster eines Empfängers ausgefüllt (Auswahlliste, Pflicht-Rotmarkierung, Hinweise), gespeichert, über „Rücklauf einlesen“ mit Vorschau übernommen, Eintrag per Suche gefunden | BESTANDEN (Screenshots 18 bis 23) |

## Prüfungen der Version 1.2.0
| ID | Prüfung | Status |
|---|---|---|
| C02 | Ankreuzspalten für Mehrfachauswahl „Zuständig“ (8 Optionen, senkrechter Titel, Auswahl „Ja“), beschriftetes breites Absenderfeld | BESTANDEN (automatisiert) |
| C05, C07 | Zwei angekreuzte Spalten ergeben die Mehrfachauswahl (Liste); Formular mit Ja-Zeile je Option | BESTANDEN (automatisiert) |
| C09 | Prüfdatei je Projekt: nur eigene Einträge vorbelegt, graue Spalten gesperrt, gelbe offen, versteckte ID, 5 leere Zeilen | BESTANDEN (automatisiert) |
| C10 bis C12 | Rücklauf: 1 geändert, 8 unverändert, 1 neu; Aktualisierung über ID ohne Dublette; neue Zeile automatisch dem Empfänger zugeordnet; zweites Einlesen erkannt | BESTANDEN (automatisiert) |
| C13 | Eine Datei mit zwei Blättern (Aufgaben, Projekte); Projektangabe über Blatt 2 aktualisiert | BESTANDEN (automatisiert) |
| UI-1 | Assistent „Vorhandene Einträge prüfen“ mit Tabelle Ressourcen, Anzeige-/Ausfüllspalten, Aufteilen nach Standort: 5 Dateien erzeugt; Datei „Bibliothek“ als Empfänger ausgefüllt (Menge, Status, neue Zeile) und eingelesen: Vorschau 1 neu/2 geändert/3 unverändert, Bürgerhaus „geprüft, keine Änderungen“; Werte in der Tabelle korrekt, neues Gerät mit Standort „Bibliothek“ | BESTANDEN (Screenshots 08 bis 10) |
| UI-2 | Automatisch speichern: Schalter ein, Zelländerung, Datei ca. 5 s später gespeichert, Ansicht und Auswahl bleiben stehen | BESTANDEN |
| UI-3 | Eintragsmaske ohne berechnete Eigenschaften, „+“ für Optionen, „Mehr“ | BESTANDEN (Screenshot 03) |
| UI-4 | Dashboard: Live-Vorschau von Stil/Palette, Anordnen per Ziehen (Reihenfolge und Breite übernommen), Element-Vorschau beim Anlegen mit Übernehmen | BESTANDEN (Screenshots 05 und 06). Dabei gefunden und behoben: Platzhalterkarte im Anordnen-Modus löste Laufzeitfehler aus |
| UI-5 | Navigation: graue Linie an der Öffnung entfernt (vergrößerte Bildschirmaufnahme geprüft), Viertelkreis-Übergänge, Titel per Klick änderbar, Navigationsfarbe per Farbfeld | BESTANDEN (Screenshots 01, 04) |
| M-1 | Programmupdate einer 1.1.0-Demodatei mit `build.ps1 -Update`, Öffnen: Format 2 → 3, alle Einträge erhalten, Erfassungsseite funktionsfähig | BESTANDEN. Dabei gefunden und behoben: Update-Skript speicherte die neuen Module nicht |

## Prüfungen der Version 1.3.0 (Oktober 2026)
| ID | Prüfung | Status |
|---|---|---|
| X01 | Excel-Formel `=[Aufwand (h)]*2` für alle Einträge in einem Rechenvorgang (66 richtig, 0 falsch) | BESTANDEN (automatisiert) |
| X02 | Deutsche Funktionsnamen (WENN, RUNDEN), Semikolon, Textverkettung | BESTANDEN (automatisiert) |
| X03 | Abgelehnt: WEBDIENST, Zellbezug A1, Blattbezug, Syntaxfehler, unbekannte Spalte | BESTANDEN (automatisiert) |
| X04 | Text „=1+1“ wird nie ausgeführt; Änderung rechnet abhängige Formelspalte neu | BESTANDEN (automatisiert) |
| X05 | Zirkelbezug bei Excel-Formel erkannt | BESTANDEN (automatisiert) |
| G01 | Gespeicherte Rasterplätze werden genau so gezeichnet | BESTANDEN (automatisiert) |
| G02 | Karte im Anordnen-Modus verschoben, „Fertig“: Platz gespeichert, verdrängte Karte rückt nach unten, Normalansicht zeigt dieselbe Lage | BESTANDEN (automatisiert) |
| P01 | Unveränderte Seite wird nur eingeblendet, nach Datenänderung neu aufgebaut | BESTANDEN (automatisiert) |
| P02 | Helle Farben: dunkle Schrift in Navigation und auf Akzentflächen, lesbarer Akzent auf Weiß | BESTANDEN (automatisiert) |
| T19 | Sicherung/Wiederherstellung mit Text „=1+1“ in den Daten: Zustand identisch | BESTANDEN (vorher Fehler gefunden und behoben) |
| UI-6 | Navigation: Symbol und Text mittig zur Markierung (vergrößerte Aufnahme), Pastellfarben Flieder/Flieder hell mit dunkler Schrift | BESTANDEN (Screenshot 01 in screenshots_130) |
| UI-7 | Formelleiste: Formel tippen, Spalte „Priorität“ in der Tabelle anklicken, Vorschau „dringend“, Übernehmen: Spalte rechts von „Status“ mit richtigen Werten | BESTANDEN (Screenshot 03). Dabei gefunden und behoben: neue Spalte landete ganz rechts; erster Klick ins mehrzeilige Formelfeld wurde im nicht modalen Fenster verschluckt (jetzt einzeilig, Schreibmarke startet in der Formel) |
| UI-8 | Anordnen: Karte „Kosten“ in Spalte 4 gezogen, Klick ins Blatt rastet ein, verdrängte Karte rückt in den freien Platz, nach „Fertig“ unverändert | BESTANDEN (Screenshot 02) |
| L-1 | Seitenwechsel Demodatei: unveränderte Seiten 63 bis 78 ms, Navigation 16 bis 63 ms (vorher 190 ms); Lasttest siehe Leistungsbericht | BESTANDEN |

Gesamt: 52 Prüfungen in `RunSelfTest` + T17, T18, T19, T01, T24 = **57 von 57 bestanden**.

## Prüfungen der Version 1.4.0 (Oktober 2026)
Automatisiert: alle 57 Prüfungen erneut **bestanden**, mit Version 1.4.0 neu gebauter Demodatei; T19-Sicherung `…_v1.4.0.xlsm`. Sichtprüfung in echter Excel-Laufzeit; die Screenshots liegen in `03_Pruefung/screenshots`.

| ID | Prüfung | Status |
|---|---|---|
| V-01 | Leere Masterdatei: „Erste Tabelle anlegen“ in der Navigation öffnet den Dialog „Neue Tabelle“ (nicht mehr „Einrichten“). „Einrichten“ unten bleibt beschriftet und unmarkiert. | BESTANDEN (Screenshot 01) |
| V-02 | Tabelle anlegen: Der vergrößerte Dialog rückt nach oben, „Fertig“ ist sichtbar. Vorher lag der Dialog unter der Taskleiste; das war die Ursache für „Tabelle lässt sich nicht anlegen“. | BESTANDEN (Screenshot 02) |
| V-03 | „Typ ändern…“ an der Hauptspalte „Projekt“ öffnet die Typauswahl (vorher Laufzeitfehler 9). | BESTANDEN (Screenshot 09) |
| V-04 | Kalender: Heute als farbiger Rahmen, keine Flächenfarbe. Drei überlappende Einträge (30.09. bis 02.10., 30.09., 01.10. bis 05.10.) liegen in getrennten Spuren als durchgehende Leisten; die Fortsetzung in der Folgewoche ist mit „→“ gekennzeichnet. | BESTANDEN (Screenshots 03, 04) |
| V-05 | Projektleiste Wochen und Monate: Die Leisten sind tagesgenau (Teilwochen und Teilmonate). „Namen auf Leisten“ zeigt die Namen klein auf der Leiste; ein zu langer Name steht rechts daneben. | BESTANDEN (Screenshots 05, 06) |
| V-06 | Formular: Das Verknüpfungsfeld ist so hoch wie „Auswählen…“. Das Register „Verknüpft aus anderen Tabellen“ zeigt eine Erklärung und die Herkunft je Feld. | BESTANDEN (Screenshots 07, 08) |
| V-07 | Rollup im Eigenschaftsdialog: Verknüpfung, Auswertung „Summe“ und Feld gewählt (nur Zahlenfelder angeboten), Klartext erscheint, gespeichert. Die Spalte zeigt Werte rechtsbündig mit einheitlich zwei Nachkommastellen. | BESTANDEN (Screenshots 10, 11, 13) |
| V-08 | Spaltenauswahl: Spalte angeklickt und zweimal „Nach oben“: Sie bleibt sichtbar, der Haken bleibt. | BESTANDEN (Screenshot 12) |
| V-09 | Formelleiste: Die Spalten- und Funktionskästchen haben den Text senkrecht mittig. | BESTANDEN (Screenshot 14) |
| V-10 | Dashboard „Anordnen“: Karte gezogen und losgelassen, sie rastet ohne weiteren Klick ein. Dabei gefunden und behoben: Der erste Klick danach ging ins Leere, weil die gezogene Form noch ausgewählt war. | BESTANDEN (Screenshot 15) |
| V-11 | Erfassungsdatei-Assistent: In Schritt 2 „← Zurück“ führt zu Schritt 1; der eingegebene Name bleibt erhalten. | BESTANDEN (Screenshot 16) |
| T25 | Leere Masterdatei: Tabelle anlegen, als einzige Tabelle löschen, erneut anlegen (dreimal angelegt, zweimal gelöscht, jeweils Tabelle geöffnet). Vorher wurde so der gemeldete Fehler 91 nachgestellt und behoben. | BESTANDEN (automatisiert, neu im Selbsttest; Gesamt jetzt **58 von 58**) |
| V-14 | Helle Navigationsfarbe (Salbei): Echte Bildschirmpixel über und unter der Markierung zeigen nur die Hintergrundfarbe, keine helle Zeile. | BESTANDEN (Screenshot 17, Pixelprüfung) |
| V-15 | Tabelle „Ressourcen“ über „Ansicht > Tabelle … löschen…“ gelöscht: Die Auswirkungen werden angezeigt (18 Einträge, Verknüpfung, 4 Dashboard-Elemente), die Bestätigung „Löschen“ wird angenommen, eine Sicherung wird erstellt; Navigation und Startseite zeigen die Tabelle nicht mehr. Die Schaltfläche „Tabelle löschen…“ im Eigenschaftsdialog ist sichtbar. | BESTANDEN (Screenshots 18 bis 21) |
| V-12 | Verknüpfung neu anlegen (Klartext-Satz), „Zurück“ bei „Neue Ansicht“ und beim Dashboard-Element | NUR CODE GEPRÜFT, nicht im Bild |
| V-13 | Lasttest mit 1.4.0: Datei neu gebaut, Messung wiederholt | siehe Leistungsbericht |

## Prüfungen der Version 1.8.2 (November 2026)
**Automatisiert:** 88 von 88 bestanden (82 in RunSelfTest, dazu T01, T17, T18, T19, T24 und T25) auf den ausgelieferten Dateien, mit „OneNote (Desktop)“ als Standarddrucker. Dazu die Bildprüfung aller Lernvideos, die Startdateien im Ordner „Mitbauen“ und die Sicherheitsprüfung. Der Lasttest ist unverändert aus 1.8.0: Die Änderungen betreffen keine gemessenen Abläufe.

| ID | Prüfung | Status |
|---|---|---|
| F-4 | Erfassungsdateien mit „OneNote (Desktop)“ als Standarddrucker: drei Prüfdateien je Spielplatz in rund zwei Sekunden angelegt. Bisher blieb Excel bei der Seiteneinrichtung stehen. Gefunden beim Aufnehmen von Video 4.4 | BESTANDEN (Aufnahmen pf0, pf1) |
| F-5 | Design „Mainz“: Beschreibung in der Stilauswahl, im Dokumentstil und in der Hilfe nennt nur Farben und Schrift, keine internen Angaben; in der Auswahl nicht abgeschnitten | BESTANDEN (Aufnahmen stile_b, hilfe) |
| F-6 | Testdaten: Telefonnummern in Testlisten, echten 2ask-Exporten und Selbsttest mit der nicht vergebenen Vorwahl 0123 statt eines echten Rufnummernschemas; Listen-Testlauf danach unverändert | BESTANDEN |
| F-7 | Farbstil-Auswahl: alle elf Beschreibungen gleich groß (bisher Wald, Erde, Salbei & Sand eine Stufe kleiner). Gefunden bei der Bildprüfung von Video 4.3 | BESTANDEN (Aufnahmen stile_a, stile_b) |
| F-8 | Kalender, Terminliste: mehrtägige Einträge, die im Vormonat beginnen und in den Monat hineinreichen, stehen in der Liste wie in der Monatsansicht; das ausgeschriebene Datum („Donnerstag, 10. Dezember 2026“, fett) passt in die Datumsspalte; das Suchfeld endet an der Spalte „Bezeichnung“, „Suchen“ und „17 Einträge“ stehen frei daneben (vorher lief das Feld darunter). Gefunden beim Aufnehmen und bei der Bildprüfung von Video 4.2 | BESTANDEN (Aufnahme k_liste) |
| F-9 | Eigenschaften-Dialog, Mehrfachliste mit zehn Einträgen (Kurzangaben der Board-Ansicht): neun ganze Zeilen, eine Zeile weiter gerollt ist „Spielplatz“ ganz zu sehen. Vorher war die Liste 9,2 Zeilen hoch, der letzte Eintrag blieb angeschnitten, und die Liste ließ sich nicht weiterrollen (MSForms zählt die angeschnittene Zeile als sichtbar). Gefunden bei der Bildprüfung von Video 4.1 | BESTANDEN (Aufnahmen v_board, v_board_2) |
| LV-4 | Modul 4 neu in vier Videos (4.1 Das Board, 4.2 Kalender und Zeitleiste, 4.3 Das Dashboard, 4.4 Daten einsammeln): neu aufgenommen mit der Tabelle „Arbeiten“ (17 Einträge), Kalender, Zeitleiste und Dashboard mit fünf Elementen | BESTANDEN |
| LV-5 | Bildprüfung aller 17 Videos: aus jedem fertigen Video Übersichtsblätter (zwei Bilder je Satz mit dem gesprochenen Text) und Ausschnitte in voller Auflösung zu jeder Markierung, jedem Klick, jeder Tippstelle und jeder Ausklappliste; Funde behoben, neu gerendert und erneut geprüft | BESTANDEN |
| LV-6 | Lernportal: 17 Videos in fünf Modulen und einem Extra, Lernwege, Kapitel, Suche, Links zu Bauplänen und Startdateien | BESTANDEN |
| S-7 | Sicherheitsprüfung (`build\sicherheit.ps1`) auf Quelltext und fertiger Masterdatei 1.8.2; Punkt B4 liest die Dateieigenschaften direkt aus der Datei | BESTANDEN (Sicherheitspruefung.md) |
| S-8 | Lieferpaket ohne Spuren des Entwicklungsrechners: in allen Excel-Dateien Autor und „zuletzt geändert von“ leer, Zeitstempel auf dem Monatsanfang; Textdateien, Lernportal, Baupläne und Videos durchsucht | BESTANDEN |

## Prüfungen der Version 1.8.1 (November 2026)
**Automatisiert:** 88 von 88 bestanden (82 in RunSelfTest, dazu T01, T17, T18, T19, T24 und T25) auf den ausgelieferten Dateien. Dazu die Prüfungen der neuen Lernreihe (Bild und Ton jedes Videos), der Startdateien im Ordner „Mitbauen“ und die Sicherheitsprüfung. Der Lasttest ist unverändert aus 1.8.0: Die Änderungen betreffen keine gemessenen Abläufe.

| ID | Prüfung | Status |
|---|---|---|
| F-1 | Jahreszahl ohne Tausenderpunkt: Feld „Pflanzjahr“ (Nummer ohne Grenzen) aus „Liste übernehmen“ zeigt 1985 statt 1.985 in der Tabelle der Bäume (Aufnahmen zu Video 3.3). Der Eingabehinweis in Erfassungsdateien nennt bei solchen Feldern eine Jahreszahl (im Code geprüft, nicht eigens fotografiert). Gefunden beim Aufnehmen von Video 3.3 | BESTANDEN |
| F-2 | Quellcode-Fenster: Titel „Quellcode (36.567 Zeilen, Katalog 1.8.1 / November 2026 / 36567 Zeilen)“, ohne Tag und Uhrzeit. Bisher mit Tag und Uhrzeit des Builds. Gefunden beim Aufnehmen des Extra-Videos | BESTANDEN (Aufnahme t_code) |
| F-3 | Aus Vorlage starten: Beschreibung „Spielplätze, Spielgeräte mit Prüffristen, Bäume mit Kontrollen, Beteiligungsprojekte“ vollständig sichtbar. Bisher abgeschnitten („… Projekte mit Bürgerb“) | BESTANDEN (Aufnahme vl_1) |
| LV-1 | Lernreihe: 16 Videos, je Satz Bild und Sprechertext abgeglichen (Kontaktbögen aus den fertigen Videos), Einsätze der Stimme gemessen (höchstens 0,3 Sekunden Abweichung, außer am Signet), Lautheit −16 LUFS, Spitze höchstens −2,9 dBTP; keine Entwicklungstage, keine Nachtzeiten, keine Pfade im Bild | BESTANDEN |
| LV-2 | Lernportal: 16 Videos in fünf Modulen und einem Extra, fünf Lernwege, Kapitel, Suche, Links zu Bauplänen und Startdateien; Ansicht in Edge geprüft | BESTANDEN |
| LV-3 | Mitbauen: fünf Baupläne (je eine Seite, PDF), Startdateien für die Module 2 bis 4, fertige Datei und Fuhrpark-Lösung über den Update-Weg auf 1.8.1; Protokoll nur „Petra Krause“ in der Beispielwelt (werktags, Arbeitszeit), Stichtag der Neuberechnung 30.11.2026; ohne Hilfsmodul der Aufnahmen | BESTANDEN |
| S-5 | Sicherheitsprüfung (`build\sicherheit.ps1`) auf Quelltext und fertiger Masterdatei 1.8.1 | BESTANDEN (Sicherheitspruefung.md) |
| S-6 | Lieferpaket ohne Spuren des Entwicklungsrechners: alle Excel-Dateien (auch Startdateien und Baumliste), Textdateien, Lernportal, Baupläne und Videos durchsucht | BESTANDEN |
| T-4 | Texte ohne Gedankenstriche (U+2013, U+2014) in Dokumentation, Lernportal, Bauplänen und Prüfberichten des Pakets | BESTANDEN |

## Prüfungen der Version 1.8.0 (November 2026)
**Automatisiert:** 88 von 88 bestanden (82 in `RunSelfTest`, neu D04 sowie L18 bis L20, dazu T01, T17, T18, T19, T24 und T25) auf den ausgelieferten Dateien. Dazu zwei Testläufe „Liste übernehmen“ (Standardsatz und neuer Satz mit Behördenlisten, gewachsenen Listen und zweiter 2ask-Befragung), Praxistest „Diagramm kopieren“ mit PowerPoint und Word, Sichtprüfung in echtem Excel (39 Aufnahmen, neu 30 bis 33 und 39), Lasttest und Sicherheitsprüfung.

| ID | Prüfung | Status |
|---|---|---|
| D04 | Diagramm kopieren im Selbsttest: erstes Diagramm des Dashboards „Auswertungen“ als Bild in der Zwischenablage | BESTANDEN (automatisiert) |
| V-52 | Diagramm kopieren in echtem PowerPoint und Word: Kachel (301 × 113 pt), Balkendiagramm und Tabelle (je 606 × 279 pt) als Bild eingefügt, Aussehen wie im Dashboard (Karte, Titel, Balken, Werte, Schrift), scharf beim Vergrößern; in Word als Bild auf Seitenbreite (454 × 208 pt). Menüknopf, „Details“ und Quellenhinweis fehlen wie vorgesehen | BESTANDEN (Screenshot 39) |
| L18 bis L20 | Hinweiszeile unter dem Kopf mit Gruppenüberschrift und Fußnote; Beträge „ca. 1.200 €“, „500,-“, „dito“, Ja/Nein-Varianten, Platzhalter, Ort mit Komma bleibt Text; unkodierter 2ask-Export mit Codeplan | BESTANDEN (automatisiert) |
| LI-10 | Vier Behördenlisten (nur lokal, Quellen in `testdaten/online/QUELLEN.md`): Gemeindeverzeichnis des Statistischen Bundesamts (15.984 Einträge, Kopf in Zeile 3 und 4 mit Hinweiszeile 5, Felder wie „Bevölkerung insgesamt“, „Reisegebiete Schlüssel“, keine Rückfragen); Schulverzeichnis Rheinland-Pfalz (vier gleich aufgebaute Blätter, 1.629 Einträge, Bezeichnungsfeld „Schule“); EFRE-Vorhaben Berlin (2.958 Einträge, englische Zeile unter dem Kopf als Hinweiszeile, Bezeichnungsfeld „Bezeichnung des Vorhabens“); Staatennamen des Auswärtigen Amts (198 Einträge). Dabei gefunden und behoben: (1) Hinweis- und Übersetzungszeilen wurden als erster Datensatz übernommen. (2) „Adenau, Stadt“ wurde eine Mehrfachauswahl mit 680 Optionen; das Anlegen dauerte 239 s, jetzt 19 s für die ganze Liste. (3) Tabulatoren am Zellende („FSP L“) wurden Teil der Option. (4) „-“ für „nichts vorhanden“ machte Zahlenspalten zu Text. (5) Feldnamen mit Fußnote, Silbentrennung oder über 80 Zeichen; Bezeichnungsfeld auf Kreis oder Projektcode | BESTANDEN |
| LI-11 | Gewachsene Listen 14 bis 16: Telefonliste (11 Einträge, verbundene Abteilungszellen an jedem Eintrag, „SB“ als „Sachbearbeitung“, ausgeblendete Spalte nicht übernommen), Veranstaltungsplanung (8 Einträge, Summenzeile und Ideenliste weggelassen, Kosten als Geldbetrag, „s.o.“ als Wert darüber, „kostenlos“ gemeldet, Ja/Nein aus „ja!“, „(ja)“, „x“), Fortbildungsbedarf (drei Ämter mit Rückfragen zu „Mitarbeiter/in“, „Wunsch“, „Kosten“; „dito“ und „siehe oben“ aufgelöst; 11 Einträge) | BESTANDEN (Screenshots 30 bis 32) |
| LI-12 | Zweite echte 2ask-Befragung (6 Fragen, 9 Bögen): kodiert, mit Umkodierung (110 fehlende Antworten leer) und unkodiert je 27 Felder ohne Hinweise; Formularfelder als Datum („3.3.26“, „2026-03-05“) und Zahl („ca. 20“, „25,5“) erkannt; Filterfrage leer, wo nicht gezeigt; Feldnamen aus langen Fragen gekürzt, Frage im Hilfetext. Dabei gefunden und behoben: Beim unkodierten Export mit Codeplan blieben alle Auswahlfragen leer (87 Hinweise) | BESTANDEN (Screenshot 33) |
| LI-13 | Standardsatz mit allen bisherigen Testlisten: 19 Tabellen übernommen, Wochenraster wie vorgesehen abgelehnt. Gegenüber 1.7.1 gewollt geändert: Kosten in Liste 07 jetzt Geldbetrag mit Rückfrage zu drei Werten, „k.A.“ ohne Rückfrage leer, in Liste 10 „Beschluss“ als Bezeichnungsfeld | BESTANDEN |
| V-53 | Aufnahmen 1.8.0 ohne Entwicklungstag (Zeitstempel der 2ask-Bögen auf den 10.03.2026 verschoben) | ERLEDIGT |
| P-2 | Lasttest 1.8.0 (20.000 Datensätze): alle Zielwerte erfüllt, u. a. Tabellenansicht 2,28 s, Filter 1,42 s, Dashboard 8,90 s (Ziel 10 s), Kalender 2,97 s; große Liste „Liste übernehmen“ (5.000 Zeilen) in 14,1 s übernommen | BESTANDEN (Leistung_Ergebnis.txt) |
| S-3 | Sicherheitsprüfung (`build\sicherheit.ps1`) auf Quelltext und fertiger Masterdatei 1.8.0 | BESTANDEN (Sicherheitspruefung.md) |
| S-4 | Lieferdateien ohne Spuren des Entwicklungsrechners: Speicherort in allen Excel-Dateien, Benutzerprofil im VBA-Verweis auf MSForms.exd (Streams dir, __SRP_0 und _VBA_PROJECT), Windows-Benutzername in Änderungsprotokoll und Systemspalte „_Benutzer“, dazu alle Textdateien des Pakets. Dabei gefunden und behoben (betraf auch frühere Versionen): alle drei Spuren in den Excel-Dateien; jetzt `build\entspuren.ps1` beim Bau und als Sperre beim Paket, bereinigte Dateien in Excel geöffnet und im Selbsttest geprüft | BESTANDEN (0 Spuren) |
| T-3 | Texte ohne Gedankenstriche (U+2013, U+2014) in Quellcode, Dokumentation und Prüfberichten des Pakets | BESTANDEN (0 Treffer) |

**Offen (bewusst nicht umgesetzt):** Farbe und Durchstreichen als Bedeutung (Liste 14: „gelb = nicht im Dienst“) werden nicht übernommen; ein Raum „2.01“, den Excel schon beim Tippen zum Datum gemacht hat, kommt als „02.01.2026“ an; das Gemeindeverzeichnis mit 16.000 Zeilen braucht rund 2,5 Minuten.

## Prüfungen der Version 1.7.1 (November 2026)
**Automatisiert:** 84 von 84 bestanden (78 in `RunSelfTest` sowie T01, T17, T18, T19, T24 und T25) auf den ausgelieferten Dateien. Dazu der **Testlauf „Liste übernehmen“** mit allen Testlisten, die **Sichtprüfung** in echtem Excel (34 Fensteraufnahmen, neu 28 bis 34) und die **Sicherheitsprüfung**. Selbsttest und Sicherheitsprüfung liefen auf den ausgelieferten Dateien, Testlauf und Sichtprüfung auf dem Stand unmittelbar davor (einzige spätere Änderung: das Beispieldatum der Selbsttestprüfung T03d). Der Lasttest wurde nicht wiederholt: Die Wartungsversion ändert weder Datenhaltung noch Seitenaufbau großer Tabellen (Leistungsbericht von 1.7.0 gilt weiter).

| ID | Prüfung | Status |
|---|---|---|
| LI-6 | Feld „Bemerkung“ (als Text erkannt) in Schritt 4 auf „Mehrfachauswahl“ umgestellt: rechts erscheint eine Häkchenliste mit allen Werten des Felds. Dem Wert „Modul 2 Nachholtermin?“ zwei vorhandene Werte und der neue Wert „Rückfrage“ zugeordnet; Ergebnis „Modul 2 Nachholtermin?; Nachmeldung; Rückfrage“, „Rückfrage“ steht danach in der Häkchenliste | BESTANDEN (Screenshots 28 und 29) |
| LI-7 | Zeilenumbrüche in einer Zelle trennen Werte wie das Semikolon | BESTANDEN (Codeprüfung; die übrigen Testlisten ergeben dieselben Tabellen wie in 1.7.0) |
| LI-8 | Meldung „Übernommen“ nach der Übernahme: Der Assistent schließt zuerst, Excel zeichnet neu und holt sein Fenster nach vorn, erst dann erscheint die Meldung | BESTANDEN (Codeprüfung; der seltene Fall aus einer echten 2ask-Übernahme ließ sich im Testlauf nicht auslösen) |
| LI-9 | Testlauf mit allen Testlisten: 19 Tabellen übernommen, Wochenraster wie vorgesehen abgelehnt; große Liste mit 5.000 Zeilen 9,4 s lesen, 10,2 s prüfen, 16,6 s übernehmen (mit Sicherung) | BESTANDEN |
| V-50 | Design „Mainz“ angewendet: Start, Tabelle „Aufgaben“, Dashboard (Balken und Linien in Mainz Gelb, Grün, Blau, Orange), Einrichten mit den neuen Farbmustern „Mainz hell“ und „Mainz Gelb“, Auswahl „Dokumentstil“ mit elf Vorlagen. Dabei gefunden und behoben: (1) Die Auswahl „Dokumentstil“ wurde mit elf Vorlagen höher als der Bildschirm; sie bekommt dann eine Bildlaufleiste. (2) Unter Einrichten lief eine lange Tabellenbeschreibung über den Tabellenrand hinaus; sie endet jetzt am Rand wie in Listen | BESTANDEN (Screenshots 30 bis 34) |
| T25-a | Selbsttest T25 scheiterte in einem von drei Läufen mit Excel-Fehler 1004 („FreezePanes kann nicht festgelegt werden“) beim Seitenaufbau nach dem Löschen. Dabei gefunden und behoben: Das Lösen der Fixierung beim Seitenaufbau ist jetzt abgesichert (die Seite setzt sie am Ende ohnehin neu). Danach 84 von 84 | BESTANDEN |
| V-51 | Aufnahmen und Berichte ohne Entwicklungstag: Stand des Dashboards in der Aufnahme auf den Tag der Beispielwelt gesetzt; das Beispieldatum der Selbsttestprüfung T03d liegt jetzt im März 2026 statt im Oktober 2026 | ERLEDIGT |
| T-2 | Texte ohne Gedankenstriche (U+2013, U+2014) in Quellcode, Dokumentation und Prüfberichten des Pakets | BESTANDEN (0 Treffer) |
| S-2 | Sicherheitsprüfung (`build\sicherheit.ps1`) auf Quelltext und fertiger Masterdatei 1.7.1 | BESTANDEN (Sicherheitspruefung.md) |

## Prüfungen der Version 1.7.0 (November 2026)
**Automatisiert:** 84 von 84 bestanden: 78 in `RunSelfTest` (neu L01 bis L17) sowie T01, T17, T18, T19, T24 und T25, auf den ausgelieferten Dateien. Dazu der **Testlauf „Liste übernehmen“** mit allen Testlisten (`build\listimport_test.ps1`; Ergebnis und übernommene Tabellen in `03_Pruefung`), die **Sichtprüfung** des Assistenten in echtem Excel (27 Fensteraufnahmen), der **Lasttest** und die **Sicherheitsprüfung**. Testlauf, Sichtprüfung und Lasttest liefen auf dem Programmstand unmittelbar vor den letzten Textänderungen (zwei Hilfetexte und ein Beispieldatum im Eingabehinweis der Erfassungsdateien); Selbsttest und Sicherheitsprüfung liefen danach auf den ausgelieferten Dateien.

| ID | Prüfung | Status |
|---|---|---|
| L01 bis L17 | Listenerkennung und Übernahme im Selbsttest: Titel, Leer- und Summenzeile; zweizeiliger Kopf mit Kreuzspalten als Mehrfachauswahl; Gruppenzeilen als Feld; Typvorschläge (Datum auch „12.3.25“, Geld, Postleitzahl mit führender Null als Text, Telefon, E-Mail, Nummer); nicht passender Wert („KW 14“) mit Rückfrage und Hinweis; Abkürzung „erl.“; mehrere Werte je Zelle; Monate als Spalten; Formular; gleich aufgebaute Blätter mit „Anzahl“ und „Menge“; CSV mit Dezimalpunkt und ISO-Datum; Raum „1.12“ bleibt Text; Wochenraster abgelehnt; Optionsnamen ohne Komma; eigene Änderung bleibt beim Neuplanen; 2ask mit Codeplan; Übernahme in eine neue Tabelle mit Optionsreihenfolge und Bezeichnungsfeld | BESTANDEN (automatisiert) |
| LI-1 | Testlauf mit 13 Verwaltungslisten (14 Tabellen): Teilnehmerliste mit zweizeiligem Kopf und Kreuzspalten, Fristenliste mit Abkürzungen und „KW 14“, Haushaltsliste mit Gruppen- und Summenzeilen, Kita-Belegung mit Monaten als Spalten, Adressliste mit Gruppen, IT-Inventar, Maßnahmenplan, Bedarfsmeldung auf 4 Blättern, zwei Tabellen auf einem Blatt, Beschlusskontrolle, CSV-Export eines Fachverfahrens, Antrag als Formular mit Kostenplan. Jede Tabelle übernommen; Felder, Typen, Optionen und Werte in `Listimport_Tabellen.txt` nachgelesen | BESTANDEN |
| LI-2 | Wochenraster (Raumbelegung): erkannt und mit verständlicher Meldung abgelehnt | BESTANDEN (Screenshot 27) |
| LI-3 | 2ask: Nachbau mit 24 Fragebögen und zwei echte Exporte aus dem 2ask-Testtarif (mit und ohne Umkodierung von -99, -98, -97): 20 Fragen, 43 Felder, Codes als Optionen in der Reihenfolge des Codeplans, Mehrfachauswahl aus Einzelspalten, fehlende Antworten leer | BESTANDEN (Screenshots 19 bis 25) |
| LI-4 | 2ask-Datei als „Sonstige Liste“ gewählt: MidPoint schlägt den 2ask-Modus vor. Unkodierter Export (75 Spalten) als sonstige Liste übernommen | BESTANDEN (Screenshot 26) |
| LI-5 | Große Liste mit 5.000 Zeilen und 12 Spalten: 9,1 s lesen und auswerten, 9,8 s prüfen, 15,4 s übernehmen (mit Sicherung). Dabei gefunden und behoben: vorher 27,2 s, 14,6 s und 31,5 s (Ursachen und Maßnahmen im Leistungsbericht) | BESTANDEN |
| V-47 | Assistent Schritt für Schritt mit der Teilnehmerliste: Art und Datei, Tabelle mit Vorschau („Kopf in Zeile 4 und 5, 1 Summenzeile“), drei Rückfragen, Umschalten auf „Je Spalte ein Ja/Nein-Feld“ und zurück, Felder in der Reihenfolge der Liste, Werte von „Verpflegung“ mit „Veg.“ als „vegetarisch“, Feld abgewählt und umbenannt, Prüfergebnis mit Hinweis, Übernahme, fertige Tabelle „Anmeldungen“. Dabei gefunden und behoben: (1) Zusammengefasste Kreuzspalten und Spalten mit mehreren Werten je Zelle standen als erstes Feld statt an ihrer Stelle. (2) Das Prüfergebnis meldete „entschuldigt“ in einer Kreuzspalte als Wert, der nicht zum Feldtyp passt und leer bleibt; jetzt „übernommen; kein Kreuz, gilt als nicht angekreuzt“. (3) Die Vorschau schnitt Texte hart ab; jetzt passend zur Spaltenbreite mit „…“. (4) Beispielwerte zeigten „erledigt, Erledigt“ doppelt, obwohl beide zusammengeführt werden | BESTANDEN (Screenshots 01 bis 11) |
| V-48 | Werte zusammenlegen („in Bearb.“ als „in Arbeit“), Monate als Spalten, Blätter zusammen mit Rückfrage zu „Anzahl“ und „Menge“, Formular mit Kostenplan | BESTANDEN (Screenshots 12 bis 18) |
| V-49 | Aufnahmen ohne Entwicklungstag und ohne Pfad des Entwicklungsrechners: Dateiname „MidPoint_Demo.xlsm“, Testlisten in einem neutralen Ordner, Sicherungszeit und Untertitel der übernommenen Tabellen auf einen Tag der Beispielwelt gesetzt. Die Zeitstempel der drei 2ask-Testbögen sind in den Testdateien auf einen Tag der Beispielwelt verschoben | ERLEDIGT |
| T-1 | Texte ohne Gedankenstriche: Suche nach U+2013 und U+2014 in Quellcode, Dokumentation und Prüfberichten des Pakets | BESTANDEN (0 Treffer) |
| P-1 | Lasttest 1.7.0 in zwei Läufen: alle Zielwerte erfüllt, Dashboard 9,3 s (Ziel 10 s). Dabei gefunden und behoben: Filter 3,15 s und Zeitverlauf 4,70 s nach einer Änderung an `cMap.SetVal`; die Änderung gilt jetzt nur noch für Importzeilen, danach 1,6 s bzw. 1,4 s | BESTANDEN (Leistungsbericht) |
| S-1 | Sicherheitsprüfung (`build\sicherheit.ps1`) auf Quelltext und fertiger Masterdatei: 12 Punkte bestanden, A8 (personenbezogene Daten) als Hinweis. Dabei angepasst: Die Vorlage wertete jedes „FollowHyperlink“ als Netzzugriff. Alle vier Stellen geprüft: Sie öffnen nur auf ausdrücklichen Klick eine eingetragene Webadresse, E-Mail-Adresse, Datei oder einen Ordner mit dem zugeordneten Windows-Programm, bei Adressen und Dateien erst nach Rückfrage. Sie stehen jetzt mit Zweck im Projektblock; jede andere Stelle bliebe NICHT BESTANDEN | BESTANDEN (Sicherheitspruefung.md) |

## Prüfungen der Version 1.6.1 (Oktober 2026)
**Automatisiert:** 67 von 67 bestanden (61 in `RunSelfTest`, neu C18, sowie T17, T18, T19, T01, T24, T25). **Sichtprüfung** in echtem Excel per Fensteraufnahme (PrintWindow) jeder Seite. Geprüft wurden 100 % Bildschirmskalierung bei Excel-Zoom 75, 100, 150 und 200 % sowie die Bildschirmskalierungen laut Zeilen V-41 bis V-43.

| ID | Prüfung | Status |
|---|---|---|
| V-36 | Alle Seiten (Start, fünf Tabellen, Dashboards, Einrichten, Hilfe, Technik) bei Zoom 100: Navigation, Titel, Status und Knöpfe ohne Überlappung. Dabei gefunden und behoben: Auf Hilfe/Technik standen 90 px leere Fläche unter den Knöpfen. | BESTANDEN (Screenshots 01 bis 04) |
| V-37 | Excel-Zoom 75 / 150 / 200: Die Seite baut sich neu auf, die Navigation bleibt 184 pt breit, Spalte A liegt genau darunter, die Knopfreihen brechen um. Hilfe/Technik passen die Textbreite an. Dabei gefunden und behoben: Bei Zoom 150 blitzte rechts neben der Navigation 1 px des Auswahlrahmens hervor; die Navigation deckt Spalte A jetzt mit 5 px Reserve. Der Untertitel brach bei Zoom 200 zweizeilig um und wurde abgeschnitten; jetzt einzeilig mit „…“. | BESTANDEN (Screenshots 05 bis 08) |
| V-38 | Seitenkopf Hilfe/Technik: Das Logo steht zwischen Untertitel und Status und verdeckt weder „Speichern“ noch „Ungespeichert“; ohne Platz entfällt es. | BESTANDEN (Screenshots 03, 04) |
| V-39 | Breite Tabelle „Veranstaltungen“ (10 Spalten): Ziehen der Bildlaufleiste ganz nach rechts, Antippen der Leiste und Mausrad nach rechts. Die Ansicht bleibt ruhig am rechten Inhaltsrand und springt nicht zurück nach links. | BESTANDEN |
| V-41 | **175 % Bildschirmskalierung** (Notebook-Höhe), alle Seiten bei Zoom 75/100/150, dazu das Erklärfenster „So funktioniert's“ und das Formular „Neuer Eintrag“. Die Kurven der Navigation sind vollständig, die Bildlaufpfeile erreichen „Hilfe & Sicherung“ und „Technische Dokumentation“, Kopfzeile und Formulare passen auf den Bildschirm. Dabei gefunden und behoben: (1) Der aktive Eintrag konnte nach einer Höhenänderung halb unter dem Pfeilstreifen liegen und wird jetzt immer nachgeführt. (2) Neben den Pfeilstreifen standen halbe Einträge; sie werden jetzt ausgeblendet. (3) Der fette Eintrag „Technische Dokumentation“ war rechts abgeschnitten; die Schrift passt sich jetzt an. (4) Bei Zoom 75 reichten Karten und „Speichern“ unter die Bildlaufleiste; UsableWidth schließt die Leiste ein, sie wird jetzt abgezogen. | BESTANDEN (Screenshots 09 bis 11) |
| V-42 | **150 % Bildschirmskalierung**: Navigation ohne Bildlauf, alle Seiten bei Zoom 100/150. Dabei gefunden und behoben: Die Kennzahlkarten des Dashboards schnitten bei Zoom 150 Werte („70.100,0“) und Titel ab. Jetzt wird der Wert verkleinert und der Titel mit „…“ gekürzt. Eine erste Lösung mit Mindestbreite verwarf ich, weil die vierte Karte schon bei Zoom 100 über den Rand ragte. | BESTANDEN (Screenshots 12, 13) |
| V-43 | **125 % Bildschirmskalierung**: alle zehn Seiten bei Zoom 100, Projekte und Dashboard bei Zoom 150. Keine Überlappung, kein Abschneiden. | BESTANDEN (Screenshots 14, 15) |
| V-34 | (aus 1.6.0) Untere Navigationseinträge bei 175 % unerreichbar: siehe V-41. | **ERLEDIGT in 1.6.1** |
| V-44 | Navigation bei geöffnetem Menüband, Klicks auf „Einrichten“, „Hilfe & Sicherung“ und „Technische Dokumentation“: Die Markierung saß zu tief, der Text des aktiven Eintrags war unsichtbar, der Schalter fehlte. Ursache gefunden: Die Fensteraufnahme (frisch gezeichnet) war korrekt, nur der Bildschirm zeigte Reste eines älteren Bildes; ein von außen erzwungenes Neuzeichnen behob es. Behoben: vollständiges Neuzeichnen nach jedem Bildwechsel. Nachprüfung: alle drei Einträge sofort und nach 1,5 s sauber. | BESTANDEN |
| V-45 | Seitenwechsel-Tempo (Demo, 100 %): Start, Dashboards, Einrichten, Hilfe und Technik wurden bei jedem Besuch neu aufgebaut, weil Excel die Fensterbreite je Blatt verschieden meldet (1424/1440 pt). Jetzt zählt die Pixelbreite des Rasterfensters. Gemessen beim zweiten Besuch: Dashboard 766 → 68 ms, Einrichten 348 → 64 ms, Start 517 → 63 ms, Hilfe 202 → 64 ms. Tabellen teilen sich ein Blatt und werden beim Wechsel neu befüllt (250 bis 470 ms). | BESTANDEN |
| V-46 | Fehlanzeige: Erfassung „Kinder mit Förderbedarf 2026“ mit 6 Kitas, 2 Rückläufe. Klick auf „Grundschule am Park“ öffnet das Menü „Fehlanzeige eintragen: … hat nichts zu melden“. Danach „3 von 6 haben geantwortet (davon 1 Fehlanzeige)“, auch in der Liste „3 von 6“. Erneuter Klick → „Fehlanzeige zurücknehmen“ → wieder 2 von 6. | BESTANDEN (Screenshot 16) |
| C18 | Fehlanzeige „Schule am Park“: geantwortet 1 → 2 (davon 1 Fehlanzeige), zurückgenommen → 1. | BESTANDEN (automatisiert) |
| V-40 | Unterkante der Navigation bei 100 %: Die Pixelmessung zeigt, dass die Navigation bis an Excels Bildlaufleiste reicht (Zeile 980, Leiste ab 978). Kein heller Spalt, keine Fixierungslinie. | BESTANDEN (Pixelmessung) |

## Prüfungen der Version 1.6.0 (Oktober 2026)
**Automatisiert:** 66 von 66 bestanden. Das sind 60 in `RunSelfTest` (neu C16, C17) sowie T17, T18, T19, T01, T24 und T25. Die Sichtprüfung lief in echtem Excel bei **175 % Bildschirmskalierung**.

| ID | Prüfung | Status |
|---|---|---|
| C16 | Dieselbe Erfassungsdatei von einer zweiten Stelle („Jugendamt“) wird angenommen (3 Einträge); dieselbe Stelle ein zweites Mal wird übersprungen (C06). | BESTANDEN (automatisiert) |
| C17 | Erfassung fehlt in der Hauptdatei (nicht gespeichert): Sie wird aus der Rücklaufdatei wiederhergestellt; der Rücklauf ist danach lesbar. | BESTANDEN (automatisiert) |
| V-28 | Navigation bei 175 %: Breite passt zu Spalte A, MidPoint-Symbol oben links. Waagerechtes Scrollen endet am Inhalt, nichts rutscht unter die Navigation. Dabei gefunden und behoben: Die Excel-Funktion PointsToScreenPixels lieferte bei 175 % falsche Breiten; jetzt zählt die DPI des Excel-Fensters. Ausgeblendete Spalten allein verhinderten das Scrollen nicht; jetzt begrenzt ein Bildlaufbereich. | BESTANDEN (Screenshot 01) |
| V-29 | Erfassungsdateien: Erklärfenster beim ersten Besuch; Ablaufbild und Liste passen ins schmale Fenster; Rückmeldungen unter der Liste. | BESTANDEN (Screenshot 02) |
| V-30 | Assistent „Kinder mit Förderbedarf 2026“: Schritt 1 und „Wer füllt aus?“ mit 6 Kitas (bzw. Tabelle „Einrichtungen“ mit 12 Einträgen). Die Datei wird erzeugt und die Hauptdatei automatisch gespeichert; nach dem Schließen und Wiederöffnen ist die Erfassung vorhanden. Das Absenderfeld der Datei hat eine Auswahlliste (=L_ABS). | BESTANDEN (Screenshots 03 bis 05) |
| V-31 | Vier Antwortdateien eingelesen (Sonnenblume, Regenbogen, Wirbelwind, Sonnenblume doppelt): Spaltenköpfe über den Spalten, Doppel erkannt („doppelt ausgewählt“), 4 neue Einträge; Rückmeldungen „3 von 6“, die offenen zuerst. Dabei gefunden und behoben: Bisher wurde dieselbe Datei nur von der ersten Stelle angenommen (Duplikatprüfung nur über die Datei-ID). Eine doppelt ausgewählte Antwort wäre zweimal übernommen worden. | BESTANDEN (Screenshots 06, 08) |
| V-32 | Sicherungsfrage: Alle drei Schaltflächen liegen im Fenster, der Text ist klar. | BESTANDEN (Screenshot 07) |
| V-33 | Fall „gehört zu einer anderen Datei“ nachgestellt (Erfassung gelöscht, Antwort eingelesen): Die Rückfrage erklärt die Ursache, „Wiederherstellen und einlesen“ funktioniert. | BESTANDEN |
| V-34 | Navigation bei geringer Fensterhöhe (175 %, Notebook-Bildschirm): Die untersten Einträge („Hilfe & Sicherung“, „Technische Dokumentation“) liegen weiter knapp unter der Bildlaufleiste, und der Bildlaufpfeil der Navigation ist nicht zu sehen. Ursache: Das Formular skaliert bei 175 % etwas anders als Excel. Die Fensterhöhe zieht jetzt die Bildlaufleiste ab; das genügt noch nicht. Umgehung: Excel-Fenster maximieren und Menüband einklappen; „Hilfe“ ist auch auf der Startseite erreichbar. | OFFEN (nächste Version) |
| V-35 | Erzeugte Erfassungsdatei und Excel-Export enthalten kein Logo und keinen Programmnamen; die Infozeile des Exports nennt keine Quelldatei mehr. | NUR CODE GEPRÜFT (keine Logo-Aufrufe in mCapture/mIO) |

## Prüfungen der Version 1.5.0 (Oktober 2026)
**Automatisiert:** 64 von 64 bestanden. Das sind 58 in `RunSelfTest` (darunter neu D01, D02, D03, E01, C14, C15) sowie T17, T18, T19, T01, T24 und T25.

| ID | Prüfung | Status |
|---|---|---|
| D01 | Neue Kennzahlen am Dashboard „Auswertungen“: Median Aufwand 13,40 und 8 verschiedene Projekte stimmen mit Nachrechnung aus den Daten überein; alle 11 Elemente ohne Fehler berechnet. | BESTANDEN (automatisiert) |
| D02 | Dashboard zeichnet Netz (82), Kreis (5), Fläche (1), Pareto (Säulen und Linie), Trichter (123) und Baumdiagramm (117). | BESTANDEN (automatisiert, Screenshots 11, 12) |
| D03 | Sankey „Status → Priorität“: Die Summe aller Flüsse entspricht der Anzahl der Aufgaben (67 = 67), 4 Ziele, 13 Bänder als Formen gezeichnet. Dabei gefunden und behoben: Einträge ohne Angabe rechts fehlten (66 statt 67); sie erscheinen jetzt als „Ohne Angabe“. | BESTANDEN (automatisiert, Screenshot 34) |
| E01 | Excel-Export zweier Tabellen: zwei Blätter, Kopfzeile in der gewählten Farbe (Petrol), Schrift Garamond. | BESTANDEN (automatisiert) |
| C14 | Prüfdatei mit Mehrfachauswahl zur Anzeige: Ankreuzspalten (vorhandene Zeilen gesperrt, neue ankreuzbar), Band „Bestehende Einträge“ oben, Trennblock (Abstand + Band „Neue Einträge“) wird beim Einlesen übersprungen, Empfänger in neuen Zeilen vorbelegt; neue Zeile mit Haken und Empfänger übernommen. | BESTANDEN (automatisiert) |
| C15 | Mehrfachauswahl „höchstens 2“: Programm lehnt 3 Werte ab, 2 werden angenommen. Erzeugte Prüfdatei: Kopfzeile „höchstens 2 ankreuzen“, Ankreuzregel vorhanden; Excel nimmt den zweiten Haken an und erkennt den dritten als ungültig. Rücklauf mit 3 Haken wird beim Einlesen nicht übernommen. Screenshot 33 zeigt die Datei mit roter Markierung. Dabei gefunden und behoben: Die Regel wurde aus VBA nur in englischer Schreibweise angenommen, und „Leere Zellen ignorieren“ hätte jede Eingabe zugelassen. | BESTANDEN (automatisiert) |
| C09 bis C13 | Bisherige Prüfdatei-Rundläufe mit der neuen Trennzeile | BESTANDEN (automatisiert) |
| V-16 | Assistent in Excel: zwei Tabellen gewählt, Zweck „Prüfen“. „Darstellung“ wird ausgeblendet; „Anzahl leerer Zeilen“ erscheint erst nach dem Häkchen; Farbe Waldgrün und Schrift Garamond gewählt; das zweite Blatt „Ressourcen“ wird direkt abgefragt. Die erzeugte Datei hat zwei Blätter, den getrennten Bereich „Neue Einträge“, „Workshop“ vorbelegt und Themen als Ankreuzspalten. Dabei gefunden und behoben: Mehrfachlisten verloren beim Schließen des Dialogs ihre Auswahl. | BESTANDEN (Screenshots 01 bis 08) |
| V-17 | Export-Dialog: „Mehrere Tabellen“ blendet die Tabellenliste ein. Export mit drei Blättern, Petrol und Garamond, getönten Zeilen. Dabei gefunden und behoben: Fehler 5 bei getönten Zeilen (Formel musste in die Excel-Sprache übersetzt werden). | BESTANDEN (Screenshots 09, 10) |
| V-18 | Trichter und Baumdiagramm: Zuerst entstanden Ersatzbalken bzw. leere Diagramme. Behoben: Die Diagramme werden aus markierten Daten angelegt und zeigen ausgeblendete Hilfsspalten an. | BESTANDEN (Screenshot 12) |
| V-19 | Tabelle „Ressourcen“: Langer Text der letzten Spalte endet an ihrem Rand. | BESTANDEN (Screenshot 13) |
| V-20 | Formular: Zahlen (Menge, Mindestbestand, Stückpreis) linksbündig. | BESTANDEN (Screenshot 14) |
| V-21 | Dialog „Element“: Feld für die Kennzahl nur bei Bedarf; Dialog „Neue Ansicht“ und „Rangliste“ im Export | NUR CODE GEPRÜFT |
| V-22 | Dashboard → Gestaltung: Farbstil mit Farbleiste, „Ändern…“ öffnet die Auswahl mit 10 Stilen und je 8 Farben. Bordeaux und Nordisch (pastell) gewählt: Vorschau sofort im Dashboard, „Abbrechen“ stellt die alten Farben her. Dabei gefunden und behoben: Der Button verdeckte die Farbleiste; große Kennzahlen waren bei pastelligen Stilen zu blass. | BESTANDEN (Screenshots 30, 31) |
| V-24 | Erfassungsdatei (Prüfen mit Ankreuzspalten): graues Band „Bestehende Einträge (10)“ mit grauem Streifen, Abstand, farbiges Band „Neue Einträge“ mit farbigem Streifen und wiederholten Spaltenköpfen; Titel und Anleitung brechen über die Kopfbreite um. Dabei gefunden und behoben: Pflichtfeld „Status“ war in leeren neuen Zeilen rot (vorbelegter Empfänger zählte als begonnen). Zusätzlich angesehen: reine Sammelliste mit Absenderfeld. | BESTANDEN (Screenshot 36) |
| V-25 | Liste „Projekte“: Budget, Beginn und Ende linksbündig; Schalter „Zeilenumbruch“ zeigt Beschreibung und Aufgaben vollständig, Zeilen wachsen mit. | BESTANDEN (Screenshot 35) |
| V-26 | Dialog „Element“: „Welches Diagramm passt?“ öffnet die Übersicht mit 14 Darstellungen (Bilder mit Beispieldaten). Doppelklick auf „Sankey“ setzt die Darstellung. Dabei gefunden und behoben: Fenster war zu hoch (Bildlauf, Titel abgeschnitten); „(wie Dashboard)“ zeigte eine Farbleiste. | BESTANDEN (Screenshot 37) |
| V-27 | Einrichten → Dokumentstil: 10 Vorlagen mit Farben und Schriftprobe. „Bordeaux elegant“ setzt Garamond und Bordeaux in Navigation, Seiten und Dialogen; „Nordisch hell“ färbt das Dashboard pastellig mit Aptos. Dabei angepasst: Bordeaux und Erde nutzen jetzt gedeckte Akzente (Weinrot, Terrakotta) statt Rot/Orange. | BESTANDEN (Screenshots 35, 38, 39) |
| V-23 | Eigenschaft „Themen“ → Weitere Einstellungen → „Höchstens auswählbar“ = 3 gespeichert. Formular: 4 Haken zeigen den Hinweis rot, Speichern wird verhindert. Ein Haken entfernt: Markierung verschwindet, Speichern mit 3 Themen gelingt. Dabei gefunden und behoben: Nach der roten Markierung waren die Haken nicht mehr sichtbar, und der rote Rahmen blieb nach der Korrektur stehen (jetzt roter Rahmen hinter der Liste). | BESTANDEN (Screenshot 32, erneut geprüft) |

## Bedienversuch mit Excel-unerfahrener Person
**NICHT AUSGEFÜHRT**: keine Testperson verfügbar. Ersatzweise dokumentierter Bedienablauf in echter Excel-Laufzeit: Vorlage anlegen → Eigenschaften (Button, ID, Dateien) hinzufügen → Einträge erfassen → Button ausführen → Ansicht/Board ändern → Erfassungsdatei erzeugen, ausfüllen und einlesen → Dashboard prüfen. Das ersetzt keinen echten Nutzertest.

## Bekannte Einschränkungen
siehe `Versionshinweise.md` und `Leistungsbericht.md`.
