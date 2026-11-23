# Hilfe zu MidPoint (aus der Anwendung exportiert)

## In zehn Schritten zur eigenen Datei

- 1. Die leere Masterdatei im Windows-Explorer kopieren und die Kopie sinnvoll umbenennen (z. B. „Materialverwaltung.xlsm“). Die Masterdatei selbst bleibt unverändert.
- 2. Die Kopie öffnen und, falls Excel fragt, „Inhalt aktivieren“ wählen. Ohne aktivierte Makros erscheint nur ein Hinweisblatt.
- 3. Oben links auf den Titel klicken (oder auf der Startseite „Titel der Datei festlegen“) und der Datei einen Namen geben, z. B. „Inventar Kitas“.
- 4. „Neue Tabelle“ (oder „Aus Vorlage“) wählen, Namen, Beschreibung und Symbol festlegen und „Tabelle anlegen“ klicken. Gibt es die Liste schon in Excel oder als 2ask-Befragung, macht „Liste übernehmen“ daraus eine fertige Tabelle mit Feldern und Einträgen.
- 5. Eigenschaften (Spalten) hinzufügen: Name eingeben, Art per Klick wählen (wie in Notion: Text, Nummer, Auswählen, Status, Datum, Person, Kontrollkästchen, Button, ID …) und Einstellungen wie Pflichtangabe, Hilfetext oder Auswahlmöglichkeiten festlegen. Mit „Vorschau der Eingabemaske“ das Ergebnis prüfen.
- 6. Bei Bedarf weitere Tabellen anlegen und über ein Feld vom Typ „Verknüpfung“ verbinden (z. B. Aufgabe → Projekt). Die Gegenseite erscheint automatisch.
- 7. In die Tabelle wechseln (linke Navigation) und mit „Neuer Eintrag“ Daten erfassen. Pflichtfelder sind mit * markiert.
- 8. Über „Ansicht“ Sortierung, Gruppierung und Spalten festlegen, über „+ Ansicht“ weitere Darstellungen anlegen (Tabelle, Kalender, Board, Projektleiste) und mit „Als Ansicht speichern“ sichern.
- 9. Unter „Dashboards“ ein Dashboard anlegen und Elemente hinzufügen (Kennzahl, Diagramm, Tabelle), ohne Formeln.
- 10. Oben rechts „Speichern“ (oder Strg+S) und unter „Hilfe & Sicherung“ eine erste Sicherung erstellen. Wer möchte, schaltet unten links „Automatisch speichern“ ein; dann wird jede Änderung sofort gespeichert.

> Tipp: Die Datei ist danach eine eigenständige Fachdatei mit Programmcode, Einstellungen und Daten. Sie braucht die Masterdatei nicht mehr.

## Aufbau der Oberfläche

Links befindet sich die feste Navigation: oben der Titel der Datei (Klick ändert ihn), Start, Ihre Tabellen, Dashboards, Erfassungsdateien sowie unten Einrichten, Hilfe & Sicherung und Technische Dokumentation. Ganz unten der Schiebeschalter „Automatisch speichern“: grün = ein (jede Änderung wird sofort in der Datei gespeichert), grau = aus (Speichern wie gewohnt mit „Speichern“ bzw. Strg+S).

Es gibt keinen getrennten Umbaumodus: Daten pflegen und Tabellen umbauen geschieht in derselben Oberfläche, wie in Notion. Die Farbe der Navigation und die Akzentfarbe wählen Sie unter „Einrichten“ per Klick auf ein Farbfeld, kräftig oder hell (ruhige Pastelltöne wie Flieder, Salbei, Himmel; dann mit dunkler Schrift). „Dokumentstil“ legt Farben, Farbstil und Schrift der ganzen Datei mit einer von 11 Vorlagen fest, oder Sie stellen sie selbst zusammen. Die Vorlage „Mainz“ ist angelehnt an das Erscheinungsbild der Stadt Mainz: Navigation in hellem Gelb, Akzent Gelb, Diagramme in Gelb und Hellgrün, Schrift Calibri.

Vergrößern und Verkleinern: mit dem Zoom-Regler unten rechts in Excel oder Strg + Mausrad. Die Seite baut sich passend neu auf, die Navigation behält ihre Größe. Der gewählte Zoom gilt für alle Seiten.

Oben rechts zeigt das Statusband, ob automatisch gespeichert wird, die letzte erfolgreiche Dateispeicherung und die letzte vorhandene Sicherung. „Ungespeicherte Änderungen“ bedeutet: Daten wurden geändert, die Datei aber noch nicht gespeichert.

Jede Tabelle hat oben Register für ihre Ansichten. Darunter liegen Suchfeld, Filter, Ansicht (Sortieren, Gruppieren, Spalten, Speichern), Export und die Aktionen für den markierten Eintrag. Aktive Filter erscheinen als Kacheln; ein Klick auf eine Kachel entfernt genau diesen Filter.

## Tägliche Arbeit

### Eintrag anlegen und bearbeiten

- „Neuer Eintrag“ öffnet die Erfassungsmaske. Sie wird automatisch aus den Feldern der Tabelle aufgebaut; Register entstehen aus den Feldgruppen. Automatisch Berechnetes (Formeln, Rollups, Erstellt/Geändert, ID) zeigt die Maske nicht; es steht in der Tabelle; die ID erscheint im Kopf der Maske.
- Selteneres erscheint erst auf Klick: „+“ neben einer Auswahl legt eine neue Option an, „Mehr“ unten bietet Duplizieren und Archivieren.
- Doppelklick auf eine Zeile oder „Bearbeiten“ öffnet den markierten Eintrag. Kurze Werte können auch direkt in der Tabellenzelle geändert werden, mit denselben Prüfregeln.
- „Speichern“ prüft alle Register. Bei einem Fehler bleibt das Fenster mit allen Eingaben offen, das betroffene Feld wird rot markiert und erklärt die Korrektur. Erst wenn alles gültig ist, wird der Eintrag vollständig gespeichert und das Fenster geschlossen.
- „Speichern & neu“ (bei neuen Einträgen) speichert und öffnet sofort eine leere Maske. „Mehr › Duplizieren“ legt eine Kopie an (eindeutige Werte werden nicht übernommen).
- „Abbrechen“ verwirft nach Rückfrage die Änderungen; gespeicherte Daten bleiben unverändert.

> Wichtig: „Speichern“ im Eintrag übernimmt die Daten in die Arbeitsmappe. Dauerhaft auf der Festplatte gesichert ist erst, wenn die Datei gespeichert wird (Statusband „Speichern“, Strg+S).

### Datum und Uhrzeit

Datumsfelder akzeptieren TT.MM.JJJJ (auch TT.MM.JJ). „Datum…“ öffnet einen Kalender mit Heute, Morgen und +7 Tage; bei Feldern mit Uhrzeit lassen sich Stunden und Minuten eintragen (Uhrzeiten werden ohne Sekunden angezeigt). „Übernehmen“ ändert nur das Eingabefeld; gespeichert wird mit „Speichern“.

### Mehrere Zellen einfügen

Werte aus der Zwischenablage können in mehrere Zellen der Tabellenansicht eingefügt werden. Alle Werte werden zuerst geprüft. Ist auch nur ein Wert ungültig, wird nichts übernommen; eine Liste zeigt Zeile, Feld und Grund. Berechnete Felder, Verknüpfungen und interne IDs sind gegen Überschreiben geschützt.

### Sammeländerung

Mehrere Zeilen markieren (Umschalt- oder Strg-Taste) und „Sammeländerung“ wählen. Nach Feld und neuem Wert zeigt eine Vorschau alle betroffenen Einträge; erst danach wird übernommen.

## Eigenschaften (Spalten) wie in Notion

Spalten werden direkt in der Tabelle angelegt und geändert, ohne Moduswechsel.

- Neue Spalte: das + rechts neben dem letzten Spaltenkopf. Name eingeben, Art anklicken, Speichern. Hilfetext, Register, Standardwert, Eindeutigkeit und Dublettenprüfung liegen unter „Weitere Einstellungen“. Die Spalte erscheint sofort in der Tabelle und in der Erfassungsmaske.
- Spalte ändern: Doppelklick auf den Spaltenkopf > „Eigenschaft bearbeiten…“ (Name, Hilfetext, Pflicht, Optionen, Farben, Standardwert) oder „Neue Eigenschaft rechts davon…“. „Eigenschaft entfernen…“ prüft vorher alle Abhängigkeiten und sichert automatisch.
- „Eigenschaften“ öffnet die Gesamtliste aller Eigenschaften der Tabelle mit Reihenfolge, Ausblenden und Vorschau der Eingabemaske.
- Tabelle löschen: in der Tabelle „Ansicht > Tabelle … löschen…“, im Dialog „Eigenschaften“ unten „Tabelle löschen…“ oder unter Einrichten (Zeile anklicken, „Tabelle löschen…“). Abhängigkeiten werden geprüft, die Auswirkungen angezeigt, zur Bestätigung wird LÖSCHEN eingegeben; vorher wird automatisch gesichert. Ohne Datenverlust: „Aus-/Einblenden“ unter Einrichten.

### Verfügbare Arten

- Text, Nummer, Geldbetrag, Prozent, Langer Text / Notiz: mit Längen- und Wertgrenzen.
- Auswählen, Mehrfachauswahl, Status: farbige Optionen; neue Optionen auch direkt in der Maske („+“).
- Mehrfachauswahl mit Grenze: unter „Weitere Einstellungen“ „Höchstens auswählbar“ eintragen (z. B. 6 Möglichkeiten, höchstens 3). Gilt in der Maske, beim Einfügen, beim Import und in Erfassungsdateien.
- Datum, Datum mit Uhrzeit: mit Kalender und „Heute“; Regel „nicht vor“ (z. B. Ende nicht vor Beginn).
- Person: gepflegte Personenliste (z. B. Windows-Kennungen), Mehrfachauswahl; „+“ in der Maske legt eine neue Person an.
- Dateien und Medien: mehrere lokale Dateien je Eintrag (nur Verweise, keine Kopien); Öffnen nach Rückfrage.
- Kontrollkästchen: Häkchen; in der Tabelle schaltet ein Doppelklick auf die Zelle um.
- URL, E-Mail, Telefon, Ort: mit Formatprüfung; Links und E-Mail-Programm werden nur nach bewusster Rückfrage geöffnet (die Anwendung selbst bleibt offline).
- Formel: eine Excel-Formel wie gewohnt, Spalten stehen in eckigen Klammern: =[Preis]*[Menge], =WENN([Frist]<HEUTE();"überfällig";""), =RUNDEN([Kosten]/[Stunden];2). Deutsche Funktionsnamen und Semikolon. Spalten einfügen per Klick: im Eigenschaftsdialog auf die Spaltenknöpfe, in der Formelleiste (Schaltfläche „fx“ im Tabellenkopf oder Spaltenmenü „Neue Formelspalte rechts davon…“) direkt auf eine Spalte der Tabelle. Vorschau und Prüfung erscheinen sofort beim Tippen. Rollup wertet verknüpfte Einträge aus (Anzahl, Summe, Min, Max, Durchschnitt); der frühere Assistent ohne Formel bleibt erreichbar.
- Verknüpfung: Beziehung zu einer anderen Tabelle (n:1 oder n:m) mit automatischer Gegenseite.
- Erstellt um, Erstellt von, Letzte Änderung um, Zuletzt bearbeitet von: werden automatisch gepflegt.
- Button: setzt per Klick bis zu drei Felder (z. B. Status = Erledigt, Ende = heute, Bearbeiter = ich). In der Maske als Schaltfläche, in der Tabelle per Doppelklick auf die Zelle.
- ID: fortlaufende, nie wiederverwendete Nummer mit optionalem Präfix (z. B. PRJ-12).
- Lokaler Pfad, Ja/Nein (auch leer): zusätzliche Arten für Ablagen und optionale Ja/Nein-Angaben.

> Hinweis: In Erfassungsmasken mit vielen Feldern blättern Sie mit der Bildlaufleiste rechts oder mit der Tabulatortaste; das Mausrad wird von Excel-Formularen nicht unterstützt.

## Erfassungsdateien: Daten von vielen einsammeln

Sie stellen einmal eine Excel-Liste zusammen, schicken sie an alle Beteiligten (z. B. 60 Kitas) und lesen die ausgefüllten Dateien wieder ein. Alle Antworten landen gesammelt in Ihrer Tabelle. Die erzeugte Datei ist eine normale Excel-Arbeitsmappe ohne Makros (.xlsx): Empfänger brauchen kein Programm und sehen keine Sicherheitswarnung. „So funktioniert's“ zeigt den Ablauf als Bild.

### Der übliche Weg: neue Einträge sammeln

- 1. „Neue Erfassung“: Name, Zieltabelle (in diese Tabelle kommen die Antworten) und „Neue Einträge sammeln“ wählen.
- 2. Spalten zum Ausfüllen wählen. Berechnete und automatische Eigenschaften werden nicht abgefragt.
- 3. Wer füllt aus? Am besten „Aus einer festen Liste wählen“ (ein Name je Zeile) oder „Aus einer Tabelle“ (z. B. alle Einrichtungen). Dann wählen die Empfänger ihren Namen aus einer Liste; alle schreiben ihn gleich, und Sie sehen später, wer noch fehlt.
- 4. Überschrift und kurzer Hinweis. Farbstil, Schrift, Zeilenzahl und Blattschutz unter „Weitere Einstellungen“.
- 5. „Datei erzeugen…“: Es entsteht eine Datei für alle. MidPoint speichert danach die Hauptdatei, damit die Rückläufe zugeordnet werden können. Die Datei per E-Mail an alle schicken.
- 6. Ausgefüllte Dateien zurückbekommen und mit „Rücklauf einlesen…“ übernehmen (mehrere auf einmal). Die Vorschau zeigt je Datei Absender, neue, geänderte und fehlerhafte Zeilen.
- 7. „Rückmeldungen“ zeigt, wer geantwortet hat (Haken mit Datum) und wer noch fehlt („offen“). „Teilnehmende…“ ändert die Liste jederzeit.
- 8. Meldet eine Stelle ohne Datei, dass sie nichts beizutragen hat (z. B. per Mail): in den „Rückmeldungen“ auf „offen“ klicken und „Fehlanzeige eintragen“ wählen. Die Stelle zählt dann als geantwortet; ein weiterer Klick nimmt die Fehlanzeige zurück.

### Erweitert: vorhandene Einträge prüfen lassen

- Ihre bestehende Liste wird vorbelegt. Graue Spalten sind nur zur Anzeige (z. B. Inventarnummer, Bezeichnung), gelbe werden geprüft bzw. ausgefüllt. Mit „Aufteilen nach“ (z. B. Kita) entsteht je Wert eine eigene Datei mit nur diesen Einträgen. Der Rücklauf aktualisiert die Einträge über ihre versteckte ID, nur geänderte Werte.
- In Prüfdateien stehen oben die bestehenden Einträge unter einem grauen Band, nach einem Abstand ein farbiges Band „Neue Einträge“ mit leeren Zeilen. Wurde ein Eintrag seit dem Versand bei Ihnen geändert, fragt MidPoint beim Einlesen nach.
- Mehrere Tabellen in einer Datei: im ersten Schritt „Erweitert: weitere Tabellen“ ankreuzen oder später „Weiteres… › Weiteres Blatt hinzufügen“.

### Gut zu wissen

- Mehrfachauswahl erscheint als Ankreuzspalten: je Option eine schmale Spalte; „Ja“ wählen oder leer lassen. Hat die Eigenschaft eine Höchstzahl, nennt die Kopfzeile sie; weitere Haken nimmt Excel nicht an.
- Doppelte Antworten: Jede Stelle wird je Datei einmal übernommen (Datei und Absender zählen). Schickt eine Stelle dieselbe Antwort zweimal, wird sie übersprungen.
- Wurde die Hauptdatei nach dem Erzeugen ohne Speichern geschlossen, kennt sie die Erfassung nicht. Beim Einlesen bietet MidPoint dann an, die Erfassung aus der Rücklaufdatei wiederherzustellen.

> Fehlerhafte Zeilen werden nicht übernommen; eine Datei mit Fehlern gilt erst nach vollständiger Übernahme als eingelesen. Lassen Sie die Datei korrigieren und lesen Sie sie erneut ein. Das „Rücklaufprotokoll“ zeigt, welche Dateien wann übernommen wurden.

## Suchen, Filtern, Sortieren, Gruppieren

- Suche: Begriff in das umrandete Suchfeld tippen und Enter oder „Suchen“. Gesucht wird in den sichtbaren Spalten; mehrere Wörter müssen alle vorkommen. Die Suche verändert keine gespeicherte Ansicht.
- Filter: „Filter“ öffnet Bedingungen passend zum Feldtyp (Text enthält, Zahl zwischen, Datum im Zeitraum, Auswahl ist, leer/nicht leer, verknüpft mit …). Bedingungen lassen sich mit „und“/„oder“ kombinieren; „und“ bindet stärker.
- „Alle Filter zurücksetzen“ im Filterdialog oder „Ansicht zurücksetzen“ stellt den gespeicherten Zustand der Ansicht wieder her.
- Sortieren und Gruppieren über „Ansicht“. Ein Doppelklick auf eine Spaltenüberschrift öffnet das Spaltenmenü: sortieren, gruppieren, filtern, in dieser Ansicht ausblenden, Eigenschaft bearbeiten, rechts daneben einfügen oder entfernen.
- Zeilenumbruch (Werkzeugleiste): aus = eine Zeile je Eintrag; an = lange Texte vollständig, die Zeilen wachsen mit. Wird je Ansicht gemerkt. Alle Werte stehen linksbündig.
- Beim Öffnen der Datei werden Suche, zusätzliche Filter, Sortierung und Gruppierung der Sitzung zurückgesetzt. Gespeicherte Ansichten und die Spaltenauswahl bleiben erhalten.

## Ansichten

Eine Tabelle kann beliebig viele gespeicherte Ansichten haben. Ansichten sind keine Datenkopien: Eine Änderung in einer Ansicht ändert den Eintrag und ist sofort in allen anderen Ansichten und Auswertungen sichtbar.

- Tabelle: Spalten, Reihenfolge, Breite, Zusammenfassungen (Summe, Durchschnitt …). Die Bezeichnung steht immer in der ersten Spalte. Es gibt bewusst keine Fensterfixierung (keine Fixierlinien).
- Kalender: Monatsansicht oder Terminliste. Doppelklick auf einen Termin öffnet ihn, Doppelklick auf einen freien Tag legt dort einen Eintrag an. Mehrtägige Einträge (Enddatum unter „Ansicht bearbeiten“) erscheinen als durchgehende Leiste in einer festen Spur; überlappende Einträge liegen untereinander. Heute ist farbig umrandet. Bei mehr als vier Terminen erscheint „+ n weitere“. Einträge ohne Datum sind über „Ohne Datum“ erreichbar.
- Board: Spalten nach Status oder Auswahlfeld mit Anzahl je Spalte. Karte anklicken und mit „Nach links/rechts“ oder „Gruppe ändern…“ verschieben; Doppelklick öffnet. Einträge ohne Gruppe stehen in einer eigenen Spalte.
- Projektleiste (Gantt): tagesgenaue Leisten von Beginn bis Ende in Tagen, Wochen oder Monaten; die rote Linie markiert heute. „Namen auf Leisten“ zeigt die Namen dezent auf der Leiste (passt ein Name nicht, steht er rechts daneben). Klick auf eine Leiste öffnet den Eintrag. Ein Ende vor dem Beginn ist ungültig. Unvollständige Einträge sind über „Ohne vollständigen Zeitraum“ erreichbar.

> Ansichtsregister sind nach Art farbig (Tabelle blau, Board orange, Kalender grün, Projektleiste lila). In Tabellen, Board und Kalender erfolgen Änderungen per Klick und Dialog; Ziehen mit der Maus gibt es im Dashboard („Anordnen“).

## Verknüpfungen

Ein Feld vom Typ „Verknüpfung“ verbindet Einträge zweier Tabellen (z. B. Aufgabe → Projekt). Gespeichert wird über interne IDs; Namensänderungen zerstören keine Zuordnung. In der Gegentabelle erscheint automatisch ein Rückverknüpfungsfeld (z. B. „Aufgaben“ beim Projekt).

- „Auswählen…“ öffnet eine Suchliste mit Zusatzmerkmalen, damit gleichnamige Einträge unterscheidbar sind. „Neu anlegen…“ erstellt einen Zieleintrag, ohne den begonnenen Eintrag zu verlieren.
- „Öffnen“ zeigt einen verknüpften Eintrag; der eigene Entwurf bleibt dabei geöffnet.
- Archivierte verknüpfte Einträge bleiben erkennbar („archiviert“). Beim endgültigen Löschen werden nur die Zuordnungen entfernt, nie andere Einträge (kein kaskadierendes Löschen).

## Berechnungen

Ein „Berechnetes Feld“ wird über Auswahl aufgebaut: Funktion wählen, dann Felder oder feste Werte zuordnen. Verfügbar sind Grundrechenarten, Runden, Prozentanteil, Wenn-dann-sonst mit UND/ODER, Datumsabstände, Tage bis Frist, Überfälligkeit, Texte verbinden sowie Auswertungen verknüpfter Einträge (Anzahl, Anzahl mit Bedingung, Summe, Minimum, Maximum, Durchschnitt).

- Leere Werte ergeben ein leeres Ergebnis und werden nicht still als 0 gezählt. Division durch 0 erscheint als „Fehler: Division durch 0“.
- Zirkelbezüge werden beim Einrichten erkannt und verhindert. Jede Berechnung speichert eine lesbare Beschreibung.
- Bei verknüpften Einträgen zählt jeder Eintrag genau einmal; archivierte sind ausgenommen. Felder mit „heute“ (Tage bis Frist, Überfällig) werden beim ersten Öffnen eines Tages neu berechnet.

## Dashboards und Auswertungen

Ein Element besteht aus Titel, Datenquelle, Kennzahl (Anzahl, ausgefüllt/leer, verschiedene Werte, Summe, Durchschnitt, Median, Minimum, Maximum, Spannweite, Anteil), optionaler Gruppierung (bei Datumsfeldern nach Monat, Quartal oder Jahr), optionalem Zeitraumfeld und Filtern. Darstellung als Kachel, Balken, Säulen, Zeitverlauf, Fläche, Anteil (Ring), Kreis, Netz (ab 3 Gruppen), Trichter, Baumdiagramm, Pareto, Rangliste, Sankey oder Tabelle. Trichter und Baumdiagramm brauchen Excel 2016/2019 oder neuer, sonst erscheint ein Balkendiagramm.

- „Welches Diagramm passt?“ (neben „Darstellung“) zeigt alle Darstellungen als Bild mit Beispieldaten und wofür sie sich eignen. Klick wählt, Doppelklick übernimmt.
- Sankey: zeigt Flüsse von einer Gruppierung (links, „Von“) zu einer zweiten (rechts, „Nach“), z. B. Status → Priorität. Die Bandbreite entspricht Anzahl oder Summe; rechts höchstens 8 Gruppen, der Rest als „Weitere“.
- Der Dashboardfilter (Zeitraum und/oder Tabellenfilter) wirkt nur auf passende Elemente; nicht betroffene Elemente sind gekennzeichnet.
- Ein Klick auf eine Kennzahl oder Tabellenzeile bzw. „Details“ zeigt die zugrunde liegenden Datensätze.
- Mehrfachzuordnungen: Ein Datensatz kann mehreren Gruppen angehören. Die Summe der Gruppen kann dann größer sein als die eindeutige Gesamtzahl; das Element weist darauf hin. Prozentanteile werden nie addiert.
- Dashboards werden beim Aufruf und nach jeder Änderung neu berechnet; „Stand“ zeigt den Zeitpunkt. „Aktualisieren“ berechnet sofort neu.
- Diagramm kopieren: Im Menü „…“ eines Elements (oder Klick auf die Karte) „Diagramm kopieren (für PowerPoint, Word)“ wählen, bei Kacheln und Tabellen entsprechend. Danach in PowerPoint oder Word mit Strg+V einfügen. Das Element kommt als Bild mit Titel, Farben und Schrift wie im Dashboard an und bleibt beim Vergrößern scharf; Menüknopf, „Details“ und Quellenhinweis werden nicht mitkopiert.
- Vorschau: Beim Anlegen oder Bearbeiten eines Elements wird es sofort im Dashboard gezeigt (blau umrandet); danach „Übernehmen“, „Nochmals bearbeiten“ oder „Verwerfen“. Selteneres (Datumsstufe, Zeitraumfeld, Reihenfolge) liegt unter „Weitere Einstellungen“.
- Gestaltung: Stil der Karten (Karten, Flach, Farbakzent, Kontur) und Farbstil. „Ändern…“ zeigt 11 Stile mit ihren acht Farben (Akzentfarbe, Ozean, Wald, Schiefer, Bordeaux, Erde, die Pastellstile Nordisch, Flieder, Salbei & Sand, Kontrast sowie Mainz mit Gelb und Hellgrün). Jede Auswahl wird sofort hinter dem Dialog gezeigt. Einzelne Elemente können einen eigenen Stil haben.
- Anordnen: Das Dashboard hat ein festes Raster aus vier Spalten (hell hinterlegt). Karte ziehen und loslassen: Sie springt sofort in die nächste Spalte und Zeile („Einrasten“ ordnet bei Bedarf erneut). Breite am Rand ziehen (¼, ½, ¾ oder ganze Breite), Diagramme unten ziehen für hoch/normal. Eine Karte in eine volle Zeile gezogen verdrängt die dortige nach unten. „Fertig“ speichert genau die gezeigte Anordnung, „Abbrechen“ verwirft sie.

## Import und Export

- Import (Einrichten > Importieren): lokale XLSX- oder CSV-Datei wählen, Vorschau prüfen, Spalten zuordnen (Typvorschläge prüfen), dann „Prüfen“. Fehlerhafte Zeilen werden mit Grund gemeldet. Übernommen werden nur gültige Zeilen, und erst nach Ihrer ausdrücklichen Bestätigung. Abbrechen ändert nichts.
- Aktualisieren statt anhängen: einen eindeutigen Schlüssel wählen (z. B. Inventarnummer). Zeilen mit vorhandenem Schlüssel aktualisieren den Eintrag.
- Export: in jeder Ansicht „Export“, als XLSX oder CSV (UTF-8, Semikolon). Beim XLSX-Export wählbar: diese Ansicht, mehrere Ansichten der Tabelle oder mehrere Tabellen (je ein Blatt), sowie „Gestaltet“ (Titel, Filterangabe, farbige Kopfzeile, Linien, wahlweise getönte Zeilen; Farbe und Schrift wählbar) oder „Schlichte Liste“ (nur Werte). Exportiert wird, was die Ansicht zeigt, ohne interne Felder.
- CSV-Schutz: Texte, die mit =, +, - oder @ beginnen, erhalten ein vorangestelltes Apostroph, damit andere Programme sie nicht als Formel ausführen.
- Dashboards lassen sich als XLSX (Werte und neu erzeugte Diagramme) oder PDF exportieren. Exporte enthalten keine Makros und keine Verbindung zur Arbeitsdatei.

## Liste übernehmen (vorhandene Listen und 2ask-Befragungen)

„Liste übernehmen“ (auf Start und unter Einrichten) macht aus einer vorhandenen Liste eine MidPoint-Tabelle, auch wenn die Liste nicht aus MidPoint stammt. Der Assistent hat fünf Schritte. Übernommen wird erst am Ende nach Ihrer Bestätigung, vorher wird automatisch gesichert. Die Quelldatei wird nur gelesen.

- 1. Art und Datei: „Befragung aus 2ask“ für einen kodierten Excel-Export mit Codeplan (in 2ask: Ergebnisse, Ergebnisse herunterladen, „Kodierte Ergebnisse/Antworten“, Format Excel und Haken bei „Codeplan hinzufügen“) oder „Sonstige Liste“ für jede andere Excel- oder CSV-Liste. Erkennt MidPoint in einer sonstigen Liste einen 2ask-Export, bietet es den passenden Modus an.
- 2. Tabelle: MidPoint sucht alle Tabellen der Datei, auch mehrere auf einem Blatt, und erkennt Titel, Kopfzeilen (auch über zwei Zeilen), Gruppen-, Summen- und Notizzeilen. Die Vorschau zeigt links, wie jede Zeile verstanden wird. Gleich aufgebaute Blätter (z. B. ein Blatt je Einrichtung) stehen zusätzlich als eine gemeinsame Tabelle zur Wahl; der Blattname wird dann ein Feld. Stimmt die Kopfzeile nicht, lässt sie sich von Hand setzen. Bei 2ask zeigt dieser Schritt die erkannten Fragen.
- 3. Rückfragen: Wo MidPoint nicht sicher ist, fragt es nach und schlägt eine Antwort vor, z. B. Kreuzspalten als eine Mehrfachauswahl, Gruppenzeilen als Feld, Summenzeilen weglassen, Monatsspalten in je einen Eintrag pro Monat umwandeln, Abkürzungen wie „erl.“ als „erledigt“, Werte, die nicht zum Typ passen („KW 14“ in einer Datumsspalte), Formelspalten und ausgeblendete Spalten. Jede Antwort wirkt sofort.
- 4. Ziel und Felder: neue Tabelle oder an eine bestehende anhängen (gleich benannte Felder werden befüllt, mit Schlüssel werden vorhandene Einträge aktualisiert). Je Feld lassen sich Name, Typ und „Übernehmen“ ändern und das Bezeichnungsfeld wählen. Bei Auswahlfeldern zeigt die Werteliste jeden Wert mit Anzahl; gleiche Bedeutungen lassen sich zusammenlegen („in Bearb.“ als „in Arbeit“). Wird ein Feld auf „Mehrfachauswahl“ umgestellt, ordnet eine Häkchenliste jeden Wert beliebig vielen Werten zu (z. B. „Kinder und Jugendliche“ zu „Kinder“ und „Jugendliche“); neue Werte lassen sich dort ergänzen. Komma, Semikolon und Zeilenumbrüche in einer Zelle trennen Werte bereits beim Lesen.
- 5. Prüfung: zeigt, wie viele Einträge übernommen werden, fehlerhafte Einträge und Hinweise zu einzelnen Werten (z. B. ein Wert, der nicht zum Typ passt und leer bleibt, oder ein Eintrag in einer Kreuzspalte, der kein Kreuz ist). Der Bericht lässt sich als CSV speichern und bleibt unter Einrichten > Letzter Importbericht verfügbar.
- 2ask: Antwortcodes werden über den Codeplan in Optionen übersetzt (Reihenfolge wie im Fragebogen). Die Einzelspalten einer Mehrfachauswahl werden ein Feld, Matrixzeilen eigene Felder, Rangfolge, Sterne, Schieberegler und Summen werden Nummern. Umkodierte fehlende Antworten (z. B. -99) bleiben leer. Von den Angaben zur Teilnahme sind Teilnahme-ID und Ende vorgewählt; Fragebögen aus dem Pretest lassen sich weglassen.

> Nicht unterstützt: Wochenraster (Uhrzeiten in den Zeilen, Wochentage in den Spalten), Zellfarben als Bedeutung und automatische Verknüpfungen zwischen zwei übernommenen Tabellen. Ein Wochenraster bitte vorher als Liste mit einer Zeile je Termin anlegen.

## Archiv, Löschen und Rückgängig

- „Archivieren“ blendet einen Eintrag aus allen aktiven Ansichten und Auswertungen aus; Verknüpfungen bleiben. „Ansicht > Archiv anzeigen“ zeigt archivierte Einträge mit „Wiederherstellen“.
- Endgültig löschen ist nur im Archiv möglich, mit Referenzprüfung, Bestätigungswort LÖSCHEN und angebotener Sicherung.
- „Rückgängig“ nimmt die letzte zusammenhängende Datenänderung dieser Sitzung zurück (bis zu 20 Schritte, z. B. Speichern eines Eintrags, Einfügen, Sammeländerung, Import, Archivieren). Strukturänderungen (Felder, Tabellen) sind davon ausgenommen; dafür gibt es Sicherungen. Nach dem Schließen der Datei ist die Rücknahme nicht mehr verfügbar.

> Hinweis: Excels eigenes Rückgängig (Strg+Z) wird durch Makros geleert. Verwenden Sie die Schaltfläche „Rückgängig“ der Anwendung.

## Speichern, Sicherung und Wiederherstellung

- Automatisch speichern (Schalter unten links): Ein = jede Änderung wird kurz danach gespeichert; aus = wie gewohnt. Die Einstellung gilt für diese Datei.
- Speichern: Statusband „Speichern“ oder Strg+S. „Gespeichert …“ wird erst nach erfolgreichem Speichern angezeigt. Bei fehlenden Schreibrechten oder nicht erreichbarem Netzlaufwerk erscheint eine klare Fehlermeldung; die Änderungen bleiben in Excel geöffnet.
- Manuelle Sicherung: „Jetzt sichern“. Ablage: Unterordner „<Dateiname>_Sicherungen“ neben der Datei. Dateiname: <Dateiname>_JJJJ-MM-TT_hhmmss_v<Version>.xlsm.
- Automatische Sicherung: beim regulären Speichern (höchstens alle 60 Minuten, einstellbar) sowie vor Import, Wiederherstellung, Zurücksetzen, Entfernen von Feldern/Tabellen und Feldtypwechsel.
- Aufbewahrung: Standardmäßig bleiben die letzten 10 Sicherungen erhalten (einstellbar). Ältere eigene Sicherungen werden entfernt, andere Dateien nie.
- Jede Sicherung wird nach dem Schreiben geprüft. Eine fehlgeschlagene Sicherung wird gemeldet und im Statusband angezeigt; sie erscheint nie als erfolgreich.
- Wiederherstellen: „Sicherungen & Wiederherstellen“ zeigt alle Sicherungen. Vor dem Wiederherstellen wird der aktuelle Stand gesichert. Übernommen werden Struktur, Daten, Verknüpfungen, Ansichten, Dashboards, Einstellungen und Protokoll; danach wird der Bestand geprüft. Anschließend die Datei speichern.
- Neue leere Kopie: erzeugt eine neue Datei mit derselben Struktur ohne Einträge. Die geöffnete Datei bleibt unverändert.
- Datei zurücksetzen: löscht alle Einträge und Verknüpfungen (Struktur und Protokoll bleiben), nur nach erfolgreicher Sicherung und dem Bestätigungswort ZURÜCKSETZEN.

> Eine Sicherung schützt den Stand zum Sicherungszeitpunkt. Änderungen danach sind erst mit der nächsten Speicherung/Sicherung geschützt. Exporte sind keine Sicherung.

## Einrichtungsschutz

Daten pflegen und Struktur ändern geschieht in derselben Oberfläche. Wer Strukturänderungen (Tabellen, Eigenschaften, Dashboards) absichern möchte, setzt unter Einrichten ein Schutzwort: Es wird vor der ersten Strukturänderung einer Sitzung abgefragt. Das ist ein Schutz gegen Fehlbedienung, keine Benutzerverwaltung und keine Zugriffskontrolle.

## Makros, Sicherheit und gemeinsame Nutzung

- Die Anwendung benötigt aktivierte Makros. Sicherheitseinstellungen werden von der Datei nie verändert. Üblich ist die Freigabe über einen vertrauenswürdigen Speicherort oder eine digitale Signatur durch die IT (siehe Technische Dokumentation).
- Alles läuft offline: keine Cloud, keine Anmeldung, keine Telemetrie, keine Downloads. Es entstehen nur bewusst erzeugte Exporte und Sicherungen.
- Eine Person bearbeitet die Datei zur gleichen Zeit. Liegt sie auf einem Netzlaufwerk und ist bereits geöffnet, öffnet Excel sie schreibgeschützt; die Anwendung zeigt dann den Lesemodus und sperrt Bearbeitungen.
- Blattschutz und ausgeblendete Blätter verhindern nur Fehlbedienung. Für vertrauliche Daten den Speicherort (Berechtigungen) und bei Bedarf den Office-Kennwortschutz der Datei nutzen.
- Das Änderungsprotokoll enthält Zeitpunkt, Windows-Kennung, Art und betroffene Felder, aber keine Feldinhalte. Die Windows-Kennung ist kein sicherer Identitätsnachweis.

## Datenzugang

Unter Einrichten > „Datenzugang“ werden die strukturierten Datenblätter (Tabellen D_T…, Konfiguration, Zuordnungen, Protokoll) schreibgeschützt eingeblendet, z. B. zur Prüfung oder für eigene Auswertungen in einer Kopie. Änderungen dort sind gesperrt, weil sie die Prüfregeln umgehen würden. Mit „Datenblätter ausblenden“ wieder schließen.

