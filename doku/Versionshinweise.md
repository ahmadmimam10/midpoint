# Versionshinweise für MidPoint

## Version 1.8.2 (November 2026, Formatversion 3)

Wartungsversion: Modul 4 der Lernreihe ist neu und ausführlicher, dazu eine Korrektur beim Anlegen von Erfassungsdateien und ein bereinigtes Lieferpaket für die Veröffentlichung als Open Source. Keine Änderung am Datenformat.

### Lernreihe
- **Modul 4 neu, jetzt in vier Videos** und ruhiger erklärt: **4.1 Das Board**, **4.2 Kalender und Zeitleiste** (eine vierte Tabelle „Arbeiten“ mit 17 Arbeiten von der Pinnwand, Kalender mit Farben je Art, Zeitleiste mit Heute-Linie), **4.3 Das Dashboard** (drei Kacheln, Balken, Ring und Säulen, die Einträge hinter einer Zahl, Filter für das ganze Dashboard, Gestaltung, Diagramm kopieren, PDF) und **4.4 Daten einsammeln**. Die Lernreihe hat jetzt 17 Videos.
- **Bildprüfung aller Videos:** Aus jedem fertigen Video wurden alle Markierungen, Klicks, Tippstellen und Ausklapplisten in voller Auflösung herausgezogen und geprüft. Korrigiert: Markierungen neben dem Ziel oder außerhalb des Bildausschnitts, Tippen, das mitten im Wort begann oder Buchstaben anschnitt, Mauszeiger, die vor dem Klick schon weiterzogen, Auswahlen, die schon vor dem Klick gewählt waren, und Bilder, die nicht zum gesprochenen Satz passten.
- **Mitbauen:** Bauplan für Modul 4 mit vier Abschnitten, neue Datei `Arbeiten_Pinnwand.xlsx` für Video 4.2, Startdateien und fertige Datei auf diesem Stand.

### Korrekturen
- **Erfassungsdateien hingen bei manchen Standarddruckern:** War ein Drucker wie „OneNote (Desktop)“ als Standard eingestellt, blieb das Anlegen der Erfassungsdateien bei der Seiteneinrichtung stehen, weil Excel bei jeder Einstellung den Drucker abfragt. Die Seiteneinrichtung läuft jetzt gebündelt ohne Druckerabfrage (Erfassungsdateien und Exporte).
- **Farbstil-Auswahl:** Jede dritte Beschreibung (Wald, Erde, Salbei & Sand) erschien eine Stufe kleiner, weil der Zeilenabstand keine ganzen Bildpunkte ergab. Jetzt sind alle Beschreibungen gleich groß.
- **Kalender, Terminliste:** Mehrtägige Einträge, die im Vormonat beginnen und in den gezeigten Monat hineinreichen, stehen jetzt auch in der Terminliste, wie schon in der Monatsansicht. Die Datumsspalte ist breiter, damit lange Daten wie „Donnerstag, 10. Dezember 2026“ nicht in die nächste Spalte ragen. Das Suchfeld der Terminliste lief unter „Suchen“ und die Anzahl der Einträge; es endet jetzt an der Spalte „Bezeichnung“.
- **Mehrfachlisten in Dialogen:** Bei mehr als neun Einträgen war der letzte Eintrag nur angeschnitten zu sehen und ließ sich nicht ganz hineinrollen, etwa „Spielplatz“ bei den Kurzangaben einer Board-Ansicht. Die Listen zeigen jetzt ganze Zeilen und rollen bis zum letzten Eintrag.
- **Design „Mainz“:** Die Beschreibung nennt nur noch das Erscheinungsbild (Gelb, Hellgrün, Blau und Orange, Schrift Calibri), keine internen Angaben.
- **Testdaten:** Telefonnummern in Testlisten, 2ask-Exporten und Selbsttest verwenden eine erfundene, nicht vergebene Vorwahl.

### Lieferpaket
- **Dateieigenschaften:** Autor und „zuletzt geändert von“ werden in allen ausgelieferten Excel-Dateien geleert, Erstell- und Änderungszeit auf den Monatsanfang gesetzt (`build\entspuren.ps1`). Die Sicherheitsprüfung liest die Eigenschaften dazu direkt aus der Datei (Punkt B4).
- **Lizenz:** MidPoint steht unter der MIT-Lizenz (`LICENSE.txt` im Paket).

### Geprüft
- Selbsttest 88 von 88 bestanden (82 in RunSelfTest, dazu T01, T17, T18, T19, T24 und T25) auf den ausgelieferten Dateien, mit „OneNote (Desktop)“ als Standarddrucker.
- Sicherheitsprüfung bestanden (`IT und Datenschutz/Pruefung/Sicherheitspruefung.md`).
- Lernvideos: Prüfblätter aus allen 17 fertigen Videos (Markierungen, Klicks, Tippstellen, Ausklapplisten), Funde behoben und neu gerendert; Lautheit −16 LUFS; Aufnahmen ohne Entwicklungstag (Beispielwelt 30.11.2026).
- Lasttest: unverändert aus Version 1.8.0 (die Änderungen betreffen keine gemessenen Abläufe).

## Version 1.8.1 (November 2026, Formatversion 3)

Wartungsversion mit der neuen **Lernreihe** und kleinen Korrekturen, die beim Aufnehmen der Videos aufgefallen sind. Keine Änderung am Datenformat.

### Lernreihe
- **16 neue Lernvideos** (zusammen rund 25 Minuten) in fünf Modulen und einem Extra. Von der leeren Masterdatei an entsteht Schritt für Schritt eine Datei für die Spielplätze eines erfundenen Grünflächenamts: die erste Tabelle mit festen Spaltenarten, Regeln und Formeln (Modul 1), Einträge erfassen und pflegen (Modul 2), Tabellen verknüpfen, eine alte Excel-Liste übernehmen und mitzählen (Modul 3), Board, Kalender, Dashboard und Erfassungsdateien (Modul 4), dazu die Übung „Ihr Fuhrpark“ und eine Ideenkiste (Modul 5) und das Extra „Für IT und Datenschutz“. Die bisherigen Lernvideos entfallen.
- **Lernportal** (`Lernen/Lernportal.html`): alle Videos mit Kapiteln, Suche und Lernwegen („Von null zur eigenen Datenbank“, „Ich trage nur ein“, „Ich sammle Daten ein“, „Ich bin in der IT oder im Datenschutz“), ohne Internet.
- **Mitbauen** (`Lernen/Mitbauen`): je Modul ein Bauplan zum Abhaken (PDF) und eine Startdatei, die alte Baumliste aus Video 3.3, die fertige Datei und die Fuhrpark-Übung mit Lösung. Die Startdateien sind über den Update-Weg (`build\build.ps1 -Update`) auf diese Version gebracht.

### Korrekturen
- **Jahreszahlen ohne Tausenderpunkt** auch bei Feldern, deren Name ein Wort auf „jahr“ ist (z. B. „Pflanzjahr“), wenn sie ohne eigene Grenzen aus „Liste übernehmen“ stammen. Bisher stand dort „1.985“. Der Eingabehinweis in Erfassungsdateien nennt bei solchen Feldern eine Jahreszahl.
- **Quellcode-Fenster:** Der Katalogstempel im Titel nannte Tag und Uhrzeit des Builds. Er nennt jetzt nur Version, Monat und Jahr (betraf auch die Exporte des Codes).
- **Aus Vorlage starten:** Die Beschreibung der Vorlage „Spielplätze & Grünflächen“ war im Fenster abgeschnitten; sie ist jetzt kürzer.
- **Seite Technik:** Der Hinweis auf den vollständigen Prüfbericht nennt den neuen Paketordner „IT und Datenschutz“.

### Lieferpaket
- Neu gegliedert: `MidPoint_Master.xlsm` ganz oben, dann `Vorstellung` (Demodatei und Video „Warum MidPoint?“), `Lernen` (Lernportal, Lernvideos, Mitbauen), `Dokumentation` (Kurzanleitung, Hilfe, Versionshinweise) und `IT und Datenschutz` (Video für IT und Datenschutz, Technische Dokumentation, `Pruefung`, `Quellcode`).

### Geprüft
- Selbsttest 88 von 88 bestanden (82 in RunSelfTest, dazu T01, T17, T18, T19, T24 und T25) auf den ausgelieferten Dateien.
- Sicherheitsprüfung bestanden (`IT und Datenschutz/Pruefung/Sicherheitspruefung.md`).
- Lernvideos: Sprechertext und Bild Satz für Satz abgeglichen, Einsätze der Stimme im fertigen Video gemessen (Abweichung höchstens 0,3 Sekunden, außer am Signet), Lautheit −16 LUFS; Aufnahmen ohne Entwicklungstag (Beispielwelt 30.11.2026).
- Startdateien, fertige Datei und Fuhrpark-Lösung: Änderungsprotokoll und Zeitstempel in der Beispielwelt, ohne Pfade und Namen des Entwicklungsrechners; Baupläne, Lernportal und Videos ebenso durchsucht.
- Lasttest: unverändert aus Version 1.8.0 (die Änderungen betreffen keine gemessenen Abläufe).

## Version 1.8.0 (November 2026, Formatversion 3)

Neue Funktion **„Diagramm kopieren“** im Dashboard und eine gründlich nachgeschärfte Funktion **„Liste übernehmen“**, erprobt mit vier Listen von Behörden, drei typisch gewachsenen Verwaltungslisten und einer zweiten 2ask-Befragung. Keine Änderung am Datenformat.

### Dashboards
- **Diagramm kopieren:** Im Menü „…“ jedes Elements (oder per Klick auf die Karte) „Diagramm kopieren (für PowerPoint, Word)“, bei Kacheln und Tabellen entsprechend. Danach in PowerPoint oder Word mit Strg+V einfügen: Das Element kommt als Vektorbild mit Titel, Farben und Schrift wie im Dashboard an und bleibt beim Vergrößern scharf. Menüknopf, „Details“ und Quellenhinweis werden nicht mitkopiert.

### Liste übernehmen
- **Hinweiszeile unter dem Kopf:** Eine Zeile mit Stichtag („31.12.2025 (Jahr)“), Übersetzung („start date“) oder Unterbegriffen („Schlüssel“, „Bezeichnung“) wird erkannt und nicht mehr als erster Datensatz übernommen. Stichtag und Übersetzung kommen in den Hilfetext des Felds, Unterbegriffe in den Feldnamen („Reisegebiete Schlüssel“).
- **Bessere Feldnamen:** Allgemeine Unterbegriffe bekommen die Gruppenüberschrift („Bevölkerung insgesamt“ statt „insgesamt“), Fußnotenzeichen fallen weg („Fläche km2 1)“ wird „Fläche km2“), Trennstriche am Zeilenende werden zusammengezogen („Umsetzungs-“ und „mechanismen“), sehr lange Überschriften und Fragetexte werden sinnvoll gekürzt („Interventionsbereich (Interventionskategorien …)“ wird „Interventionsbereich“). Der vollständige Text steht im Hilfetext des Felds.
- **Mehrfachauswahl nur bei echten Auswahlwerten:** „Adenau, Stadt“ oder „Bürgerhaus, gr. Saal“ bleiben Text. Bisher entstand daraus eine Mehrfachauswahl mit Hunderten Optionen.
- **Platzhalter** wie „-“, „k. A.“ oder „n. v.“ in Zahlen- und Datumsspalten bleiben leer, ohne die Spalte zu Text zu machen.
- **Beträge und Ja/Nein wie in Verwaltungslisten:** „ca. 1.200 €“, „rd. 300“ und „500,-“ werden Zahlen; „ja!“, „(ja)“ und „Ja.“ gelten als Ja.
- **„dito“ und „siehe oben“:** Rückfrage, ob der Wert der Zeile darüber gemeint ist.
- **Blätter mit leicht abweichenden Spalten** („Kosten“ und „Kosten ca.“) werden zusammen angeboten; abweichende Spaltennamen erscheinen als Rückfrage.
- **Bezeichnungsfeld:** Vorgeschlagen wird eine Namens- oder Titelspalte, nicht mehr eine Spalte mit vielen gleichen Werten (Kreis, Amt), eine Bemerkung oder eine selten ausgefüllte Spalte; eine Kennung nur, wenn es keinen Namen gibt.
- Tabulatoren am Zellende (z. B. „FSP L“ aus einem Fachverfahren) werden entfernt.
- **2ask:** Beim unkodierten Export mit Codeplan werden die Antworttexte den Optionen zugeordnet; bisher blieben alle Auswahlfragen leer. Formular- und Textfelder ohne Typangabe werden als Datum oder Zahl erkannt, wenn die Antworten danach aussehen („3.3.26“, „ca. 20“).
- **Schneller bei vielen Optionen:** Felder mit Hunderten Optionen entstehen in einem Schritt. Gilt auch für „Daten importieren“.

### Testdaten
- Drei neue gewachsene Listen 14 bis 16 (Telefonliste, Veranstaltungsplanung, Fortbildungsbedarf mehrerer Ämter), erzeugt mit `build\testlisten_unprof.ps1`.
- Zweite echte 2ask-Befragung „Bürgerbefragung Bürgerservice“ mit sechs Fragen (Matrix mit Ausweichantwort, Mehrfachauswahl mit Textfeld und exklusiver Antwort, Rangfolge, Filterfrage, Formular, Gruppenmatrix) in `03_Pruefung/Testlisten/2ask_komplex`.

### Lieferdateien
- **Ohne Spuren des Entwicklungsrechners:** Die Excel-Dateien enthielten bisher den Ordner, in dem sie gebaut wurden, das VBA-Projekt den Pfad des Windows-Benutzerprofils (Verweis auf MSForms.exd) und die Demodatei den Windows-Benutzernamen im Änderungsprotokoll. Das betraf auch frühere Versionen. Der Bau entfernt diese Angaben jetzt (`build\entspuren.ps1`), Einträge aus dem Aufbau heißen „Aufbau“, und das Paket entsteht nur, wenn keine Spur übrig ist.

### Geprüft
- Selbsttest 88 von 88 bestanden auf den ausgelieferten Dateien, neu: Diagramm kopieren und drei Prüfungen zu „Liste übernehmen“.
- Praxistest „Diagramm kopieren“: Kachel, Diagramm und Tabelle in PowerPoint und Word eingefügt, Aussehen wie im Dashboard (Aufnahme 39).
- „Liste übernehmen“ mit vier Behördenlisten (Statistisches Bundesamt, Statistisches Landesamt Rheinland-Pfalz, Berliner Senatsverwaltung, Auswärtiges Amt), den drei neuen Listen, der zweiten 2ask-Befragung und allen bisherigen Testlisten. Das Schulverzeichnis Rheinland-Pfalz (1.629 Schulen) braucht jetzt 19 statt 239 Sekunden.
- Lasttest mit 20.000 Datensätzen: alle Zielwerte erfüllt (Dashboard 8,9 s).
- Sicherheitsprüfung bestanden (`03_Pruefung/Sicherheitspruefung.md`).
- Lieferdateien ohne Pfade und Namen des Entwicklungsrechners: alle Excel-Dateien samt VBA-Projekt und alle Textdateien des Pakets durchsucht.

## Version 1.7.1 (November 2026, Formatversion 3)

Wartungsversion mit Wünschen aus der ersten Nutzung von „Liste übernehmen“ und dem neuen Design „Mainz“. Keine Änderung am Datenformat.

### Liste übernehmen
- **Mehrfachauswahl per Häkchen:** Wird ein Feld in Schritt 4 auf „Mehrfachauswahl“ umgestellt, zeigt der Assistent neben der Werteliste eine Häkchenliste. Jeder Wert der Liste lässt sich damit beliebig vielen Werten zuordnen (z. B. „Kinder und Jugendliche“ zu „Kinder“ und „Jugendliche“); unter „Neuer Wert“ lassen sich weitere ergänzen.
- Zeilenumbrüche in einer Zelle (Alt+Eingabe) trennen bei Mehrfachauswahl die Werte, wie Komma und Semikolon.
- **Meldung nach der Übernahme sichtbar:** Die Meldung „Übernommen: …“ erscheint jetzt, wenn der Assistent geschlossen und die neue Tabelle geöffnet ist. Nach einer langen Übernahme (z. B. einer 2ask-Befragung) blieb sie bisher hinter dem Excel-Fenster verborgen, und Excel schien zu hängen.

### Design „Mainz“
- Neue Vorlage **„Mainz“** unter Einrichten > Dokumentstil, angelehnt an das Erscheinungsbild der Stadt Mainz: Navigation in hellem Gelb (20 %), Schaltflächen und Markierungen in kräftigem Gelb mit dunkler Schrift, Schrift Calibri.
- Neuer **Farbstil „Mainz“** für Dashboards, Erfassungsdateien und Export: Gelb und Hellgrün, danach Blau, Orange, Türkis und weitere kräftige Farben.
- Die Farben „Mainz hell“ (Navigation) und „Mainz Gelb“ (Akzent) stehen auch einzeln unter „Eigene Auswahl“ und auf der Seite Einrichten.
- Die Auswahl der Farbstile rollt, wenn das Fenster nicht genug Platz hat.

### Dabei gefunden und behoben
- Unter Einrichten lief eine lange Tabellenbeschreibung über den Rand der Tabellenliste hinaus; sie endet jetzt am Rand.
- Beim Seitenaufbau konnte Excel selten das Lösen der Fixierung ablehnen (Fehler 1004, einmal im Selbsttest nach dem Löschen einer Tabelle). Das ist jetzt abgefangen.

### Geprüft
- Selbsttest 84 von 84 bestanden auf den ausgelieferten Dateien.
- Testlauf „Liste übernehmen“ mit allen Testlisten: 19 Tabellen übernommen, Wochenraster wie vorgesehen abgelehnt; 5.000 Zeilen in 16,6 s übernommen.
- Sichtprüfung in echtem Excel: Häkchenliste bei Mehrfachauswahl und Design „Mainz“ auf Start, Tabelle, Dashboard, Einrichten und in der Auswahl „Dokumentstil“ (Aufnahmen 28 bis 34 in `03_Pruefung/screenshots`).
- Sicherheitsprüfung bestanden (`03_Pruefung/Sicherheitspruefung.md`).

## Version 1.7.0 (November 2026, Formatversion 3)

Neue Funktion **„Liste übernehmen“**: vorhandene Listen und 2ask-Befragungen werden zu MidPoint-Tabellen, auch wenn sie nicht aus MidPoint stammen. Keine Änderung am Datenformat; bestehende Dateien lassen sich weiterverwenden.

### Liste übernehmen
- Neu auf **Start** und unter **Einrichten**: „Liste übernehmen“. Ein Assistent führt in fünf Schritten von der Datei zur fertigen Tabelle: Art und Datei, Tabelle, Rückfragen, Ziel und Felder, Prüfung.
- **Zwei Arten:** „Befragung aus 2ask“ (kodierter Excel-Export mit Codeplan) und „Sonstige Liste“ (jede andere Excel- oder CSV-Liste). Wird eine 2ask-Datei als sonstige Liste gewählt, schlägt MidPoint den passenden Modus vor.
- **Tabellen finden:** MidPoint sucht alle Tabellen der Datei, auch mehrere auf einem Blatt, und erkennt Titel, Kopfzeilen (auch über zwei Zeilen mit verbundenen Zellen), Gruppen-, Summen- und Notizzeilen. Die Vorschau zeigt links, wie jede Zeile verstanden wird. Die Kopfzeile lässt sich von Hand setzen.
- **Gleich aufgebaute Blätter** (z. B. ein Blatt je Einrichtung) lassen sich als eine Tabelle übernehmen. Der Blattname wird ein Feld; abweichende Spaltennamen („Anzahl“ statt „Menge“) werden nach Rückfrage zusammengelegt.
- **Typvorschlag je Spalte:** Datum (auch „12.3.25“ und „2025-01-14“), Datum mit Uhrzeit, Nummer, Geldbetrag, Prozent, Ja/Nein (auch „x“, „J/N“), Auswählen, Mehrfachauswahl, E-Mail, Telefon, URL und langer Text. Nummern mit führenden Nullen (Postleitzahl, Fallnummer) und Raumnummern wie „1.12“ bleiben Text. In CSV-Dateien wird ein Dezimalpunkt („1450.00“) richtig gelesen. Technische Spaltenköpfe wie „FALLNR“ werden lesbar („Fallnummer“).
- **Rückfragen mit Vorschlag**, wo MidPoint nicht sicher ist: Kreuzspalten als eine Mehrfachauswahl, Gruppenzeilen als Feld, Summenzeilen weglassen, Monate als Spalten in je einen Eintrag pro Monat umwandeln (Jahr aus dem Titel), mehrere Werte in einer Zelle („Sport, Gesundheit“), Abkürzungen („erl.“ als „erledigt“, „JHA“ als „Jugendhilfeausschuss“), Werte, die nicht zum Typ passen („KW 14“ in einer Datumsspalte), Formelspalten und ausgeblendete Spalten. Jede Antwort wirkt sofort.
- **Felder anpassen:** Name, Typ, „Übernehmen“ und Bezeichnungsfeld. Die Werte eines Auswahlfelds stehen mit Anzahl in einer Liste und lassen sich zusammenlegen („in Bearb.“ als „in Arbeit“) oder bei Mehrfachauswahl auf mehrere Werte verteilen („Kinder; Jugendliche“).
- **Ziel:** neue Tabelle oder an eine bestehende anhängen; gleich benannte Felder werden befüllt, mit einem Schlüssel werden vorhandene Einträge aktualisiert.
- **Prüfung vor der Übernahme:** Zahl der Einträge, fehlerhafte Einträge und Hinweise zu einzelnen Werten (z. B. ein Wert, der nicht zum Typ passt und leer bleibt, oder „entschuldigt“ in einer Kreuzspalte), jeweils mit Quellzeile. Der Bericht lässt sich als CSV speichern und bleibt unter Einrichten > Letzter Importbericht verfügbar. Vorher wird automatisch gesichert; die Quelldatei wird nur gelesen (schreibgeschützt, Makros gesperrt).
- **2ask:** Antwortcodes werden über den Codeplan zu Optionen in der Reihenfolge des Fragebogens. Die Einzelspalten einer Mehrfachauswahl werden ein Feld, der Zusatztext („Sonstiges, und zwar“) ein eigenes Textfeld, Matrixzeilen und Gruppenmatrix eigene Felder. Skalen zwischen zwei Gegensätzen werden wahlweise Nummern, Rangfolge, Sterne, Schieberegler und Summen werden Nummern, Datumsfelder echte Daten. Umkodierte fehlende Antworten (-99, -98, -97) bleiben leer. Teilnahme-ID und Ende sind vorgewählt, weitere Angaben zur Teilnahme (Browser, Gerät, Dauer) abwählbar; Fragebögen aus dem Pretest lassen sich weglassen.
- **Grenzen:** Ein Wochenraster (Uhrzeiten in den Zeilen, Wochentage in den Spalten) wird erkannt und mit einer verständlichen Meldung abgelehnt. Zellfarben haben keine Bedeutung. Verknüpfungen zwischen zwei übernommenen Tabellen werden danach wie gewohnt angelegt.

### Dabei gefunden und behoben
- „Daten importieren“: Wurde beim Import ein neues Feld der Art E-Mail, Telefon oder URL angelegt, blieben die Werte leer, weil die Prüfung die Feldart nicht kannte. Jetzt werden sie geprüft und übernommen.
- Hilfe-Export (`build\exportdocs.ps1`): Die Überschrift der Hilfe war falsch kodiert.
- **Texte ohne Gedankenstriche:** In Oberfläche, Meldungen, Hilfe, erzeugten Dateien und Dokumentation stehen statt Gedankenstrichen jetzt Doppelpunkt, Komma, Semikolon oder Klammern, Spannen mit „bis“ (z. B. „Einträge 1 bis 50 von 200“, Sortierung „A bis Z“). Leere Kennzahlen zeigen „keine Werte“, leere Summen „Summe (leer)“, leere Felder auf Boardkarten „(leer)“. Exportblätter heißen „Tabelle, Ansicht“. In der Demodatei heißen die Altersgruppen „0 bis 3 Jahre“, „3 bis 6 Jahre“ und „6 bis 12 Jahre“.
- Prüfberichte nennen den Entwicklungsstand nur mit Monat und Jahr (Selbsttest; der Sicherungsname im Test T19 wird ohne Tag und Uhrzeit gezeigt). Der Eingabehinweis für „Datum mit Uhrzeit“ in Erfassungsdateien nennt als Beispiel jetzt den 03.06.2025 statt eines Tages aus der Entwicklungszeit.
- **Schnellerer Import großer Listen**, auch bei „Daten importieren“: Die Prüfung legt die Feldbeschreibung einmal je Spalte an statt für jede Zelle, die Übernahme findet Auswahlwerte über eine vorbereitete Tabelle. Gemessen mit 5.000 Zeilen und 12 Spalten: Prüfen 14,6 → 9,8 s, Übernehmen 31,5 → 15,4 s (mit Sicherung); beim Lesen und Auswerten der Liste 27,2 → 9,1 s.

### Geprüft
- **Selbsttest:** 78 von 78 Prüfungen in `RunSelfTest` (darunter neu L01 bis L17 für „Liste übernehmen“) sowie T01, T17, T18, T19, T24 und T25 bestanden.
- **Testlauf „Liste übernehmen“** (`build\listimport_test.ps1`): 20 Fälle. 19 Tabellen übernommen (13 Verwaltungslisten mit 14 Tabellen, 2ask-Nachbau kodiert und unkodiert, zwei echte 2ask-Exporte, eine Liste mit 5.000 Zeilen), das Wochenraster wie vorgesehen abgelehnt. Felder, Optionen und erste Einträge jeder übernommenen Tabelle stehen in `03_Pruefung/Listimport_Tabellen.txt`.
- **Sichtprüfung** des Assistenten in echtem Excel: 27 Aufnahmen aller fünf Schritte mit sechs verschiedenen Listen und zwei übernommenen Tabellen (`03_Pruefung/screenshots`).
- **Lasttest** mit 10.000 Einträgen: alle Zielwerte erfüllt (`03_Pruefung/Leistungsbericht.md`).
- **Sicherheitsprüfung:** bestanden (`03_Pruefung/Sicherheitspruefung.md`).

### Sicherheitsprüfung vor jedem Release
- Neu: `build\sicherheit.ps1` prüft vor jedem Release Quelltext und fertige Datei (Netz und Internet, Programme und Befehle, Windows-Schnittstellen, COM-Objekte, Zugriff auf den eigenen Code, Sicherheitseinstellungen, Dateizugriffe, personenbezogene Daten, externe Verknüpfungen, Datenverbindungen, ausgeblendete Datenblätter, Dateieigenschaften, Verhalten ohne Makros). Das Ergebnis steht auf der Seite Technik direkt unter dem Programmstand und ganz oben in der Technischen Dokumentation; der Bericht liegt im Paket unter `03_Pruefung`.
- Ergebnis für 1.7.0: bestanden. MidPoint baut selbst keine Verbindung auf. Webadressen, E-Mail-Adressen, Dateien und Ordner öffnet es nur auf ausdrücklichen Klick mit dem zugeordneten Windows-Programm (Browser, E-Mail-Programm, Explorer), bei Adressen und Dateien erst nach Rückfrage; der Bericht nennt diese vier Stellen mit Zweck.

## Version 1.6.1 (Oktober 2026, Formatversion 3)

Wartungsversion für Bildschirmskalierung, Excel-Zoom, sauberes Scrollen und schnellere Seitenwechsel, dazu Fehlanzeigen bei den Rückmeldungen. Keine Änderung am Datenformat.

### Rückmeldungen: Fehlanzeige per Klick
- Meldet eine Stelle ohne Datei, dass sie nichts beizutragen hat (z. B. per Mail „bei uns keine Kinder mit Förderbedarf“), klicken Sie in den „Rückmeldungen“ auf „offen ▾“ und wählen **„Fehlanzeige eintragen“**. Ein Formular ist nicht nötig.
- Die Stelle zählt dann als geantwortet („3 von 6 haben geantwortet (davon 1 Fehlanzeige)“) und steht mit „Fehlanzeige · Datum“ in der Liste. Ein weiterer Klick nimmt die Fehlanzeige zurück.
- Jede Fehlanzeige wird im Änderungsprotokoll vermerkt.

### Schnellere Seitenwechsel
- Unveränderte Seiten werden nur eingeblendet statt neu aufgebaut. Das gilt jetzt auch für Start, Dashboards, Einrichten und Hilfe. Bisher passte der Seitenschlüssel wegen einer je Blatt leicht abweichenden Fensterbreite nie, und diese Seiten wurden bei jedem Besuch neu gezeichnet (200 bis 800 ms).
- Jeder Seitenaufbau läuft wie bisher mit ausgeschalteter Bildschirmaktualisierung, ausgeschalteten Ereignissen und manueller Berechnung. Alles wird auch im Fehlerfall wiederhergestellt.

### Excel-Zoom (Schieberegler unten rechts, Strg + Mausrad)
- **Zoomen funktioniert jetzt überall:** Die Seite wird beim Zoomen neu aufgebaut. Die Navigation behält ihre Größe und liegt immer genau neben dem Inhalt. Sie wird nicht mehr abgeschnitten und verdeckt nichts.
- Der gewählte Zoom gilt für alle Seiten und bleibt beim Seitenwechsel erhalten. Das hilft z. B. beim Vergrößern für eine bessere Lesbarkeit.
- Knopfreihen brechen bei starker Vergrößerung sauber um. Der Text von „Hilfe & Sicherung“ und „Technische Dokumentation“ passt sich der Fensterbreite an.

### Waagerechtes Scrollen
- **Kein Zurückspringen und kein Stottern mehr:** In breiten Tabellen (z. B. 10 Spalten) bleibt die Ansicht beim Scrollen nach rechts ruhig am rechten Rand des Inhalts. Das Antippen der Bildlaufleiste springt nicht mehr ganz nach links. Ursache war ein fester Bildlaufbereich mit ausgeblendeten Spalten aus 1.6.0; er ist entfallen.

### Navigation bei 125 bis 175 % Bildschirmskalierung
- Die Kurven des aktiven Eintrags werden in jeder Skalierung scharf und vollständig gezeichnet. Bisher waren sie bei 175 % abgeschnitten.
- Die Navigation endet genau an Excels Bildlaufleiste, ohne hellen Spalt.
- Alle Einträge bis „Technische Dokumentation“ sind erreichbar. Bei wenig Höhe erscheinen Bildlaufpfeile in der Navigation. Der aktive Eintrag bleibt immer sichtbar, auch nach einer Größenänderung. Neben den Pfeilstreifen steht nie ein halb abgeschnittener Eintrag.
- Lange Einträge (z. B. „Technische Dokumentation“ in Fettschrift) werden bei hoher Skalierung minimal kleiner gesetzt statt rechts abgeschnitten.
- Logo und „Demo · MidPoint“ stehen mittig nebeneinander.
- Bei geöffnetem Menüband wirkte die weiße Markierung nach einem Klick auf „Einrichten“, „Hilfe & Sicherung“ oder „Technische Dokumentation“ verrutscht, und der Schalter „Automatisch speichern“ fehlte. Excel hatte Teile eines älteren Bildes stehen lassen. Die Navigation wird jetzt nach jedem Wechsel vollständig neu gezeichnet.
- Der Tooltip zeigt „Hilfe & Sicherung“ statt „Hilfe_Sicherung“.

### Seitenkopf
- Das MidPoint-Logo auf „Hilfe & Sicherung“ und „Technische Dokumentation“ verdeckt weder „Speichern“ noch „Ungespeichert“. Bei wenig Platz wird es weggelassen.
- Der Statustext rechts überlagert nie den Seitentitel. Bei wenig Platz wird er kleiner oder entfällt.
- Der Untertitel bleibt immer einzeilig. Ist er zu lang, wird er an einer Wortgrenze mit „…“ gekürzt.
- Auf „Hilfe & Sicherung“ und „Technische Dokumentation“ ist die leere Fläche zwischen den Schaltflächen und dem Text entfernt.
- Rechts neben der Navigation blitzt bei Zoom kein Rest des Auswahlrahmens mehr hervor.
- Der rechte Seitenrand berücksichtigt Excels senkrechte Bildlaufleiste. Karten, Status und „Speichern“ reichen nicht mehr darunter.

### Dashboards
- Kennzahlen passen in jede Kartenbreite (z. B. „70.100,00 €“ bei starkem Zoom). Die Zahl wird bei Bedarf kleiner gesetzt, nie abgeschnitten. Lange Kartentitel enden einzeilig mit „…“.

## Version 1.6.0 (Oktober 2026, Formatversion 3)

### Neuer Name: MidPoint
- Das Programm heißt jetzt **MidPoint**. Das Logo steht oben links in der Navigation (weiß auf dunkler, farbig auf heller Navigation), auf den Seiten „Hilfe & Sicherung“ und „Technische Dokumentation“ sowie auf dem Hinweisblatt.
- Erzeugte Dateien (Erfassungsdateien, Excel-Exporte) bleiben ganz normale Excel-Dateien: ohne Logo, ohne Programmnamen. Die Infozeile des Exports nennt keine Quelldatei mehr.
- Dateinamen: `MidPoint_Master.xlsm`, `MidPoint_Demo.xlsm`. Bestehende Dateien lassen sich wie gewohnt weiterverwenden.

### Erfassungsdateien: einfacher und übersichtlicher
- **Seite neu aufgebaut:**
  - Oben ein Ablaufbild in vier Schritten: Liste festlegen → Datei verschicken → Rücklauf einlesen → Gesamtliste.
  - Darunter „Ihre Erfassungen“ mit der Spalte „Rückmeldungen“ (z. B. „12 von 60“).
  - Daneben bzw. darunter die Übersicht **„Rückmeldungen“**: wer hat geantwortet (✓ mit Datum) und wer fehlt noch („offen“), mit Fortschrittsbalken.
- **„So funktioniert's“:** ein Fenster mit Ablaufbild und den beiden Arten:
  - Neue Einträge sammeln: am häufigsten.
  - Bestand prüfen: erweitert.
  - Es öffnet sich beim ersten Besuch der Seite von selbst.
- **Assistent in vier klaren Schritten:**
  1. Name, Zieltabelle, „Was möchten Sie tun?“. Neue Einträge sammeln ist voreingestellt; weitere Tabellen gibt es nur als erweiterte Option.
  2. Spalten zum Ausfüllen.
  3. **Wer füllt aus?** (siehe nächster Punkt)
  4. Überschrift und Hinweis. Farbstil, Schrift, Zeilenzahl und Blattschutz liegen unter „Weitere Einstellungen“.
- **Einheitliche Absender:** In der Datei wählen die Empfänger ihren Namen aus einer Liste. Die Liste ist entweder selbst eingetragen (ein Name je Zeile) oder kommt aus einer Tabelle der Datei, z. B. alle Einträge von „Einrichtungen“. Freie Eingabe bleibt möglich. Die Liste lässt sich jederzeit unter „Teilnehmende…“ ändern.
- **Eine Datei für alle:** Beim Erzeugen fragt MidPoint nicht mehr nach mehreren Dateien. Es entsteht eine Datei, die Sie an alle schicken. Danach wird die Hauptdatei automatisch gespeichert, damit die Rückläufe zugeordnet werden können. Eine Meldung erklärt die nächsten Schritte.
- Seltenes (Rücklauf-Übersicht je Empfängerdatei, weiteres Blatt, Entfernen) liegt unter „Weiteres…“.

### Fehler behoben beim Einlesen
- **Dieselbe Datei von mehreren Stellen:** Bisher wurde nur die erste Rückmeldung angenommen, alle weiteren galten als „bereits eingelesen“. Jetzt zählen Datei **und** Absender: jede Stelle wird einmal übernommen, eine doppelt geschickte Antwort derselben Stelle weiterhin übersprungen.
- **„gehört zu einer anderen Datei“:** Wurde die Hauptdatei nach dem Erzeugen ohne Speichern geschlossen, war die Erfassung danach unbekannt. Jetzt erklärt MidPoint die Ursache und stellt die Erfassung auf Wunsch aus der Rücklaufdatei wieder her; danach wird normal eingelesen. Zusätzlich speichert MidPoint die Hauptdatei direkt nach dem Erzeugen.
- Eine Rücklaufdatei, die gerade in Excel geöffnet ist, wird direkt gelesen statt mit einer Fehlermeldung abgelehnt.
- **Vorschau beim Einlesen:** echte Spaltenköpfe (Datei, Absender, neu, geändert, unverändert, Fehler, Hinweis) genau über den Spalten; das Fenster ist breit genug für den Hinweis.
- **Sicherungsfrage:** Die Schaltflächen ragen nicht mehr über den linken Fensterrand; allgemein verbreitern sich Meldungsfenster bei langen Schaltflächen. Klarer Text („Eine Sicherung vor dem Einlesen der Rückläufe wird empfohlen“).
- Nach dem Einlesen: Wahl zwischen „Tabelle anzeigen“ und „Bei den Rückmeldungen bleiben“.

### Bildschirme und Skalierung (z. B. 175 %)
- **Breite der Navigation:** Sie richtet sich jetzt nach der Skalierung des Bildschirms, auf dem Excel gerade läuft. Bei einem Wechsel des Bildschirms oder Zooms passt sie sich an.
- **Kein Scrollen ins Leere mehr:** Waagerechtes Scrollen endet am Rand des Inhalts.
- **Nichts verschwindet unter der Navigation:** Die Spalte unter der Navigation bleibt beim Scrollen stehen, Inhalte rutschen nicht mehr darunter.
- **Erfassungsseite bei schmalen Fenstern:** Das Ablaufbild wird schmaler, und die Rückmeldungen stehen unter der Liste statt daneben.

## Version 1.5.0 (Oktober 2026, Formatversion 3)

### Dokumentstil für die ganze Datei
- Unter **Einrichten → „Dokumentstil“** wählen Sie eine von 10 Vorlagen. Jede Vorlage legt Navigationsfarbe, Akzentfarbe, Farbstil und Schrift gemeinsam fest. Die Auswahl zeigt jede Vorlage mit ihren Farben und den Namen in ihrer Schrift:
  - Klassisch Blau, Petrol & Sand, Wald, Schiefer
  - Bordeaux elegant (Garamond), Erde warm (Georgia)
  - Nordisch hell, Flieder hell, Salbei hell
  - Kontrast klar (Verdana)
- **„Eigene Auswahl“:** Navigation, Akzent, Farbstil und Schrift einzeln wählen.
- Die Schrift gilt für Oberfläche, Dialoge, Dashboards und erzeugte Dateien. Der Farbstil ist die Voreinstellung für Dashboards, neue Erfassungsdateien und den Excel-Export; dort lässt er sich weiterhin einzeln ändern.
- Neue gedeckte Akzentfarben: Weinrot und Terrakotta.

### „Welches Diagramm passt?“ und Sankey
- Im Dialog „Element“ steht neben „Darstellung“ die Schaltfläche **„Welches Diagramm passt?“**. Sie öffnet eine Übersicht aller Darstellungen als Bild mit Beispieldaten und einem Satz, wofür sie sich eignet. Klick wählt, Doppelklick übernimmt.
- Die Bilder erzeugt Excel selbst (offline) im eigenen Farbstil. Sie werden beim ersten Öffnen einmal angelegt.
- **Neu: Sankey-Diagramm.** Es zeigt, wie sich Einträge von einer Gruppierung auf eine zweite verteilen, z. B. Status → Priorität.
  - Die Breite der Bänder entspricht der Anzahl bzw. Summe.
  - Einträge ohne Angabe erscheinen als „Ohne Angabe“.
  - Rechts stehen höchstens 8 Gruppen, der Rest als „Weitere“.
  - Ein Klick auf eine Gruppe links zeigt die Einträge.
  - Das Diagramm wird aus Excel-Formen gezeichnet und braucht keine Erweiterung.

### Listen
- **Schalter „Zeilenumbruch“** in der Werkzeugleiste jeder Tabellenansicht. Aus (Standard): eine Zeile je Eintrag. An: lange Texte werden vollständig gezeigt, die Zeilen wachsen mit. Die Einstellung wird je Ansicht gespeichert.
- **Alle Werte linksbündig**, auch Zahlen, Beträge und Datumswerte (Listen, Summenzeile, Erfassungsdateien).

### Erfassungsdateien übersichtlicher
- **Bestehende und neue Einträge klar getrennt, auf einem Blatt:**
  - Oben ein graues Band „Bestehende Einträge (n): bitte prüfen“ und links ein grauer Streifen.
  - Nach einem Abstand folgt ein farbiges Band „Neue Einträge: hier ergänzen, was oben noch fehlt“ mit farbigem Streifen.
  - Bei längeren Listen werden die Spaltenköpfe im Bereich für neue Einträge wiederholt.
- **Passt sich an:** Titelband, Einleitung und Anleitung sind mindestens so breit wie ein gut lesbarer Absatz und brechen sauber um, auch bei wenigen Spalten. Das Namensfeld endet am rechten Rand.
- Lange Spaltennamen, Gruppentitel und Optionsnamen werden umgebrochen statt abgeschnitten; die Kopfhöhe richtet sich danach. Lange Texte in bestehenden Einträgen sind vollständig lesbar.
- Fehler behoben: Pflichtfelder in leeren neuen Zeilen wurden rot, weil der vorbelegte Empfänger als „begonnen“ zählte.

### Mehrfachauswahl mit Höchstzahl
- In „Eigenschaft bearbeiten“ → „Weitere Einstellungen“ gibt es bei der Mehrfachauswahl das Feld **„Höchstens auswählbar“**. Beispiel: 6 Möglichkeiten, höchstens 3. Leer bedeutet beliebig viele.
- Die Grenze gilt überall:
  - **Formular:** Unter der Liste steht „Höchstens 3 auswählbar“. Bei zu vielen Haken wird der Hinweis rot und Speichern wird verhindert. Die Haken bleiben dabei sichtbar.
  - **Einfügen, Sammeländerung und Import** lehnen zu viele Werte ab.
- **Erfassungsdateien mit Ankreuzspalten:** Die Kopfzeile nennt die Grenze („höchstens 3 ankreuzen“). Sind in einer Zeile schon 3 Haken gesetzt, nimmt Excel keinen weiteren an; erst muss ein anderer Haken entfernt werden. Hineinkopierte zu viele Haken werden rot markiert und beim Einlesen nicht übernommen. Im Formular-Layout gilt dasselbe für die senkrechte Ankreuzliste.

### Farbstile mit Farbmustern
- Statt einer Namensliste zeigt die Auswahl jeden Stil mit kurzer Beschreibung und seinen **acht Farben**. Klick wählt, Doppelklick übernimmt.
- **10 Stile**, alle gedeckt und ohne grelle Farben:
  - kräftig: Akzentfarbe, Ozean, Wald, Schiefer, Bordeaux, Erde
  - pastell: Nordisch, Flieder, Salbei & Sand
  - Kontrast (barrierearm), gut unterscheidbar auch bei Farbsehschwäche
- Die Auswahl gibt es für Dashboards (Gestaltung), einzelne Elemente („wie Dashboard“ oder eigener Stil), Erfassungsdateien und den Excel-Export.
- Helle Stile erhalten automatisch dunkle Schrift. Große Zahlen auf Kennzahlkacheln werden bei hellen Stilen abgedunkelt, damit sie lesbar bleiben.
- Bisher gewählte Farben und Paletten (z. B. Petrol, Sonnenuntergang) bleiben gültig.

### Erfassungsdateien
- **Mehrere Tabellen auf einmal:** Im ersten Schritt lassen sich mehrere Tabellen ankreuzen. Jede wird ein eigenes Blatt in derselben Datei. Für jedes weitere Blatt fragt der Assistent direkt die Spalten ab; die Titel zeigen „(Blatt ‚Ressourcen‘)“. „Blatt hinzufügen“ bleibt für spätere Ergänzungen.
- **Klare Trennung beim Prüfen:**
  - Über den vorhandenen Einträgen steht „Vorhandene Einträge prüfen“.
  - Darunter folgt eine farbige Zeile „Neue Einträge ergänzen“ mit eigenem Bereich für neue Zeilen. Diese Zeile wird beim Einlesen übersprungen.
- **Empfänger vorbelegt:** In den neuen Zeilen steht der Aufteilungswert (z. B. „Workshop“) bereits in der festen Spalte. Die Spalte ist hell statt grau.
- **Ankreuzspalten auch beim Prüfen:** Eine Mehrfachauswahl (z. B. Themen) erscheint auch als Spalte zur Anzeige als Ankreuzspalten. In vorhandenen Zeilen sind die Haken grau und gesperrt, in neuen Zeilen wird angekreuzt; die Haken werden beim Einlesen übernommen.
- **Gestaltung:**
  - Farbstil aus der Auswahl mit Farbmustern (siehe oben). Helle Stile erhalten automatisch dunkle Schrift.
  - Schriftarten: Segoe UI, Calibri, Aptos, Arial, Verdana, **Garamond**, Georgia, Cambria.
- **Einfachere Dialoge:**
  - „Darstellung“ erscheint nur beim Sammeln.
  - „Anzahl leerer Zeilen“ erscheint nur, wenn neue Einträge erlaubt sind.
  - Beschriftung und Breite des Namensfeldes erscheinen nur, wenn es abgefragt wird.
  - Die Schrittanzeige (z. B. „2/4“) zählt nur die tatsächlich nötigen Schritte.

### Excel-Export
- **Mehrere Ansichten oder Tabellen in einer Datei:** „Diese Ansicht“, „Mehrere Ansichten dieser Tabelle“ oder „Mehrere Tabellen“. Jede wird ein eigenes Blatt.
- **Darstellung wählbar:**
  - **Gestaltet:** Titel, Infozeile mit Filter, Kopfzeile in der gewählten Farbe, feine Linien, wahlweise leicht getönte Zeilen, Filterpfeile, fixierte Kopfzeile, Querformat beim Drucken.
  - **Schlichte Liste:** nur Spaltenköpfe und Werte, gut zum Weiterrechnen.
  - Farbe und Schrift wie bei den Erfassungsdateien. Die letzte Wahl wird gemerkt.
- Texte, die mit =, +, - oder @ beginnen, werden auch im Excel-Export nie als Formel ausgeführt.

### Auswertungen (Dashboards)
- **Neue Darstellungen:**
  - **Netz** (Netzdiagramm, ab 3 Gruppen)
  - **Kreis**
  - **Fläche**
  - **Trichter**
  - **Baumdiagramm**
  - **Pareto** (Säulen absteigend plus kumulierter Anteil)
  - **Rangliste** (beschriftete Fortschrittsbalken)
  
  Trichter und Baumdiagramm sind Excel-Diagrammarten ab Excel 2016/2019; in älteren Versionen erscheint stattdessen ein Balkendiagramm.
- **Neue Kennzahlen:** Median, Spannweite (Maximum − Minimum; bei Datumsfeldern in Tagen), Anzahl verschiedener Werte.
- Im Dialog „Element“:
  - Eine kurze Erklärung zeigt jede Darstellung.
  - Das Feld für die Kennzahl erscheint nur, wenn die gewählte Kennzahl eines braucht.
- Die Demodatei enthält ein Dashboard „Auswertungen“ mit allen neuen Darstellungen und Kennzahlen.

### Tabellen und Formulare
- **Letzte Spalte:** Langer Text endet an ihrem Rand, statt über die Tabelle hinauszulaufen.
- **Zahlen im Formular linksbündig** wie alle anderen Eingaben (inzwischen auch in Listen, siehe oben).
- **Allgemeine Dialoge:** Angaben, die nur bei einer bestimmten Wahl gebraucht werden, bleiben verborgen, bis sie nötig sind; der Dialog rückt automatisch zusammen.
- **Fehler behoben:** Mehrfachlisten in Dialogen (z. B. Spalten für Board-Karten) verloren beim Schließen ihre Auswahl. Die Auswahl wird jetzt gemerkt und auch nach einem Hinweis wieder angezeigt.

## Version 1.4.0 (Oktober 2026, Formatversion 3)

### Fehler behoben
- **„Erste Tabelle anlegen“** in der Navigation öffnet direkt den Dialog „Neue Tabelle“. Vorher führte der Eintrag zur Seite „Einrichten“, beide Einträge wurden gleichzeitig markiert und die Beschriftung „Einrichten“ verschwand.
- **Neue Tabelle:** Nach „Tabelle anlegen“ wächst der Dialog. Er rückt jetzt nach oben, sodass „Fertig“ und alle Schaltflächen sichtbar bleiben. Vorher lagen sie bei kleineren Bildschirmen unter der Taskleiste.
- **Neue Tabelle nach dem Löschen der einzigen Tabelle:**
  - Bisher kam „Fehler 91“, auch nach „Datei zurücksetzen“.
  - Ursache: Nach dem Löschen der letzten Zeile blieb eine Konfigurationsliste ganz ohne Datenzeile zurück, und das Anhängen der nächsten Zeile schlug fehl.
  - Jetzt bleibt eine leere Platzhalterzeile stehen. Das zentrale Anhängen von Zeilen (Konfiguration, Einträge, Zuordnungen, Protokoll) legt bei Bedarf selbst eine Datenzeile an; betroffene Dateien heilen sich so beim nächsten Anlegen.
- **Helle Navigationsfarben:** Über und unter dem ausgewählten Eintrag erscheinen keine feinen weißen Linien mehr. Die Ursache war die Verkleinerung des fein gezeichneten Streifens, die an den Rändern Pixel von außerhalb des Bildes einmischte.
- **Typ der Hauptspalte ändern:** Statt „Fehler 9“ erscheint die Typauswahl. Automatisch gefüllte Arten wie „Erstellt um“, Rollup und Button werden dort nicht mehr angeboten.

### Tabelle löschen
- Eine Tabelle lässt sich jetzt gut auffindbar löschen:
  - in der Tabelle über **„Ansicht > Tabelle … löschen…“**
  - im Dialog „Eigenschaften“ über die Schaltfläche **„Tabelle löschen…“**
  - unter **Einrichten**: Zeile anklicken, dann „Tabelle löschen…“ (bisher „Entfernen…“)
- Ablauf:
  - Abhängigkeiten werden geprüft (blockierende werden genannt), die Auswirkungen angezeigt.
  - Zur Bestätigung wird „LÖSCHEN“ eingegeben; vorher wird automatisch gesichert.
  - Danach erscheint ein Hinweis auf die Wiederherstellung.
- Ohne Datenverlust bleibt weiterhin „Aus-/Einblenden“ unter Einrichten.
- Im Eigenschaftsdialog heißt die Schaltfläche für Spalten jetzt eindeutig „Spalte entfernen…“.

### Zahlen einheitlich
- Berechnungen ohne feste Kommastellen zeigen ganze Zahlen ohne Komma („50“ statt „50,“). Andere Werte erhalten einheitlich zwei Nachkommastellen, in der Tabelle, im Formular und beim Export.
- Zahlen stehen überall rechtsbündig: Tabelle, Formular (auch berechnete Werte) und die grauen Spalten in Erfassungsdateien.

### Kalender
- **Heute** ist mit einem farbigen Rahmen und „Heute“ an der Tageszahl markiert, nicht mehr mit einer Flächenfarbe. Er lässt sich so nicht mehr mit Einträgen verwechseln.
- **Mehrtägige Einträge** sind durchgehende Leisten mit fester Spur je Woche. Überlappende Einträge liegen in getrennten Spuren und vermischen sich nicht. Der Name steht am Leistenanfang; „→“ zeigt eine Fortsetzung aus der Vorwoche. Bei vollen Tagen erscheint „+ n weitere“.

### Projektleiste (Gantt)
- **Tagesgenaue Leisten** auch in der Wochen- und Monatsansicht: Ein Eintrag über drei Tage füllt nur diesen Teil der Woche. Die Leisten sind abgerundet; der Fortschritt erscheint als dunklerer Anteil.
- **Heute-Linie** genau am heutigen Tag, auch in Wochen und Monaten.
- **Neu: „Namen auf Leisten“** (Schalter in der Werkzeugleiste und unter „Ansicht bearbeiten“):
  - Die Namen stehen klein und dezent auf den Leisten.
  - Passt ein Name nicht hinein, steht er rechts daneben, bei Platzmangel gekürzt mit „…“.
- Ein Klick auf eine Leiste öffnet den Eintrag.

### Einrichtung verständlicher
- **Verknüpfung anlegen:**
  - kurze Erklärung mit Beispiel
  - drei nummerierte Schritte: Tabelle, Anzahl, Name auf der Gegenseite
  - Klartext, was entsteht, z. B. „Jeder Eintrag in ‚Aufgaben‘ wählt genau einen Eintrag aus ‚Projekte‘ …“
- **Rollup direkt im Dialog** statt über den Formelassistenten:
  1. Über welche Verknüpfung?
  2. Was berechnen? (Anzahl, Summe, Durchschnitt, kleinster/größter Wert, Anzahl ausgefüllt)
  3. Von welchem Feld? Es werden nur passende Felder angeboten.
  
  Dazu kommen Klartext zum Ergebnis und ein Vorschlag für den Ergebnistyp. Bedingungen gibt es weiterhin über „Mit Bedingung (erweitert)…“.
- **Register „Verknüpft aus anderen Tabellen“** (früher „Verknüpfte Einträge“):
  - Eine Erklärung steht direkt im Register; jedes Feld nennt seine Herkunft, z. B. „Einträge aus ‚Aufgaben‘, die im Feld ‚Projekt‘ diesen Eintrag ausgewählt haben“.
  - Das Register zeigt die Gegenrichtung einer Verknüpfung.

### Formulare und Dialoge
- **Verknüpfungsfeld im Formular** ist genauso hoch wie „Auswählen…“. Listen werden allgemein nicht mehr auf ganze Zeilen gekürzt.
- **Formelleiste:** Die Texte in den Spalten- und Funktionskästchen stehen senkrecht mittig.
- **Spaltenauswahl:** Ein Klick auf eine Spalte wählt sie nur aus; „Nach oben“/„Nach unten“ blenden nichts mehr ein oder aus. Sichtbarkeit über „Sichtbar“, Doppelklick oder Leertaste; ein Haken zeigt sichtbare Spalten.
- **„Zurück“ in mehrstufigen Dialogen:** Die Angaben bleiben beim Zurückgehen erhalten. Das gilt für:
  - Erfassungsdatei festlegen (alle Schritte)
  - Dashboard-Element (Schritt 2 zu 1)
  - Neue Ansicht (Einstellungen zu Name/Darstellung)

### Dashboard
- **Einrasten sofort nach dem Loslassen** einer Karte, ohne zusätzlichen Klick ins Blatt. Ohne Zeitgeber: Die Anwendung reagiert auf die Aktualisierung der Excel-Oberfläche (`CommandBars.OnUpdate`), nur solange „Anordnen“ aktiv ist.

## Version 1.3.0 (Oktober 2026, Formatversion 3)

### Schneller und ruhiger
- **Seitenwechsel ohne Neuaufbau, wenn sich nichts geändert hat.** Jede Seite merkt sich ihren Stand (Daten, Struktur, Ansichtszustand, Fensterbreite, Datum). Unveränderte Seiten werden nur eingeblendet: rund 60 bis 80 ms statt 200 bis 450 ms (Demodatei). Bei 10.000 Einträgen öffnet die zuletzt gezeigte Tabelle in 0,04 s statt 1,9 s und das Dashboard in 0,05 s statt 9,5 s.
- **Navigation:** Ein Seitenwechsel setzt nur noch die Markierung um, statt alle Einträge neu anzulegen (rund 20 bis 50 ms statt 190 ms). Neu gezeichnet wird nur bei geänderten Einträgen, Titel oder Farben, und dann ohne Zwischenbild.
- **Kein Stottern beim Farbwechsel:** Die Navigation wird nicht mehr geschlossen und neu geöffnet. Bei der Navigationsfarbe werden auf der Seite nur die Farbfelder neu gezeichnet.
- **Navigation:** Symbol und Beschriftung stehen genau auf der Mitte der Markierung (gemessene Zeilenhöhe statt fester Werte).
- Erster Aufbau von Seiten schneller: Schaltflächenbreiten werden je Beschriftung nur einmal gemessen, Optionsfarben werden für gleichfarbige Zellen gemeinsam gesetzt, Spaltenköpfe in einem Zug geschrieben, Bezeichnungen verknüpfter Einträge zwischengespeichert, Startseiten-Karten mit weniger Formen.

### Helle Farben (Pastell)
- Für Navigation und Akzent gibt es neben den kräftigen Farben ruhige helle Töne: Navigation Flieder, Salbei, Himmel, Sand, Rosé, Nebel; Akzent Flieder hell, Mint, Himmel hell, Pfirsich, Rosé hell.
- Bei hellen Farben wird die Schrift automatisch dunkel (Navigation, Hauptschaltflächen, Kalender). Ein heller Akzent wird für Schrift und Symbole auf Weiß abgedunkelt, damit alles lesbar bleibt.

### Formelspalten mit echten Excel-Formeln
- Eine Formelspalte ist jetzt eine **Excel-Formel** in gewohnter deutscher Schreibweise; Spalten stehen in eckigen Klammern: `=[Preis]*[Menge]`, `=WENN([Priorität]="Hoch";"dringend";"")`, `=RUNDEN([Kosten]/[Stunden];2)`.
- **Spalten per Klick einfügen:** In der Formelleiste (Schaltfläche „fx“ neben „+“ im Tabellenkopf oder Spaltenmenü „Neue Formelspalte rechts davon…“) genügt ein Klick auf eine Spalte der Tabelle. Im Eigenschaftsdialog stehen alle Spalten als Knöpfe bereit, dazu häufige Funktionen.
- **Sofortige Prüfung und Vorschau** an echten Einträgen beim Tippen; der Ergebnistyp (Zahl, Text, Ja/Nein) wird vorgeschlagen.
- Berechnet wird spaltenweise in einem Rechenvorgang (10.000 Einträge in einem Zug). Spaltennamen dürfen sich später ändern (gespeichert werden interne IDs).
- **Sicherheit:** Bezüge auf andere Blätter, Dateien oder Zellen sowie Funktionen mit Netzwerk- oder Systemzugriff (z. B. WEBDIENST, INDIREKT) sind gesperrt. Eingegebene Texte werden nie als Formel ausgeführt.
- Bestehende Berechnungen (Assistent) und Rollups funktionieren unverändert; der Assistent bleibt über „Assistent statt Formel…“ erreichbar.

### Dashboard: Anordnen im Raster
- Das Dashboard hat ein festes Raster aus **vier Spalten**. Karte ziehen, loslassen, ins Blatt klicken (oder „Einrasten“): Sie springt auf den nächsten Rasterplatz. Breite ¼, ½, ¾ oder ganze Breite; Diagramme normal oder hoch.
- Eine Karte, die auf einen belegten Platz gezogen wird, verdrängt die dortige nach rechts bzw. unten; leere Zeilen entfallen.
- **„Fertig“ speichert genau die gezeigte Anordnung** (Spalte, Zeile, Breite, Höhe); sie bleibt dauerhaft so. „Abbrechen“ verwirft alles. Schaltflächen und Rasterhilfe sind in diesem Modus nicht verschiebbar.

### Behoben
- Wiederherstellen einer Sicherung konnte Text, der mit „=“ beginnt, als Formel auswerten; jetzt bleibt er Text (gilt für alle Schreibwege in Blätter).

### Bekannte Einschränkungen (1.3.0)
- Excel-Formeln behandeln leere Zellen wie Excel (leer × 2 = 0). Wer „leer bleibt leer“ möchte, schreibt z. B. `=WENN(ISTLEER([Menge]);"";[Menge]*2)`.
- In der Formelleiste wirkt ein Klick auf dieselbe Tabellenzelle nur einmal hintereinander; danach eine andere Zelle der Spalte anklicken (oder den Spaltenknopf).
- Ziehen beim Anordnen: Die Karte rastet erst nach dem Loslassen und einem Klick ins Blatt bzw. „Einrasten“ ein (Excel meldet das Loslassen einer Form nicht).
- Tabellen teilen sich ein Blatt: Beim Wechsel zwischen zwei Tabellen wird neu aufgebaut; die Wiederverwendung greift beim Zurückkehren zur zuletzt gezeigten Tabelle und bei allen anderen Seiten.

## Version 1.2.0 (Oktober 2026, Formatversion 3)

### Einfacher wie in Notion
- **Kein Umbaumodus mehr:** Daten pflegen und Tabellen umbauen in derselben Oberfläche. Neue Spalte über „+“ im Tabellenkopf, Doppelklick auf einen Spaltenkopf öffnet das Spaltenmenü. Ein gesetzter Einrichtungsschutz wird einmal je Sitzung vor der ersten Strukturänderung abgefragt.
- **Automatisch speichern** als Schiebeschalter unten links: ein = jede Änderung wird kurz danach gespeichert (die Seite bleibt dabei stehen), aus = wie bisher.
- **Schlankere Formulare:** Die Eintragsmaske zeigt keine automatisch berechneten Eigenschaften (Formel, Rollup, Erstellt/Geändert, ID; die ID steht im Kopf); Einheiten statt Erklärtexte, „+“ statt „Neue Option…“, Duplizieren/Archivieren unter „Mehr“, Registerleiste nur bei mehreren Registern. Der Eigenschaften-Dialog zeigt Name, Art und Pflicht; Hilfetext, Register, Standardwert, Eindeutigkeit und Dublettenprüfung unter „Weitere Einstellungen“. Dashboard-Elemente: Datumsstufe, Zeitraumfeld, Reihenfolge und Gruppenanzahl unter „Weitere Einstellungen“.
- **Uhrzeiten ohne Sekunden** (Anzeige, Eingabe, Kalenderdialog).
- **Titel der Datei** jederzeit änderbar: Klick oben links in der Navigation, auf der Startseite und unter „Einrichten“.
- **Farben:** Farbfelder für Navigation (zehn Farben) und Akzent direkt auf „Einrichten“; farbige Ansichtsregister je Art (Tabelle blau, Board orange, Kalender grün, Projektleiste lila).
- **Navigation:** Übergänge der aktiven Markierung als glatter, großer Viertelkreis (6-fache Kantenglättung); die feine graue Linie an der Öffnung ist entfernt.
- **Bildschirm- und Fenstergrößen:** Die Navigation endet genau über Excels unterer Leiste (vorher lagen die untersten Punkte darunter und wirkten abgeschnitten), passt sich beim Ein-/Ausklappen des Menübands an, verdichtet sich bei wenig Höhe stufenweise und bietet danach Pfeile zum Blättern; der aktive Eintrag bleibt sichtbar. Seiten brechen ab 600 pt Breite um statt waagerecht zu scrollen. Geprüft im Vollbild (1920 × 1080) und in einem kleinen Fenster (760 × 480 pt), mit ein- und ausgeklapptem Menüband.

### Dashboards
- **Gestaltung:** Stile (Karten, Flach, Farbakzent, Kontur) und Farbpaletten (Akzentfarbe, Ozean, Wald, Sonnenuntergang, Pastell, Kontrast, Graustufen) mit Live-Vorschau hinter dem Dialog; eigene Palette je Element möglich.
- **Anordnen per Ziehen:** Karten mit der Maus verschieben und an den Rändern breiter/schmaler bzw. Diagramme höher ziehen; „Fertig“ richtet am Raster aus.
- **Vorschau beim Anlegen:** Das neue Element erscheint sofort (blau umrandet) im Dashboard; „Übernehmen“, „Nochmals bearbeiten“ oder „Verwerfen“.
- Diagramme ohne doppelte Diagrammtitel.

### Erfassungsdateien
- **Mehrfachauswahl als Ankreuzspalten:** je Option eine schmale Spalte mit senkrechtem Titel, darüber der Name der Eigenschaft; Auswahl „Ja“ oder leer. Im Formular eine Ja-Zeile je Option. Beim Einlesen wird daraus wieder die Mehrfachauswahl (deaktivierte Optionen eines Eintrags bleiben erhalten).
- **Vorhandene Einträge prüfen lassen (vorbelegte Dateien):** Spalten nur zur Anzeige (grau, gesperrt) und zum Ausfüllen (gelb, mit vorhandenen Werten), Einträge wie eine gespeicherte Ansicht, **je Empfänger eine eigene Datei** („Aufteilen nach“, z. B. Standort oder Kita) mit nur dessen Einträgen. Optional leere Zeilen für neue Einträge, die beim Einlesen automatisch dem Empfänger zugeordnet werden.
- **Rücklauf aktualisiert über die ID:** nur geänderte Werte; Vorschau mit neu/geändert/unverändert/Fehler; Konfliktabfrage, wenn ein Eintrag seit dem Versand geändert wurde; Dateien ohne Änderungen gelten als eingegangen.
- **Mehrere Tabellen in einer Datei:** „Blatt hinzufügen…“ ergänzt weitere Blätter (eigene Tabelle, Spalten und Zuordnung zum Empfänger).
- **Bessere Gestaltung:** Beschriftung und Breite des Namensfeldes (z. B. „Einrichtung / prüfende Person“), Farbe der Kopfzeile, Kurzanleitung an/aus, Blattschutz, Spaltenbreiten passend zum Inhalt.
- **Rücklauf-Übersicht:** zeigt je Empfänger, ob die Datei eingegangen ist.
- Dateien der Version 1.1 werden weiterhin eingelesen.

### Migration und Build
- Formatversion 3: Beim ersten Öffnen werden die Spalten „Optionen“ (_Erfassungen) und „Empfaenger“ (_Ruecklauf) ergänzt; Daten bleiben unverändert. Ältere Sicherungen werden beim Wiederherstellen ebenso ergänzt.
- `build.ps1 -Update` (Programmcode einer bestehenden Fachdatei erneuern) korrigiert: speichert jetzt zuverlässig alle Module. Geprüft mit der Demodatei 1.1.0 (alle Einträge erhalten, Migration auf Format 3).

### Bekannte Einschränkungen (1.2.0)
- Ziehen mit der Maus gibt es im Dashboard („Anordnen“), nicht in Board, Kalender und Projektleiste.
- Mehrfachauswahl mit mehr als 30 aktiven Optionen erscheint in Erfassungsdateien weiterhin als Text „A; B“.
- Fügen Empfänger Werte per Kopieren/Einfügen ein, kann Excel die Auswahllisten der Zielzellen überschreiben; beim Einlesen wird trotzdem jeder Wert geprüft.
- Die Konfliktprüfung vergleicht den Änderungszeitpunkt des ganzen Eintrags (nicht einzelner Spalten).

## Version 1.1.0 (Oktober 2026, Formatversion 2)

Erste Auslieferung. Enthält den vollständigen Auftragsumfang sowie die nachträglich gewünschten Erweiterungen.

### Grundumfang (Projektauftrag)
- Tabellen und Eigenschaften frei zusammenstellen, Verknüpfungen 1:n und n:m mit automatischer Gegenseite.
- Dynamische Erfassungsmaske mit Registern, Pflichtangaben, Hilfetexten, Fehlern direkt am Feld; direkte Zellbearbeitung, Mehrzellen-Einfügen und Sammeländerung mit denselben Prüfregeln.
- Gespeicherte Ansichten: Tabelle, Kalender, Board, Projektleiste (Gantt); Suche, Filterkacheln, Sortieren, Gruppieren, Spaltenauswahl.
- Berechnungsbaukasten ohne Excel-Formeln (RECHNEN, RUNDEN, ANTEIL, WENN, PRÜFEN, Datumsfunktionen, ROLLUP) mit Zirkelbezugsprüfung.
- Dashboards (Kennzahl, Balken, Säulen, Zeitverlauf, Anteil, Tabelle) mit Filter und Detailaufruf; eindeutige Zählung bei n:m.
- Import (XLSX/CSV) mit Vorschau und Fehlerbericht, Export von Ansichten und Dashboards (XLSX, CSV, PDF).
- Vorlagen, Sicherung/Wiederherstellung, Änderungsprotokoll, Rückgängig, Archiv, Lesemodus, Einrichtungsschutz.
- Integrierte Hilfe und Technische Dokumentation, vollständiger Codekatalog mit Export (ZIP, HTML, Markdown, Text).

### Erweiterungen nach Rückmeldung
- **Keine Fensterfixierung mehr** (keine Fixierlinien) und keine gestrichelten Seitenumbruchlinien.
- **Umbaumodus sichtbar:** „Tabelle umbauen“ in jeder Tabelle; im Umbaumodus blau hinterlegte Spaltenköpfe, „Eigenschaft hinzufügen“, „+“ neben dem letzten Spaltenkopf, Spaltenmenü per Doppelklick (sortieren, gruppieren, filtern, ausblenden, bearbeiten, rechts einfügen, entfernen), „Umbau beenden“.
- **Eigenschaftsarten wie in Notion:** Text, Nummer, Auswählen, Mehrfachauswahl, Status, Datum, Person, Dateien und Medien, Kontrollkästchen, URL, E-Mail, Telefon, Formel, Verknüpfung, Rollup, Erstellt um, Erstellt von, Letzte Änderung um, Zuletzt bearbeitet von, Button, Ort, ID, dazu Geldbetrag, Prozent, Langer Text, Datum mit Uhrzeit, Ja/Nein (auch leer), Lokaler Pfad. Auswahl im Eigenschaftsdialog als Kachelraster mit Symbolen.
- **Erfassungsdateien:** eigene Seite „Erfassungsdateien“. Aus einer Tabelle wird eine makrofreie XLSX-Datei (Liste oder Formular) mit Auswahllisten, Datums-/Zahlenprüfung, Pflichtmarkierung, Eingabehinweisen und Blattschutz erzeugt, auf Wunsch je Empfänger eine eigene Datei. „Rücklauf einlesen“ prüft beliebig viele ausgefüllte Dateien, zeigt eine Vorschau mit Fehlerliste, übernimmt gültige Einträge in einer Transaktion und verhindert doppeltes Einlesen (Datei-ID).
- Vorlagen nutzen die neuen Arten (ID, Person, Kontrollkästchen, Button, E-Mail, Telefon, URL).

### Behobene Fehler während der Abnahme
- Wiederherstellung einer Sicherung schlug bei gemischten Spaltenformaten fehl und hinterließ einen unvollständigen Stand. Jetzt: korrekte Formatübernahme und automatische Rücknahme auf den Stand unmittelbar vor dem Versuch, falls eine Wiederherstellung scheitert.
- Zwei Sicherungen in derselben Sekunde überschrieben sich; jetzt eindeutige Dateinamen.
- E-Mail-Prüfung löste bei Eingaben ohne „@“ einen Laufzeitfehler aus.
- Kontrollkästchen ohne ausdrücklichen Wert wurden bei neuen Einträgen als „fehlend“ gemeldet.
- Symbole in Spaltenköpfen und das Entfernen-Zeichen in Filterkacheln wurden als Platzhalter angezeigt; Pfeilzeichen in Texten erschienen als „?“.
- Hilfe- und Technikseite waren nach dem Öffnen teilweise von der Navigation verdeckt.
- Statusleiste von Excel zeigte nach manchen Aktionen „FALSE“.

### Bekannte Einschränkungen
- Keine gleichzeitige Bearbeitung durch mehrere Personen (Excel-Dateisperre; zweite Person arbeitet im Lesemodus).
- Kein Drag-and-drop in Board, Kalender und Projektleiste (bewusst; alle Änderungen per Klick und Dialog).
- Mausrad scrollt in Formularen nicht innerhalb von Registerseiten (Grenze von Excel-Formularen); Bildlaufleiste und Tabulator funktionieren.
- Erfassungsdateien: Mehrfachauswahl wird als Text „A; B“ erfasst (Excel-Auswahllisten erlauben nur einen Wert); Verknüpfungen nur als Einfachverknüpfung über die Bezeichnung des Zieleintrags.
