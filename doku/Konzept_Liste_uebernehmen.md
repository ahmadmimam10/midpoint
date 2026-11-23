# Konzept: „Liste übernehmen“ (MidPoint 1.7.0)

Stand: November 2026

## Ziel

Vorhandene Listen, die nicht aus MidPoint stammen, sollen ohne Vorarbeit in eine saubere MidPoint-Tabelle übernommen werden. MidPoint liest die Datei, erkennt Aufbau und Spaltentypen, schlägt alles vor und **fragt nur dort nach, wo es unsicher ist**. Erst nach der Prüfung und einer ausdrücklichen Bestätigung wird etwas angelegt; vorher entsteht automatisch eine Sicherung, und die Übernahme lässt sich in der Sitzung mit „Rückgängig“ zurücknehmen.

Der bisherige Import (Einrichten > Importieren) bleibt für einfache Tabellen mit einer Kopfzeile, etwa eigene MidPoint-Exporte.

## Einstieg

Schaltfläche **„Liste übernehmen“** auf der Startseite und unter Einrichten, direkt neben „Daten importieren“. Es öffnet sich ein Assistent mit Schrittanzeige.

## Zwei Arten

| Art | Wofür | Was MidPoint braucht |
|---|---|---|
| **Befragung aus 2ask** | Ergebnisse einer 2ask-Umfrage | den kodierten Excel-Export mit Codeplan (2ask: Ergebnisse herunterladen, „Kodierte Ergebnisse“, Excel, „Codeplan hinzufügen“) |
| **Sonstige Liste** | jede andere Excel- oder CSV-Liste | nur die Datei |

Erkennt MidPoint beim Lesen einen 2ask-Export (Blätter „results“ und „code_plan“), schlägt es die passende Art selbst vor. Fehlt beim 2ask-Export der Codeplan, erklärt der Assistent, wie man ihn mit exportiert.

## Ablauf

1. **Quelle:** Art wählen, Datei wählen. Die Datei wird nur gelesen (schreibgeschützt, Makros aus).
2. **Tabelle** (sonstige Liste): MidPoint zeigt alle gefundenen Tabellen (Blatt, Bereich, Zeilen, Spalten) mit Vorschau. Man wählt eine aus. Gleich aufgebaute Blätter lassen sich zusammen übernehmen.
   **Befragung** (2ask): Übersicht über Fragen, Fragearten und Teilnehmende.
3. **Rückfragen:** Nur die unsicheren Punkte, jeweils mit einem Vorschlag, den man übernehmen oder ändern kann. Ist alles eindeutig, steht dort „Keine Rückfragen“.
4. **Felder:** Alle Felder mit Name, Typ, Beispielwerten und Haken „übernehmen“. Name und Typ lassen sich ändern; bei Auswahlfeldern lassen sich Werte zusammenführen. Ziel: neue Tabelle oder bestehende Tabelle (mit Schlüssel zum Aktualisieren, z. B. für neue 2ask-Rückläufe).
5. **Prüfung:** Wie beim bisherigen Import: gültige Zeilen, Fehler mit Grund, Hinweise auf Werte, die nicht übernommen werden konnten. Erst „Übernehmen“ legt Tabelle, Felder, Optionen und Einträge an.

## Erkennung bei sonstigen Listen

| Erkennung | Beispiel aus den Testlisten | Rückfrage |
|---|---|---|
| Tabellen auf allen Blättern finden, Titel- und Hinweiszeilen überspringen | Titel, „Stand“, Hinweisblatt | nein |
| Kopfzeile, auch über zwei Zeilen mit verbundenen Zellen | „Teilnahme“ über „Modul 1 bis 3“ | nur wenn unklar |
| Summen- und Zwischensummenzeilen | „Summe“, „Gesamt“, „Zwischensumme“, Formeln | ja |
| Gruppenüberschriften als eigenes Feld | „Produkt 36.1 …“, „Freie Träger“ | ja |
| Senkrecht verbundene Zellen nach unten auffüllen | „Gebäude“ | ja |
| Ausgeblendete Spalten | „intern“ | ja |
| Kreuze über mehrere Spalten als eine Mehrfachauswahl | Modul 1, 2, 3 mit „x“ | ja |
| Mehrere Werte in einer Zelle | „Sport, Gesundheit“, „Bildung; Sprache“ | ja |
| Typ je Spalte: Text, langer Text, Nummer, Geldbetrag, Prozent, Datum, Datum mit Uhrzeit, Ja/Nein, Auswählen, Mehrfachauswahl, E-Mail, Telefon | Status, Frist, Betrag | nur wenn nicht eindeutig |
| Nummern mit führenden Nullen, Aktenzeichen, PLZ bleiben Text | „0004711“, „01067“ | nein |
| Werte, die nicht zum Typ passen | „KW 14“ im Datum, „k.A.“ im Betrag | ja: leer lassen und melden oder als Text übernehmen |
| Ähnliche Werte zusammenführen | „erl.“ wie „erledigt“, „Veg.“ wie „vegetarisch“, „JHA“ wie „Jugendhilfeausschuss“ | ja |
| Formelspalten | Rest, Abfluss in % | ja: weglassen oder als Werte |
| Gleich aufgebaute Blätter zusammenführen, Blattname als Feld | ein Blatt je Kita | ja |
| Spalten, die nie zusammen vorkommen, zusammenlegen | „Menge“ und „Anzahl“ | ja |
| Monate oder Jahre als Spalten in Einträge umwandeln | Jan bis Dez | ja |
| Formular statt Liste (Bezeichnung: Wert) als einen Eintrag | Antrag | ja |
| CSV: Zeichensatz, Trennzeichen, ISO-Datum, Dezimalpunkt | Fachverfahren-Export | nein |

## 2ask

Der Codeplan liefert für jede Spalte Frageart, Datentyp, Codes und Texte. Daraus entstehen die Felder ohne Raten:

| 2ask | Feld in MidPoint |
|---|---|
| Einfachauswahl, Dropdown | Auswählen (Optionen in Code-Reihenfolge) |
| Mehrfachauswahl (`Q2_01` …) | eine Mehrfachauswahl; Zusatztext („Sonstiges, und zwar“) als eigenes Textfeld |
| Matrix (Likert, Dropdown) | je Zeile ein Auswahlfeld |
| Matrix Mehrfachauswahl | je Zeile eine Mehrfachauswahl |
| Matrix Textfeld | je Zelle ein Textfeld |
| Bipolar | je Zeile eine Nummer 1 bis 5 (Rückfrage: oder Auswahl) |
| Sterne, Schieberegler, Summe, Zahlfeld | Nummer |
| Rangfolge | je Element eine Nummer „Rang: …“ |
| Gruppenmatrix | je Zeile und Gruppe ein Auswahlfeld |
| Formular | je Feld Text, Dropdown als Auswählen |
| Datumsfeld | Datum (aus Text) |
| Zeitfeld | Text („07:00“) |
| Teilnahme-ID, Beendet am | übernommen (Teilnahme-ID ist Schlüssel für spätere Nachlieferungen) |
| übrige Metadaten (Browser, Gerät, Dauer …) | vorgeschlagen „nicht übernehmen“ |

Umkodierte fehlende Antworten (-99, -98, -97, auch als Text) werden leer übernommen. Feldnamen sind eindeutig und lassen sich im Schritt „Felder“ ändern.

## Grenzen dieser Version

- Wochenraster (Uhrzeit × Wochentag) wird nicht in Termine zerlegt.
- Verknüpfungen zwischen zwei übernommenen Tabellen (z. B. Aufgaben zu Projekten) entstehen nicht automatisch; dafür die zweite Liste in eine bestehende Tabelle mit Verknüpfungsfeld übernehmen.
- Zellfarben (z. B. Ampel ohne Legende) werden nicht ausgewertet.

## Prüfung

- Selbsttest in der Datei: Erkennung an eingebauten Beispielen (Kopfzeile über zwei Zeilen, Summen- und Gruppenzeilen, Kreuzspalten, Mehrfachwerte, Abkürzungen, 2ask-Codeplan).
- Testlauf mit allen Testlisten aus `testdaten/importlisten` (Skript `build/listimport_test.ps1`): Erkennung und Übernahme in eine Testkopie, Abgleich mit der Soll-Erkennung aus der LIESMICH.
- Sichtprüfung jedes Assistentenschritts in Excel.
