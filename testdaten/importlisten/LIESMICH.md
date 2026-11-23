# Testlisten für den Listen-Import

Testsatz für die geplante Funktion „Excel-Liste übernehmen“: MidPoint soll den Aufbau einer vorhandenen Liste erkennen (Kopfzeile, Gruppen, Summen, Spaltentypen, Auswahl und Mehrfachauswahl), bei Unsicherheit nachfragen und die Daten in MidPoint-Tabellen übernehmen. Die Listen sind bewusst **normale Verwaltungslisten** ohne Datenüberprüfung, jede mit einem anderen Aufbau und den typischen kleinen Unsauberkeiten.

Alle Namen, Adressen, Aktenzeichen und Beträge sind erfunden. E-Mail-Adressen enden auf `.example`.

**Erzeugen:** `powershell -ExecutionPolicy Bypass -File build\testlisten.ps1` (braucht Excel, dauert etwa 30 Sekunden). Die erzeugten XLSX- und CSV-Dateien liegen nicht im Repository, weil sie jederzeit gleich neu entstehen (feste Zufallszahlen für die 2ask-Antworten). Im Repository liegen nur die echten 2ask-Exporte im Ordner `2ask_echt/`, denn die lassen sich nicht neu erzeugen.

## Überblick

| Datei | Art der Liste | Was die Erkennung hier lernen muss |
|---|---|---|
| `01_Teilnehmerliste_Fortbildung.xlsx` | Anmeldeliste | Titelzeilen, zweizeiliger Kopf, Kreuze über drei Spalten = Mehrfachauswahl, Summenzeile |
| `02_Fristenliste_Wiedervorlage.xlsx` | Wiedervorlage | Datum teils als Text, Status in Schreibweisen, Kürzel, Ampel nur als Zellfarbe |
| `03_Haushalt_Mittelabfluss.xlsx` | Haushaltsübersicht | Gruppenüberschriften als Feld, Zwischensummen, Formeln, Fehlerwerte |
| `04_Belegung_Kitas_Monate.xlsx` | Monatsübersicht | Kreuztabelle (Monate als Spalten) in Zeilen umwandeln |
| `05_Adressliste_Ansprechpersonen.xlsx` | Telefonliste | Gruppenzeilen, Namen in vielen Formen, Telefon und PLZ als Text |
| `06_Inventar_IT.xlsx` | Inventar | senkrecht verbundene Zellen, ausgeblendete Spalte, Ja/Nein in Varianten |
| `07_Massnahmenplan.xlsx` | Maßnahmenplan | Mehrfachwerte in einer Zelle mit wechselnden Trennzeichen, Freitext |
| `08_Bedarfsmeldung_je_Kita.xlsx` | Rücklauf je Einrichtung | mehrere gleichartige Blätter zusammenführen, Blattname als Feld |
| `09_Projekte_und_Aufgaben.xlsx` | zwei Blätter | Verknüpfung über den Namen, abweichende Schreibweisen |
| `10_Beschlusskontrolle.xlsx` | Beschlussliste | Datum mit Uhrzeit, Abkürzungen, vage Fristen, lange Texte |
| `11_Fachverfahren_Export.csv` | Export eines Fachverfahrens | Windows-1252, ISO-Datum, Dezimalpunkt, führende Nullen, J/N |
| `12_Antrag_Formular.xlsx` | Antragsformular | Formular statt Liste: ein Datensatz untereinander |
| `13_Raumbelegung_Wochenplan.xlsx` | Wochenraster | Uhrzeit × Wochentag, Buchungen als verbundene Zellen |
| `14_Telefonliste_Amt_FINAL_neu.xlsx` | gewachsene Telefonliste | senkrecht verbundene Abteilungszellen, Durchwahlen in allen Schreibweisen, zwei Personen in einer Zelle, Farbe und Durchstreichen statt Status, Legende darunter, ausgeblendete Spalte, Raum „2.01“ von Excel zum Datum gemacht |
| `15_Veranstaltungsplanung_2026_Kopie.xlsx` | Planung im Entwurf | Kopf mit Zeilenumbrüchen, Datum als „KW 31-33“, „Mai 2026“, „tbd“, Kosten als „ca. 1.200 €“, „500,-“, „s.o.“, Ja/Nein als „ja!“, „(ja)“, „x“, Summenzeile mitten drin, Ideenliste darunter |
| `16_Fortbildungsbedarf_Rueckmeldungen.xlsx` | Rückmeldungen dreier Ämter | Blätter mit abweichender Spaltenfolge und -benennung, mehrere Wünsche per Zeilenumbruch, Kästchenzeichen, „dito“ und „siehe oben“, Auswertungsblatt mit Formeln, ausgeblendetes Altblatt |
| `2ask_komplex/Buergerbefragung_*.xlsx` | echte 2ask-Exporte der zweiten Testbefragung (9 Bögen) | Matrix mit Ausweichantwort, Mehrfachauswahl mit „Sonstiges“-Textfeld und exklusiver Antwort, Rangfolge, Filterfrage mit Anzeigelogik, Formular, Gruppenmatrix mit Dropdown; kodiert, mit Umkodierung und unkodiert |
| `2ask_Export_kodiert.xlsx` | Befragungsexport (Nachbau, 24 Teilnehmende) | vier Kopfzeilen, Codes statt Texte, Blatt „code_plan“, alle Fragearten |
| `2ask_Export_unkodiert.xlsx` | Befragungsexport (Nachbau) | Antworttexte statt Codes |
| `2ask_Export_kodiert.csv`, `2ask_Export_unkodiert.csv` | Befragungsexport (Nachbau) | CSV mit anderer Reihenfolge der Kopfzeilen, Dezimalkomma |
| `2ask_echt/Original_*.xlsx`, `*.csv` | echte 2ask-Exporte (3 Testbögen) | Referenz für den Nachbau, dazu eine Fassung mit Umkodierung fehlender Antworten |

Gemeinsame Regel für alle Listen: Werte, die nicht in den erkannten Typ passen, werden **nicht geraten**. Sie bleiben leer und erscheinen in der Prüfliste mit Zeile, Spalte und Originalwert; auf Wunsch wird der Originalwert in eine Bemerkung übernommen.

## Die Listen im Einzelnen

### 01 Teilnehmerliste Fortbildung (Blatt „Anmeldungen“)

**Aufbau:** Titel in Zeile 1, „Stand“ in Zeile 2, Kopf über zwei Zeilen (4 und 5) mit verbundenen Zellen: „Teilnahme“ steht über den drei Spalten „Modul 1 (12.05.)“, „Modul 2 (16.06.)“, „Modul 3 (08.09.)“. 20 Anmeldungen in Zeile 6 bis 25, Summenzeile 26 mit Formeln, Kopf fixiert.

**Fallen:** Kreuze als „x“ und „X“; einmal „entschuldigt“ statt Kreuz (die Summenformel zählt es mit); „Jugendamt “ mit Leerzeichen am Ende; Verpflegung als „vegetarisch“, „Veg.“, „vegan“, „keine Angabe“ oder leer.

**Soll:** Tabelle „Anmeldungen“ mit 20 Einträgen

| Feld | Typ | Hinweis |
|---|---|---|
| Nr. | Zahl | oder weglassen (nur laufende Nummer) |
| Name, Vorname | Text | Name als Bezeichnungsfeld |
| Dienststelle | Auswahl | 7 Werte, „Jugendamt “ = „Jugendamt“ |
| E-Mail | Text | |
| Teilnahme | Mehrfachauswahl | aus den drei Kreuzspalten: Modul 1, Modul 2, Modul 3 |
| Verpflegung | Auswahl | vegetarisch, vegan, keine Angabe |
| Bemerkung | Langtext | |

**Rückfragen:** Kopf über zwei Zeilen richtig erkannt? Summenzeile weglassen? Die drei Kreuzspalten zu einer Mehrfachauswahl „Teilnahme“ zusammenfassen (oder drei Ja/Nein-Felder)? „entschuldigt“ als „nicht ausgewählt“ werten und in die Bemerkung schreiben? „Veg.“ = „vegetarisch“?

### 02 Fristenliste Wiedervorlage (Blatt „WV 2025“)

**Aufbau:** Kopf in Zeile 1 mit Autofilter, 22 Vorgänge, Schrift Arial 10.

**Fallen:**
- Aktenzeichen „51.2-0101/25“ (Text, eindeutig: Schlüssel).
- „Eingang“ meist echtes Datum, aber auch „12.3.25“ als Text, „KW 14“, „Anfang Mai“ und leer.
- „erledigt am“ einmal „s. Vermerk“.
- Status in Schreibweisen: „erledigt“, „Erledigt“, „erl.“, „in Arbeit“, „in Bearb.“, „offen“, „zurückgestellt“.
- Kürzel „Mü“, „MÜ“, „Sch“, „Kl“, „Ba“ ohne Auflösung.
- Ampel nur als Füllfarbe der Spalte „Frist“, ohne Legende (gemeint: Stand 01.07.2025, rot = überfällig, gelb = in drei Wochen fällig, grün = erledigt oder später).

**Soll:** Tabelle „Wiedervorlagen“ mit 22 Einträgen: Az. (Text, Schlüssel), Betreff (Text, Bezeichnungsfeld), Eingang (Datum; „12.3.25“ wird 12.03.2025, die anderen Texte bleiben leer mit Hinweis), Frist (Datum), zust. (Auswahl Mü, Sch, Kl, Ba), Status (Auswahl offen, in Arbeit, zurückgestellt, erledigt), erledigt am (Datum), Bemerkung (Langtext).

**Rückfragen:** „erl.“ = „erledigt“ und „in Bearb.“ = „in Arbeit“? Was tun mit „KW 14“ und „Anfang Mai“? Die Ampelfarbe übernehmen (eher nicht, MidPoint färbt Fristen selbst)? Kürzel als Personen anlegen?

Bereits heute erledigt: Groß- und Kleinschreibung („erledigt“/„Erledigt“, „Mü“/„MÜ“) und Leerzeichen am Rand führt MidPoint beim Import schon zusammen.

### 03 Haushalt Mittelabfluss (Blatt „Übersicht“)

**Aufbau:** Titel, Hinweis „Beträge in Euro“, Kopf in Zeile 4. Drei Gruppenüberschriften als verbundene, farbige Zeilen („Produkt 36.1 Förderung von Kindern in Tageseinrichtungen“ usw.), darunter die Kontozeilen, je eine Zwischensumme, Leerzeilen, Gesamtzeile. „Rest“ und „Abfluss in %“ sind Formeln.

**Fallen:** Gruppenzeilen sind keine Datensätze, gehören aber als Wert an jede Kontozeile darunter. „k.A.“ im Ist; dadurch zeigen Rest und Abfluss in dieser Zeile „#WERT!“. Ein negativer Ist-Wert (Korrekturbuchung). Konto als siebenstellige Zahl.

**Soll:** Tabelle „Haushalt“ mit 10 Konten: Produkt (Auswahl aus den 3 Gruppenzeilen), Konto (Text, Schlüssel), Bezeichnung (Text), Ansatz 2025 (Geld), Ist bis 30.06. (Geld; „k.A.“ wird leer mit Hinweis). Rest und Abfluss nicht übernehmen, sondern in MidPoint als Berechnung anlegen (oder als Wert, wenn gewünscht). Zwischensummen, Leerzeilen und Gesamt entfallen.

**Rückfragen:** Gruppenzeilen als Feld „Produkt“ übernehmen? Summenzeilen weglassen? Formelspalten neu berechnen oder als Wert übernehmen? „#WERT!“ und „k.A.“ als leer?

### 04 Belegung Kitas nach Monaten

**Aufbau:** Kreuztabelle: Einrichtungen in den Zeilen, Monate Jan bis Dez in den Spalten, Kopf in Zeile 3, Summenzeile 14.

**Fallen:** Monate sind Spalten statt Zeilen. Okt bis Dez noch leer. „Umbau“ als Text in zwei Zahlenzellen. Die Summenzeile zeigt 0 für die leeren Monate.

**Soll (empfohlen):** Stammdaten „Einrichtungen“ (10 Einträge: Name, Plätze lt. Betriebserlaubnis) und Tabelle „Belegung“ in Zeilenform: Einrichtung (Verknüpfung), Monat (Datum, z. B. 01.01.2025), belegte Plätze (Zahl), Bemerkung („Umbau“). Das ergibt 90 Einträge (10 Einrichtungen × 9 Monate). **Alternative:** eine Tabelle mit 12 Zahlenfeldern Jan bis Dez (einfacher, aber schlecht auszuwerten).

**Rückfragen:** Monatsspalten in Zeilen umwandeln? Leere Monate auslassen? Welches Jahr (aus dem Titel: 2025)? „Umbau“ leer lassen und als Bemerkung übernehmen?

### 05 Adressliste Ansprechpersonen

**Aufbau:** Titel, „Stand: Mai 2025“, Kopf in Zeile 4, Gruppenzeilen „Stadtverwaltung“, „Freie Träger“, „Gesundheitswesen“, „Externe Fachstellen“.

**Fallen:** Namen als „Becker, Anna“, „Dr. Martin Keller“, „Fr. Özdemir“ oder Einrichtung statt Person. Telefon in vielen Schreibweisen („0123 45-3456“, „0123/45-2000“, „+49 123 555012“, „0611 98 76 54“). PLZ als Text mit führender Null („01067“).

**Soll:** Tabelle „Ansprechpersonen“ mit 10 Einträgen: Bereich (Auswahl aus den 4 Gruppenzeilen), Name (Text, nicht automatisch zerlegen), Funktion, Telefon (Text, nie Zahl), E-Mail, Straße, PLZ (Text), Ort (Auswahl oder Text).

**Rückfragen:** Gruppenzeilen als Feld „Bereich“? Namen in Vor- und Nachname trennen, wo es eindeutig ist, oder als ein Feld lassen?

### 06 Inventar IT

**Aufbau:** Kopf in Zeile 1, „Gebäude“ als senkrecht verbundene Zellen über mehrere Zeilen, Spalte K „intern“ ausgeblendet.

**Fallen:** Das Gebäude steht nur in der ersten Zeile des Blocks (beim Lesen nach unten auffüllen). Raum „1.12“ ist Text und darf nicht zu 1,12 werden. Seriennummern mit führenden Nullen. „defekt“ als „ja“, „nein“, „x“ oder leer. „Garantie bis“ als Datum oder „abgelaufen“. „Laptop“ und „Notebook“ meinen dasselbe.

**Soll:** Tabelle „Inventar“ mit 17 Geräten: Gebäude (Auswahl, aufgefüllt), Raum (Text), Inventar-Nr. (Text, Schlüssel), Gerät (Auswahl), Hersteller / Modell (Text), Seriennummer (Text), Anschaffung (Datum), Wert (Geld), defekt (Ja/Nein), Garantie bis (Datum; „abgelaufen“ leer mit Hinweis). Die ausgeblendete Spalte nur nach Rückfrage.

**Rückfragen:** Verbundene Zellen auffüllen? Ausgeblendete Spalte „intern“ übernehmen? „x“ = ja und leer = nein? „Laptop“ und „Notebook“ zusammenlegen?

### 07 Maßnahmenplan

**Aufbau:** Titel, Kopf in Zeile 3, 10 Maßnahmen mit viel Freitext und Zeilenumbrüchen in den Zellen.

**Fallen:**
- Mehrfachwerte in **einer** Zelle mit wechselnden Trennzeichen: „Sport, Gesundheit“, „Bildung; Sprache“, „Freizeit / Prävention“.
- Zielgruppe teils mehrfach: „Kinder, Jugendliche“ und „Kinder und Jugendliche“.
- Zeitraum als Freitext: „Sommerferien 2025“, „03/2025 bis 06/2025“, „Q2 2025“, „laufend“.
- Kosten als Text: „ca. 5.000 €“, „3.000 bis 4.000 €“, „800“, „k.A.“, „noch offen“, „15.000 € pro Jahr“.

**Soll:** Tabelle „Maßnahmen“: Maßnahme (Text), Beschreibung (Langtext), Themen (Mehrfachauswahl mit 14 Werten: Sport, Gesundheit, Bildung, Sprache, Freizeit, Prävention, Integration, Begegnung, Inklusion, Sicherheit, Stadtplanung, Medien, Kultur, Ernährung), Zielgruppe (Mehrfachauswahl: Familien, Kinder, Jugendliche, Eltern, Senioren, alle), Zeitraum (Text), Kosten (Text; eine Zahl nur auf Wunsch und nur bei eindeutigem Betrag), verantwortlich (Auswahl), Stand (Langtext).

**Rückfragen:** Themen als Mehrfachauswahl mit den Trennzeichen Komma, Semikolon und Schrägstrich richtig? „Kinder und Jugendliche“ = Kinder + Jugendliche? Kosten als Text lassen?

### 08 Bedarfsmeldung je Kita

**Aufbau:** Blatt „Hinweise“ ohne Daten, dazu vier Blätter je Einrichtung mit (fast) gleichem Aufbau, Kopf jeweils in Zeile 3.

**Fallen:** Der Blattname ist ein Datenwert (Einrichtung). „Menge“ heißt auf einem Blatt „Anzahl“. Das Blatt „Kita Kunterbunt“ hat eine zusätzliche Spalte „Lieferant“ mittendrin. Einheiten als „Stück“, „Stk.“, „Packung“, „Pck.“.

**Soll:** **eine** Tabelle „Bedarfsmeldungen“ mit 16 Einträgen: Einrichtung (Auswahl aus den Blattnamen), Artikel, Menge (Zahl), Einheit (Auswahl; Stk. = Stück, Pck. = Packung), Lieferant (Text, meist leer), geschätzter Preis (Geld), Dringlichkeit (Auswahl gering, mittel, hoch, sehr hoch), Begründung (Langtext).

**Rückfragen:** Alle Einrichtungsblätter zusammenführen und den Blattnamen als Feld „Einrichtung“ übernehmen? Blatt „Hinweise“ überspringen? „Anzahl“ = „Menge“? Abkürzungen der Einheiten zusammenführen?

### 09 Projekte und Aufgaben

**Aufbau:** Blatt „Projekte“ (6) und Blatt „Aufgaben“ (14). Aufgaben nennen ihr Projekt über den Namen.

**Fallen:** Zwei Projektnamen weichen ab: „Umbau Bibliothek“ statt „Umbau Stadtteilbibliothek“, „Website“ statt „Barrierefreie Website“. „erledigt“ als „x“ oder leer. Eine Frist „vor Infoabend“.

**Soll:** zwei verknüpfte Tabellen. Projekte: Nr., Projekt (Bezeichnungsfeld), Leitung, Beginn, Ende (Datum), Budget (Geld), Status (Auswahl aktiv, geplant, abgeschlossen). Aufgaben: Aufgabe, Projekt (Verknüpfung), zuständig (Auswahl), Frist (Datum), erledigt (Ja/Nein).

**Rückfragen:** Die beiden Blätter als verknüpfte Tabellen übernehmen? Zwei Projektnamen nicht gefunden: „Umbau Bibliothek“ gehört zu „Umbau Stadtteilbibliothek“, „Website“ zu „Barrierefreie Website“ (ähnlichster Name als Vorschlag)? Leitung und zuständig als gemeinsame Personenliste?

### 10 Beschlusskontrolle

**Fallen:** Sitzung als Datum mit Uhrzeit, einmal nur „Mai 2025“. TOP gemischt („TOP 4“, „5.1“, „7“), daher Text. Gremium als Langform und Abkürzung („Jugendhilfeausschuss“ und „JHA“). „Umsetzung bis“ als Datum oder Text („nach den Sommerferien“, „fortlaufend“). Zuständig einmal mehrfach („Amt 40, Amt 61“). Lange Texte mit Zeilenumbruch.

**Soll:** Tabelle „Beschlüsse“ mit 8 Einträgen: Sitzung (Datum mit Uhrzeit), TOP (Text), Gremium (Auswahl Jugendhilfeausschuss, Sozialausschuss, Stadtrat), Beschluss (Langtext), zuständig (Mehrfachauswahl Amt 40, Amt 50, Amt 51, Amt 61), Umsetzung bis (Datum; Texte leer mit Hinweis), Stand der Umsetzung (Langtext).

**Rückfragen:** „JHA“ = „Jugendhilfeausschuss“? „Mai 2025“ als 01.05.2025 oder leer? Zuständig als Mehrfachauswahl?

### 11 Fachverfahren Export (CSV)

**Aufbau:** technischer Export: Semikolon, Windows-1252 ohne BOM, Kopf in Großbuchstaben mit Unterstrich, ISO-Datum („2025-01-14“), Dezimalpunkt („1450.00“), Fallnummern mit führenden Nullen („0004711“), Kennzeichen J/N, Kürzel der Sachbearbeitung, ein Feld mit Semikolon in Anführungszeichen.

**Soll:** Tabelle „Fälle“ mit 12 Einträgen: Fallnummer (Text, Schlüssel, Nullen bleiben), Eingang (Datum), Leistungsart (Auswahl, 5 Werte), Kostenstelle (Text), Betrag (Geld: „1450.00“ ist 1.450,00 €, nicht 145.000), aktiv (Ja/Nein aus J/N), Sachbearbeitung (Auswahl), Bemerkung (Text).

**Rückfragen:** Feldnamen lesbar machen (FALLNR wird „Fallnummer“)? „0.00“ als 0 oder leer?

### 12 Antrag (Formular)

**Aufbau:** kein Listenblatt, sondern ein Formular: Bezeichnung mit Doppelpunkt in Spalte A, Wert in Spalte B, darunter ein kleiner „Kostenplan“ mit Summe, Ort und Datum, Unterschriftszeile.

**Soll:** Erkennen „Formular, keine Liste“: **ein** Eintrag in der Tabelle „Anträge“ mit Antragsteller, Ansprechperson, Telefon, E-Mail, Maßnahme, Zeitraum (Text oder zerlegt in „von“ und „bis“), Teilnehmende (Zahl), davon mit Behinderung (Zahl), Beantragter Zuschuss (Geld). Kostenplan als zweite Tabelle mit 4 Positionen und Verknüpfung zum Antrag, oder weglassen. Praktischer Nutzen: viele gleich aufgebaute Antragsdateien in einem Durchgang einsammeln, ähnlich wie die Rückläufer der Erfassungsdateien.

**Rückfragen:** Formular erkannt, Felder so übernehmen? Zeitraum in „von“ und „bis“ zerlegen? Kostenplan übernehmen?

### 13 Raumbelegung Wochenplan

**Aufbau:** Raster mit Uhrzeiten 08:00 bis 18:00 in den Zeilen und Montag bis Freitag in den Spalten. Buchungen sind verbundene, farbige Zellen über eine oder mehrere Stunden. Woche und Datum stehen nur im Titel („KW 20 (12. bis 16. Mai 2025)“).

**Soll:** Tabelle „Belegungen“ mit 8 Einträgen: Veranstaltung (Text), Beginn und Ende (Datum mit Uhrzeit, aus Titel, Wochentag und Zeile), Raum (aus dem Titel: Saal). Beispiel: „Sprachkurs A1“ am Mittwoch über drei Zeilen ab 10:00 wird 14.05.2025 10:00 bis 13:00.

**Rückfragen:** Belegungsraster (Uhrzeit × Tag) erkannt? Datum der Woche aus dem Titel richtig? Ende = letzte Zeile plus eine Stunde? Farben ignorieren?

Dieser Fall ist der schwierigste. Wenn die Erkennung nicht sicher ist, ist „nicht unterstützt, bitte als Liste anlegen“ eine ehrliche Antwort.

### 14 bis 16: gewachsene Listen

Erzeugen mit `powershell -ExecutionPolicy Bypass -File build\testlisten_unprof.ps1`. Diese drei Listen sind absichtlich so, wie Listen in Ämtern über Jahre entstehen: viele Hände, keine Vorgaben.

**14 Telefonliste:** Abteilung als senkrecht verbundene Zelle über mehrere Zeilen, Leerzeile zwischen den Abteilungen, „SB“ für Sachbearbeitung, Durchwahlen als „0123 45-3401“, „-3403“, „intern 3404“, „12-3411 / -3412“, zwei Personen in einer Zelle, gelb hinterlegt (nicht im Dienst) und durchgestrichen (ausgeschieden) mit Legende unter der Tabelle, ausgeblendete Spalte „alte Durchwahl“, ein Kommentar. **Soll:** 11 Einträge, Abteilung an jedem Eintrag, „SB“ als „Sachbearbeitung“, Durchwahl als Text, ausgeblendete Spalte nicht übernehmen.

**15 Veranstaltungsplanung:** Kopfzellen mit Zeilenumbruch, Datum als echtes Datum oder als „jeden 1. Samstag“, „KW 31-33“, „Mai 2026“, „11.11.“, „tbd“, Kosten als „ca. 1.200 €“, „kostenlos“, „500,-“, „s.o.“, Ja/Nein als „nein“, „ja!“, „x“, „(ja)“, „j“, Summenzeile mitten in der Tabelle, darunter „Ideen (noch offen):“. **Soll:** 8 Einträge, Datum als Text (zu viele freie Angaben), Kosten als Geldbetrag mit Rückfrage zu „kostenlos“, „s.o.“ als Wert der Zeile darüber, „Anmeldung nötig?“ als Ja/Nein, Ort und Zuständigkeit als Text (keine Mehrfachauswahl), Summenzeile und Ideenliste weglassen.

**16 Fortbildungsbedarf:** drei Blätter je Amt, Spalten in anderer Reihenfolge und anders benannt („Name“, „Mitarbeiter/in“; „Fortbildungswunsch“, „Wunsch“; „Kosten ca.“, „Kosten“), eine zusätzliche Spalte „Genehmigt AL“, mehrere Wünsche per Zeilenumbruch in einer Zelle, „☑“ und „☐“, „dito“ und „siehe oben“, Blatt „Auswertung“ mit Formeln, ausgeblendetes Blatt „alt“. **Soll:** alle drei Ämter zusammen (11 Einträge) mit Rückfragen zu den abweichenden Spaltennamen, Blattname als Feld, „dito“ als Wert der Zeile darüber, Auswertung und Altblatt nicht als Tabelle.

### Behördenlisten aus dem Internet (nur lokal)

Vier öffentliche Excel-Listen liegen zum Testen in `testdaten/online/` (nicht im Lieferpaket und nicht im Repository; Quellen und Bezug in `testdaten/online/QUELLEN.md`): Gemeindeverzeichnis des Statistischen Bundesamts (rund 16.000 Zeilen, dreizeiliger Kopf mit Stichtagszeile), Schulverzeichnis des Statistischen Landesamts Rheinland-Pfalz (acht Blätter), Liste der EFRE-Vorhaben der Berliner Senatsverwaltung (zweisprachiger Kopf, rund 3.000 Vorhaben) und Liste der Staatennamen des Auswärtigen Amts. Testlauf: `build\listimport_test.ps1 -Set neu`.

## 2ask: so sehen die Exporte aus (geprüft)

Geprüft mit einer echten Testbefragung im kostenlosen Testtarif von 2ask (November 2026): 20 Fragen mit allen dort verfügbaren Fragearten, drei Pretest-Durchgänge (einer vollständig, einer lückenhaft, einer mit Sonderzeichen wie „Ayşe“, Anführungszeichen, Semikolon und Zeilenumbruch) und fünf Exporte. Die Originale liegen in `2ask_echt/`; nur die Zeitstempel der drei Testbögen sind dort auf einen Tag der Beispielwelt verschoben (10.03.2026, Uhrzeiten unverändert). Der Nachbau `2ask_Export_*` hat genau denselben Aufbau, aber 24 erfundene Teilnehmende. Das Skript wurde automatisch gegen die Originale verglichen: alle vier Kopfzeilen, Form der Import-JSON, Datentyp jeder Spalte, Codeplan (148 Zeilen) und die Anführungszeichen im CSV stimmen überein.

### Exportoptionen (Ergebnisse, Ergebnisse herunterladen)

- **Fragebögen:** abgeschlossene oder nicht abgeschlossene.
- **Teilnehmer:** „Pretest Teilnehmer“ oder „Teilnehmer (ohne Pretest)“, jeweils mit Anzahl.
- **Antwort Ausgabe:** „Kodierte Ergebnisse/Antworten“ oder „Unkodierte Ergebnisse (Antworttext statt Kodierung)“, dazu „Umkodierung von nicht beantworteten Fragen“ (siehe unten).
- **Export Format:** Excel (XLSX) mit „Codeplan hinzufügen“, CSV (Dezimaltrenner Komma, Feldtrenner Strichpunkt, Textbegrenzer doppelte Anführungszeichen, alle drei einstellbar) oder SPSS.
- **Filter** aus „Reports & Diagramme“.
- Der Export wird erst erzeugt (Benachrichtigung mit Fortschritt) und dann heruntergeladen. Dateiname: Zeitstempel, dann `_export_survey_sessions_` und die Umfrage-ID, also `JJJJMMTThhmmss_export_survey_sessions_<ID>.xlsx` bzw. `.csv`.

### Aufbau XLSX

- Erzeugt mit Apache POI. Blatt **„results“**, beim kodierten Export mit Codeplan zusätzlich Blatt **„code_plan“**.
- **Vier Kopfzeilen:** 1 Fragetext (bei Metadaten englisch, z. B. „Started at“), 2 Unterbeschriftung („Answer-Option: Fachberatung“, „Answer: Sprache; Answer-Option:unter 3 Jahren“, „Answer: Fachberatung; Group: Wichtigkeit“), 3 technischer Import-Schlüssel als JSON (`{"import_id":"…"}`), 4 Variablenname (`Q1`, `Q2_01` …).
- **Ab Zeile 4 eine Excel-Tabelle „results“** mit den Variablennamen als Spaltenköpfen (Filterknöpfe, Format „TableStyleMedium2“). Ein Import kann diese Tabelle direkt lesen.
- Codes und Zahlen als Zahl, Texte als Text, Start und Ende als echtes Datum mit Uhrzeit (Format TT.MM.JJJJ hh:mm:ss). Antworten aus Datums- und Zeitfeldern dagegen als **Text** („01.08.2015“, „07:00“). Mehrzeilige Antworten mit Zeilenumbruch in der Zelle.
- **24 Metadatenspalten** vor den Fragen: `survey_session_id` (13 Zeichen, z. B. „WVWYXS9DKJA60“), `survey_session_type` („pretest“), `started_at`, `ended_at`, `last_interaction_at`, `progress_percent`, `duration_in_seconds`, `items_answered_count`, `items_answered_percent`, `locale`, `source` („weblink“), `country`, `device_type`, `browser`, `browser_locale`, `browser_version`, `operating_system` und sieben `contact.*`-Spalten (leer, außer bei Versand über die Kontaktliste).

### Aufbau CSV

- UTF-8 mit BOM, Semikolon. Datensätze sind mit **LF** getrennt (nicht CRLF); mehrzeilige Antworten enthalten CRLF innerhalb des Feldes.
- **Kopfzeilen in anderer Reihenfolge als im XLSX:** 1 Variablenname, 2 Fragetext, 3 Unterbeschriftung, 4 Import-JSON.
- Texte, Codes, Datumsangaben und Dezimalzahlen stehen in Anführungszeichen (`"2"`, `"4,0"`, `"10.03.2026 18:01:02"`), die ganzen Zahlen der Metadaten nicht (`100;32;17`). Leere Felder sind ganz leer (`;;`).
- Dezimalkomma; Werte aus Sternen, Schieberegler, Rangfolge, Summe und Zahlfeldern immer mit Nachkommastelle („70,0“).

### Spalten je Frageart

| Frageart | Variablen (Testbefragung) | kodiert | unkodiert | Codeplan `sub_type` |
|---|---|---|---|---|
| Einfachauswahl | `Q1` | Code 1 bis n | Optionstext | radio |
| Dropdown | `Q3` | Code | Optionstext | select |
| Mehrfachauswahl | `Q2_01` bis `Q2_05` | **Code der Option** (1, 3, 5), nicht 1/0; nicht gewählt = leer | Optionstext oder leer | checkbox |
| Zusatztext an einer Option („Sonstiges, und zwar“) | `Q2t_05`, steht **vor** `Q2_01` | Text | Text | text_entry |
| Matrix Likert | `Q4_1`, `Q4_2`, `Q4_4` | Code | Text | matrix_table, radio |
| Matrix Mehrfachauswahl | `Q5_1_01` bis `Q5_3_03` (Zeile_Spalte) | Code der Spalte | Spaltentext | checkbox |
| Matrix Textfeld | `Q6_1_01` bis `Q6_2_02` | Text | Text | input |
| Matrix Dropdown | `Q7_1`, `Q7_2` | Code | Text | select |
| Matrix Bipolar | `Q8_1`, `Q8_2` | Code 1 bis 5 | „[4]“ (Optionen ohne Text) | radio |
| Sterne | `Q9_1`, `Q9_2` | Zahl („4,0“) | Zahl | slider, range |
| Schieberegler | `Q10` | Zahl in der Schrittweite des Reglers (35 wurde 40) | Zahl | slider, range |
| Formular | `Q11_1` bis `Q11_3` | Text; Dropdown im Formular als Code | Text | form, input, select_coded |
| Gruppenmatrix | `Q12_1_01`, `Q12_2_01`, `Q12_1_02`, `Q12_2_02` (Zeile_Gruppe, gruppenweise sortiert) | Code je Gruppe | Text | side_by_side, group, radio |
| Rangfolge | `Q13_01` bis `Q13_04` (je Element) | Rang als Zahl („2,0“) | Rang | rank_order, drag_drop |
| Summe | `Q14_1` bis `Q14_3` | Zahl | Zahl | constant_sum |
| Eingabefeld einzeilig, mehrzeilig, E-Mail | `Q15`, `Q16`, `Q20` | Text | Text | text_entry (string) |
| Eingabefeld Zahl | `Q17` | Zahl („48,5“) | Zahl | text_entry (number) |
| Eingabefeld Datum | `Q18` | Text „01.08.2015“ | Text | text_entry (localdate) |
| Eingabefeld Zeit | `Q19` | Text „07:00“ | Text | text_entry (localtime) |

Die Nummern zählen in der Reihenfolge des Anlegens und schließen Lücken nicht: In der Testbefragung wurde die dritte Zeile von Q4 gelöscht, deshalb gibt es `Q4_4`, aber kein `Q4_3`.

### Codeplan (Blatt „code_plan“)

Langformat mit den Spalten `export_key`, `type`, `sub_type`, `data_type`, `coding`, `de`. Je Frage eine Zeile `type = question` (mit der Frageart als `sub_type`), je Matrixzeile eine Zeile `answer`, je Option eine Zeile `answer_option` mit Code und Text, bei der Gruppenmatrix zusätzlich `group`. `data_type` ist `string`, `number`, `date`, `localdate` oder `localtime`. Damit lässt sich jede Spalte ohne Raten einem MidPoint-Typ zuordnen: Auswahl, Mehrfachauswahl, Zahl, Datum, Uhrzeit oder Text.

### Fehlende Antworten

- **Standard: leer.** Ein fester Code wie -99 erscheint nur, wenn „Umkodierung von nicht beantworteten Fragen“ eingeschaltet ist.
- Einstellbar sind getrennte Codes für Checkbox- und Radiobox-Antworten „gesehen, aber nicht beantwortet“ (Vorgabe -99), abweichend für Mehrfachauswahl, „nicht gesehen“ (Anzeigelogik oder Sprünge) und dieselben zwei Fälle für Eingabefelder.
- Test mit -99, -98 und -97 (Datei `2ask_echt/Original_XLSX_kodiert_Codeplan_Umkodierung_99_98_97.xlsx`): Einfachauswahl, Dropdown und Matrix bekamen -99; **jede nicht gewählte Option** einer Mehrfachauswahl bekam -98, auch wenn andere Optionen gewählt waren, ebenso die Gruppenmatrix; Texte, Sterne, Schieberegler und Rangfolge bekamen -97, und zwar **als Text** statt als Zahl.

### Was das für den Import heißt

- **2ask erkennen** an `survey_session_id` in der Kopfzeile, an der Tabelle „results“ oder am Blatt „code_plan“.
- Die übrigen Kopfzeilen überspringen (XLSX Zeile 1 bis 3, CSV Zeile 2 bis 4) und die Variablennamen als Schlüssel nehmen; Fragetext und Unterbeschriftung ergeben die Feldbezeichnung.
- **Mit Codeplan** stehen Typ, Optionen und Codes fest; es bleibt nur die Rückfrage, welche Fragen und Metadaten übernommen werden. **Ohne Codeplan** ist der unkodierte Export besser, denn Codes allein lassen sich nicht deuten.
- Spalten einer Mehrfachauswahl (`Qn_01`, `Qn_02` …) zu einem Feld zusammenfassen, `Qnt_xx` als Zusatztext dazu.
- Umkodierte Werte (-99, -98, -97, auch als Text) als leer behandeln.
- `survey_session_id` eignet sich als Schlüssel für eine spätere Aktualisierung; Browser, Gerät und Kontaktfelder in der Regel weglassen.

### Zweite Testbefragung „Bürgerbefragung Bürgerservice“ (Ordner `2ask_komplex/`)

Sechs Fragen, die mehrere Besonderheiten verbinden: Matrix mit fünf Zeilen und der Ausweichantwort „kann ich nicht beurteilen“, Mehrfachauswahl mit mehrzeiligem Textfeld bei „Sonstiges, und zwar:“ und exklusiver Antwort „Keinen dieser Wege“, Rangfolge mit fünf Punkten, Seitenwechsel, Filterfrage (nur sichtbar, wenn „Online-Serviceportal“ gewählt wurde), Formular mit vier Textfeldern (Datum, Wartezeit, Anzahl, Anliegen) und Gruppenmatrix (Ja/Nein je Dienstleistung und Schulnote als Dropdown). Neun Pretest-Bögen über die Vorschau, darunter Antworten mit Semikolon und Zeilenumbruch im Textfeld, „ca. 20“ und „25,5“ als Wartezeit, Datum als „3.3.26“ und „2026-03-05“, ein fast leerer Bogen und einer nur mit „Keinen dieser Wege“. Zeitstempel auf den 10.03.2026 verschoben.

Neu gelernt:
- Die Rangfolge nimmt nur alle Ränge oder keinen („Alle Ränge müssen eindeutig vergeben werden.“).
- Nicht gezeigte Filterfragen sind im Export leer; mit Umkodierung bekommen sie den Code für „nicht gesehen“ (hier -97).
- Ein in der Vorschau abgebrochener Bogen erscheint nicht unter „nicht abgeschlossene Fragebögen“ (Anzahl 0).
- Im Formular sind alle Felder Textfelder; Datum und Zahl stehen deshalb als Text im Export und werden beim Übernehmen aus den Antworten erkannt.
- Unkodierter Export mit Codeplan: Antworttexte statt Codes, der Codeplan liefert dazu die Optionen in Reihenfolge.

### Noch nicht geprüft

- Wert von `survey_session_type` bei echten Teilnehmenden (gesehen nur „pretest“; dafür hätte die Umfrage öffentlich freigeschaltet werden müssen).
- Weiterempfehlungsfrage und Datei-Upload (im Testtarif gesperrt), nicht abgeschlossene Fragebögen echter Teilnehmender, SPSS, mehrsprachige Umfragen.

### Quellen

- Eigene Testbefragung „Testexport Fragearten“ im 2ask-Testtarif, Exporte vom November 2026 (Ordner `2ask_echt/`)
- [2ask-Hilfe: Ergebnisse herunterladen](https://www.2ask.com/docs/umfragen/ergebnisse-herunterladen/)
- [2ask-Hilfe: Fragen editieren](https://www.2ask.com/docs/umfragen/fragebogeneditor/frageelemente/fragen-editieren/)
- [2ask-Hilfe: OData-Feed für Power BI und Tableau](https://www.2ask.com/docs/umfragen/datenanbindung/odata-feed-power-bi-tableau/)
- [2ask-Preise: Funktionsumfang des Testtarifs](https://www.2ask.com/preise/)
