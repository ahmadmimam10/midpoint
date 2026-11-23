# Projektauftrag: Excel-Datenbankbaukasten

Version 1.0 · Oktober 2026

## 1. Entwicklungsauftrag und Ziel

Entwickle eine vollständig offline nutzbare Excel-/VBA-Anwendung mit dem Arbeitstitel **Excel-Datenbankbaukasten**. Übernimm dabei Architektur, Oberflächengestaltung, Implementierung und Qualitätssicherung.

Die Anwendung ermöglicht Menschen ohne vertiefte Excel- oder Programmierkenntnisse, eigene strukturierte Erfassungs- und Verwaltungsanwendungen zusammenzustellen. Vorbild ist das Bedienprinzip von Notion-Datenbanken: frei definierbare Felder, verknüpfte Tabellen, mehrere Ansichten derselben Daten und einfach konfigurierbare Auswertungen. Eine technische Verbindung zu Notion ist nicht vorgesehen.

Typische Einsätze sind Aufgabenverwaltung, Projektübersichten, Schülerlisten, Dokumentation, Kontakte, Ressourcen- und Materiallisten. Die Anwendung bleibt fachlich neutral. Keine Fachbereichsstruktur wird fest einprogrammiert.

**Nutzungsprinzip:** Die Person kopiert die leere Masterdatei, benennt die Kopie um, richtet darin ihre Tabellen und Ansichten ein und speichert sie als eigenständige Fachdatei. Jede Kopie enthält ihren vollständigen Programmcode, ihre Konfiguration und ihre Daten. Sie benötigt die ursprüngliche Masterdatei anschließend nicht mehr.

Ziel ist eine tatsächlich funktionsfähige Anwendung. Dieser Auftrag umfasst die Entwicklung; das Vorliegen dieses Dokuments bedeutet noch keine Umsetzung oder technische Prüfung.

## 2. Verbindliche Rahmenbedingungen

- Zielplattform: Microsoft Excel 365 Desktop unter Windows; Unterstützung von 32- und 64-Bit-Office ohne zusätzliche Installation.
- Dateiformat: XLSM mit tatsächlich enthaltenem VBA-Code und den erforderlichen Formularen.
- Vollständiger Offlinebetrieb: Einrichtung, Erfassung, Suche, Ansichten, Berechnungen, Auswertung, Hilfe, Import, Export und Sicherung funktionieren ohne Internet.
- Keine Cloud, Anmeldung, Onlineaktivierung, Telemetrie, Webdienste, externen APIs, Downloads zur Laufzeit oder nachgeladenen Schriften, Icons und Bibliotheken.
- Keine erforderlichen Add-ins, Plugins, zusätzlichen Apps, lokalen Server oder externen Datenbankdienste. Kein Browser als Voraussetzung.
- Makros müssen nach den Regeln der Organisation zugelassen sein. Sicherheitsrichtlinien werden nicht automatisch verändert oder umgangen. Eine notwendige Freigabe oder Signierung ist in der Betriebsanleitung zu erläutern.
- Eine Person bearbeitet eine Datei zur gleichen Zeit. Speicherung lokal oder auf einem gewöhnlichen Netzlaufwerk; kein gleichzeitiges Schreiben mehrerer Personen.
- Außerhalb der Arbeitsmappe dürfen nur bewusst erzeugte Exporte, Sicherungen und ausdrücklich gewählte lokale Dateiverweise entstehen.
- Oberfläche und Hilfe sind deutschsprachig. Datums- und Zahlenformate folgen den deutschen Konventionen.
- Kein dauerhaft erforderlicher Zugriff auf das VBA-Projektobjektmodell auf Anwenderrechnern. Die Anlage neuer Tabellen oder Felder darf keine Codegenerierung oder Änderung des VBA-Projekts erfordern.

## 3. Prioritäten und Umfang

Verbindlich sind alle nachfolgend beschriebenen Kernfunktionen: Tabellen- und Feldbaukasten, dynamische Erfassung, Beziehungen, gespeicherte Ansichten einschließlich Kalender, Board und Gantt, Berechnungen, Dashboards, Import/Export und zuverlässige Datenpflege.

Die Entwicklung erfolgt in funktionsfähigen Etappen. Eine Etappe darf klar benannte Funktionen noch nicht enthalten; die abschließende Lieferung erfüllt den gesamten Kernumfang. Keine leeren Schaltflächen oder nur optisch dargestellten Funktionen.

Zusätzlich wünschenswert sind Drag-and-drop in Ansichten, komfortable Rücknahme mehrerer Arbeitsschritte und weitere visuelle Darstellungen. Sie dürfen die Stabilität der Kernfunktionen nicht beeinträchtigen.

Nicht beauftragt sind Cloud-Synchronisierung, echte parallele Mehrbenutzerbearbeitung, mobile Nutzung, KI-Funktionen, eine freie Notion-Seiten-/Blockbearbeitung, komplexe Projektplanung mit Ressourcenoptimierung oder eine vollständige Unternehmensdatenbank.

## 4. Oberfläche und Bedienprinzip

### 4.1 Einheitliche Arbeitsoberfläche

Gestalte eine ruhige, moderne Oberfläche mit klarer Navigation, großzügigen Abständen und zurückhaltenden Farben. Verwende einen hellen Hintergrund, dunkle gut lesbare Schrift und eine konsistente Akzentfarbe. Status- und Auswahlwerte erhalten unterscheidbare Farben und zusätzlich lesbare Bezeichnungen.

Die Navigation enthält Start, die vorhandenen Tabellen, Dashboards, Einrichten und Hilfe. Zur aktuellen Tabelle erscheinen ihre Ansichten als eindeutig benannte Register oder gleichwertige Umschalter. Suche, Filter, neuer Eintrag und Ansichtsverwaltung befinden sich an konsistenten Stellen.

Nutze gestaltete Arbeitsblätter, native Excel-Tabellen, Shapes und passende VBA-UserForms. Wähle eine robuste Kombination anhand eines frühen technischen Prototyps. Standardmäßig soll die Person mit der gestalteten Arbeitsoberfläche arbeiten; strukturierte Daten bleiben über einen ausdrücklich gewählten Datenzugang erreichbar.

Bei Programmende und Fehlern müssen von der Anwendung geänderte Excel-Einstellungen wiederhergestellt werden. Andere geöffnete Arbeitsmappen dürfen nicht dauerhaft beeinflusst werden.

### 4.2 Zwei Modi

**Einrichten:** Tabellen, Felder, Auswahlmöglichkeiten, Beziehungen, Ansichten, Berechnungen und Dashboards konfigurieren.

**Arbeiten:** Daten erfassen und pflegen, vorhandene Ansichten nutzen, suchen, filtern, auswerten und exportieren. Strukturänderungen werden nicht versehentlich ausgelöst.

Ein optionaler Einrichtungsschutz dient gegen Fehlbedienung. Er wird nicht als sichere Benutzerverwaltung oder Zugriffskontrolle bezeichnet.

### 4.3 Konkrete Gestaltungsanforderungen

- Alle Bereiche einschließlich Dialogen, Leerzuständen, Fehlermeldungen und Auswertungen folgen demselben Design.
- Hauptaktionen tragen verständliche deutsche Beschriftungen; Icons sind ergänzend.
- Pflichtfelder sind erkennbar. Fehlermeldungen stehen möglichst am betroffenen Feld und erklären die Korrektur.
- Dynamische Masken sind bei vielen Feldern scrollbar, sinnvoll gruppiert und per Tastatur bedienbar.
- Ungespeicherte Eingaben gehen beim Wechsel zu einem verknüpften Eintrag nicht unbemerkt verloren.
- Prüfung mindestens bei 1366 × 768 und 1920 × 1080 sowie 100 % und 125 % Windows-Skalierung.
- Keine unlesbaren Texte, abgeschnittenen Bedienelemente oder überdeckten Schaltflächen. Breite Tabellen dürfen horizontal scrollen.

## 5. Tabellen und Felder zusammenstellen

Ein Assistent führt durch Name und Beschreibung der neuen Tabelle und das Hinzufügen ihrer Felder. Felddefinitionen sind später veränderbar. Das Ergebnis wird unmittelbar in einer Vorschau sichtbar.

Jede Tabelle erhält eine interne unveränderliche ID. Jeder Datensatz und jedes Feld erhält ebenfalls eine stabile interne ID. Zusätzlich definiert die Person ein gut lesbares Bezeichnungsfeld für Suche und Verknüpfungen. Gleiche Bezeichnungen dürfen vorkommen, müssen in Auswahlfenstern aber anhand weiterer Angaben unterscheidbar sein.

### 5.1 Feldtypen

| Feldtyp | Verhalten und Einstellungen |
|---|---|
| Kurzer Text | Bezeichnung, optionale maximale Länge, optional eindeutiger Wert |
| Langer Text / Notiz | Mehrzeilige Eingabe ohne aufwendigen Rich-Text-Editor |
| Zahl | Ganzzahl oder Dezimalzahl, Nachkommastellen, optional Minimum und Maximum |
| Geldbetrag / Prozent | Fachlich eindeutige Speicherung und passende Anzeige; Prozentformat ohne uneindeutige Faktoren |
| Datum / Datum mit Uhrzeit | Plausibilitätsprüfung und komfortable Eingabe, optionaler Standardwert |
| Ja/Nein | Klare Auswahl; optionale Angabe, ob leer erlaubt ist |
| Einfachauswahl | Genau eine Option aus einer bearbeitbaren Liste |
| Mehrfachauswahl | Mehrere Optionen mit übersichtlicher Auswahl und farbiger Kennzeichnung |
| Status | Geordnete, frei benennbare Statuswerte mit Farben und Kennzeichnung abgeschlossener Zustände |
| Verknüpfung | Ein oder mehrere Datensätze einer anderen Tabelle auswählen |
| Berechnetes Feld | Ergebnis einer gespeicherten Berechnungsregel, gegen manuelle Änderung geschützt |
| Lokaler Pfad | Lokalen Datei- oder Ordnerpfad hinterlegen; nur nach bewusster Aktion öffnen |

Automatische Systemangaben sind Datensatz-ID, erstellt am und geändert am. Die lokale Windows-Kennung kann als Änderungsmerkmal protokolliert werden, gilt aber nicht als sichere Identitätsfeststellung.

### 5.2 Gemeinsame Feldeinstellungen

Feldname, Hilfetext, Pflichtangabe, sinnvoller Standardwert, Reihenfolge, Darstellung und Sichtbarkeit sind konfigurierbar. Auswahloptionen lassen sich umbenennen, einfärben, sortieren und für neue Eingaben deaktivieren, ohne vorhandene Zuordnungen zu zerstören.

Nach dem Ergänzen eines Feldes erscheint es ohne Programmierung in der Erfassungsmaske und in passenden Such-, Filter-, Ansichts- und Auswertungsdialogen. Beim Ergänzen eines Pflichtfeldes in einer gefüllten Tabelle muss die Behandlung bestehender leerer Werte ausdrücklich festgelegt werden.

### 5.3 Sichere Strukturänderungen

Umbenennen oder Verschieben verändert keine internen IDs. Vor dem Wechsel eines Feldtyps prüft die Anwendung alle betroffenen Werte und zeigt eine Vorschau mit Konflikten. Keine stille Umwandlung ungültiger Werte in leere Felder oder Null.

Vor dem Entfernen eines Feldes oder einer Tabelle werden Beziehungen, Formeln, Ansichten und Dashboards auf Abhängigkeiten geprüft. Abhängigkeiten müssen kontrolliert aufgelöst werden; unbemerkte Folgeschäden sind unzulässig. Vor strukturellen Eingriffen wird eine Sicherung angeboten beziehungsweise gemäß Sicherungskonzept erstellt. Entfernen und Ausblenden sind getrennte Aktionen.

## 6. Erfassen, suchen und pflegen

Jede Tabelle erlaubt die Bearbeitung direkt in einer Tabellenansicht und über eine dynamisch aus ihrer Konfiguration erzeugte Detailmaske.

Die Detailmaske zeigt alle freigegebenen Felder, Notizen und verknüpften Einträge. Sie unterstützt Speichern, Speichern und neu, Abbrechen, Duplizieren und Archivieren. Verknüpfte Einträge sind von dort aus erreichbar.

Direkte Zelländerungen, Mehrzellen-Einfügen und Maskeneingaben verwenden dieselben Validierungsregeln. Ungültige Sammeländerungen werden vor Übernahme abgefangen oder vollständig kontrolliert zurückgenommen. Das Einfügen darf keine berechneten Felder oder internen IDs überschreiben.

Weitere Funktionen:

- Schnellsuche über sichtbare beziehungsweise ausdrücklich ausgewählte Felder; keine versehentliche Änderung der gespeicherten Ansicht.
- Fachlich passende Filter: Text enthält, Zahl zwischen, Datum im Zeitraum, Auswahl enthält, leer/nicht leer sowie verknüpfter Datensatz.
- Kombinierbare UND-/ODER-Filter mit sichtbarer Zusammenfassung und „Alle Filter zurücksetzen“.
- Mehrstufige Sortierung und Gruppierung nach geeigneten Feldern.
- Mehrere Einträge auswählen und Felder gesammelt ändern; Vorschau vor Übernahme.
- Dublettenprüfung anhand konfigurierbarer Felder; Warnung mit Möglichkeit, den bestehenden Eintrag zu öffnen. Eindeutige Felder verhindern echte Verletzungen der Eindeutigkeitsregel.
- Archiv mit Suche und Wiederherstellung; archivierte Einträge werden standardmäßig aus aktiven Ansichten ausgeschlossen.
- Endgültiges Löschen nur als gesonderte bewusste Aktion mit Referenzprüfung und klarer Bestätigung.

## 7. Beziehungen zwischen Tabellen

Unterstütze 1:n- und n:m-Beziehungen innerhalb derselben Arbeitsmappe. Beim Anlegen eines Verknüpfungsfeldes wählt die Person Zieltabelle, angezeigte Bezeichnung und Einfach- oder Mehrfachzuordnung.

Verknüpfungen werden über IDs und separate Zuordnungen gespeichert, nicht über zusammengesetzte Namen oder durch Kommas getrennte Mehrfachwerte als Datenbasis. Auf der Gegenseite sind zugehörige Einträge ohne doppelte Erfassung sichtbar. Änderungen werden referenzsicher übernommen.

Die Auswahl bietet Suche, zusätzliche Identifikationsmerkmale und eine Schnellanlage eines neuen Zieldatensatzes. Nach der Schnellanlage kehrt die Person ohne Verlust ihrer Eingaben zurück.

Beispiel: Projekte und Aufgaben. Eine Aufgabe gehört zu einem Projekt. Beim Projekt werden seine Aufgaben sowie deren Anzahl, offene Anzahl und Kostensumme angezeigt. Ein weiteres Beispiel sind Veranstaltungen, denen mehrere Einrichtungen zugeordnet werden können.

Archivierte verknüpfte Datensätze bleiben erkennbar. Das Löschen eines referenzierten Datensatzes darf keine verwaisten Verknüpfungen erzeugen. Automatisches kaskadierendes Löschen wird nicht als Standard verwendet.

## 8. Mehrere Ansichten derselben Daten

Eine Tabelle kann beliebig benannte gespeicherte Ansichten innerhalb der praktisch unterstützten Kapazität besitzen. Jede Ansicht speichert Darstellungstyp, Feldzuordnung, Filter, Sortierung, Gruppierung und Darstellungsoptionen. Ansichten sind keine Datenkopien.

Änderungen in einer Ansicht aktualisieren die zugrunde liegenden Datensätze und anschließend alle betroffenen Ansichten und Auswertungen. Einträge behalten ihre Identität beim Sortieren, Filtern und Wechseln der Darstellung.

### 8.1 Tabelle

Konfigurierbare Spaltenreihenfolge, Breite und Sichtbarkeit; farbige Auswahl-/Statuskennzeichnung, fixierte Orientierungsspalten, lesbare Datums-/Zahlenformate und optionale Zusammenfassungen. Lange Inhalte dürfen die Übersicht nicht zerstören; die Detailmaske zeigt den vollständigen Text.

### 8.2 Kalender

Monatsansicht und ergänzende Listenansicht für Termine. Die Person wählt Datumsfeld, optionales Enddatum und Beschriftung. Termine lassen sich öffnen, bearbeiten und für einen ausgewählten Tag neu anlegen. Einträge ohne Datum sind separat erreichbar. Mehrere Termine am Tag werden ohne verdeckte Einträge dargestellt oder über einen klaren Mehr-Einträge-Zugang geöffnet.

### 8.3 Board

Spalten nach Status oder Einfachauswahl, Karten mit konfigurierbaren Kurzangaben und sichtbaren Gruppenzahlen. Karten lassen sich öffnen, neu anlegen und einer anderen Gruppe zuordnen. Einträge ohne Gruppenzuordnung erhalten einen eigenen Bereich. Bei aktiven Filtern bleibt die Filterwirkung erkennbar.

### 8.4 Gantt / Projektleiste

Konfigurierbare Zuordnung von Bezeichnung, Beginn, Ende und optional Status/Fortschritt. Horizontale Zeitachse mit Tages-, Wochen- und Monatsdarstellung, sichtbarer Heute-Markierung und sinnvoller Zeitraumsauswahl. Balken lassen sich öffnen und ihre Zeitangaben komfortabel bearbeiten. Beginn nach Ende ist ungültig. Einträge ohne vollständige Zeitangaben werden separat angezeigt. Keine Abhängigkeits- oder kritischer-Pfad-Engine erforderlich.

### 8.5 Interaktion und technische Grenzen

Bearbeitung per Klick und Dialog ist in allen Ansichten verbindlich. Drag-and-drop für Statuswechsel oder Terminverschiebung sowie das Ziehen von Gantt-Balken wird nur umgesetzt, wenn ein früher Prototyp stabile Funktion in der Zielumgebung nachweist. Ein vollständiger gleichwertiger Bedienweg ohne Drag-and-drop bleibt vorhanden. Fehlende Drag-and-drop-Funktion darf nicht durch eine funktionslose Attrappe kaschiert werden.

## 9. Berechnungsbaukasten

Berechnungen werden mit Feld- und Funktionsauswahl aufgebaut; eigene Formelsyntax ist für Standardfälle nicht erforderlich. Eine Ergebnisvorschau und verständliche Fehlermeldungen sind Pflicht.

Unterstützte Kernfälle:

- Grundrechenarten zwischen numerischen Feldern und Konstanten, Rundung und Prozentberechnungen.
- Bedingung „Wenn … dann … sonst …“, Vergleich sowie UND/ODER.
- Datumsdifferenzen, Tage bis zur Frist und Überfälligkeitskennzeichnung.
- Verknüpfen von Texten.
- Anzahl verknüpfter Einträge, Anzahl mit Filter sowie Summe, Minimum, Maximum und Durchschnitt geeigneter verknüpfter Werte.

Formeln referenzieren Feld-IDs statt sichtbarer Namen. Abhängigkeiten werden geordnet neu berechnet; direkte und indirekte Zirkelbezüge werden beim Einrichten verhindert. Division durch Null und ungültige Datumsoperationen erzeugen erkennbare Fehlerzustände. Leere Werte werden nicht still als Null interpretiert.

Mehrfachverknüpfungen dürfen keine unbemerkten Doppelzählungen erzeugen. Tabellenübergreifende Summen müssen ihre Zähleinheit eindeutig ausweisen. Jede Berechnung speichert eine lesbare Beschreibung ihrer Logik.

## 10. Auswertungsassistent und Dashboards

Die Person erstellt eine Auswertung durch Auswahl von Datenquelle, Kennzahl, Aggregation, Gruppierung, optionalem Zeitraum und Filtern. Technische Pivot-Konfiguration oder Formeleingabe ist für die Standardfälle nicht erforderlich.

Unterstütze Anzahl eindeutiger Datensätze, Anzahl ausgefüllter/leerer Werte, Summe, Durchschnitt, Minimum, Maximum sowie Anteile mit ausdrücklich definierter Grundgesamtheit. Datumswerte lassen sich nach Monat, Quartal und Jahr gruppieren.

Dashboards können mehrere konfigurierbare Elemente enthalten: Kennzahlkachel, Balken-/Säulendiagramm, Zeitverlauf, Anteilsdarstellung und Ergebnistabelle. Ein geführtes Raster mit wählbarer Anordnung genügt; ein frei programmierbarer Layouteditor ist nicht erforderlich.

Jedes Element besitzt Titel, Datenquelle, Berechnungsregel und Filter. Ein gemeinsamer Dashboardfilter wirkt nur auf kompatible Elemente; abweichende oder unbeeinflusste Elemente werden gekennzeichnet. Kennzahlen oder Gruppen lassen sich zur zugrunde liegenden Datensatzliste öffnen.

Auswertungen aktualisieren sich nach Änderungen; bei aufwendigen Berechnungen zeigt die Anwendung den Aktualisierungsstatus und bietet eine eindeutige Aktualisierungsschaltfläche. Veraltete Ergebnisse dürfen nicht wie aktuelle erscheinen.

Fehlende Werte, ausgeschlossene Archiveinträge und überlappende Mehrfachkategorien werden transparent behandelt. Bei Mehrfachauswahl darf die Summe der Kategorien die eindeutige Gesamtzahl überschreiten; dieser Umstand wird erklärt. Keine irreführenden Summen von Prozenten oder doppelten Ressourcenwerten.

## 11. Import, Export und Vorlagen

### Import

Assistent für lokale XLSX- und CSV-Dateien mit Auswahl der Quelltabelle, Vorschau, Überschriftenzuordnung und Feldtypvorschlägen. Neue Tabelle erzeugen oder Daten an eine bestehende anhängen. Updates bestehender Datensätze nur bei ausdrücklich gewähltem und geprüftem eindeutigen Schlüssel.

Datums-/Dezimalformate, Trennzeichen und leere Werte berücksichtigen. Fehlerhafte Zeilen werden mit Begründung gemeldet. Keine stille Teilübernahme: gültige Zeilen dürfen erst nach Vorschau und ausdrücklicher Entscheidung übernommen werden; der Fehlerbericht bleibt verfügbar. Quelldateien werden nicht verändert und keine Makros aus Quellen ausgeführt.

### Export

Aktuelle gefilterte Ansicht als saubere XLSX- oder CSV-Datei; ausgewählte Auswertung beziehungsweise Dashboard als XLSX oder PDF mit geeigneter Druckgestaltung. Exporte enthalten Titel, Erstellungszeitpunkt und bei Bedarf Filterbeschreibung. Archive oder ausgeblendete interne Felder dürfen nicht versehentlich mitexportiert werden.

Exporte ohne Makros enthalten keine Verbindung zurück zur Arbeitsdatei. CSV-Export schützt standardmäßig gegen die Interpretation textlicher Inhalte als Formeln und erläutert gegebenenfalls notwendige Darstellungsänderungen.

### Vorlagen

Liefere eine leere Masterdatei und innerhalb der Anwendung auswählbare, konfigurationsbasierte Starter für Aufgaben/Projekte, Kontakte, Ressourcen und Dokumentation. Vorlagen verwenden dieselbe Engine wie selbst angelegte Tabellen; keine parallelen Sonderimplementierungen.

Eigene Strukturen können innerhalb einer Datei dupliziert werden: nur Struktur oder Struktur mit Daten. Bei Datenkopien sind neue IDs und korrekt umgesetzte interne Beziehungen erforderlich. Eine vollständige Fachdatei wird einfach im Dateisystem kopiert. „Neue leere Kopie erstellen“ darf die aktuell geöffnete Originaldatei nicht leeren.

## 12. Datenmodell und technische Architektur

Entwickle eine konfigurationsgesteuerte Engine mit getrennter Speicherung von Tabellenbeschreibung, Feldbeschreibung, Auswahloptionen, Datensätzen, Zuordnungen, Ansichten, Formeln und Dashboarddefinitionen.

Bevorzugt werden je Fachtabelle native Excel-ListObjects auf internen Datenblättern sowie separate Konfigurations- und Zuordnungstabellen. Technische Namen bleiben stabil; sichtbare Beschriftungen sind frei änderbar. Begründe Abweichungen und prüfe Typstabilität, Geschwindigkeit und verständlichen Datenzugang. Versteckte Blätter sind kein Sicherheitsmechanismus.

Trenne mindestens Oberfläche, Datenzugriff, Konfiguration, Validierung, Beziehungen, Suche/Filter, Berechnung, Ansichten, Dashboard, Import/Export, Sicherung und Fehlerbehandlung in wartbare Module und Klassen. Identifiziere Datensätze niemals dauerhaft anhand ihrer sichtbaren Zeilennummer.

Alle Schreibwege verwenden gemeinsame Prüf- und Speicherfunktionen. Neue Werte und Beziehungsänderungen werden zunächst geprüft und anschließend kontrolliert übernommen. Da Excel keine Transaktionsdatenbank ist, ist ein eigenes Rücknahme-/Wiederherstellungskonzept für mehrteilige Änderungen erforderlich.

Keine erforderlichen externen ActiveX-Komponenten oder zusätzlichen COM-Bibliotheken. Dynamische Formulare arbeiten mit vorhandenen Standardsteuerelementen. Keine freien VBA-Ausdrücke aus Nutzereingaben ausführen; der Berechnungsbaukasten verarbeitet ausschließlich erlaubte Operationen.

Verarbeite größere Datenmengen in Arrays und verwende gezielte Aktualisierung statt vollständigem Neuaufbau bei jeder Eingabe. Definiere Formatversion, Programmversion und eine migrationsfähige Konfiguration. Künftige Updates bestehender Fachdateien dürfen keine manuelle Neuerfassung voraussetzen; ein automatisches Updatewerkzeug ist kein Pflichtumfang dieser ersten Version.

## 13. Sicherung, Verlauf und Betrieb

- Sichtbarer Speicherstatus und klare Meldungen bei Schreibschutz oder Speicherfehlern. „Gespeichert“ erst nach erfolgreichem Speichern anzeigen.
- Beim Öffnen Erkennung von Nur-Lesen; Bearbeitungsaktionen sperren und verständlich erläutern. Gewöhnliche Excel-Dateisperren unterstützen sequenziellen Zugriff, ersetzen aber keine Mehrbenutzerdatenbank.
- Manuelle Sicherung per Schaltfläche und automatische Sicherung vor riskanten Strukturänderungen. Nach erstmaligem Speichern ist eine konfigurierbare automatische Sicherung bei regulären Speichervorgängen vorgesehen.
- Sicherungen standardmäßig in einem lokalen oder netzlaufwerkbasierten Unterordner neben der Fachdatei, mit Zeitstempel und Versionsangabe. Standardmäßig die letzten zehn erfolgreichen Sicherungen behalten. Fehlgeschlagene Sicherungen dürfen nicht als erfolgreich erscheinen.
- Wiederherstellung aus einer sichtbaren Sicherungsliste. Aktuelle Daten vor Überschreiben sichern; Wiederherstellung kontrolliert durchführen und validieren. Eine geöffnete Datei nicht unkontrolliert durch sich selbst ersetzen.
- Änderungsverlauf für Anlage, Bearbeitung, Archivierung, Wiederherstellung und Strukturänderungen mit Zeitpunkt, Objekt-ID und Änderungsart. Keine unnötigen vollständigen Kopien sensibler Feldinhalte im Protokoll.
- „Letzte Änderung rückgängig“ mindestens für die letzte zusammenhängende Datenänderung der laufenden Sitzung. VBA leert häufig Excels normalen Rückgängig-Verlauf; deshalb eine eigene, getestete Lösung vorsehen und deren Grenzen anzeigen.
- Arbeitsblattschutz dient gegen Fehlbedienung. Zugriffsschutz und gegebenenfalls Dateiverschlüsselung beruhen auf den vorhandenen Office-/Windows-Möglichkeiten und organisatorischen Regeln. Keine selbst erfundene Verschlüsselung.
- Bei deaktivierten Makros zeigt ein statisches Startblatt verständlich, warum die Anwendung nicht bedienbar ist. Keine angeblich laufenden Funktionen ohne Makros.

## 14. Datenmengen und Leistungsziele

Prüfe als Zielgröße zehn Fachtabellen, bis zu 50 Felder pro Tabelle und insgesamt 20.000 Datensätze einschließlich einer Tabelle mit 10.000 Datensätzen. Zusätzlich ein gesonderter Test mit mindestens 40.000 Beziehungszuordnungen. Dies ist ein Prüfprofil, kein Versprechen unbegrenzter Skalierung.

Zielwerte auf einem dokumentierten Windows-/Excel-Testrechner: Einzelnen Datensatz öffnen oder speichern innerhalb von zwei Sekunden; einfache Suche/Filter in 10.000 Einträgen innerhalb von drei Sekunden; ein Dashboard mit sechs Standardelementen innerhalb von zehn Sekunden aktualisieren. Kalender, Board und Gantt dürfen zeitlich beziehungsweise mengenmäßig begrenzte Ausschnitte anzeigen, müssen Einschränkungen aber offen darstellen und alle Daten erreichbar halten.

Messwerte mit Hardware, Office-Bitness, Testdatenumfang und Messmethode dokumentieren. Bei Überschreitungen zuerst optimieren und anschließend verbleibende Grenzen konkret benennen. Keine erfundenen Leistungsnachweise.

## 15. Entwicklung und verbindliche Qualitätsprüfungen

### Etappen

1. Architektur, Datenmodell, modulare Struktur und Bedienabläufe festlegen. Technische Risiken für dynamische Masken, direkte Tabellenbearbeitung, Mehrfachauswahl und Ansichten früh anhand echter Excel-Prototypen prüfen.
2. Durchgängigen Kernprozess implementieren: leere Kopie → Tabelle/Felder anlegen → Datensatz erfassen → speichern → neu öffnen → suchen → auswerten.
3. Beziehungen, sichere Strukturänderungen und Berechnungen ergänzen.
4. Kalender, Board und Gantt mit gemeinsamer Datenbasis fertigstellen.
5. Dashboards, Import/Export, Vorlagen und Sicherung vervollständigen.
6. Integration, Fehlertests, Leistung und visuelle Bedienung in echter Excel-Laufzeit prüfen; Fehler beheben; vollständiges Paket übergeben.

Routineentscheidungen eigenständig anhand dieses Auftrags treffen. Nur tatsächliche fachliche Widersprüche oder nicht erfüllbare Rahmenbedingungen zur Entscheidung vorlegen. Keine Funktionen stillschweigend aus dem Auftrag entfernen.

### Abnahmetests

| ID | Test | Erwartetes Ergebnis |
|---|---|---|
| T01 | Masterdatei kopieren, umbenennen und unabhängig öffnen | Alle Funktionen ohne Zugriff auf die Originaldatei verfügbar |
| T02 | Neue Tabelle mit allen Kernfeldtypen anlegen | Passende Eingabeelemente, Filter und Auswertungsangebote ohne Codeänderung |
| T03 | Zahl außerhalb des Bereichs, falsches Datum, fehlendes Pflichtfeld eingeben | Verständliche Meldung, keine ungültige Speicherung |
| T04 | Mehrere Zeilen mit gültigen und ungültigen Werten einfügen | Kontrollierte Prüfung, keine Teilkorruption oder überschriebenen Systemfelder |
| T05 | Feld umbenennen, umordnen und neues Pflichtfeld ergänzen | Ansichten, Beziehungen und Berechnungen bleiben konsistent; Altwerte werden behandelt |
| T06 | Inkompatiblen Feldtypwechsel oder Löschen eines abhängigen Feldes versuchen | Konfliktvorschau und sichere Ablehnung beziehungsweise kontrollierte Auflösung |
| T07 | Einfach-/Mehrfachverknüpfung, gleichnamige Ziele und Schnellanlage nutzen | Richtige IDs, korrekte Rückverweise, kein Verlust des begonnenen Eintrags |
| T08 | Zwei Tabellenansichten mit verschiedenen Filtern speichern und neu öffnen | Einstellungen bleiben getrennt erhalten; Daten werden nicht dupliziert |
| T09 | Datum im Kalender, Status im Board und Zeitraum im Gantt bearbeiten | Derselbe Datensatz wird in Tabelle, Ansichten und Auswertungen konsistent aktualisiert |
| T10 | Undatierte Einträge, viele Termine am Tag, leere Boardgruppe und ungültigen Gantt-Zeitraum prüfen | Erreichbare Einträge, lesbare Darstellung, korrekte Validierung |
| T11 | Formel, Rollup, leere Werte, Division durch Null und Zirkelbezug prüfen | Nachvollziehbare Ergebnisse oder klarer Fehler; kein Rechenkreislauf |
| T12 | n:m-Auswertung mit überlappenden Kategorien prüfen | Eindeutige Gesamtzahl und erklärbare Gruppenzahlen ohne stille Doppelzählung |
| T13 | Dashboard mit sechs Elementen, Filtern und Detailaufruf nutzen | Korrekte Kennzahlen, erkennbare Filterwirkung, passende Quelldatensätze |
| T14 | XLSX/CSV importieren, fehlerhafte Zeilen prüfen und Import abbrechen | Vorschau und Fehlerbericht; Abbruch verändert keine Bestandsdaten |
| T15 | Gefilterte Ansicht und Dashboard exportieren | Lesbare Dateien mit richtigem Umfang, ohne unbeabsichtigte interne Daten |
| T16 | Archivieren, wiederherstellen und letzte Datenänderung zurücknehmen | Beziehungen und Datenzustand bleiben korrekt |
| T17 | Speichern, Excel vollständig schließen und Datei erneut öffnen | Daten, Konfiguration, Ansichten und Dashboarddefinitionen vollständig erhalten |
| T18 | Nur-Lesen, fehlende Schreibrechte und fehlgeschlagene Sicherung prüfen | Keine falsche Erfolgsmeldung; sichere Bedienführung |
| T19 | Sicherung erstellen und tatsächlich wiederherstellen | Prüffähig identischer fachlicher Zustand der gesicherten Version |
| T20 | Offline-Test bei getrennter Netzwerkverbindung und technische Code-/Verbindungsprüfung | Gesamter Kernumfang funktioniert; keine von der Anwendung ausgelösten Netzwerkzugriffe oder externen Verbindungen |
| T21 | Definiertes Lastprofil und Anzeigegrößen prüfen | Dokumentierte Messwerte; bedienbare und nicht abgeschnittene Oberfläche |
| T22 | Gezielten Laufzeitfehler und Wechsel zu einer anderen Arbeitsmappe prüfen | Fehler sauber behandelt; Excel-Einstellungen werden wiederhergestellt |
| T23 | Struktur oder Struktur mit Daten duplizieren | Unabhängige neue IDs; Beziehungen korrekt umgesetzt; Original unverändert |
| T24 | Makros deaktiviert öffnen | Statischer verständlicher Hinweis statt vorgetäuschter Funktionsfähigkeit |

Offline bedeutet keine Netzwerkzugriffe durch die entwickelte Anwendung. Vorhandene Office-Lizenzierung oder andere Hintergrunddienste des Betriebssystems sind davon getrennt zu beurteilen und dürfen nicht als Anwendungsfunktion erforderlich sein.

Für jeden Test sind Status **BESTANDEN**, **NICHT BESTANDEN** oder **NICHT AUSGEFÜHRT**, Prüfumgebung und nachvollziehbare Evidenz festzuhalten. VBA-Projekt in echter Excel-Laufzeit kompilieren; Hauptabläufe dort ausführen. Echte Screenshots jeder Hauptansicht und wichtiger Dialoge erstellen und auf Layoutmängel prüfen.

Zusätzlich einen kurzen Bedienversuch mit einer Excel-unerfahrenen Person vorsehen: Tabelle aus einer Beschreibung bauen, fünf Einträge erfassen, Ansicht speichern und eine Auswertung erstellen. Ist keine Testperson verfügbar, als nicht ausgeführt kennzeichnen und einen dokumentierten Bedienablauf durchführen; das ersetzt keinen Nutzertest.

## 16. Liefergegenstände und Abschluss

1. **Excel_Datenbankbaukasten_Master.xlsm:** leere, kopierbare und vollständig konfigurierte Masteranwendung ohne produktive Beispieldaten.
2. **Excel_Datenbankbaukasten_Demo.xlsm:** separate Beispieldatei mit eindeutig fiktiven Projekten, Aufgaben, Einrichtungen und Ressourcen; demonstriert Beziehungen, alle vier Ansichten, Berechnungen und Dashboards.
3. Vollständiger VBA-Quellcode einschließlich Klassen, Formulardateien und benötigter Formulardaten; nachvollziehbare Zuordnung zur ausgelieferten Version.
4. Kurzanleitung „In zehn Schritten zur eigenen Datei“ sowie ausführliche integrierte Offlinehilfe.
5. Technische Dokumentation zu Datenmodell, Konfiguration, Modulen, unterstützten Grenzen, Sicherung/Wiederherstellung und Vorbereitung späterer Versionsmigrationen.
6. Testmatrix mit tatsächlichen Ergebnissen, Leistungsbericht, echten UI-Screenshots und bekannten Einschränkungen.
7. Versionshinweise und ein geordnetes ZIP-Paket mit allen Liefergegenständen. Falls ein Build erforderlich ist: reproduzierbare Build-Anleitung mit eindeutig benannten Voraussetzungen.

Steht bei der Umsetzung keine echte Excel-Laufzeit zur Verfügung, dürfen Generierung einer Datei, statische Codeprüfung oder Tests in einer anderen Tabellenkalkulation nicht als erfolgreiche Excel-Abnahme ausgegeben werden. Liefere dann den möglichst vollständigen reproduzierbaren Entwicklungsstand, kennzeichne nicht ausgeführte Prüfungen und benenne die verbleibenden Schritte. Eine XLSX-Datei oder lose VBA-Module sind kein funktionsgleicher Ersatz für die beauftragte Masteranwendung.

**Erfolgskriterium:** Eine Person mit normalen Office-Grundkenntnissen kann aus einer kopierten Masterdatei eine eigene kleine Datenbankanwendung zusammenstellen, Daten komfortabel pflegen, dieselben Einträge in mehreren Ansichten bearbeiten und aussagekräftige Auswertungen erstellen – vollständig offline und ohne eigene Programmierung.
