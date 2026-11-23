# Technische Dokumentation (aus der Anwendung exportiert)

## MidPoint: Technische Dokumentation

Programmversion 1.8.2, Stand November 2026 · Formatversion 3 · Entwickler: Ahmed Imam mit Claude Code.

> Sicherheitsprüfung vor jedem Release (Version 1.8.2, Stand November 2026): bestanden. Die Anwendung arbeitet vollständig lokal: kein eigener Netz- oder Internetzugriff (Links, E-Mail-Adressen, Dateien und Ordner öffnet nur ein ausdrücklicher Klick mit dem zugeordneten Windows-Programm, bei Adressen und Dateien nach Rückfrage), keine fremden Programme oder Befehle, nur Windows-Systembibliotheken (user32, gdi32, gdiplus, oleaut32), keine Änderung von Office-, Windows- oder Makroeinstellungen, kein Zugriff auf den eigenen VBA-Code zur Laufzeit. Geschrieben wird nur in den Sicherungsordner neben der Datei, in den TEMP-Ordner (Bilder) und an selbst gewählte Ziele; Quelldateien für Import und „Liste übernehmen“ werden nur gelesen. Hinweis: Protokoll und Einträge speichern die Windows-Kennung; Fachdaten können personenbezogen sein. Der vollständige Bericht liegt im Lieferpaket im Ordner „IT und Datenschutz“.

### Sicherheitsprüfung vor jedem Release

Stand: November 2026 · geprüft: 69 Quelldateien (36724 Zeilen) und MidPoint_Master.xlsm

**Ergebnis: bestanden.** Die Anwendung arbeitet vollständig lokal, ohne Netz- oder Internetzugriff, ohne fremde Programme und ohne Änderung von Sicherheitseinstellungen.

| Nr. | Prüfpunkt | Ergebnis | Befund |
|---|---|---|---|
| A1 | Netz und Internet | BESTANDEN | Kein eigener Netz- oder Internetzugriff im Code (keine HTTP-, Download-, Mail- oder Datenbankverbindung). Nur auf ausdrücklichen Klick übergibt die Anwendung eine eingetragene Adresse, Datei oder einen Ordner an das zugeordnete Windows-Programm (Browser, E-Mail-Programm, Explorer), bei Adressen und Dateien erst nach Rückfrage (4 Stellen). 8 Adresse(n) nur als Text, z. B. Lizenzhinweis; sie werden nie aufgerufen. |
| A2 | Programme und Befehle | BESTANDEN | Die Anwendung startet keine Programme, Befehlszeilen oder Skripte und sendet keine Tastendrücke. |
| A3 | Windows-Schnittstellen | BESTANDEN | Nur Windows-Systembibliotheken: gdi32, gdiplus, oleaut32, user32. Keine fremden Bibliotheken oder Add-ins. |
| A4 | COM-Objekte | BESTANDEN | Nur bekannte Windows- und Office-Bausteine (1), jeder mit festem Zweck; kein Laden von Adressen. |
| A5 | VBA-Projekt | BESTANDEN | Die Anwendung verändert ihren Code nicht und braucht keinen „Zugriff auf das VBA-Projektobjektmodell“. Der Codekatalog liest eine beim Build gespeicherte Textkopie. |
| A6 | Sicherheitseinstellungen | BESTANDEN | Keine Office-, Windows- oder Registry-Einstellung wird geändert. Beim Lesen von Import- und Quelldateien und beim Prüfen von Sicherungen werden Makros der geöffneten Datei ausdrücklich gesperrt (AutomationSecurity = ForceDisable) und danach der vorherige Wert wiederhergestellt. |
| A7 | Dateien und Laufwerke | BESTANDEN | Geschrieben wird nur an festen, eigenen Orten oder an Zielen, die die Person selbst wählt. Gelöscht werden nur eigene Sicherungen (Namensmuster) und Zwischendateien; keine Platzhalter, keine Ordner. 42 Fundstellen. |
| A8 | Personenbezogene Daten | HINWEIS | Die Datei speichert Windows-Kennungen der Bearbeitenden und, je nach Einrichtung, personenbezogene Fachdaten. Für das Verzeichnis der Verarbeitungstätigkeiten melden; Zugriff über die Ordnerrechte regeln. |
| B1 | Externe Verknüpfungen | BESTANDEN | Keine Verknüpfungen zu anderen Arbeitsmappen oder OLE-Quellen. |
| B2 | Datenverbindungen und Links | BESTANDEN | Keine Datenverbindungen, Abfragen oder Links nach außen; 0 Links führen nur innerhalb der Datei. |
| B3 | Datenblätter | BESTANDEN | 16 Datentabellen sind sehr ausgeblendet und geschützt; bearbeitet wird nur über die Formulare. |
| B4 | Dateieigenschaften | BESTANDEN | Keine Namen in den Dateieigenschaften. |
| B5 | Ohne Makros | BESTANDEN | Ohne Makros erscheint das Blatt „Hinweis“ mit der Anleitung zum Freischalten; Daten bleiben verborgen und geschützt. |

### Zweck und Umfang

Offline nutzbare Excel-/VBA-Anwendung, mit der Personen ohne Programmierkenntnisse eigene strukturierte Datenbankanwendungen zusammenstellen: frei definierbare Tabellen und Felder, Verknüpfungen (1:n, n:m), mehrere gespeicherte Ansichten (Tabelle, Kalender, Board, Projektleiste), Berechnungen, Dashboards, Import/Export, Sicherung und Änderungsprotokoll. Fachlich neutral; keine Fachstruktur ist fest einprogrammiert.

### Voraussetzungen und Betrieb

- Microsoft Excel 365 Desktop unter Windows, 32 oder 64 Bit (Declare-Anweisungen mit PtrSafe/LongPtr). Dateiformat XLSM. Keine Add-ins, keine zusätzlichen Bibliotheken, keine ActiveX-Steuerelemente, kein Zugriff auf das VBA-Projektobjektmodell zur Laufzeit.
- Genutzte Systembestandteile: MSForms (UserForms), user32/gdi32 (Andocken der Navigation, DPI), gdiplus/oleaut32 (Zeichnen der Navigationsform), Office-CommandBars (temporäre Kontextmenüs, werden nicht dauerhaft gespeichert). Kein Scripting.Dictionary/FileSystemObject (eigene Klasse cMap, VBA-Dateibefehle).
- Netzwerk: Die Anwendung baut keine Verbindungen auf. Dateizugriffe nur auf die Arbeitsdatei, den Sicherungsordner, gewählte Import-/Exportdateien und bewusst geöffnete lokale Pfade.
- Makrofreigabe: Datei an einem vertrauenswürdigen Speicherort ablegen oder das VBA-Projekt durch die IT mit einem Codesignaturzertifikat signieren und das Zertifikat per Richtlinie als vertrauenswürdigen Herausgeber verteilen. Die Anwendung ändert keine Sicherheitseinstellungen. Ohne Makros zeigt das Blatt „Hinweis“ eine statische Erklärung.

### Architektur (Module)

- mStart, ThisWorkbook: Start/Schließen, Ereignisse, Lesemodus, Formatprüfung, Integritätsprüfung, Hinweisblatt vor dem Speichern.
- mApp: ruhige Ausführung (ScreenUpdating/Events/Berechnung werden erfasst und, auch im Fehlerpfad, wiederhergestellt), zentrale Fehlergrenze, Modus, Lese-/Schreibstatus.
- mCfg: Konfiguration (Tabellen, Felder, Optionen, Ansichten, Dashboards, Systemwerte) mit Puffer. mData: Datenzugriff, Zuordnungen, Transaktionen und Rückgängig. mValid: gemeinsame Prüfregeln aller Schreibwege. mFmt: deutsche Darstellung.
- mCalc: Berechnungsbaukasten (eigener Interpreter, Abhängigkeiten, Zirkelbezugsprüfung, gezielte Neuberechnung). mQuery: Suche, Filter, Sortierung, Gruppierung.
- mNav, frmNav, mRender: Navigation (als Kindfenster im Arbeitsmappenfenster verankert, kein Fensterfixieren), Seitenaufbau, Schaltflächen, Filterkacheln, Kontextmenüs. mViewTable, mViewCal, mViewBoard, mViewGantt: Darstellungen. mDash: Auswertungen und Dashboards. mPages: Start, Einrichten, Suche, Hilfe, Technik.
- Schneller Seitenwechsel: Jede Seite ist ein eigenes Blatt. mNav merkt sich je Blatt den Stand beim letzten Aufbau (Daten- und Strukturstand, Ansichtszustand, Fensterbreite, Datum, Lesemodus). Ist er unverändert, wird die Seite nur eingeblendet und das Statusband aktualisiert (mNav.TryReuse). Schaltflächenbreiten werden je Beschriftung einmal gemessen, gleichfarbige Zellen gemeinsam formatiert, Bezeichnungen verknüpfter Einträge zwischengespeichert.
- mStruct: Strukturänderungen, Abhängigkeitsprüfung, Typwechsel, Duplizieren, Vorlagen. mIO: Import/Export. mBackup: Speichern, Sicherung, Wiederherstellung, leere Kopie, Zurücksetzen. mLog: Protokoll. mCodeCat: Codekatalog und Codeexport. mCmd: alle Bedienbefehle mit Fehlergrenze. mBuild: Build-, Vorlagen- und Demodaten (nur Entwicklung).
- mListImport (seit 1.7.0): Liste übernehmen. Liest die Quelle schreibgeschützt (Werte, Formeln, verbundene Zellen, ausgeblendete Spalten, Zahlenformate, Excel-Tabellen), findet Tabellen, deutet Kopf-, Gruppen- und Summenzeilen, schlägt je Spalte einen Typ vor, erzeugt Rückfragen mit Vorschlag und baut daraus einen Importplan, den mIO.CheckImport und mIO.ExecuteImport prüfen und übernehmen. 2ask-Exporte werden über das Blatt „code_plan“ gedeutet.
- Formulare: frmRecord (dynamische Maske), frmPicker (Verknüpfungsauswahl), frmDate (Kalender), frmFilter, frmColumns, frmFormula, frmField, frmTableEditor, frmImport, frmListImport (Assistent „Liste übernehmen“), frmProps (allgemeiner Dialog), frmMsg (Meldungen), frmCode (Codeansicht), frmNav. Alle Steuerelemente entstehen zur Laufzeit; es gibt keine binären Formularressourcen.
- Klassen: cMap (Schlüssel/Wert), cField (Felddefinition), cTableData (Arrayabbild einer Tabelle), cViewState (Sitzungszustand einer Ansicht), cEvt (Ereignisweiterleitung dynamischer Steuerelemente), cLiSheet, cLiTab, cLiCol, cLiQ (gelesenes Blatt, gefundene Tabelle, geplantes Feld und Rückfrage der Listenübernahme).

### Datenmodell

- _Tabellen: ID (T0001), Name, Beschreibung, Symbol, Bezeichnungsfeld, Reihenfolge, Status, Dublettenfelder, Erstellt, Blatt.
- _Felder: ID (F0001), Tabelle, Name, Typ, Hilfetext, Pflicht, Standard, Reihenfolge, Sichtbar, Gruppe, Eindeutig, MaxLaenge, Dezimalstellen, Minimum, Maximum, LeerErlaubt, Zieltabelle, Mehrfach, Gegenfeld, Formel, Beschreibung, Ergebnistyp, NichtVor, Status, Erstellt.
- _Optionen: ID (O0001), Feld, Name, Farbe, Reihenfolge, Aktiv, Abgeschlossen. Einfach-/Mehrfachauswahl und Status speichern Options-IDs (Mehrfach: „O0001;O0004“); Umbenennen ändert keine Daten.
- Datenblätter D_T0001 …: je Fachtabelle ein ListObject mit Systemspalten _ID (R0000001, dateiweit eindeutig), _Erstellt, _Geaendert, _Benutzer, _Archiviert und je Feld eine Spalte mit der Feld-ID als Überschrift.
- _Zuordnungen: Feld, Von, Nach, Erstellt; Verknüpfungen ausschließlich über IDs. Rückverknüpfungsfelder (Typ BACKLINK) speichern nichts selbst, sondern lesen dieselben Zuordnungen in Gegenrichtung.
- _Ansichten: ID (V0001), Tabelle, Name, Typ, Reihenfolge, Spalten („F0001:220;F0003:90“), Filter, Sortierung („F0003:A“), Gruppierung, Optionen (z. B. datum=F0005;ende=F0006). _Dashboards/_Elemente: Dashboard- und Elementdefinitionen. _Protokoll: Änderungsprotokoll. _System: Versionen, Zähler, Einstellungen. _Listen: Hilfslisten für Zell-Auswahllisten. _Vorlagen: konfigurationsbasierte Starter. _Code: Codekatalog. _Importbericht: letzter Fehlerbericht.
- Filterformat: „Verknüpfung|FeldID|Operator|Wert1|Wert2“ je Bedingung, durch Semikolon getrennt; Werte maskiert, Zahlen/Daten sprachunabhängig (Punkt bzw. JJJJ-MM-TT).
- Formelformat: „FUNKTION|Argument|…“ mit Argumenten F:FeldID, N:Zahl, T:Text, D:Datum, B:Ja/Nein, O:OptionID, K:Schlüsselwort. Ausgewertet werden die Funktionen RECHNEN, RUNDEN, ANTEIL, WENN, PRUEFEN, DATDIFF, TAGEBIS, UEBERFAELLIG, DATPLUS, TEXT und ROLLUP. Excel-Formeln (seit 1.3.0): „XL|=RUNDEN({F0003}*{F0004};2)“ mit Feld-IDs in { }. mCalc übersetzt die deutsche Formel einmal über FormulaLocal in die interne Form, schreibt die Spaltenwerte aller Einträge in Hilfsspalten des ausgeblendeten Blattes _Rechnen, füllt die Formel daneben, berechnet nur diesen Bereich und liest das Ergebnis in einem Zug (mCalc.XlEvalRows). Gesperrt: Bezüge auf Blätter/Dateien/Zellen, WEBSERVICE, FILTERXML, RTD, CALL, REGISTER.ID, IMAGE, HYPERLINK, INDIRECT, OFFSET, INFO, CELL. Texte werden immer als Text übergeben (nie ausgeführt); Daten, die mit = + - @ beginnen, werden beim Schreiben in Blätter maskiert (CellText). Keine Ausführung von Nutzereingaben als VBA.

### Eigenschaftsarten (Notion-ähnlich)

Die Spalte „Typ“ in _Felder speichert die Art (z. B. KONTROLL, PERSON, URL, AUTOID). Jede Art beruht auf einem Grundtyp (mConst.BaseTypeOf), der Speicherung, Filter, Sortierung und Auswertung bestimmt; dadurch funktionieren alle Ansichten und Dashboards ohne Sonderlogik.

- KONTROLL: Ja/Nein ohne Leerwert; PERSON: Mehrfachauswahl (Personenliste als Optionen); DATEIEN: Pfadliste „Pfad1|Pfad2“; URL, EMAIL, TELEFON, ORT: Text mit Formatprüfung (mValid.KindFormatOk); ROLLUP: Berechnung mit Funktion ROLLUP.
- Automatisch gepflegt (nicht editierbar, mValid lehnt manuelle Werte ab): ERSTELLT_UM/_VON und GEAENDERT_UM/_VON aus den Systemspalten; AUTOID mit Präfix im Feld „Standard“ und Zähler „IDZaehler_<FeldID>“ in _System (Nummern werden nie wiederverwendet).
- BUTTON: Feld „Formel“ = „AKTION|F:FeldID|Wert|F:FeldID|Wert …“ (Werte maskiert; Schlüsselwörter heute, jetzt, ich, leer, umschalten). Ausführung über mForms.RunButton mit derselben Prüfung, Transaktion, Protokollierung und Rücknahme wie jede andere Änderung.
- Seit 1.2.0 kein getrennter Einrichtungsmodus (mApp.IsSetupMode = schreibbar; RequireSetup fragt nur ein gesetztes Schutzwort einmal je Sitzung ab). Spaltenkopf-Doppelklick öffnet ein temporäres Kontextmenü (sortieren, gruppieren, filtern, ausblenden, Eigenschaft bearbeiten/einfügen/entfernen); „+“ legt neue Spalten an. Ohne Fensterfixierung und ohne Seitenumbruchlinien.
- Eintragsmaske (frmRecord.ShowInForm): Formeln, Rollups und automatische Arten werden nicht angezeigt; Buttons erst bei gespeicherten Einträgen. Fehler an ausgeblendeten Eigenschaften erscheinen im Fuß der Maske.
- Automatisch speichern: _System „AutoSpeichern“ (Ja/Nein). mApp.MarkDirty/MarkCfgDirty planen per Application.OnTime (1 s) mBackup.AutoSaveTick; solange ein Dialog offen ist oder eine Transaktion läuft, wird neu geplant. Gespeichert wird über ControlledSave; danach wird die aktuelle Seite ohne Neuaufbau wieder gezeigt.
- Navigation: Titel/Untertitel aus _System (AppName, Untertitel, TitelFestgelegt); Navigations- und Akzentfarbe als Farbfelder auf „Einrichten“ (mCmd.Cmd_Swatch). Die Aktivmarkierung und der Schiebeschalter werden per GDI+ gezeichnet (mGdi.NavPicture, 6-fache Abtastung, Viertelkreis-Übergänge, Bild rechts breiter als sichtbar gegen Kantenartefakte). Höhe = sichtbare Höhe des Fensters (UsableHeight, ohne Excels untere Leiste); bei wenig Höhe wird stufenweise verdichtet, danach eigener Bildlauf mit Pfeilen. Das Ein-/Ausklappen des Menübands meldet Excel nicht; die Navigation prüft ihre Höhe daher bei Mausbewegung, Zellauswahl, Fensteraktivierung und Seitenaufbau (bewusst ohne Application.OnTime, damit Excel die Datei nach dem Schließen nie erneut öffnet).
- Seitenwechsel bauen die Navigation nicht neu auf, sondern setzen nur die Markierung um (frmNav.Update); neu aufgebaut wird nur bei geänderten Einträgen, Titel oder Farben, dabei ohne Zwischenzeichnen (WM_SETREDRAW). Gezeichnet werden nur die Streifen mit Rundungen mehrfach abgetastet. Symbol und Beschriftung stehen exakt auf der Mitte der Markierung (gemessene Zeilenhöhe). Helle Navigationsfarben schalten die Schrift auf dunkel (IsLightColor, CNavText/CNavStrong); helle Akzentfarben werden für Schrift abgedunkelt (CAccent) und für Flächen hell verwendet (CAccentFill, COnAccent).
- Dashboards: Stil und Palette je Dashboard in _System (DashStil_<ID>, DashPalette_<ID>), eigene Palette je Element in _Elemente.Optionen (pal=…), Höhe (hoch=1). „Anordnen“ zeichnet Rasterspalten (gesperrt) und Platzhalterkarten ohne Makro (nicht gesperrt, daher ziehbar). Raster je Element in _Elemente.Optionen (gx = Spalte 0 bis 3, gy = Zeile) und Groesse (1 bis 4 Spalten); mDash.LayoutGrid ordnet überschneidungsfrei an (gezogene Karten zuerst, verdrängte rücken nach rechts bzw. unten, leere Zeilen entfallen). Einrasten bei Zellauswahl oder „Einrasten“ (SnapArrangement, nur im Speicher), „Fertig“ speichert (ApplyArrangement). Vorschau beim Anlegen: Element wird gespeichert, markiert (gHighlight) gezeigt und erst nach „Übernehmen“ behalten; „Verwerfen“ löscht es wieder.

### Erfassungsdateien (mCapture)

- _Erfassungen: ID (C0001), Name, Tabelle, Felder (FeldIDs zum Ausfüllen), Modus (Liste/Formular), Titel, Einleitung, Zeilen, Absender, AbsenderFeld, Schutz, Erstellt, Version, Optionen. Optionen (Schlüssel=Wert): art=pruefen, anzeige=FeldIDs (grau), ansicht=ViewID (Einträge), teilen=FeldID (je Empfänger eine Datei), neuzeilen=n, eltern=ErfassungID (weiteres Blatt derselben Datei), abslabel, absbreite, farbe, anleitung. _Ruecklauf: DateiID, Datei, Erfassung, Absender, Zeilen, Importiert, Benutzer, Empfaenger. Versand je Datei in _System „Versand_<ID>“ (Zeitpunkt, Empfänger, DateiID) für die Rücklauf-Übersicht.
- Erzeugte Datei (.xlsx, ohne Makros): ein Blatt je Festlegung (Hauptblatt plus weitere Blätter), sehr verstecktes Blatt „Listen“ (Auswahlwerte, Namen L_<FeldID>, L_JA) und sehr verstecktes Blatt „_Info“ (Kennung, ErfassungID, DateiID, DatenbankID, Empfaenger, Aufbau=2; je Blatt eine Zeile „Blatt“ mit Name, ErfassungID, TabelleID, erster/letzter Datenzeile, „Spalten“ mit ID- und Stand-Spalte, sowie je Spalte „Spalte|Art|FeldID|OptionID|Spalte“; Art ANZ = grau, FELD = gelb, OPT = Ankreuzspalte einer Mehrfachauswahl, ID/STAND versteckt). Mehrfachauswahl bis 30 Optionen als Ankreuzspalten (senkrechter Titel, Datenüberprüfung „Ja“), darüber als Text „A; B“.
- Einlesen: Datei schreibgeschützt mit deaktivierten Makros öffnen, Kennung, Datenbank und doppeltes Einlesen prüfen, je Blatt die Zeilen lesen. Zeilen mit ID aktualisieren den Eintrag (nur geänderte Werte; Konflikt, wenn _Geaendert jünger als der Stand beim Versand), Zeilen ohne ID legen neue Einträge an (Aufteilungsfeld wird auf den Empfänger gesetzt). Prüfung über mValid.ParseCell/ParseInput und ValidateRecord (dieselben Regeln wie im Formular), Verknüpfungen über eindeutige Bezeichnungen, Übernahme in einer Transaktion, DateiID in _Ruecklauf vermerken. Dateien des Aufbaus 1 (Version 1.1) werden weiterhin gelesen.

### Speicherwege, Transaktionen, Rückgängig

Alle Schreibwege (Formular, Zelle, Einfügen, Sammeländerung, Board, Import, Rücknahme) nutzen mValid und mData. Eine Transaktion zeichnet Zeilen- und Zuordnungsoperationen mit Vorher-Stand auf; schlägt ein Teil fehl, werden alle Operationen in umgekehrter Reihenfolge zurückgenommen. Abhängige Berechnungen laufen innerhalb derselben Transaktion (betroffene Datensätze und über Rollups verknüpfte Datensätze; bei Massenänderungen spaltenweise Neuberechnung). Nach Abschluss wird die Transaktion als Rückgängig-Schritt der Sitzung abgelegt (max. 20).

### Oberfläche, Ereignisse, Fehler

- Ein Oberflächenblatt je Seitentyp (Start, Tabelle, Kalender, Board, Gantt, Dashboard, Einrichten, Hilfe, Technik, Suche); jeweils nur das aktive ist sichtbar. Seiten werden bei jedem Aufruf frisch aufgebaut; Sitzungsfilter liegen in cViewState.
- Arbeitsmappenereignisse (SheetChange, SelectionChange, BeforeDoubleClick, WindowResize, BeforeSave, BeforeClose) leiten an die Ansichtsmodule; Wiedereintritt wird verhindert.
- Alle Bedienbefehle laufen über mCmd mit Fehlergrenze: begonnene Transaktionen werden zurückgenommen, Excel-Zustände wiederhergestellt, eine verständliche Meldung erscheint; der VBA-Editor wird nicht geöffnet.
- Einstellungen anderer Arbeitsmappen werden nicht verändert: Gitternetz, Köpfe und Register werden nur für das Fenster dieser Datei ausgeblendet; Menüband und Bearbeitungsleiste bleiben unverändert.

### Sicherung, Wiederherstellung, Migration

- Sicherung per SaveCopyAs in „<Datei>_Sicherungen“, Prüfung auf Größe und ZIP-Signatur, Aufbewahrung der letzten N eigenen Sicherungen (Standard 10).
- Wiederherstellung: Sicherung schreibgeschützt mit deaktivierten Makros öffnen, Formatversion prüfen, Konfigurations- und Datenblätter übernehmen, Integrität prüfen, neu berechnen. Die geöffnete Datei wird nicht durch sich selbst ersetzt.
- Migration: _System enthält Formatversion und Programmversion. Beim Öffnen prüft mStart die Formatversion; künftige Versionen ergänzen fehlende Spalten/Blätter per Migrationsroutine, ohne Daten neu erfassen zu müssen. Programmupdate: Code der neuen Masterversion in eine Kopie der Fachdatei übernehmen (build.ps1 -Update „Pfad“ erzeugt „…_aktualisiert.xlsm“; beim ersten Öffnen ergänzt die Migration fehlende Spalten, z. B. von Format 2 auf 3).

### Leistung und Grenzen

- Prüfprofil: 10 Tabellen, bis 50 Felder je Tabelle, 20.000 Datensätze (eine Tabelle mit 10.000) und 40.000 Zuordnungen. Gemessene Werte stehen im mitgelieferten Prüfbericht.
- Tabellenansicht zeigt seitenweise 100 Einträge (einstellbar); Kalender bis 4 Termine je Tag plus „weitere“; Board bis 40 Karten je Spalte; Projektleiste bis 300 Zeilen je Zeitraum. Alle Daten bleiben über Filter, Seiten und Listen erreichbar.
- Mausrad scrollt in Formularen nicht innerhalb von Registerseiten (Grenze von MSForms); Bildlaufleiste und Tabulator/Enter funktionieren. Drag-and-drop nur beim Anordnen von Dashboards.
- Keine gleichzeitige Bearbeitung durch mehrere Personen; Excel-Dateisperre auf Netzlaufwerken ermöglicht nacheinander arbeiten (zweite Person: Lesemodus).

### Build und Reproduzierbarkeit

Quellen liegen als Textdateien vor (src\modules, src\classes, src\forms, src\ThisWorkbook.cls). Das PowerShell-Skript build\build.ps1 erzeugt mit einer lokalen Excel-Installation die Master- und Demodatei: neue Mappe, Module/Klassen importieren, Formulare anlegen und Code einfügen, mBuild.Build_Init (Blätter, Konfiguration, Vorlagen, Hilfe, Codekatalog), speichern. Voraussetzung nur auf dem Entwicklungsrechner: „Zugriff auf das VBA-Projektobjektmodell vertrauen“. Anwenderrechner benötigen diese Einstellung nicht.

- Prüfskripte: build\lint.ps1 (statische Prüfung: reservierte Namen, Doppeldeklarationen, Zeichensatz Windows-1252, verschmolzene Zeilen), build\selftest.ps1 (automatisierte Abnahmeprüfungen mit mSelfTest in einer Demokopie inkl. Sicherung/Wiederherstellung und Speichern/Wiederöffnen), build\perf.ps1 (Lasttestdatei bauen und Laufzeiten messen). mSelfTest läuft nur in Demokopien (Sperre über Systemwert „Demo“).

### Datenschutz

Alle Daten liegen ausschließlich in der Datei bzw. in bewusst erzeugten Exporten und Sicherungen. Das Protokoll speichert keine Feldinhalte. Die Anwendung bietet keine eigene Verschlüsselung; Vertraulichkeit über Ablageort/Berechtigungen und gegebenenfalls den Office-Dateikennwortschutz.

