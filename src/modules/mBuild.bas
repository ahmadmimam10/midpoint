Attribute VB_Name = "mBuild"
'==============================================================================
' mBuild - Nur fuer den reproduzierbaren Build (build\build.ps1): legt Blaetter,
' Konfigurationstabellen, Vorlagen, Hinweisblatt, Hilfe und Codekatalog an und
' erzeugt Demo- bzw. Belastungsdaten. Wird im Betrieb nicht aufgerufen.
'==============================================================================
Option Explicit

Private mRnd As Long
Private mTplRow As Long

'==============================================================================
' Grundgeruest
'==============================================================================
Public Sub Build_Init()
    Dim sh As Variant, ws As Worksheet, i As Long, hdr As Variant, lo_ As ListObject
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' erstes Blatt wird Hinweisblatt
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Name = SH_HINWEIS
    Do While ThisWorkbook.Worksheets.Count > 1
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(2).Delete
        Application.DisplayAlerts = True
    Loop
    WriteHinweis ws
    ' Oberflaechenblaetter
    For Each sh In Array(SH_START, SH_TABLE, SH_CAL, SH_BOARD, SH_GANTT, SH_DASH, SH_SETUP, SH_HELP, SH_TECH, SH_SEARCH, SH_CAPTURE)
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CStr(sh)
        ws.Cells.Font.Name = UI_FONT
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.Zoom = 100
    Next
    ' Konfiguration
    For Each sh In mCfg.CfgSheets()
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CStr(sh)
        hdr = mCfg.CfgHeaders(CStr(sh))
        For i = 0 To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next
        ws.Columns(1).NumberFormat = "@"
        If sh = SH_REL Then ws.Columns(2).NumberFormat = "@": ws.Columns(3).NumberFormat = "@"
        If sh = SH_LOG Then ws.Columns(1).NumberFormat = "dd.mm.yyyy hh:mm:ss": ws.Columns(5).NumberFormat = "@": ws.Columns(6).NumberFormat = "@"
        If sh = SH_SYS Then ws.Columns(2).NumberFormat = "@"
        Set lo_ = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(hdr) + 1)), , xlYes)
        lo_.Name = Mid$(CStr(sh), 2)
        lo_.TableStyle = "TableStyleLight1"
    Next
    For Each sh In Array(SH_LST, SH_TPL, SH_CODE, SH_IMP)
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CStr(sh)
    Next
    ThisWorkbook.Worksheets(SH_CODE).Range("A1:D1").Value = Array("Modul", "Art", "Nr", "Text")
    mCfg.InvalidateAll
    ' Systemwerte
    SysDefault "AppName", APP_TITLE
    SysDefault "Untertitel", "Offline-Datenbank in Excel"
    SysDefault "Formatversion", CStr(FORMAT_VERSION)
    SysDefault "Programmversion", PROG_VERSION
    SysDefault "AutoSpeichern", "Nein"
    SysDefault "Akzentfarbe", CStr(RGB(37, 99, 235))
    SysDefault "Navigationsfarbe", CStr(RGB(30, 41, 59))
    SysDefault "SicherungBeimSpeichern", "Ja"
    SysDefault "SicherungAbstandMinuten", "60"
    SysDefault "SicherungenBehalten", "10"
    SysDefault "LetzteBerechnung", DateInv(Date)
    WriteTemplates
    mStruct.RebuildLists
    ' Hilfe und Technik einmal schreiben
    mPages.WriteHelp ThisWorkbook.Worksheets(SH_HELP)
    mPages.WriteTech ThisWorkbook.Worksheets(SH_TECH)
    ' interne Blaetter ausblenden
    For Each ws In ThisWorkbook.Worksheets
        If mProtect.IsInternalSheet(ws.Name) Then ws.Visible = xlSheetVeryHidden
    Next
    Application.EnableEvents = True
End Sub

' Nach dem Einbetten der Bilder (build.ps1): Logo auf dem Hinweisblatt
Public Sub Build_Logos()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_HINWEIS)
    ws.Unprotect Password:=""
    ws.Range("B2").Value = ""
    ws.Rows(2).RowHeight = 52
    mLogo.PlaceSheetLogo ws, ws.Range("B2").Left, ws.Range("B2").Top + 6, 40
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True
End Sub

Private Sub SysDefault(ByVal key As String, ByVal value As String)
    If Len(mCfg.SysGet(key, "")) = 0 Then mCfg.SysSet key, value
End Sub

Private Sub WriteHinweis(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Cells.Font.Name = UI_FONT
    ws.Columns(1).ColumnWidth = 3
    ws.Columns(2).ColumnWidth = 110
    ws.Range("B2").Value = APP_TITLE
    ws.Range("B2").Font.Size = 20: ws.Range("B2").Font.Bold = True
    ws.Range("B4").Value = "Diese Datei ist eine Excel-Anwendung und benötigt aktivierte Makros."
    ws.Range("B4").Font.Size = 13: ws.Range("B4").Font.Bold = True: ws.Range("B4").Font.Color = RGB(185, 28, 28)
    ws.Range("B6").Value = "Sie sehen dieses Blatt, weil die Makros beim Öffnen nicht ausgeführt wurden. Ohne Makros sind Erfassung, Suche, Ansichten und Auswertungen nicht bedienbar; es werden auch keine Funktionen vorgetäuscht."
    ws.Range("B8").Value = "So geht es weiter:"
    ws.Range("B8").Font.Bold = True
    ws.Range("B9").Value = "1. Erscheint oben eine gelbe Leiste „Inhalt aktivieren“ bzw. „Bearbeitung aktivieren“: anklicken."
    ws.Range("B10").Value = "2. Erscheint eine rote Leiste „Sicherheitsrisiko: Makros wurden blockiert“: Die Datei stammt aus dem Internet oder einer nicht vertrauenswürdigen Quelle. Bitte die IT bzw. die zuständige Stelle fragen."
    ws.Range("B11").Value = "3. In Organisationen werden Makros üblicherweise über einen vertrauenswürdigen Speicherort oder eine digitale Signatur der IT freigegeben. Diese Datei verändert keine Sicherheitseinstellungen."
    ws.Range("B13").Value = "Die Anwendung arbeitet vollständig offline: keine Cloud, keine Anmeldung, keine Internetverbindung."
    ws.Range("B14").Value = AUTHOR_NOTE & " · Version " & PROG_VERSION
    ws.Range("B6:B14").WrapText = True
    ws.Range("B13:B14").Font.Color = RGB(107, 114, 128)
    ws.Rows.AutoFit
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True
End Sub

'==============================================================================
' Vorlagen (Konfigurationszeilen; "~" steht innerhalb von Formeln/Filtern fuer "|")
'==============================================================================
Private Sub TL(ByVal ws As Worksheet, ByVal s As String)
    mTplRow = mTplRow + 1
    ws.Cells(mTplRow, 1).Value = s
End Sub

Private Sub WriteTemplates()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_TPL)
    ws.Cells.Clear
    ws.Columns(1).NumberFormat = "@"
    mTplRow = 0
    TL ws, "VORLAGE|aufgaben|Aufgaben & Projekte|Projekte mit Aufgaben, Status, Prioritäten, Fristen, Kosten und Fortschritt"
    TL ws, "TABELLE|P|Projekte|Projekte und Vorhaben mit Laufzeit, Budget und Aufgabenstand|59621"
    TL ws, "FELD|P|Projekt|TEXT|pflicht=1;bez=1"
    TL ws, "FELD|P|Status|STATUS|optionen=Geplant:grau,Aktiv:blau,Pausiert:gelb,Abgeschlossen:gruen:x;std=Geplant"
    TL ws, "FELD|P|Verantwortlich|TEXT|"
    TL ws, "FELD|P|Beginn|DATUM|"
    TL ws, "FELD|P|Ende|DATUM|"
    TL ws, "FELD|P|Budget|GELD|min=0"
    TL ws, "FELD|P|Beschreibung|LTEXT|gruppe=Beschreibung"
    TL ws, "NICHTVOR|P|Ende|Beginn"
    TL ws, "TABELLE|A|Aufgaben|Aufgaben mit Status, Priorität, Frist und Aufwand|59581"
    TL ws, "FELD|A|Aufgabe|TEXT|pflicht=1;bez=1"
    TL ws, "FELD|A|Nr.|AUTOID|std=AUF"
    TL ws, "FELD|A|Status|STATUS|optionen=Offen:grau,In Arbeit:blau,Wartet:orange,Erledigt:gruen:x;std=Offen;pflicht=1"
    TL ws, "FELD|A|Priorität|AUSWAHL|optionen=Niedrig:gruen,Mittel:gelb,Hoch:rot;std=Mittel"
    TL ws, "FELD|A|Zuständig|PERSON|hilfe=Eine oder mehrere Personen; neue Personen über „+“"
    TL ws, "FELD|A|Wichtig|KONTROLL|"
    TL ws, "FELD|A|Beginn|DATUM|std=heute"
    TL ws, "FELD|A|Frist|DATUM|"
    TL ws, "FELD|A|Aufwand (h)|ZAHL|dez=1;min=0"
    TL ws, "FELD|A|Kosten|GELD|min=0"
    TL ws, "FELD|A|Fortschritt|PROZENT|min=0;max=100"
    TL ws, "FELD|A|Notizen|LTEXT|gruppe=Notizen"
    TL ws, "BUTTON|A|Als erledigt markieren|{Status}=Erledigt,{Fortschritt}=100"
    TL ws, "LINK|A|Projekt|P|0|Aufgaben"
    TL ws, "NICHTVOR|A|Frist|Beginn"
    TL ws, "FORMEL|A|Tage bis Frist|TAGEBIS~F:{Frist}|ZAHL|0"
    TL ws, "FORMEL|A|Überfällig|UEBERFAELLIG~F:{Frist}~F:{Status}|JANEIN|0"
    TL ws, "FORMEL|P|Aufgaben gesamt|ROLLUP~F:{Aufgaben}~K:ANZAHL|ZAHL|0|Auswertung"
    TL ws, "FORMEL|P|Offene Aufgaben|ROLLUP~F:{Aufgaben}~K:ANZAHLWENN~~F:{A.Status}~K:<>~O:{A.Status=Erledigt}|ZAHL|0|Auswertung"
    TL ws, "FORMEL|P|Erledigte Aufgaben|ROLLUP~F:{Aufgaben}~K:ANZAHLWENN~~F:{A.Status}~K:=~O:{A.Status=Erledigt}|ZAHL|0|Auswertung"
    TL ws, "FORMEL|P|Erledigt in %|ANTEIL~F:{Erledigte Aufgaben}~F:{Aufgaben gesamt}|PROZENT|0|Auswertung"
    TL ws, "FORMEL|P|Kosten der Aufgaben|ROLLUP~F:{Aufgaben}~K:SUMME~F:{A.Kosten}|GELD|2|Auswertung"
    TL ws, "FORMEL|P|Budget verbleibend|RECHNEN~F:{Budget}~K:-~F:{Kosten der Aufgaben}|GELD|2|Auswertung"
    TL ws, "ANSICHT|P|Alle Projekte|Tabelle||||{Beginn}:A|"
    TL ws, "ANSICHT|P|Projektleiste|Gantt|start={Beginn};ende={Ende};farbe={Status};fortschritt={Erledigt in %};massstab=Woche||||"
    TL ws, "ANSICHT|A|Alle Aufgaben|Tabelle||||{Frist}:A|"
    TL ws, "ANSICHT|A|Offene Aufgaben|Tabelle|||~{Status}~OFFEN~~|{Frist}:A|f={Projekt}"
    TL ws, "ANSICHT|A|Board nach Status|Board|gruppe={Status};karte={Priorität},{Frist},{Projekt}||||"
    TL ws, "ANSICHT|A|Kalender|Kalender|datum={Frist};farbe={Status}||||"
    TL ws, "VORLAGE|kontakte|Kontakte|Personen und Organisationen mit Kategorien, Kontaktdaten und Dublettenprüfung"
    TL ws, "TABELLE|K|Kontakte|Personen und Organisationen|59671"
    TL ws, "FELD|K|Name|TEXT|pflicht=1;bez=1"
    TL ws, "FELD|K|Organisation|TEXT|"
    TL ws, "FELD|K|Funktion|TEXT|"
    TL ws, "FELD|K|E-Mail|EMAIL|"
    TL ws, "FELD|K|Telefon|TELEFON|"
    TL ws, "FELD|K|Website|URL|"
    TL ws, "FELD|K|Kategorie|MEHRFACH|optionen=Partner:blau,Kunde:gruen,Lieferant:orange,Behörde:lila,Intern:grau"
    TL ws, "FELD|K|Aktiv|KONTROLL|std=Ja"
    TL ws, "FELD|K|Letzter Kontakt|DATUM|"
    TL ws, "FELD|K|Notizen|LTEXT|gruppe=Notizen"
    TL ws, "DUBLETTE|K|{Name},{Organisation}"
    TL ws, "ANSICHT|K|Alle Kontakte|Tabelle||||{Name}:A|"
    TL ws, "ANSICHT|K|Aktive Kontakte|Tabelle|||~{Aktiv}~JA~~|{Organisation}:A|"
    TL ws, "VORLAGE|ressourcen|Ressourcen & Material|Inventar, Bestände, Standorte, Ausleihe und Wert"
    TL ws, "TABELLE|R|Ressourcen|Geräte, Material, Räume und Fahrzeuge|59621"
    TL ws, "FELD|R|Bezeichnung|TEXT|pflicht=1;bez=1"
    TL ws, "FELD|R|Inventarnummer|TEXT|eindeutig=1"
    TL ws, "FELD|R|Kategorie|AUSWAHL|optionen=Gerät:blau,Material:braun,Raum:lila,Fahrzeug:orange,Software:tuerkis"
    TL ws, "FELD|R|Standort|TEXT|"
    TL ws, "FELD|R|Status|STATUS|optionen=Verfügbar:gruen,Ausgeliehen:orange,In Wartung:gelb,Ausgemustert:grau:x;std=Verfügbar"
    TL ws, "FELD|R|Menge|ZAHL|min=0;std=1"
    TL ws, "FELD|R|Mindestbestand|ZAHL|min=0"
    TL ws, "FELD|R|Stückpreis|GELD|min=0"
    TL ws, "FELD|R|Ausgeliehen bis|DATUM|"
    TL ws, "FELD|R|Dokument|PFAD|"
    TL ws, "FORMEL|R|Gesamtwert|RECHNEN~F:{Menge}~K:*~F:{Stückpreis}|GELD|2"
    TL ws, "FORMEL|R|Nachbestellen|PRUEFEN~F:{Menge}~K:<~F:{Mindestbestand}|JANEIN|0"
    TL ws, "ANSICHT|R|Alle Ressourcen|Tabelle||||{Bezeichnung}:A|"
    TL ws, "ANSICHT|R|Nach Status|Board|gruppe={Status};karte={Kategorie},{Standort}||||"
    TL ws, "VORLAGE|doku|Dokumentation|Dokumente, Protokolle und Nachweise mit Ablageort und Wiedervorlage"
    TL ws, "TABELLE|D|Dokumentation|Protokolle, Berichte, Verträge und Nachweise|59681"
    TL ws, "FELD|D|Titel|TEXT|pflicht=1;bez=1"
    TL ws, "FELD|D|Art|AUSWAHL|optionen=Protokoll:blau,Bericht:lila,Vertrag:orange,Nachweis:gruen,Sonstiges:grau"
    TL ws, "FELD|D|Datum|DATUM|std=heute;pflicht=1"
    TL ws, "FELD|D|Verfasst von|TEXT|"
    TL ws, "FELD|D|Status|STATUS|optionen=Entwurf:grau,In Prüfung:gelb,Freigegeben:gruen:x;std=Entwurf"
    TL ws, "FELD|D|Schlagworte|MEHRFACH|optionen=Intern:grau,Extern:blau,Vertraulich:rot,Archivrelevant:braun"
    TL ws, "FELD|D|Wiedervorlage|DATUM|"
    TL ws, "FELD|D|Ablageort|PFAD|"
    TL ws, "FELD|D|Inhalt / Zusammenfassung|LTEXT|gruppe=Inhalt"
    TL ws, "FORMEL|D|Tage bis Wiedervorlage|TAGEBIS~F:{Wiedervorlage}|ZAHL|0"
    TL ws, "ANSICHT|D|Alle Dokumente|Tabelle||||{Datum}:D|"
    TL ws, "ANSICHT|D|Wiedervorlagen|Kalender|datum={Wiedervorlage};farbe={Status}||||"
    TL ws, "VORLAGE|veranst|Veranstaltungen & Einrichtungen|Veranstaltungen mit mehreren beteiligten Einrichtungen (n:m-Beziehung)"
    TL ws, "TABELLE|E|Einrichtungen|Beteiligte Einrichtungen und Partner|59647"
    TL ws, "FELD|E|Einrichtung|TEXT|pflicht=1;bez=1"
    TL ws, "FELD|E|Art|AUSWAHL|optionen=Schule:blau,Kita:gelb,Verein:gruen,Verwaltung:grau,Beratungsstelle:lila"
    TL ws, "FELD|E|Stadtteil|AUSWAHL|optionen=Nord:blau,Süd:gruen,Ost:orange,West:lila,Mitte:grau"
    TL ws, "FELD|E|Ansprechperson|TEXT|"
    TL ws, "FELD|E|Plätze|ZAHL|min=0"
    TL ws, "TABELLE|V|Veranstaltungen|Termine mit Format, Themen und beteiligten Einrichtungen|59271"
    TL ws, "FELD|V|Veranstaltung|TEXT|pflicht=1;bez=1"
    TL ws, "FELD|V|Datum|DATUMZEIT|pflicht=1"
    TL ws, "FELD|V|Ende|DATUMZEIT|"
    TL ws, "FELD|V|Format|AUSWAHL|optionen=Workshop:blau,Infoabend:gelb,Fachtag:lila,Netzwerktreffen:gruen"
    TL ws, "FELD|V|Themen|MEHRFACH|optionen=Bildung:blau,Gesundheit:gruen,Integration:orange,Sport:rot,Kultur:lila,Digitales:tuerkis"
    TL ws, "FELD|V|Teilnehmende|ZAHL|min=0"
    TL ws, "FELD|V|Status|STATUS|optionen=Geplant:grau,Bestätigt:blau,Durchgeführt:gruen:x,Abgesagt:rot:x;std=Geplant"
    TL ws, "LINK|V|Einrichtungen|E|1|Veranstaltungen"
    TL ws, "NICHTVOR|V|Ende|Datum"
    TL ws, "FORMEL|E|Anzahl Veranstaltungen|ROLLUP~F:{Veranstaltungen}~K:ANZAHL|ZAHL|0"
    TL ws, "FORMEL|V|Beteiligte Einrichtungen|ROLLUP~F:{Einrichtungen}~K:ANZAHL|ZAHL|0"
    TL ws, "ANSICHT|V|Alle Veranstaltungen|Tabelle||||{Datum}:A|"
    TL ws, "ANSICHT|V|Kalender|Kalender|datum={Datum};ende={Ende};farbe={Format}||||"
    TL ws, "ANSICHT|E|Alle Einrichtungen|Tabelle||||{Einrichtung}:A|"
    ' Grünflächenamt: Spielplätze, Spielgeräte mit Prüfintervallen, Bäume, Projekte mit Bürgerbeteiligung
    TL ws, "VORLAGE|gruen|Spielplätze & Grünflächen|Spielplätze, Spielgeräte mit Prüffristen, Bäume mit Kontrollen, Beteiligungsprojekte"
    TL ws, "TABELLE|S|Spielplätze|Alle Spielplätze mit Lage, Altersgruppen und Zustand|59667"
    TL ws, "FELD|S|Spielplatz|TEXT|pflicht=1;bez=1;eindeutig=1;hilfe=Name des Spielplatzes, z. B. Spielplatz Lindenhof"
    TL ws, "FELD|S|Nr.|AUTOID|std=SP"
    TL ws, "FELD|S|Stadtteil|AUSWAHL|pflicht=1;optionen=Nord:blau,Süd:gruen,Ost:orange,West:lila,Mitte:grau"
    TL ws, "FELD|S|Altersgruppen|MEHRFACH|optionen=0 bis 3 Jahre:rosa,3 bis 6 Jahre:gelb,6 bis 12 Jahre:gruen,ab 12 Jahre:blau;hilfe=Für welche Altersgruppen ist der Platz ausgelegt?"
    TL ws, "FELD|S|Fläche (m²)|ZAHL|min=0;hilfe=Spielfläche in Quadratmetern"
    TL ws, "FELD|S|Zustand|STATUS|optionen=Gut:gruen,Mängel:gelb,Gesperrt:rot;std=Gut;pflicht=1"
    TL ws, "FELD|S|Adresse|ORT|"
    TL ws, "FELD|S|Zuständig|PERSON|hilfe=Zuständige Person im Grünflächenamt"
    TL ws, "TABELLE|G|Spielgeräte|Spielgeräte mit Prüfintervall, Prüfergebnis und Mängeln|59663"
    TL ws, "FELD|G|Gerät|TEXT|pflicht=1;bez=1;hilfe=Bezeichnung des Spielgeräts"
    TL ws, "FELD|G|Gerätenr.|AUTOID|std=GER"
    TL ws, "FELD|G|Geräteart|AUSWAHL|pflicht=1;optionen=Schaukel:blau,Rutsche:orange,Klettergerät:lila,Wippe:gelb,Sandspiel:braun,Karussell:rosa,Seilbahn:tuerkis"
    TL ws, "FELD|G|Baujahr|ZAHL|min=1980;max=2030;hilfe=Jahr der Aufstellung (vierstellig)"
    TL ws, "FELD|G|Letzte Prüfung|DATUM|pflicht=1;hilfe=Datum der letzten Hauptuntersuchung (TT.MM.JJJJ)"
    TL ws, "FELD|G|Prüfintervall (Tage)|ZAHL|min=1;max=730;std=365;hilfe=Abstand der Hauptuntersuchungen in Tagen"
    TL ws, "FELD|G|Prüfergebnis|STATUS|pflicht=1;optionen=Ohne Mängel:gruen,Geringe Mängel:gelb,Sofort beheben:rot,Gesperrt:grau;std=Ohne Mängel;hilfe=Ergebnis der letzten Prüfung"
    TL ws, "FELD|G|Mängel|MEHRFACH|optionen=Verschleiß:gelb,Fallschutz:orange,Splitter:braun,Lockere Teile:rot,Vandalismus:lila;hilfe=Alle zutreffenden Mängel ankreuzen"
    TL ws, "FELD|G|Bemerkung|TEXT|max_len=200"
    TL ws, "FELD|G|Erfasst von|TEXT|gruppe=Rücklauf;hilfe=Person/Stelle aus der Erfassungsdatei"
    TL ws, "LINK|G|Spielplatz|S|0|Spielgeräte"
    TL ws, "FORMEL|G|Nächste Prüfung|DATPLUS~F:{Letzte Prüfung}~F:{Prüfintervall (Tage)}|DATUM|0"
    TL ws, "FORMEL|G|Tage bis Prüfung|TAGEBIS~F:{Nächste Prüfung}|ZAHL|0"
    TL ws, "FORMEL|G|Prüfung überfällig|UEBERFAELLIG~F:{Nächste Prüfung}|JANEIN|0"
    TL ws, "TABELLE|B|Bäume|Baumkataster der Spielplätze mit Vitalität und Kontrollen|60426"
    TL ws, "FELD|B|Baum|TEXT|pflicht=1;bez=1;hilfe=Baumart und Plakettennummer, z. B. Linde 1042"
    TL ws, "FELD|B|Baumart|AUSWAHL|pflicht=1;optionen=Linde:gruen,Ahorn:orange,Eiche:braun,Platane:tuerkis,Kastanie:rot,Birke:grau"
    TL ws, "FELD|B|Pflanzjahr|ZAHL|min=1850;max=2030"
    TL ws, "FELD|B|Stammumfang (cm)|ZAHL|min=0"
    TL ws, "FELD|B|Vitalität|STATUS|pflicht=1;optionen=Vital:gruen,Leicht geschädigt:gelb,Stark geschädigt:orange,Abgängig:rot;std=Vital"
    TL ws, "FELD|B|Letzte Baumkontrolle|DATUM|pflicht=1"
    TL ws, "FELD|B|Maßnahme|STATUS|optionen=Keine:gruen:x,Totholz entfernen:gelb,Kronenpflege:orange,Fällung prüfen:rot;std=Keine;hilfe=Erforderliche Pflegemaßnahme"
    TL ws, "FELD|B|Verkehrssicher|KONTROLL|std=Ja"
    TL ws, "FELD|B|Erfasst von|TEXT|gruppe=Rücklauf"
    TL ws, "LINK|B|Spielplatz|S|0|Bäume"
    TL ws, "TABELLE|P|Projekte|Neubau, Sanierung und Beteiligungsverfahren|60612"
    TL ws, "FELD|P|Projekt|TEXT|pflicht=1;bez=1"
    TL ws, "FELD|P|Art|AUSWAHL|pflicht=1;optionen=Neubau:blau,Sanierung:orange,Erweiterung:lila,Pflanzung:gruen"
    TL ws, "FELD|P|Status|STATUS|optionen=Idee:grau,Planung:blau,Beteiligung:lila,Umsetzung:orange,Fertig:gruen:x;std=Idee"
    TL ws, "FELD|P|Beginn|DATUM|"
    TL ws, "FELD|P|Ende|DATUM|"
    TL ws, "FELD|P|Budget|GELD|min=0"
    TL ws, "FELD|P|Beteiligung|MEHRFACH|optionen=Kinderbeteiligung:gelb,Bürgerversammlung:blau,Online-Umfrage:tuerkis,Ortsbeirat:lila;hilfe=Welche Beteiligungsformate gab es?"
    TL ws, "FELD|P|Teilnehmende|ZAHL|min=0"
    TL ws, "LINK|P|Spielplatz|S|0|Projekte"
    TL ws, "NICHTVOR|P|Ende|Beginn"
    TL ws, "FORMEL|S|Geräte|ROLLUP~F:{Spielgeräte}~K:ANZAHL|ZAHL|0|Auswertung"
    TL ws, "FORMEL|S|Geräte mit Mängeln|ROLLUP~F:{Spielgeräte}~K:ANZAHLWENN~~F:{G.Prüfergebnis}~K:<>~O:{G.Prüfergebnis=Ohne Mängel}|ZAHL|0|Auswertung"
    TL ws, "FORMEL|S|Bäume gesamt|ROLLUP~F:{Bäume}~K:ANZAHL|ZAHL|0|Auswertung"
    TL ws, "ANSICHT|S|Alle Spielplätze|Tabelle||||{Spielplatz}:A|"
    TL ws, "ANSICHT|S|Nach Zustand|Board|gruppe={Zustand};karte={Stadtteil},{Geräte},{Geräte mit Mängeln}||||"
    TL ws, "ANSICHT|G|Alle Geräte|Tabelle||||{Nächste Prüfung}:A|"
    TL ws, "ANSICHT|G|Prüfung überfällig|Tabelle|||~{Prüfung überfällig}~JA~~|{Nächste Prüfung}:A|"
    TL ws, "ANSICHT|G|Prüfkalender|Kalender|datum={Nächste Prüfung};farbe={Prüfergebnis}||||"
    TL ws, "ANSICHT|G|Nach Prüfergebnis|Board|gruppe={Prüfergebnis};karte={Spielplatz},{Nächste Prüfung}||||"
    TL ws, "ANSICHT|B|Alle Bäume|Tabelle||||{Baum}:A|"
    TL ws, "ANSICHT|B|Nach Vitalität|Board|gruppe={Vitalität};karte={Baumart},{Spielplatz},{Maßnahme}||||"
    TL ws, "ANSICHT|P|Alle Projekte|Tabelle||||{Beginn}:A|"
    TL ws, "ANSICHT|P|Projektleiste|Gantt|start={Beginn};ende={Ende};farbe={Status};massstab=Monat;namen=Ja||||"
End Sub

'==============================================================================
' Codekatalog (nur beim Build: liest das VBA-Projekt der Mappe)
'==============================================================================
Public Sub Build_CodeCatalog()
    Dim comp As Object, ws As Worksheet, n As Long, i As Long, a() As Variant, total As Long, kind As String, cm As Object, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_CODE)
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("Modul", "Art", "Nr", "Text")
    ws.Columns(4).NumberFormat = "@"
    For Each comp In ThisWorkbook.VBProject.VBComponents
        total = total + comp.CodeModule.CountOfLines
    Next
    ReDim a(1 To total + 1, 1 To 4)
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Set cm = comp.CodeModule
        If cm.CountOfLines > 0 Then
            Select Case comp.Type
                Case 1: kind = "Modul"
                Case 2: kind = "Klasse"
                Case 3: kind = "Formular"
                Case Else: kind = "Dokument"
            End Select
            If kind <> "Dokument" Or comp.Name = "ThisWorkbook" Or comp.Name = ThisWorkbook.CodeName Then
                For i = 1 To cm.CountOfLines
                    n = n + 1
                    a(n, 1) = comp.Name: a(n, 2) = kind: a(n, 3) = i
                    a(n, 4) = CellText(cm.Lines(i, 1))
                Next
            End If
        End If
    Next
    If n > 0 Then ws.Range("A2").Resize(n, 4).Value = mData.SliceRows(a, n, 4)
    ' Entwicklungsstand nur mit Monat und Jahr (das Quellcode-Fenster zeigt den Katalogstempel im Titel)
    mCfg.SysSet "CodeKatalog", PROG_VERSION & " / " & PROG_DATE & " / " & n & " Zeilen"
End Sub

' Endzustand fuer Auslieferung (Startseite, Hinweisblatt vorbereitet)
Public Sub Build_Finalize()
    Dim ws As Worksheet
    Application.EnableEvents = False
    NeutralBuildUser
    mCfg.InvalidateAll
    mData.InvalidateAllData
    For Each ws In ThisWorkbook.Worksheets
        If mProtect.IsInternalSheet(ws.Name) Then
            If ws.ProtectContents Then ws.Unprotect ""
            ws.Visible = xlSheetVeryHidden
        End If
    Next
    ThisWorkbook.Worksheets(SH_HINWEIS).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(SH_HINWEIS).Activate
    ActiveWindow.DisplayWorkbookTabs = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SH_HINWEIS And ws.Visible = xlSheetVisible Then ws.Visible = xlSheetHidden
    Next
    Application.EnableEvents = True
End Sub

' Benutzername des Rechners, auf dem gebaut wurde, gehoert nicht in die Lieferdatei: Protokoll, Systemspalte
' "_Benutzer" und Felder "Erstellt von"/"Geändert von" bekommen "Aufbau". Formelspalten bleiben unberuehrt.
Private Sub NeutralBuildUser()
    Dim ws As Worksheet, lo_ As ListObject, colRng As Range, v As Variant
    Dim r As Long, c As Long, a As Long, n As Long, hit As Boolean, wasProt As Boolean, u1 As String, u2 As String
    u1 = LCase$(Environ$("USERNAME")): u2 = LCase$(Application.UserName)
    For Each ws In ThisWorkbook.Worksheets
        wasProt = ws.ProtectContents
        For Each lo_ In ws.ListObjects
            If Not lo_.DataBodyRange Is Nothing Then
                For c = 1 To lo_.ListColumns.Count
                    Set colRng = lo_.ListColumns(c).DataBodyRange
                    If colRng.Cells.CountLarge = 1 Then
                        ReDim v(1 To 1, 1 To 1): v(1, 1) = colRng.Value
                    Else
                        v = colRng.Value
                    End If
                    n = UBound(v, 1): a = 0
                    ' zusammenhaengende Treffer als ein Block schreiben (in Demo und Lasttest ganze Spalten)
                    For r = 1 To n + 1
                        hit = False
                        If r <= n Then hit = IsBuildUser(v(r, 1), u1, u2)
                        If hit Then
                            If a = 0 Then a = r
                        ElseIf a > 0 Then
                            NeutralRun ws, colRng, a, r - 1
                            a = 0
                        End If
                    Next
                Next
            End If
        Next
        If wasProt And Not ws.ProtectContents And Not mProtect.IsInternalSheet(ws.Name) Then mProtect.ProtectUi ws
    Next
End Sub

Private Function IsBuildUser(ByVal x As Variant, ByVal u1 As String, ByVal u2 As String) As Boolean
    Dim s As String
    If VarType(x) <> vbString Then Exit Function
    s = LCase$(Trim$(x))
    If Len(s) = 0 Then Exit Function
    IsBuildUser = (s = u1 Or s = u2)
End Function

Private Sub NeutralRun(ByVal ws As Worksheet, ByVal colRng As Range, ByVal a As Long, ByVal b As Long)
    Dim rg As Range, cel As Range
    If ws.ProtectContents Then ws.Unprotect ""
    Set rg = ws.Range(colRng.Cells(a, 1), colRng.Cells(b, 1))
    ' HasFormula ist bei gemischten Bereichen Null: dann Zelle fuer Zelle, Formeln bleiben stehen
    If rg.HasFormula = False Then
        rg.Value = "Aufbau"
    Else
        For Each cel In rg.Cells
            If Not cel.HasFormula Then cel.Value = "Aufbau"
        Next
    End If
End Sub

'==============================================================================
' Demodatei (fiktive Beispieldaten)
'==============================================================================
Private Function Rnd01() As Double
    ' Park-Miller-Generator mit Gleitkommaarithmetik (reproduzierbar, ohne Ueberlauf)
    Dim x As Double
    If mRnd <= 0 Then mRnd = 1
    x = CDbl(mRnd) * 16807#
    x = x - Int(x / 2147483647#) * 2147483647#
    mRnd = CLng(x)
    Rnd01 = x / 2147483647#
End Function

Private Function Pick(ByVal a As Variant) As Variant
    Pick = a(Int(Rnd01() * (UBound(a) + 1)))
End Function

Private Function FidOf(ByVal tid As String, ByVal nm As String) As String
    FidOf = mCfg.FieldIdByName(tid, nm)
    If Len(FidOf) = 0 Then Err.Raise vbObjectError + 600, , "Feld " & nm & " fehlt"
End Function

Private Function OidOf(ByVal tid As String, ByVal fieldName As String, ByVal optName As String) As String
    OidOf = mCfg.OptIdByName(FidOf(tid, fieldName), optName)
End Function

' Legt einen Datensatz an; pairs: Array("Feldname", Wert, ...)
Private Function AddRec(ByVal tid As String, ByVal pairs As Variant) As String
    Dim rid As String, a As Variant, t As cTableData, i As Long, c As Long
    rid = mData.NewRecordId()
    a = mData.NewRowArray(tid, rid)
    Set t = mData.TData(tid)
    For i = 0 To UBound(pairs) Step 2
        c = t.ColOf(FidOf(tid, CStr(pairs(i))))
        a(c) = pairs(i + 1)
    Next
    mData.InsertRecord tid, a
    AddRec = rid
End Function

Public Sub Build_Demo()
    Dim tP As String, tA As String, tE As String, tV As String, tR As String, i As Long, j As Long, rid As String, d0 As Date
    Dim projects() As String, ress() As String, einr() As String, did As String, props As cMap, linkF As String, k As Long
    Dim pNames As Variant, aVerbs As Variant, aObjs As Variant, people As Variant, n As Long, st As String, dStart As Date, dEnd As Date
    Dim rNames As Variant, eNames As Variant, vNames As Variant, lst As String, themes As Variant, ids As Variant
    mRnd = 20261025
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mApp.gStarted = False
    mCfg.SysSet "Modus", "Arbeiten"
    mCfg.SysSet "Demo", "Ja"
    mCfg.SysSet "AppName", "Demo · MidPoint"
    mCfg.SysSet "TitelFestgelegt", "Ja"
    mCfg.SysSet "Untertitel", "Fiktive Beispieldaten"
    tP = mStruct.ApplyTemplate("aufgaben")
    tA = mCfg.TableIdByName("Aufgaben")
    tE = mStruct.ApplyTemplate("veranst")
    tV = mCfg.TableIdByName("Veranstaltungen")
    tR = mStruct.ApplyTemplate("ressourcen")
    ' zusaetzliche n:m-Verknuepfung Aufgabe -> benoetigte Ressourcen
    Set props = New cMap
    props.SetVal "Zieltabelle", tR
    props.SetVal "Mehrfach", "Ja"
    props.SetVal "Gegenname", "Eingesetzt in Aufgaben"
    props.SetVal "Gruppe", "Ressourcen"
    linkF = mStruct.AddField(tA, "Benötigte Ressourcen", FT_LINK, props)
    people = Array("A. Berger", "M. Kaya", "S. Nowak", "J. Lindqvist", "T. Okafor", "L. Romano", "F. Hartmann", "D. Yilmaz")
    For i = 0 To UBound(people)
        mStruct.AddOption FidOf(tA, "Zuständig"), CStr(people(i)), "", False
    Next
    mStruct.RebuildLists
    d0 = DateSerial(2026, 10, 1)

    mData.TxBegin "Demodaten"
    ' Projekte
    pNames = Array("Umbau Stadtteilbibliothek", "Digitalisierung Antragswesen", "Sommerprogramm Jugend", "Neues Leitsystem Rathaus", _
                   "Energieberatung Quartier Süd", "Ehrenamtsbörse 2027", "Barrierefreie Website", "Fortbildungsreihe Kita")
    people = Array("A. Berger", "M. Kaya", "S. Nowak", "J. Lindqvist", "T. Okafor", "L. Romano", "F. Hartmann", "D. Yilmaz")
    ReDim projects(0 To UBound(pNames))
    For i = 0 To UBound(pNames)
        dStart = d0 + Int((i - 3) * 21 + Rnd01() * 10)
        dEnd = dStart + 60 + Int(Rnd01() * 120)
        st = Choose(1 + (i Mod 4), "Aktiv", "Aktiv", "Geplant", IIf(i = 3, "Abgeschlossen", "Pausiert"))
        projects(i) = AddRec(tP, Array("Projekt", pNames(i), "Status", OidOf(tP, "Status", st), "Verantwortlich", people(i), _
                                       "Beginn", dStart, "Ende", dEnd, "Budget", CDbl(5000 + Int(Rnd01() * 40) * 1000), _
                                       "Beschreibung", "Fiktives Beispielprojekt zur Demonstration von Beziehungen, Berechnungen und Ansichten."))
    Next
    ' Ressourcen
    rNames = Array("Beamer Konferenzraum", "Laptop Pool 1", "Laptop Pool 2", "Moderationskoffer", "Stellwände (Set)", "Lastenrad", _
                   "Dienstwagen E-Kombi", "Großer Sitzungssaal", "Werkraum", "Kamera-Set", "Faltpavillon", "Tablet-Koffer (10 Stk.)", _
                   "Kopierpapier A4 (Karton)", "Flipchart-Papier", "Erste-Hilfe-Set", "Lautsprecheranlage", "Visualisierungssoftware", "Etikettendrucker")
    ReDim ress(0 To UBound(rNames))
    For i = 0 To UBound(rNames)
        st = Choose(1 + (i Mod 5), "Verfügbar", "Verfügbar", "Ausgeliehen", "Verfügbar", "In Wartung")
        ress(i) = AddRec(tR, Array("Bezeichnung", rNames(i), "Inventarnummer", "INV-" & Format$(1000 + i * 7, "0000"), _
                                   "Kategorie", OidOf(tR, "Kategorie", Choose(1 + (i Mod 5), "Gerät", "Gerät", "Material", "Raum", "Fahrzeug")), _
                                   "Standort", Pick(Array("Rathaus EG", "Rathaus 2. OG", "Bürgerhaus", "Lager Nord", "Bibliothek")), _
                                   "Status", OidOf(tR, "Status", st), "Menge", CDbl(1 + Int(Rnd01() * 12)), "Mindestbestand", CDbl(Int(Rnd01() * 6)), _
                                   "Stückpreis", CDbl(Int(20 + Rnd01() * 1800)), "Ausgeliehen bis", IIf(st = "Ausgeliehen", d0 + 20 + i, Empty)))
    Next
    ' Aufgaben
    aVerbs = Array("Angebote einholen für", "Konzept abstimmen:", "Termin planen:", "Material bestellen für", "Bericht schreiben:", _
                   "Rückmeldung einholen zu", "Ausschreibung vorbereiten:", "Präsentation erstellen:", "Abnahme durchführen:", "Checkliste prüfen:")
    aObjs = Array("Möblierung", "Beschilderung", "Schulungstermine", "Technik", "Öffentlichkeitsarbeit", "Haushaltsmittel", "Barrierefreiheit", "Raumplanung")
    n = 0
    For i = 0 To UBound(projects)
        For j = 1 To 6 + Int(Rnd01() * 4)
            n = n + 1
            dStart = d0 + Int((i - 3) * 21 + j * 5 + Rnd01() * 6)
            st = Pick(Array("Offen", "Offen", "In Arbeit", "In Arbeit", "Wartet", "Erledigt", "Erledigt"))
            rid = AddRec(tA, Array("Aufgabe", Pick(aVerbs) & " " & Pick(aObjs), "Status", OidOf(tA, "Status", st), _
                                   "Priorität", OidOf(tA, "Priorität", Pick(Array("Niedrig", "Mittel", "Mittel", "Hoch"))), _
                                   "Zuständig", OidOf(tA, "Zuständig", Pick(people)), "Wichtig", (Rnd01() < 0.25), "Beginn", dStart, "Frist", dStart + 3 + Int(Rnd01() * 25), _
                                   "Aufwand (h)", Round(1 + Rnd01() * 24, 1), "Kosten", CDbl(Int(Rnd01() * 40) * 50), _
                                   "Fortschritt", IIf(st = "Erledigt", 1#, Round(Rnd01() * 0.9, 1))))
            mData.SetLinks FidOf(tA, "Projekt"), rid, projects(i)
            lst = ""
            For k = 1 To Int(Rnd01() * 3)
                lst = ListAdd(lst, ress(Int(Rnd01() * (UBound(ress) + 1))))
            Next
            If Len(lst) > 0 Then mData.SetLinks linkF, rid, lst
        Next
    Next
    ' Aufgabe ohne Frist und eine im Archiv
    rid = AddRec(tA, Array("Aufgabe", "Ideen sammeln für Jahresplanung", "Status", OidOf(tA, "Status", "Offen"), "Zuständig", OidOf(tA, "Zuständig", "A. Berger")))
    rid = AddRec(tA, Array("Aufgabe", "Alte Inventarliste bereinigen", "Status", OidOf(tA, "Status", "Erledigt"), "Zuständig", OidOf(tA, "Zuständig", "M. Kaya"), "Beginn", d0 - 40, "Frist", d0 - 30))
    mData.SetArchived tA, rid, True
    ' Einrichtungen
    eNames = Array("Grundschule am Lindenpark", "Kita Sonnenblume", "Sportverein Blau-Weiß", "Bürgeramt Mitte", "Familienberatung Nord", _
                   "Gesamtschule Westend", "Kita Regenbogen", "Kulturverein Altstadt", "Jugendtreff Ost", "Seniorenbüro Süd", "Musikschule Mitte", "Integrationsbüro")
    ReDim einr(0 To UBound(eNames))
    For i = 0 To UBound(eNames)
        einr(i) = AddRec(tE, Array("Einrichtung", eNames(i), "Art", OidOf(tE, "Art", Choose(1 + (i Mod 5), "Schule", "Kita", "Verein", "Verwaltung", "Beratungsstelle")), _
                                   "Stadtteil", OidOf(tE, "Stadtteil", Pick(Array("Nord", "Süd", "Ost", "West", "Mitte"))), _
                                   "Ansprechperson", Pick(people), "Plätze", CDbl(20 + Int(Rnd01() * 200))))
    Next
    ' Veranstaltungen (mit mehreren Einrichtungen und Themen)
    vNames = Array("Netzwerktreffen Bildung", "Infoabend Ganztag", "Fachtag Inklusion", "Workshop Medienkompetenz", "Sportfest im Park", _
                   "Elternabend Übergang", "Kulturnacht", "Digitalsprechstunde", "Gesundheitstag", "Sprachcafé", "Vorlesetag", "Berufsorientierung", _
                   "Nachbarschaftsfest", "Workshop Antragstellung", "Fachtag Kinderschutz", "Infoabend Ehrenamt", "Netzwerktreffen Quartier", "Erste-Hilfe-Kurs")
    themes = Array("Bildung", "Gesundheit", "Integration", "Sport", "Kultur", "Digitales")
    For i = 0 To UBound(vNames)
        dStart = d0 + Int(Rnd01() * 50) - 5 + TimeSerial(9 + Int(Rnd01() * 9), Pick(Array(0, 30)), 0)
        lst = ""
        For k = 0 To Int(Rnd01() * 3)
            lst = ListAdd(lst, OidOf(tV, "Themen", CStr(Pick(themes))))
        Next
        st = IIf(dStart < Date, "Durchgeführt", Pick(Array("Geplant", "Bestätigt", "Bestätigt")))
        If i = 7 Then st = "Abgesagt"
        rid = AddRec(tV, Array("Veranstaltung", vNames(i), "Datum", dStart, "Ende", dStart + TimeSerial(2 + Int(Rnd01() * 4), 0, 0), _
                               "Format", OidOf(tV, "Format", Pick(Array("Workshop", "Infoabend", "Fachtag", "Netzwerktreffen"))), _
                               "Themen", lst, "Teilnehmende", CDbl(10 + Int(Rnd01() * 120)), "Status", OidOf(tV, "Status", st)))
        lst = ""
        For k = 0 To 1 + Int(Rnd01() * 3)
            lst = ListAdd(lst, einr(Int(Rnd01() * (UBound(einr) + 1))))
        Next
        mData.SetLinks FidOf(tV, "Einrichtungen"), rid, lst
    Next
    ' mehrere Termine am selben Tag (Kalendertest)
    For i = 1 To 6
        rid = AddRec(tV, Array("Veranstaltung", "Beratungstermin " & i, "Datum", DateSerial(2026, 10, 27) + TimeSerial(8 + i, 0, 0), _
                               "Format", OidOf(tV, "Format", "Infoabend"), "Status", OidOf(tV, "Status", "Bestätigt"), "Teilnehmende", CDbl(5 + i)))
    Next
    mData.TxCommit
    mData.ClearUndo
    mCalc.RecalcAll

    ' Dashboard mit sechs Elementen
    did = mDash.CreateDashboard("Projektübersicht", "Fiktive Beispielauswertung: Aufgaben, Kosten, Veranstaltungen")
    DemoElement did, "Offene Aufgaben", "Kachel", tA, "ANZAHL", "", "", "", "1", "~" & FidOf(tA, "Status") & "~OFFEN~~"
    DemoElement did, "Kosten aller Aufgaben", "Kachel", tA, "SUMME", FidOf(tA, "Kosten"), "", "", "1", ""
    DemoElement did, "Überfällige Aufgaben", "Kachel", tA, "ANZAHL", "", "", "", "1", "~" & FidOf(tA, "Überfällig") & "~JA~~"
    DemoElement did, "Durchschnittlicher Fortschritt", "Kachel", tA, "MITTEL", FidOf(tA, "Fortschritt"), "", "", "1", ""
    DemoElement did, "Aufgaben nach Status", "Balken", tA, "ANZAHL", "", FidOf(tA, "Status"), "", "2", ""
    DemoElement did, "Aufgaben nach Frist (Monat)", "Zeitverlauf", tA, "ANZAHL", "", FidOf(tA, "Frist"), FidOf(tA, "Frist"), "2", ""
    DemoElement did, "Veranstaltungen nach Format", "Anteil", tV, "ANZAHL", "", FidOf(tV, "Format"), FidOf(tV, "Datum"), "2", ""
    DemoElement did, "Veranstaltungen nach Themen (Mehrfachzuordnung)", "Tabelle", tV, "ANZAHL", "", FidOf(tV, "Themen"), FidOf(tV, "Datum"), "2", ""
    did = mDash.CreateDashboard("Ressourcen", "Bestand und Wert")
    DemoElement did, "Gesamtwert Inventar", "Kachel", tR, "SUMME", FidOf(tR, "Gesamtwert"), "", "", "1", ""
    DemoElement did, "Nachzubestellen", "Kachel", tR, "ANZAHL", "", "", "", "1", "~" & FidOf(tR, "Nachbestellen") & "~JA~~"
    DemoElement did, "Ressourcen nach Status", "Säulen", tR, "ANZAHL", "", FidOf(tR, "Status"), "", "2", ""
    DemoElement did, "Wert nach Kategorie", "Tabelle", tR, "SUMME", FidOf(tR, "Gesamtwert"), FidOf(tR, "Kategorie"), "", "4", ""
    ' Diagrammvielfalt (ab 1.5): Netz, Kreis, Fläche, Trichter, Baumdiagramm, Pareto, Rangliste, Sankey
    did = mDash.CreateDashboard("Auswertungen", "Weitere Darstellungen und Kennzahlen an den Beispieldaten")
    DemoElement did, "Median Aufwand (h)", "Kachel", tA, "MEDIAN", FidOf(tA, "Aufwand (h)"), "", "", "1", ""
    DemoElement did, "Projekte mit Aufgaben", "Kachel", tA, "EINDEUTIG", FidOf(tA, "Projekt"), "", "", "1", ""
    DemoElement did, "Spannweite Kosten", "Kachel", tA, "SPANNE", FidOf(tA, "Kosten"), "", "", "1", ""
    DemoElement did, "Projekte", "Kachel", tP, "ANZAHL", "", "", "", "1", ""
    DemoElement did, "Aufwand je Zuständigem (Profil)", "Netz", tA, "SUMME", FidOf(tA, "Aufwand (h)"), FidOf(tA, "Zuständig"), "", "2", ""
    DemoElement did, "Kosten nach Priorität", "Kreis", tA, "SUMME", FidOf(tA, "Kosten"), FidOf(tA, "Priorität"), "", "2", ""
    DemoElement did, "Kosten je Projekt (Pareto)", "Pareto", tA, "SUMME", FidOf(tA, "Kosten"), FidOf(tA, "Projekt"), "", "2", ""
    DemoElement did, "Aufgaben je Zuständigem", "Rangliste", tA, "ANZAHL", "", FidOf(tA, "Zuständig"), "", "2", ""
    DemoElement did, "Aufgaben nach Status (Trichter)", "Trichter", tA, "ANZAHL", "", FidOf(tA, "Status"), "", "2", ""
    DemoElement did, "Kosten nach Projekt (Baumdiagramm)", "Baumdiagramm", tA, "SUMME", FidOf(tA, "Kosten"), FidOf(tA, "Projekt"), "", "2", ""
    DemoElement did, "Aufgaben: Status " & ChrW(&H2192) & " Priorität (Sankey)", "Sankey", tA, "ANZAHL", "", FidOf(tA, "Status"), "", "2", "", "ziel=" & FidOf(tA, "Priorität")
    DemoElement did, "Teilnehmende über die Zeit", "Fläche", tV, "SUMME", FidOf(tV, "Teilnehmende"), FidOf(tV, "Datum"), FidOf(tV, "Datum"), "4", ""
    mLog.Add "Demodaten", "", "", "", "Fiktive Beispieldaten erzeugt"
    mLog.Flush
    Application.EnableEvents = True
End Sub

'==============================================================================
' Beispiel Grünflächenamt (fiktive Daten): vier Tabellen, Dashboards, Erfassungsdateien
'==============================================================================
Public Sub Build_Gruen()
    Dim tS As String, tG As String, tB As String, tP As String, i As Long, j As Long, rid As String, d0 As Date, sp() As String
    Dim sNames As Variant, people As Variant, gTypes As Variant, gNames As Variant, lst As String, k As Long, st As String, ty As String
    Dim did As String, m As cMap, cid As String, sel As String, nm As Variant, last As Date, ages As Variant, art As String
    mRnd = 20261028
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mApp.gStarted = False
    mCfg.SysSet "Modus", "Arbeiten"
    mCfg.SysSet "Demo", "Ja"
    mCfg.SysSet "AppName", "Grünflächenamt"
    mCfg.SysSet "TitelFestgelegt", "Ja"
    mCfg.SysSet "Untertitel", "Spielplätze · fiktive Beispieldaten"
    mCfg.SysSet "Navigationsfarbe", CStr(mCmd.NavValue("Waldgrün"))
    mCfg.SysSet "Akzentfarbe", CStr(mCmd.AccentValue("Grün"))
    tS = mStruct.ApplyTemplate("gruen")
    tG = mCfg.TableIdByName("Spielgeräte")
    tB = mCfg.TableIdByName("Bäume")
    tP = mCfg.TableIdByName("Projekte")
    people = Array("K. Brandt", "S. Demir", "L. Vogel", "R. Ahmadi")
    For i = 0 To UBound(people)
        mStruct.AddOption FidOf(tS, "Zuständig"), CStr(people(i)), "", False
    Next
    mStruct.RebuildLists
    d0 = DateSerial(2026, 10, 28)

    mData.TxBegin "Beispieldaten Grünflächenamt"
    sNames = Array("Spielplatz Am Stadtpark", "Spielplatz Lindenhof", "Spielplatz Wasserturm", "Spielplatz Rosengarten", "Spielplatz Kiesgrube", _
                   "Spielplatz Schillerplatz", "Spielplatz Auenwiese", "Spielplatz Mühlbach", "Spielplatz Am Sportfeld", "Spielplatz Sonnenhang", _
                   "Spielplatz Kastanienallee", "Spielplatz Burgwiese", "Spielplatz Feldrand", "Spielplatz Bahnhofstraße")
    ages = Array("0 bis 3 Jahre", "3 bis 6 Jahre", "6 bis 12 Jahre", "ab 12 Jahre")
    ReDim sp(0 To UBound(sNames))
    For i = 0 To UBound(sNames)
        lst = ""
        For k = 0 To 3
            If Rnd01() < 0.6 Or k = 1 Then lst = ListAdd(lst, OidOf(tS, "Altersgruppen", CStr(ages(k))))
        Next
        st = Choose(1 + (i Mod 7), "Gut", "Gut", "Mängel", "Gut", "Gut", "Mängel", IIf(i = 6, "Gesperrt", "Gut"))
        sp(i) = AddRec(tS, Array("Spielplatz", sNames(i), "Stadtteil", OidOf(tS, "Stadtteil", Choose(1 + (i Mod 5), "Nord", "Süd", "Ost", "West", "Mitte")), _
                                 "Altersgruppen", lst, "Fläche (m²)", CDbl(400 + Int(Rnd01() * 30) * 100), "Zustand", OidOf(tS, "Zustand", st), _
                                 "Adresse", Replace(sNames(i), "Spielplatz ", "") & " " & (2 + Int(Rnd01() * 40)), _
                                 "Zuständig", OidOf(tS, "Zuständig", CStr(people(i Mod 4)))))
    Next
    ' Spielgeräte: 4-7 je Platz, Prüfungen über das Jahr verteilt (einige überfällig)
    gTypes = Array("Schaukel", "Rutsche", "Klettergerät", "Wippe", "Sandspiel", "Karussell", "Seilbahn")
    gNames = Array(Array("Doppelschaukel", "Nestschaukel", "Kleinkindschaukel"), Array("Hangrutsche", "Turmrutsche", "Wellenrutsche"), _
                   Array("Kletterkombination", "Kletternetz", "Boulderwand"), Array("Federwippe", "Balkenwippe"), Array("Sandbaustelle", "Matschanlage"), _
                   Array("Drehkarussell", "Drehteller"), Array("Seilbahn"))
    For i = 0 To UBound(sp)
        For j = 1 To 4 + Int(Rnd01() * 4)
            k = Int(Rnd01() * (UBound(gTypes) + 1))
            last = d0 - Int(Rnd01() * 420)
            st = Pick(Array("Ohne Mängel", "Ohne Mängel", "Ohne Mängel", "Ohne Mängel", "Geringe Mängel", "Geringe Mängel", "Sofort beheben"))
            If i = 6 And j = 1 Then st = "Gesperrt"
            lst = ""
            If st <> "Ohne Mängel" Then
                lst = OidOf(tG, "Mängel", CStr(Pick(Array("Verschleiß", "Fallschutz", "Splitter", "Lockere Teile", "Vandalismus"))))
                If Rnd01() < 0.4 Then lst = ListAdd(lst, OidOf(tG, "Mängel", CStr(Pick(Array("Verschleiß", "Fallschutz", "Splitter")))))
            End If
            rid = AddRec(tG, Array("Gerät", Pick(gNames(k)), "Geräteart", OidOf(tG, "Geräteart", CStr(gTypes(k))), "Baujahr", CDbl(2004 + Int(Rnd01() * 21)), _
                                   "Letzte Prüfung", last, "Prüfintervall (Tage)", CDbl(IIf(gTypes(k) = "Seilbahn", 180, 365)), _
                                   "Prüfergebnis", OidOf(tG, "Prüfergebnis", st), "Mängel", lst))
            mData.SetLinks FidOf(tG, "Spielplatz"), rid, sp(i)
        Next
    Next
    ' Bäume: 2-5 je Platz
    Dim bArt As Variant, bNo As Long
    bArt = Array("Linde", "Ahorn", "Eiche", "Platane", "Kastanie", "Birke")
    bNo = 1000
    For i = 0 To UBound(sp)
        For j = 1 To 2 + Int(Rnd01() * 4)
            art = CStr(Pick(bArt)): bNo = bNo + 1 + Int(Rnd01() * 9)
            st = Pick(Array("Vital", "Vital", "Vital", "Leicht geschädigt", "Leicht geschädigt", "Stark geschädigt", "Abgängig"))
            rid = AddRec(tB, Array("Baum", art & " " & bNo, "Baumart", OidOf(tB, "Baumart", art), "Pflanzjahr", CDbl(1950 + Int(Rnd01() * 65)), _
                                   "Stammumfang (cm)", CDbl(60 + Int(Rnd01() * 260)), "Vitalität", OidOf(tB, "Vitalität", st), _
                                   "Letzte Baumkontrolle", d0 - Int(Rnd01() * 300), _
                                   "Maßnahme", OidOf(tB, "Maßnahme", Choose(1 + Int(Rnd01() * 3) + IIf(st = "Vital", 0, 1), "Keine", "Keine", "Totholz entfernen", IIf(st = "Abgängig", "Fällung prüfen", "Kronenpflege"))), _
                                   "Verkehrssicher", (st <> "Abgängig")))
            mData.SetLinks FidOf(tB, "Spielplatz"), rid, sp(i)
        Next
    Next
    ' Projekte mit Bürgerbeteiligung
    Dim pr As Variant
    pr = Array(Array("Neubau Wasserspielplatz Auenwiese", "Neubau", "Beteiligung", -60, 240, 180000, "Kinderbeteiligung;Bürgerversammlung;Online-Umfrage", 214, 6), _
               Array("Sanierung Fallschutz Lindenhof", "Sanierung", "Umsetzung", -30, 45, 38000, "", 0, 1), _
               Array("Inklusive Spielgeräte Am Stadtpark", "Erweiterung", "Planung", 10, 200, 95000, "Kinderbeteiligung;Ortsbeirat", 63, 0), _
               Array("Kinderbeteiligung Schillerplatz", "Neubau", "Beteiligung", -20, 150, 120000, "Kinderbeteiligung;Online-Umfrage", 148, 5), _
               Array("Baumersatzpflanzung Kastanienallee", "Pflanzung", "Fertig", -150, -40, 22000, "Ortsbeirat", 18, 10), _
               Array("Erweiterung Bolzplatz Am Sportfeld", "Erweiterung", "Idee", 90, 300, 60000, "", 0, 8), _
               Array("Schattenplätze Rosengarten", "Pflanzung", "Umsetzung", -45, 30, 31000, "Bürgerversammlung", 42, 3), _
               Array("Sanierung Sandflächen Nord", "Sanierung", "Planung", 30, 110, 27000, "", 0, 2))
    For i = 0 To UBound(pr)
        lst = ""
        For Each nm In Split(pr(i)(6), ";")
            If Len(nm) > 0 Then lst = ListAdd(lst, OidOf(tP, "Beteiligung", CStr(nm)))
        Next
        rid = AddRec(tP, Array("Projekt", pr(i)(0), "Art", OidOf(tP, "Art", CStr(pr(i)(1))), "Status", OidOf(tP, "Status", CStr(pr(i)(2))), _
                               "Beginn", d0 + pr(i)(3), "Ende", d0 + pr(i)(4), "Budget", CDbl(pr(i)(5)), "Beteiligung", lst, _
                               "Teilnehmende", IIf(pr(i)(7) > 0, CDbl(pr(i)(7)), Empty)))
        mData.SetLinks FidOf(tP, "Spielplatz"), rid, sp(pr(i)(8))
    Next
    mData.TxCommit
    mData.ClearUndo
    mCalc.RecalcAll

    ' Dashboards
    did = mDash.CreateDashboard("Spielplatz-Cockpit", "Prüfungen, Mängel und Bäume auf einen Blick")
    DemoElement did, "Spielplätze", "Kachel", tS, "ANZAHL", "", "", "", "1", ""
    DemoElement did, "Spielgeräte", "Kachel", tG, "ANZAHL", "", "", "", "1", ""
    DemoElement did, "Prüfung überfällig", "Kachel", tG, "ANZAHL", "", "", "", "1", "~" & FidOf(tG, "Prüfung überfällig") & "~JA~~"
    DemoElement did, "Offene Baumpflege", "Kachel", tB, "ANZAHL", "", "", "", "1", "~" & FidOf(tB, "Maßnahme") & "~OFFEN~~"
    DemoElement did, "Spielgeräte nach Prüfergebnis", "Balken", tG, "ANZAHL", "", FidOf(tG, "Prüfergebnis"), "", "2", ""
    DemoElement did, "Fällige Prüfungen je Monat", "Zeitverlauf", tG, "ANZAHL", "", FidOf(tG, "Nächste Prüfung"), FidOf(tG, "Nächste Prüfung"), "2", ""
    DemoElement did, "Bäume nach Vitalität", "Säulen", tB, "ANZAHL", "", FidOf(tB, "Vitalität"), "", "2", ""
    DemoElement did, "Mängel (Mehrfachnennung)", "Tabelle", tG, "ANZAHL", "", FidOf(tG, "Mängel"), "", "2", ""
    did = mDash.CreateDashboard("Projekte & Beteiligung", "Bauvorhaben, Budget und Bürgerbeteiligung")
    DemoElement did, "Projekte", "Kachel", tP, "ANZAHL", "", "", "", "1", ""
    DemoElement did, "Budget gesamt", "Kachel", tP, "SUMME", FidOf(tP, "Budget"), "", "", "1", ""
    DemoElement did, "Beteiligte Personen", "Kachel", tP, "SUMME", FidOf(tP, "Teilnehmende"), "", "", "1", ""
    DemoElement did, "Projekte nach Status", "Balken", tP, "ANZAHL", "", FidOf(tP, "Status"), "", "2", ""
    DemoElement did, "Beteiligungsformate", "Anteil", tP, "ANZAHL", "", FidOf(tP, "Beteiligung"), "", "2", ""

    ' Erfassungsdateien: Prüfdateien je Spielplatz (vorbelegt, nur eigene Einträge)
    mCapture.EnsureSheets
    Set m = New cMap
    cid = mCfg.NextId("C", 4)
    sel = ""
    For Each nm In Array("Letzte Prüfung", "Prüfergebnis", "Mängel", "Bemerkung")
        sel = ListAdd(sel, FidOf(tG, CStr(nm)))
    Next
    m.SetVal "ID", cid: m.SetVal "Name", "Jahresprüfung Spielgeräte": m.SetVal "Tabelle", tG: m.SetVal "Felder", sel
    m.SetVal "Modus", "Liste": m.SetVal "Titel", "Jahresprüfung Spielgeräte 2026"
    m.SetVal "Einleitung", "Bitte prüfen Sie die Geräte Ihres Spielplatzes, tragen Sie Datum und Ergebnis ein, kreuzen Sie Mängel an und senden Sie die Datei zurück."
    m.SetVal "Zeilen", "0": m.SetVal "Absender", "Ja": m.SetVal "AbsenderFeld", FidOf(tG, "Erfasst von"): m.SetVal "Schutz", "Ja": m.SetVal "Version", "1"
    m.SetVal "Erstellt", Now
    m.SetVal "Optionen", CapOpts(ListAdd(ListAdd(ListAdd(FidOf(tG, "Gerätenr."), FidOf(tG, "Gerät")), FidOf(tG, "Geräteart")), FidOf(tG, "Baujahr")), FidOf(tG, "Spielplatz"), "3", "Prüfer/in")
    mCfg.AddRow SH_CAP, m
    Set m = New cMap
    cid = mCfg.NextId("C", 4)
    sel = ""
    For Each nm In Array("Letzte Baumkontrolle", "Vitalität", "Maßnahme", "Verkehrssicher")
        sel = ListAdd(sel, FidOf(tB, CStr(nm)))
    Next
    m.SetVal "ID", cid: m.SetVal "Name", "Baumkontrolle": m.SetVal "Tabelle", tB: m.SetVal "Felder", sel
    m.SetVal "Modus", "Liste": m.SetVal "Titel", "Baumkontrolle Herbst 2026"
    m.SetVal "Einleitung", "Bitte kontrollieren Sie die Bäume am Spielplatz und tragen Sie Vitalität und Maßnahme ein."
    m.SetVal "Zeilen", "0": m.SetVal "Absender", "Ja": m.SetVal "AbsenderFeld", FidOf(tB, "Erfasst von"): m.SetVal "Schutz", "Ja": m.SetVal "Version", "1"
    m.SetVal "Erstellt", Now
    m.SetVal "Optionen", CapOpts(ListAdd(ListAdd(FidOf(tB, "Baum"), FidOf(tB, "Baumart")), FidOf(tB, "Stammumfang (cm)")), FidOf(tB, "Spielplatz"), "2", "Kontrolleur/in")
    mCfg.AddRow SH_CAP, m
    mLog.Add "Beispieldaten", "", "", "", "Beispiel Grünflächenamt erzeugt"
    mLog.Flush
    Application.EnableEvents = True
End Sub

' Optionen einer Prüfdatei (wie der Assistent: über KVSet, Feldlisten werden maskiert)
Private Function CapOpts(ByVal anz As String, ByVal splitFid As String, ByVal newRows As String, ByVal absLabel As String) As String
    Dim o As String
    o = KVSet(o, "art", "pruefen")
    o = KVSet(o, "anzeige", anz)
    o = KVSet(o, "teilen", splitFid)
    o = KVSet(o, "neuzeilen", newRows)
    o = KVSet(o, "abslabel", absLabel)
    o = KVSet(o, "absbreite", "Breit")
    o = KVSet(o, "farbe", "Akzentfarbe")
    o = KVSet(o, "anleitung", "1")
    CapOpts = o
End Function

Private Sub DemoElement(ByVal did As String, ByVal ttl As String, ByVal typ As String, ByVal tid As String, ByVal measure As String, ByVal fid As String, _
                        ByVal gfid As String, ByVal dfid As String, ByVal sz As String, ByVal filt As String, Optional ByVal extraOpt As String = "")
    Dim m As New cMap
    m.SetVal "Titel", ttl
    m.SetVal "Typ", typ
    m.SetVal "Tabelle", tid
    m.SetVal "Kennzahl", measure
    m.SetVal "Feld", fid
    m.SetVal "Gruppierung", gfid
    m.SetVal "Datumsstufe", "Monat"
    m.SetVal "Datumsfeld", dfid
    m.SetVal "Groesse", sz
    m.SetVal "Basis", "ALLE"
    m.SetVal "Optionen", "sort=" & IIf(typ = "Zeitverlauf" Or typ = "Fläche" Or typ = "Netz", "name", "wert") & ";top=12" & IIf(Len(extraOpt) > 0, ";" & extraOpt, "")
    m.SetVal "Filter", Replace(filt, "~", "|")
    mDash.CreateElement did, m
End Sub

'==============================================================================
' Belastungsdaten (getrennte Pruefkopie): 10 Tabellen, 20.000 Datensaetze, 40.000 Zuordnungen
'==============================================================================
Public Sub Build_Stress()
    Dim tA As String, tB As String, tid As String, i As Long, j As Long, k As Long, f As cField, t As cTableData, rows() As Variant, baseNo As Long
    Dim props As cMap, linkF As String, rels() As Variant, n As Long, statusF As String, oids As Collection, bIds() As String, nB As Long, c As Long
    Dim types As Variant, ty As String, words As Variant
    mRnd = 4242
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mApp.gStarted = False
    words = Array("Alpha", "Beta", "Gamma", "Delta", "Nord", "Süd", "Ost", "West", "Projekt", "Auftrag", "Gerät", "Termin", "Bericht")
    types = Array(FT_TEXT, FT_ZAHL, FT_GELD, FT_DATUM, FT_AUSWAHL, FT_JANEIN, FT_PROZENT, FT_LTEXT, FT_STATUS, FT_MEHRFACH)
    StepLog "", True
    ' Tabelle B (Ziel der Zuordnungen), 9 kleinere Tabellen
    For k = 1 To 9
        tid = mStruct.CreateTable("Lasttabelle " & k, "Belastungstest", &HE8FD&)
        mStruct.AddField tid, "Name", FT_TEXT
        mStruct.SetLabelField tid, mCfg.FieldIdByName(tid, "Name")
        For j = 1 To 9
            ty = types(j Mod (UBound(types) + 1))
            Set props = New cMap
            If ty = FT_AUSWAHL Or ty = FT_STATUS Or ty = FT_MEHRFACH Then props.SetVal "Optionen", "Eins:blau,Zwei:gruen,Drei:gelb,Vier:rot:x"
            mStruct.AddField tid, "Feld " & j, ty, props
        Next
        mStruct.CreateDefaultView tid
        If k = 1 Then tB = tid
        FillTable tid, IIf(k = 1, 1112, 1111)
        StepLog "Lasttabelle " & k & " angelegt und gefüllt"
    Next
    ' Tabelle A mit 50 Feldern und 10.000 Datensaetzen
    tA = mStruct.CreateTable("Lasttabelle Groß", "Belastungstest: 50 Felder, 10.000 Datensätze", &HE9D5&)
    mStruct.AddField tA, "Name", FT_TEXT
    mStruct.SetLabelField tA, mCfg.FieldIdByName(tA, "Name")
    For j = 1 To 47
        ty = types(j Mod (UBound(types) + 1))
        Set props = New cMap
        If ty = FT_AUSWAHL Or ty = FT_STATUS Or ty = FT_MEHRFACH Then props.SetVal "Optionen", "Eins:blau,Zwei:gruen,Drei:gelb,Vier:rot:x"
        mStruct.AddField tA, "Feld " & j, ty, props
    Next
    Set props = New cMap
    props.SetVal "Zieltabelle", tB
    props.SetVal "Mehrfach", "Ja"
    props.SetVal "Gegenname", "Große Einträge"
    linkF = mStruct.AddField(tA, "Verknüpft", FT_LINK, props)
    Set props = New cMap
    props.SetVal "Formel", "ROLLUP|F:" & linkF & "|K:ANZAHL"
    props.SetVal "Ergebnistyp", FT_ZAHL
    props.SetVal "Beschreibung", "Anzahl verknüpfter Einträge"
    mStruct.AddField tA, "Anzahl verknüpft", FT_FORMEL, props
    mStruct.CreateDefaultView tA
    mStruct.CreateView tA, "Board", VT_BOARD, "gruppe=" & mCfg.FieldIdByName(tA, "Feld 8") & ";karte=" & mCfg.FieldIdByName(tA, "Feld 1")
    mStruct.CreateView tA, "Kalender", VT_CAL, "datum=" & mCfg.FieldIdByName(tA, "Feld 3")
    mStruct.RebuildLists
    StepLog "Große Tabelle angelegt"
    FillTable tA, 10000
    StepLog "10.000 Datensätze eingefügt"
    ' 40.000 Zuordnungen (je 4)
    Set t = mData.TData(tB)
    nB = t.RowCount
    ReDim bIds(1 To nB)
    For i = 1 To nB: bIds(i) = t.IdAt(i): Next
    Set t = mData.TData(tA)
    ReDim rels(1 To t.RowCount * 4, 1 To 4)
    For i = 1 To t.RowCount
        For j = 0 To 3
            n = n + 1
            rels(n, 1) = linkF: rels(n, 2) = t.IdAt(i): rels(n, 3) = bIds(1 + ((i * 7 + j * 131) Mod nB)): rels(n, 4) = Now
        Next
    Next
    StepLog "Zuordnungsarray vorbereitet"
    mData.TxBegin "Lastzuordnungen"
    mData.RelAddBulk rels
    StepLog "40.000 Zuordnungen geschrieben"
    mData.TxCommit
    StepLog "Transaktion abgeschlossen"
    mData.ClearUndo
    mCalc.RecalcAll
    StepLog "Alle Berechnungen neu"
    Application.EnableEvents = True
End Sub

' Zeitprotokoll des Lasttest-Aufbaus (Datei Lasttest_Aufbau.txt im Ordner der Arbeitsmappe bzw. TEMP)
Private Sub StepLog(ByVal what As String, Optional ByVal reset As Boolean = False)
    Static t0 As Double, tl As Double
    Dim h As Integer, p As String
    If reset Then t0 = Timer: tl = t0
    p = Environ$("TEMP") & "\Lasttest_Aufbau.txt"
    h = FreeFile
    If reset Then Open p For Output As #h Else Open p For Append As #h
    If Len(what) > 0 Then Print #h, Format$(Timer - tl, "0.0") & " s" & vbTab & "(gesamt " & Format$(Timer - t0, "0") & " s)" & vbTab & what
    Close #h
    tl = Timer
End Sub

Private Sub FillTable(ByVal tid As String, ByVal count As Long)
    Dim t As cTableData, rows() As Variant, i As Long, f As cField, c As Long, baseNo As Long, oids As Collection, v As Variant
    Set t = mData.TData(tid)
    ReDim rows(1 To count, 1 To t.ColCount)
    baseNo = mCfg.ReserveIds("R", count)
    For i = 1 To count
        rows(i, 1) = "R" & Format$(baseNo + i - 1, "0000000")
        rows(i, 2) = Now: rows(i, 3) = Now: rows(i, 4) = "build"
        For Each f In mCfg.Fields(tid)
            c = t.ColOf(f.ID)
            Select Case f.FType
                Case FT_TEXT: rows(i, c) = "Eintrag " & i & " " & Choose(1 + (i Mod 5), "Nord", "Süd", "Ost", "West", "Mitte")
                Case FT_LTEXT: If i Mod 3 = 0 Then rows(i, c) = "Längere Notiz zum Eintrag " & i & " mit etwas Text für die Suche."
                Case FT_ZAHL: rows(i, c) = CDbl(i Mod 997)
                Case FT_GELD: rows(i, c) = Round((i Mod 500) * 3.75, 2)
                Case FT_PROZENT: rows(i, c) = (i Mod 100) / 100
                Case FT_DATUM: rows(i, c) = DateSerial(2026, 1, 1) + (i Mod 700)
                Case FT_JANEIN: rows(i, c) = (i Mod 2 = 0)
                Case FT_AUSWAHL, FT_STATUS
                    Set oids = mCfg.OptionIds(f.ID)
                    rows(i, c) = oids(1 + (i Mod oids.Count))
                Case FT_MEHRFACH
                    Set oids = mCfg.OptionIds(f.ID)
                    rows(i, c) = oids(1 + (i Mod oids.Count)) & IIf(i Mod 3 = 0, ";" & oids(1 + ((i + 1) Mod oids.Count)), "")
            End Select
        Next
    Next
    mData.TxBegin "Lastdaten"
    mData.InsertRecordsBulk tid, rows
    mData.TxCommit
End Sub
