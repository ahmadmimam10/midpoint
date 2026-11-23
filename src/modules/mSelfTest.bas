Attribute VB_Name = "mSelfTest"
'==============================================================================
' mSelfTest - Automatisierte Pruefungen fuer die Testmatrix (nur Entwicklung).
' Laeuft in echter Excel-Laufzeit gegen eine KOPIE der Demodatei und liefert
' je Pruefung "BESTANDEN"/"NICHT BESTANDEN" mit Detail. Veraendert Daten!
'==============================================================================
Option Explicit

Private mOut As String
Private mPass As Long
Private mFail As Long

Private Function NzD(ByVal v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then NzD = CDbl(v)
End Function

Private Sub Check(ByVal id As String, ByVal desc As String, ByVal ok As Boolean, Optional ByVal detail As String = "")
    If ok Then mPass = mPass + 1 Else mFail = mFail + 1
    mOut = mOut & id & vbTab & IIf(ok, "BESTANDEN", "NICHT BESTANDEN") & vbTab & desc & IIf(Len(detail) > 0, ": " & detail, "") & vbLf
End Sub

Private Function Fid(ByVal tid As String, ByVal nm As String) As String
    Fid = mCfg.FieldIdByName(tid, nm)
End Function

Private Function FirstActive(ByVal tid As String, Optional ByVal skip As Long = 0) As String
    Dim t As cTableData, r As Long, n As Long
    Set t = mData.TData(tid)
    For r = 1 To t.RowCount
        If Not t.IsArchived(r) Then
            If n = skip Then FirstActive = t.IdAt(r): Exit Function
            n = n + 1
        End If
    Next
End Function

Public Function RunSelfTest() As String
    Dim tA As String, tP As String, f As cField, v As Variant, msg As String, ok As Boolean, vals As cMap, errs As cMap
    Dim rid As String, p1 As String, p2 As String, back As String, deps As Collection, path As String, nt As String
    Dim t As cTableData, t2 As cTableData, r As Long, seen As cMap, dupId As Boolean, s As String, vid As String, c As Variant, found As Boolean
    Dim oldName As String, cnt0 As Long, i As Long, startT As Double
    mOut = "": mPass = 0: mFail = 0
    On Error GoTo crash
    mApp.gStarted = True
    tA = mCfg.TableIdByName("Aufgaben")
    tP = mCfg.TableIdByName("Projekte")
    If Len(tA) = 0 Or Len(tP) = 0 Or mCfg.SysGet("Demo", "") <> "Ja" Then RunSelfTest = "Nur in einer Kopie der Demodatei ausführbar": Exit Function

    '--- T03: Validierung ------------------------------------------------------
    Set f = mCfg.Field(Fid(tA, "Fortschritt"))
    ok = mValid.ParseInput(f, "150", v, msg)
    If ok Then ok = mValid.CheckValue(f, v, msg)
    Check "T03a", "Prozent über Maximum (150 %) wird abgelehnt", Not ok, msg
    Set f = mCfg.Field(Fid(tA, "Frist"))
    msg = ""
    ok = mValid.ParseInput(f, "31.02.2026", v, msg)
    Check "T03b", "Ungültiges Datum 31.02.2026 wird abgelehnt", Not ok, msg
    Set vals = New cMap
    vals.SetVal Fid(tA, "Aufgabe"), Empty
    ok = mValid.ValidateRecord(tA, "", vals, New cMap, errs)
    Check "T03c", "Fehlendes Pflichtfeld verhindert das Speichern", (Not ok) And errs.Exists(Fid(tA, "Aufgabe")), IIf(ok, "", CStr(errs.GetVal(Fid(tA, "Aufgabe"), "")))
    Set vals = New cMap
    vals.SetVal Fid(tA, "Aufgabe"), "Test"
    vals.SetVal Fid(tA, "Status"), mCfg.OptIdByName(Fid(tA, "Status"), "Offen")
    vals.SetVal Fid(tA, "Beginn"), DateSerial(2026, 3, 20)
    vals.SetVal Fid(tA, "Frist"), DateSerial(2026, 3, 10)
    ok = mValid.ValidateRecord(tA, "", vals, New cMap, errs)
    Check "T03d", "Frist vor Beginn wird abgelehnt (Regel „nicht vor“)", Not ok, IIf(ok, "", Join(errs.Items, " "))

    '--- Neue Eigenschaftsarten ----------------------------------------------------
    Check "N01", "E-Mail-Format: „abc“ abgelehnt, „a.b@amt.de“ akzeptiert", (Not mValid.KindFormatOk(FT_EMAIL, "abc", msg)) And mValid.KindFormatOk(FT_EMAIL, "a.b@amt.de", msg)
    Check "N02", "Telefon „+49 (0)611 123-45“ akzeptiert, „12ab“ abgelehnt", mValid.KindFormatOk(FT_TELEFON, "+49 (0)611 123-45", msg) And Not mValid.KindFormatOk(FT_TELEFON, "12ab", msg)
    Set t = mData.TData(tA)
    Set seen = New cMap
    ok = True
    For r = 1 To t.RowCount
        s = NzS(t.Val_(r, Fid(tA, "Nr.")))
        If Len(s) = 0 Or seen.Exists(s) Then ok = False
        seen.SetVal s, True
    Next
    Check "N03", "ID-Eigenschaft „Nr.“: alle " & t.RowCount & " Einträge haben eine eindeutige Nummer", ok, "z. B. " & NzS(t.Val_(1, Fid(tA, "Nr.")))
    rid = FirstActive(tA, 2)
    mForms.RunButton tA, rid, Fid(tA, "Als erledigt markieren")
    ok = (CStr(mData.GetValue(tA, rid, Fid(tA, "Status"))) = mCfg.OptIdByName(Fid(tA, "Status"), "Erledigt")) And (Abs(CDbl(mData.GetValue(tA, rid, Fid(tA, "Fortschritt"))) - 1) < 0.0001)
    Check "N04", "Button „Als erledigt markieren“ setzt Status und Fortschritt (mit Protokoll und Rückgängig)", ok And mData.UndoAvailable
    Set f = mCfg.Field(Fid(tA, "Nr."))
    msg = ""
    ok = mValid.ParseInput(f, "X-1", v, msg)
    Check "N05", "Automatische Eigenschaft (ID) kann nicht von Hand überschrieben werden (alle Eingabewege)", Not ok, msg
    Set f = mCfg.Field(Fid(tA, "Wichtig"))
    Check "N06", "Kontrollkästchen ist nie leer (Standard „nicht angehakt“)", Not IsBlank(mValid.DefaultValue(f)) Or (f.FType = FT_JANEIN And Not f.AllowEmpty)

    '--- T05: Umbenennen / Umordnen ------------------------------------------------
    Set f = mCfg.Field(Fid(tA, "Priorität"))
    oldName = f.Name
    Set vals = New cMap
    vals.SetVal "Name", "Dringlichkeit"
    mStruct.UpdateField f.ID, vals
    mCfg.InvalidateAll
    found = False
    For Each c In mCfg.ViewIds(tA)
        If InStr(1, mCfg.ViewGet(CStr(c), "Optionen"), f.ID) > 0 Then found = True
    Next
    Check "T05a", "Umbenennen: Ansichten verweisen weiter über die Feld-ID (Board-Karte)", (mCfg.FieldName(f.ID) = "Dringlichkeit") And found
    vals.SetVal "Name", oldName
    mStruct.UpdateField f.ID, vals
    mStruct.MoveField f.ID, 1
    mStruct.MoveField f.ID, -1
    Check "T05b", "Umordnen und Zurückbenennen ohne Datenverlust", mCfg.FieldName(f.ID) = oldName And Not IsBlank(mData.GetValue(tA, FirstActive(tA), f.ID))

    '--- T06: Abhaengigkeiten --------------------------------------------------------
    Set deps = mStruct.FieldDependencies(Fid(tA, "Frist"))
    Check "T06a", "Löschen von „Frist“ wird blockiert (von Berechnung „Tage bis Frist“ verwendet)", mStruct.HasBlocking(deps), deps.Count & " Abhängigkeit(en)"
    Dim newVals As Variant, createOpts As Collection, conflicts As Collection
    Set conflicts = mStruct.ConversionPreview(Fid(tA, "Aufgabe"), FT_ZAHL, newVals, createOpts)
    Check "T06b", "Inkompatibler Typwechsel Text nach Zahl zeigt Konfliktvorschau (keine stille Umwandlung)", conflicts.Count > 0, conflicts.Count & " Konflikte"
    Check "T06c", "Verknüpfungen und Berechnungen können nicht per Typwechsel umgewandelt werden", Not mStruct.CanConvert(FT_LINK, FT_TEXT) And Not mStruct.CanConvert(FT_TEXT, FT_FORMEL)

    '--- T07: Verknuepfungen ---------------------------------------------------------
    rid = FirstActive(tA, 3)
    p1 = FirstActive(tP, 0): p2 = FirstActive(tP, 1)
    back = mCfg.Field(Fid(tA, "Projekt")).PairField
    mData.TxBegin "Selbsttest Verknüpfung"
    mData.SetLinks Fid(tA, "Projekt"), rid, p1
    mData.TxCommit
    ok = ListHas(mData.LinkTargets(back, p1), rid)
    mData.TxBegin "Selbsttest Verknüpfung 2"
    mData.SetLinks Fid(tA, "Projekt"), rid, p2
    mData.TxCommit
    ok = ok And ListHas(mData.LinkTargets(back, p2), rid) And Not ListHas(mData.LinkTargets(back, p1), rid)
    Check "T07a", "Einfachverknüpfung: Rückverweis wandert mit (kein doppelter Eintrag)", ok
    Check "T07b", "Verknüpfung speichert interne IDs", Left$(mData.LinkTargets(Fid(tA, "Projekt"), rid), 1) = "R"

    '--- T11: Berechnungen -----------------------------------------------------------
    Set f = mCfg.Field(Fid(tA, "Tage bis Frist"))
    ok = mCalc.HasCycle(f.ID, "RECHNEN|F:" & f.ID & "|K:+|F:" & Fid(tA, "Aufwand (h)"), path)
    Check "T11a", "Zirkelbezug wird erkannt und abgelehnt", ok, path
    mData.TxBegin "Selbsttest leeres Projekt"
    rid = mData.NewRecordId()
    v = mData.NewRowArray(tP, rid)
    v(mData.TData(tP).ColOf(Fid(tP, "Projekt"))) = "Selbsttest ohne Aufgaben"
    mData.InsertRecord tP, v
    mData.TxCommit
    mCalc.RecalcRecord tP, rid
    v = mData.GetValue(tP, rid, Fid(tP, "Erledigt in %"))
    Check "T11b", "Division durch null (Projekt ohne Aufgaben) ergibt klaren Fehler/leer statt Absturz", IsErrVal(v) Or IsBlank(v), mFmt.DisplayText(mCfg.Field(Fid(tP, "Erledigt in %")), v, rid)
    v = mData.GetValue(tP, p2, Fid(tP, "Aufgaben gesamt"))
    Check "T11c", "Rollup zählt verknüpfte Aufgaben", Val(NzS(v)) >= 1, "Aufgaben gesamt = " & NzS(v)

    '--- T16: Archiv und Rueckgaengig -------------------------------------------------
    rid = FirstActive(tA, 5)
    mForms.SetArchived tA, rid, True
    ok = mData.IsArchivedRec(tA, rid)
    mApp.BeginWork
    mData.UndoLast
    mApp.EndWork
    Check "T16", "Archivieren und Rücknahme stellen den Zustand wieder her", ok And Not mData.IsArchivedRec(tA, rid)

    '--- T23: Duplizieren --------------------------------------------------------------
    cnt0 = mData.TData(tP).RowCount
    startT = Timer
    nt = mStruct.DuplicateTable(tP, "Projekte Kopie", True)
    Set t = mData.TData(tP): Set t2 = mData.TData(nt)
    ok = (Len(nt) > 0 And nt <> tP And t2.RowCount = t.RowCount And t.RowCount = cnt0)
    If ok Then
        For r = 1 To t2.RowCount
            If t.RowOf(t2.IdAt(r)) > 0 Then dupId = True
        Next
    End If
    Check "T23", "Struktur mit Daten duplizieren: neue IDs, Original unverändert", ok And Not dupId, t2.RowCount & " Einträge kopiert in " & Format$(Timer - startT, "0.0") & " s"

    '--- T12: n:m-Auswertung (eindeutige Zaehlung) -----------------------------------
    Dim did As String, eid As String, props As cMap, res As cMap
    did = mDash.CreateDashboard("Selbsttest", "")
    Set props = New cMap
    props.SetVal "Titel", "Aufgaben je Ressource"
    props.SetVal "Typ", "Balken"
    props.SetVal "Tabelle", tA
    props.SetVal "Kennzahl", "ANZAHL"
    props.SetVal "Gruppierung", Fid(tA, "Benötigte Ressourcen")
    eid = mDash.CreateElement(did, props)
    Set res = mDash.Compute(eid, did)
    s = ""
    For Each c In res.Keys
        s = s & CStr(c) & ";"
    Next
    Check "T12", "n:m-Gruppierung liefert Gruppenwerte mit Hinweis auf Mehrfachzuordnung", res.Count > 0, Left$(s, 120)

    '--- T02: Tabelle mit allen Eigenschaftsarten -----------------------------------
    Dim codes As Variant, nt2 As String, pr As cMap, linkFid As String, chkFid As String, made As Long, fr As frmRecord, noFilter As String, fx As cField
    codes = FieldTypeCodes()
    nt2 = mStruct.CreateTable("Selbsttest alle Arten", "Alle Eigenschaftsarten", &HE8FD&)
    mStruct.SetLabelField nt2, mStruct.AddField(nt2, "Name", FT_TEXT)
    For i = 0 To UBound(codes)
        Set pr = New cMap
        Select Case CStr(codes(i))
            Case FT_AUSWAHL, FT_STATUS, FT_MEHRFACH, FT_PERSON: pr.SetVal "Optionen", "Eins:blau,Zwei:gruen,Drei:rot:x"
            Case FT_LINK: pr.SetVal "Zieltabelle", tP: pr.SetVal "Mehrfach", "Ja": pr.SetVal "Gegenname", "Selbsttest Arten"
            Case FT_FORMEL: pr.SetVal "Formel", "RECHNEN|N:1|K:+|N:2": pr.SetVal "Ergebnistyp", FT_ZAHL
            Case FT_ROLLUP: pr.SetVal "Formel", "ROLLUP|F:" & linkFid & "|K:ANZAHL": pr.SetVal "Ergebnistyp", FT_ZAHL
            Case FT_BUTTON: pr.SetVal "Formel", "AKTION|F:" & chkFid & "|umschalten"
        End Select
        s = mStruct.AddField(nt2, "Art " & FieldTypeName(CStr(codes(i))), CStr(codes(i)), pr)
        If CStr(codes(i)) = FT_LINK Then linkFid = s
        If CStr(codes(i)) = FT_KONTROLL Then chkFid = s
        made = made + 1
    Next
    mStruct.CreateDefaultView nt2
    mStruct.RebuildLists
    Set fr = New frmRecord
    fr.PrepareOnly nt2, ""
    Unload fr
    For Each fx In mCfg.Fields(nt2)
        If fx.FType <> FT_BACKLINK Then
            If ArrLen(mQuery.OpsFor(mQuery.OpKind(fx))) = 0 Then noFilter = noFilter & fx.Name & " "
        End If
    Next
    Check "T02", "Neue Tabelle mit allen " & made & " Eigenschaftsarten: Eingabemaske, Filterangebote und Ansicht entstehen ohne Codeänderung", _
          made = UBound(codes) + 1 And Len(noFilter) = 0, IIf(Len(noFilter) > 0, "ohne Filter: " & noFilter, mCfg.Fields(nt2).Count & " Eigenschaften")

    '--- T04: Mehrzellen-Einfuegen in der Tabellenansicht ---------------------------
    Dim wsT As Worksheet, colA As Long, ridA1 As String, ridA2 As String, v1 As Variant, v2 As Variant, cols As Collection
    vid = mCfg.ViewIds(tA)(0)
    mNav.State(vid).SortSpec = ""
    mNav.GoTable tA, vid
    Set wsT = mRender.Sheet(SH_TABLE)
    Set cols = mViewTable.ViewColumns(vid)
    For i = 1 To cols.Count
        If cols(i)(0) = Fid(tA, "Aufwand (h)") Then colA = COL_F1 + i - 1
    Next
    ridA1 = mViewTable.RidAtRow(R_DATA): ridA2 = mViewTable.RidAtRow(R_DATA + 1)
    v1 = mData.GetValue(tA, ridA1, Fid(tA, "Aufwand (h)")): v2 = mData.GetValue(tA, ridA2, Fid(tA, "Aufwand (h)"))
    mUI.gQuiet = True: mUI.gQuietLog = ""
    Application.EnableEvents = True
    wsT.Range(wsT.Cells(R_DATA, colA), wsT.Cells(R_DATA + 1, colA)).Value = Application.Transpose(Array("7", "abc"))
    ok = (mData.GetValue(tA, ridA1, Fid(tA, "Aufwand (h)")) = v1) And (mData.GetValue(tA, ridA2, Fid(tA, "Aufwand (h)")) = v2) And InStr(mUI.gQuietLog, "Bericht") > 0
    Check "T04a", "Einfügen mit einem ungültigen Wert: nichts wird übernommen, Fehlerliste erscheint", ok
    mUI.gQuietLog = ""
    Set wsT = mRender.Sheet(SH_TABLE)
    wsT.Range(wsT.Cells(R_DATA, colA), wsT.Cells(R_DATA + 1, colA)).Value = Application.Transpose(Array("7", "8,5"))
    ok = (CDbl(mData.GetValue(tA, ridA1, Fid(tA, "Aufwand (h)"))) = 7) And (Abs(CDbl(mData.GetValue(tA, ridA2, Fid(tA, "Aufwand (h)"))) - 8.5) < 0.0001)
    Check "T04b", "Einfügen gültiger Werte in zwei Zeilen wird vollständig übernommen", ok
    Set wsT = mRender.Sheet(SH_TABLE)
    Err.Clear
    On Error Resume Next
    wsT.Cells(R_DATA, COL_ID).Value = "R9999999"
    i = Err.Number
    On Error GoTo crash
    Check "T04c", "Systemspalte (interne ID) ist gegen Überschreiben geschützt", (i <> 0 Or NzS(mRender.Sheet(SH_TABLE).Cells(R_DATA, COL_ID).Value) = ridA1) And mData.RecordExists(tA, ridA1)
    mUI.gQuiet = False

    '--- Suche: Ergebnisse stimmen mit Nachzaehlung ueberein ------------------------
    Dim exp1 As Long, exp2 As Long, exp3 As Long, got1 As Long, got2 As Long, got3 As Long, pid As String
    Set t = mData.TData(tA)
    pid = mData.TData(tP).IdAt(1)
    For r = 1 To t.RowCount
        If Not t.IsArchived(r) Then
            If InStr(1, NzS(t.Val_(r, Fid(tA, "Aufgabe"))), "Möblierung", vbTextCompare) > 0 Then exp1 = exp1 + 1
            If CStr(t.Val_(r, Fid(tA, "Priorität"))) = mCfg.OptIdByName(Fid(tA, "Priorität"), "Hoch") Then exp2 = exp2 + 1
            If ListHas(mData.LinkTargets(Fid(tA, "Projekt"), t.IdAt(r)), pid) Then exp3 = exp3 + 1
        End If
    Next
    vid = mCfg.ViewIds(tA)(0)
    mNav.ResetState vid
    mNav.State(vid).Search = "möblierung": mViewTable.ViewRows vid, got1
    mNav.State(vid).Search = "hoch": mViewTable.ViewRows vid, got2
    mNav.State(vid).Search = LCase$(mData.RecordLabel(tP, pid)): mViewTable.ViewRows vid, got3
    mNav.State(vid).Search = ""
    Check "S01", "Suche findet Text (" & got1 & "/" & exp1 & "), Auswahlwert „Hoch“ (" & got2 & "/" & exp2 & ") und verknüpften Projektnamen (" & got3 & "/" & exp3 & ")", _
          got1 = exp1 And got2 >= exp2 And got3 >= exp3 And exp1 > 0 And exp3 > 0

    '--- T08: zwei Ansichten mit unterschiedlichen Filtern ---------------------------
    Dim vA As String, vB As String, nA As Long, nB As Long, recs0 As Long
    recs0 = mData.TData(tA).RowCount
    vA = mStruct.CreateView(tA, "Selbsttest offen", VT_TABLE, "", "", "|" & Fid(tA, "Status") & "|OFFEN||")
    vB = mStruct.CreateView(tA, "Selbsttest hoch", VT_TABLE, "", "", "|" & Fid(tA, "Priorität") & "|IST|" & mCfg.OptIdByName(Fid(tA, "Priorität"), "Hoch") & "|")
    mNav.ResetState vA: mNav.ResetState vB
    mViewTable.ViewRows vA, nA
    mViewTable.ViewRows vB, nB
    Check "T08", "Zwei gespeicherte Ansichten mit eigenen Filtern (" & nA & " bzw. " & nB & " Einträge); Daten werden nicht dupliziert", _
          nA > 0 And nB > 0 And nA <> nB And mData.TData(tA).RowCount = recs0 And mCfg.ViewGet(vA, "Filter") <> mCfg.ViewGet(vB, "Filter"), "Fortbestand nach Speichern/Wiederöffnen: siehe T17"

    '--- T22: gezielter Laufzeitfehler und andere Arbeitsmappe ---------------------
    Dim calc0 As Long, wbX As Workbook, gridOk As Boolean
    calc0 = Application.Calculation
    mUI.gQuiet = True: mUI.gQuietLog = ""
    ForcedError
    ok = (mApp.WorkDepth = 0) And Application.EnableEvents And Application.Calculation = calc0 And InStr(mUI.gQuietLog, "Fehler") > 0 And mData.TxDepth = 0 And _
         NzS(mData.GetValue(tA, mData.TData(tA).IdAt(1), Fid(tA, "Aufgabe"))) <> "WIRD ZURÜCKGENOMMEN"
    mUI.gQuiet = False
    Set wbX = Application.Workbooks.Add
    wbX.Activate
    gridOk = ActiveWindow.DisplayGridlines And ActiveWindow.DisplayHeadings And ActiveWindow.DisplayWorkbookTabs
    wbX.Close False
    ThisWorkbook.Activate
    Check "T22", "Laufzeitfehler: Meldung statt Debugger, Transaktion zurückgenommen, Excel-Zustände wiederhergestellt; andere Arbeitsmappe unverändert", ok And gridOk

    '--- X: Formelspalten mit Excel-Formeln (1.3.0) ------------------------------------
    Dim xf As String, xdef As String, xmsg As String, aw As Variant, xv As Variant, nOk As Long, nBad As Long, xf2 As String, xr As String
    xdef = mCalc.XlFromDisplay(tA, "=[Aufwand (h)]*2", xmsg)
    Set pr = New cMap
    pr.SetVal "Formel", xdef: pr.SetVal "Ergebnistyp", FT_ZAHL: pr.SetVal "Dezimalstellen", 1
    xf = mStruct.AddField(tA, "Selbsttest Formel", FT_FORMEL, pr)
    mCalc.RecalcAll
    Set t = mData.TData(tA)
    For r = 1 To t.RowCount
        aw = t.Val_(r, Fid(tA, "Aufwand (h)"))
        If Not IsBlank(aw) Then
            xv = t.Val_(r, xf)
            If VarType(xv) = vbDouble Then
                If Abs(xv - 2 * CDbl(aw)) < 0.0000001 Then nOk = nOk + 1 Else nBad = nBad + 1
            Else
                nBad = nBad + 1
            End If
        End If
    Next
    Check "X01", "Excel-Formel =[Aufwand (h)]*2 wird für alle Einträge in einem Rechenvorgang berechnet", Len(xmsg) = 0 And nOk > 0 And nBad = 0, nOk & " richtig, " & nBad & " falsch"
    xdef = mCalc.XlFromDisplay(tA, "=WENN([Aufwand (h)]>5;""viel"";""wenig"")&"" (""&RUNDEN([Aufwand (h)];0)&"")""", xmsg)
    Set pr = New cMap
    pr.SetVal "Formel", xdef: pr.SetVal "Ergebnistyp", FT_TEXT
    xf2 = mStruct.AddField(tA, "Selbsttest Formel Text", FT_FORMEL, pr)
    mCalc.RecalcAll
    xr = FirstActive(tA)
    xv = mData.GetValue(tA, xr, xf2)
    aw = mData.GetValue(tA, xr, Fid(tA, "Aufwand (h)"))
    Check "X02", "Deutsche Funktionsnamen (WENN, RUNDEN), Semikolon und Textverkettung", VarType(xv) = vbString And (Left$(NzS(xv), 4) = "viel" Or Left$(NzS(xv), 5) = "wenig"), NzS(aw) & " " & ChrW(&H2192) & " " & NzS(xv)
    ok = Len(mCalc.XlCheck(mCalc.XlFromDisplay(tA, "=WEBDIENST(""http://example.com"")", xmsg))) > 0
    ok = ok And Len(mCalc.XlCheck(mCalc.XlFromDisplay(tA, "=A1*2", xmsg))) > 0
    ok = ok And Len(mCalc.XlCheck(mCalc.XlFromDisplay(tA, "=Start!B2", xmsg))) > 0
    ok = ok And Len(mCalc.XlCheck(mCalc.XlFromDisplay(tA, "=SUMME([Aufwand (h)];1", xmsg))) > 0
    mCalc.XlFromDisplay tA, "=[Gibt es nicht]*2", xmsg
    Check "X03", "Gesperrt bzw. abgelehnt: Netzwerkfunktion, Zellbezug, Blattbezug, Syntaxfehler, unbekannte Spalte", ok And Len(xmsg) > 0, xmsg
    ' Text, der wie eine Formel aussieht, bleibt Text; Aenderung rechnet abhaengige Formel neu
    Dim vals2 As cMap
    xdef = mCalc.XlFromDisplay(tA, "=[Aufgabe]&""!""", xmsg)
    Set pr = New cMap
    pr.SetVal "Formel", xdef: pr.SetVal "Ergebnistyp", FT_TEXT
    xf2 = mStruct.AddField(tA, "Selbsttest Formel Kette", FT_FORMEL, pr)
    mCalc.RecalcAll
    Set vals2 = New cMap
    vals2.SetVal Fid(tA, "Aufgabe"), "=1+1"
    mApp.BeginWork
    mData.TxBegin "Selbsttest Formeltext"
    mData.UpdateRecord tA, xr, vals2
    mData.TxCommit
    mApp.EndWork
    Check "X04", "Eingaben werden nie als Formel ausgeführt; Änderung rechnet abhängige Formelspalte neu", NzS(mData.GetValue(tA, xr, xf2)) = "=1+1!", NzS(mData.GetValue(tA, xr, xf2))
    Check "X05", "Zirkelbezug bei Excel-Formel wird erkannt", mCalc.HasCycle(xf, mCalc.XlFromDisplay(tA, "=[Selbsttest Formel]+1", xmsg), path), path

    '--- G: Dashboard-Raster (Anordnen) -------------------------------------------------
    Dim gd As String, ge(1 To 3) As String, wsD As Worksheet, sA As Shape, gOk As Boolean
    gd = mDash.CreateDashboard("Selbsttest Raster", "")
    For i = 1 To 3
        Set pr = New cMap
        pr.SetVal "Titel", "Raster " & i: pr.SetVal "Typ", IIf(i = 3, "Tabelle", "Kachel"): pr.SetVal "Tabelle", tA: pr.SetVal "Kennzahl", "ANZAHL"
        If i = 3 Then pr.SetVal "Gruppierung", Fid(tA, "Status")
        ge(i) = mDash.CreateElement(gd, pr)
    Next
    mCfg.SetV SH_ELM, ge(1), "Groesse", "2": mCfg.SetV SH_ELM, ge(1), "Optionen", "gx=2;gy=0"
    mCfg.SetV SH_ELM, ge(2), "Groesse", "2": mCfg.SetV SH_ELM, ge(2), "Optionen", "gx=0;gy=0"
    mCfg.SetV SH_ELM, ge(3), "Groesse", "4": mCfg.SetV SH_ELM, ge(3), "Optionen", "gx=0;gy=1"
    Check "G01", "Gespeicherte Rasterplätze werden genau so gezeichnet", mDash.LayoutInfo(gd) = ge(1) & ":2,0,2;" & ge(2) & ":0,0,2;" & ge(3) & ":0,1,4;", mDash.LayoutInfo(gd)
    ' Karte 1 im Anordnen-Modus unter Karte 2 ziehen -> nach "Fertig" dort gespeichert, Karte 3 rutscht nach unten
    mNav.GoDash gd
    mDash.BeginArrange
    mNav.RenderCurrent
    Set wsD = mRender.Sheet(SH_DASH)
    Set sA = wsD.Shapes("arr_" & ge(1))
    sA.Left = wsD.Shapes("arr_" & ge(2)).Left + 3
    sA.Top = wsD.Shapes("arr_" & ge(3)).Top + 4
    mDash.ApplyArrangement
    mDash.EndArrange
    mNav.RenderCurrent
    gOk = (mDash.LayoutInfo(gd) = ge(1) & ":0,1,2;" & ge(2) & ":0,0,2;" & ge(3) & ":0,2,4;")
    gOk = gOk And Abs(wsD.Shapes("card_" & ge(1)).Left - wsD.Shapes("card_" & ge(2)).Left) < 1 And wsD.Shapes("card_" & ge(1)).Top > wsD.Shapes("card_" & ge(2)).Top + 50
    Check "G02", "Anordnen: gezogene Karte rastet ein und bleibt nach „Fertig“ an ihrem Platz; verdrängte Karte rückt nach unten", gOk, mDash.LayoutInfo(gd)
    mDash.EndArrange

    '--- P: Seitenwechsel, Farben --------------------------------------------------------
    mNav.GoPage "Start"
    mNav.GoPage "Einrichten"
    mRender.gProf = ""
    mNav.GoPage "Start"
    ok = (Len(mRender.gProf) = 0)
    mApp.MarkDirty
    mRender.gProf = ""
    mNav.GoPage "Start"
    Check "P01", "Unveränderte Seite wird beim Wechsel nur eingeblendet, nach einer Änderung neu aufgebaut", ok And Len(mRender.gProf) > 0
    Dim nav0 As String, acc0 As String
    nav0 = mCfg.SysGet("Navigationsfarbe", ""): acc0 = mCfg.SysGet("Akzentfarbe", "")
    mCfg.SysSet "Navigationsfarbe", CStr(mCmd.NavValue("Flieder"))
    mCfg.SysSet "Akzentfarbe", CStr(mCmd.AccentValue("Mint"))
    ok = NavIsLight() And Not IsLightColor(CNavText()) And Not IsLightColor(CNavStrong()) And COnAccent() = CText() And Not IsLightColor(CAccent())
    mNav.RefreshNav True
    mCfg.SysSet "Navigationsfarbe", nav0: mCfg.SysSet "Akzentfarbe", acc0
    mNav.RefreshNav True
    Check "P02", "Helle Pastellfarben: dunkle Schrift in Navigation und auf Akzentflächen, lesbarer Akzent auf Weiß", ok

    '--- D: Auswertungen (ab 1.5): neue Kennzahlen und Darstellungen ---------------------
    Dim dd As Variant, dashA As String, el As Variant, rs As cMap, med As Variant, uniq As Variant, nErrEl As Long, nEl As Long
    Dim tAuf As cTableData, medVals() As Double, nv As Long, rr As Long, vv As Variant, expMed As Double, seenP As New cMap, k2 As Long, tmpD As Double
    For Each dd In mCfg.DashIds()
        If mCfg.GetS(SH_DSB, CStr(dd), "Name") = "Auswertungen" Then dashA = CStr(dd)
    Next
    If Len(dashA) > 0 Then
        For Each el In mCfg.ElementIds(dashA)
            nEl = nEl + 1
            Set rs = mDash.Compute(CStr(el), dashA)
            If rs.Exists("error") Then nErrEl = nErrEl + 1
            If mDash.ElemGet(CStr(el), "Kennzahl") = "MEDIAN" Then med = rs.GetVal("totalValue")
            If mDash.ElemGet(CStr(el), "Kennzahl") = "EINDEUTIG" Then uniq = rs.GetVal("totalValue")
        Next
        ' Vergleichswerte direkt aus den Daten
        Set tAuf = mData.TData(mCfg.TableIdByName("Aufgaben"))
        ReDim medVals(1 To tAuf.RowCount)
        For rr = 1 To tAuf.RowCount
            If Not tAuf.IsArchived(rr) Then
                vv = tAuf.Val_(rr, mCfg.FieldIdByName(tAuf.TableID, "Aufwand (h)"))
                If IsNumeric(vv) And Not IsBlank(vv) Then nv = nv + 1: medVals(nv) = CDbl(vv)
                vv = mData.LinkTargets(mCfg.FieldIdByName(tAuf.TableID, "Projekt"), tAuf.IdAt(rr))
                If Len(vv) > 0 Then seenP.SetVal LCase$(CStr(vv)), True
            End If
        Next
        For rr = 2 To nv
            tmpD = medVals(rr): k2 = rr - 1
            Do While k2 >= 1
                If medVals(k2) <= tmpD Then Exit Do
                medVals(k2 + 1) = medVals(k2): k2 = k2 - 1
            Loop
            medVals(k2 + 1) = tmpD
        Next
        If nv Mod 2 = 1 Then expMed = medVals((nv + 1) \ 2) Else expMed = (medVals(nv \ 2) + medVals(nv \ 2 + 1)) / 2
    End If
    Check "D01", "Neue Kennzahlen: Median und Anzahl verschiedener Werte stimmen mit Nachrechnung überein; alle " & nEl & " Elemente berechnet", _
          Len(dashA) > 0 And nErrEl = 0 And Abs(CDbl(NzD(med)) - expMed) < 0.0001 And CLng(NzD(uniq)) = seenP.Count, _
          "Median " & FmtNum(CDbl(NzD(med)), 2) & " (erwartet " & FmtNum(expMed, 2) & "), verschiedene Projekte " & NzD(uniq) & " (erwartet " & seenP.Count & ")"
    Dim chTypes As String, co As ChartObject, okTypes As Boolean
    Dim skV As Variant, skF As Variant, skSum As Double, skTot As Double, skBands As Long, skNB As Long, skShp As Shape
    Dim cpEid As String, cpOk As Boolean, cpFmt As Variant, cpN As Long
    If Len(dashA) > 0 Then
        mNav.GoDash dashA
        For Each co In ThisWorkbook.Worksheets(SH_DASH).ChartObjects
            chTypes = chTypes & co.Chart.ChartType & ";"
        Next
        ' Netz (xlRadarFilled 82), Kreis (5), Fläche (1), Pareto (Säulen 51), Trichter (123), Baumdiagramm (117)
        okTypes = InStr(";" & chTypes, ";82;") > 0 And InStr(";" & chTypes, ";5;") > 0 And InStr(";" & chTypes, ";1;") > 0 And _
                  InStr(";" & chTypes, ";123;") > 0 And InStr(";" & chTypes, ";117;") > 0
        ' Sankey: Summe aller Flüsse = Anzahl der Aufgaben; Bänder als Formen gezeichnet
        For Each el In mCfg.ElementIds(dashA)
            If mDash.ElemGet(CStr(el), "Typ") = "Sankey" Then
                Set rs = mDash.Compute(CStr(el), dashA)
                skV = rs.GetVal("sankey", Empty)
                If IsArray(skV) Then
                    For Each skF In skV(1)
                        skSum = skSum + skF(2)
                    Next
                    skNB = UBound(skV(0)) + 1
                End If
                skTot = CDbl(NzD(rs.GetVal("totalValue")))
                For Each skShp In ThisWorkbook.Worksheets(SH_DASH).Shapes
                    If skShp.Name Like "es_" & CStr(el) & "_f*" Then skBands = skBands + 1
                Next
            End If
        Next
        ' Diagramm kopieren (1.8.0): erstes Diagramm als Bild in die Zwischenablage
        For Each co In ThisWorkbook.Worksheets(SH_DASH).ChartObjects
            cpEid = mDash.ElementFromShape(co.Name)
            If Len(cpEid) > 0 Then Exit For
        Next
        If Len(cpEid) > 0 Then
            cpOk = mDash.CopyElement(cpEid)
            cpFmt = Application.ClipboardFormats
            If IsArray(cpFmt) Then cpN = UBound(cpFmt) - LBound(cpFmt) + 1
        End If
        mNav.GoPage "Start"
    End If
    Check "D02", "Dashboard „Auswertungen“ zeichnet Netz, Kreis, Fläche, Pareto, Trichter und Baumdiagramm", okTypes, "Diagrammtypen: " & chTypes
    Check "D03", "Sankey „Status " & ChrW(&H2192) & " Priorität“: Summe der Flüsse = Anzahl der Aufgaben, Bänder gezeichnet", _
          skTot > 0 And Abs(skSum - skTot) < 0.0001 And skBands > 0 And skNB > 0, _
          "Flüsse " & skSum & ", Aufgaben " & skTot & ", Ziele " & skNB & ", Bänder " & skBands
    Check "D04", "Diagramm kopieren: Element „" & mDash.ElemGet(cpEid, "Titel") & "“ als Bild in der Zwischenablage (für PowerPoint und Word)", _
          cpOk And cpN > 0, "Formate in der Zwischenablage: " & cpN

    '--- E: Excel-Export mit mehreren Blättern und Gestaltung ---------------------------
    Dim ePath As String, eJobs As New Collection, eWb As Workbook, eOk As Boolean, eN As Long, eInfo As String
    ePath = ThisWorkbook.Path & "\Export_Selbsttest.xlsx"
    If FileExists(ePath) Then Kill ePath
    eJobs.Add mNav.FirstViewOf(mCfg.TableIdByName("Projekte"))
    eJobs.Add mNav.FirstViewOf(mCfg.TableIdByName("Aufgaben"))
    eN = mIO.ExportForTest(ePath, eJobs)
    Set eWb = Workbooks.Open(ePath, ReadOnly:=True)
    eOk = (eWb.Worksheets.Count = 2) And (eWb.Worksheets(1).Cells(5, 1).Interior.Color = mDocStyle.DocColor("Petrol")) And _
          (eWb.Worksheets(2).Cells(6, 1).Font.Name = "Garamond") And eWb.Worksheets(1).Cells(5, 1).Value = mCfg.FieldName(mCfg.LabelField(mCfg.TableIdByName("Projekte")))
    eInfo = eWb.Worksheets(1).Name & " / " & eWb.Worksheets(2).Name & ", " & eN & " Einträge"
    eWb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Check "E01", "Excel-Export: zwei Tabellen als zwei Blätter, Kopfzeile in gewählter Farbe, gewählte Schrift", eOk, eInfo

    '--- Erfassungsdateien (Rundlauf) ---------------------------------------------
    Dim cap As String, ln As Variant
    cap = mCapture.SelfTestRoundTrip(ThisWorkbook.Path)
    For Each ln In Split(cap, vbLf)
        If Len(ln) > 0 Then
            mOut = mOut & ln & vbLf
            If InStr(ln, vbTab & "BESTANDEN") > 0 Then mPass = mPass + 1 Else mFail = mFail + 1
        End If
    Next

    '--- L: Liste übernehmen (1.7.0) ------------------------------------------------
    ListImportChecks

    mOut = mOut & vbLf & "Ergebnis: " & mPass & " bestanden, " & mFail & " nicht bestanden"
    RunSelfTest = mOut
    WriteResult
    Exit Function
crash:
    mOut = mOut & vbLf & "ABBRUCH: " & Err.Number & " " & Err.Description
    RunSelfTest = mOut
    mData.TxAbortIfOpen
    mApp.ResetWork
    WriteResult
End Function

' T21: Laufzeitmessung in der Lasttestdatei (Build_Stress). Ergebnis in Leistung_Ergebnis.txt
Public Function Measure() As String
    Dim tA As String, vid As String, st As cViewState, n As Long, rows() As Long, t0 As Double, s As String, f As frmRecord
    Dim rid As String, vals As cMap, did As String, props As cMap, i As Long, fidName As String, vb As Variant, totalRecs As Long, ids As Variant
    Dim spec As Variant, h As Integer
    On Error GoTo crash
    mApp.gStarted = True
    tA = mCfg.TableIdByName("Lasttabelle Groß")
    If Len(tA) = 0 Then Measure = "Lasttestdatei erforderlich": Exit Function
    ids = mCfg.TableIds()
    For i = 0 To ArrLen(ids) - 1
        totalRecs = totalRecs + mData.TData(CStr(ids(i))).RowCount
    Next
    s = "Lastprofil: " & ArrLen(ids) & " Tabellen, " & totalRecs & " Datensätze gesamt, größte Tabelle " & mData.TData(tA).RowCount & " Datensätze mit " & _
        mCfg.Fields(tA).Count & " Feldern, " & mCfg.TD(SH_REL).RowCount & " Zuordnungen" & vbLf
    s = s & "Rechner: Excel " & Application.Version & " " & IIf(InStr(Application.OperatingSystem, "64") > 0, "64", "32") & " Bit, " & Application.OperatingSystem & vbLf
    s = s & "Messung mit VBA-Timer (Sekunden), jeweils erster Aufruf nach dem Öffnen (kalt) und Wiederholung (warm)" & vbLf & vbLf
    vid = mCfg.ViewIds(tA)(0)

    t0 = Timer: mNav.GoTable tA, vid: s = s & "Tabellenansicht öffnen (10.000 DS, Seite 1)" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    t0 = Timer: mNav.GoTable tA, vid: s = s & "Tabellenansicht erneut (warm)" & vbTab & Format$(Timer - t0, "0.00") & vbLf

    Set st = mNav.State(vid)
    st.Search = "Eintrag 12 Nord"
    t0 = Timer: rows = mViewTable.ViewRows(vid, n): s = s & "Erste Suche „Eintrag 12 Nord“ in 10.000 DS × 50 Spalten (" & n & " Treffer, baut Suchindex)" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    s = s & "   Suchindex je Feldart: " & mQuery.gSrchProfile & vbLf
    t0 = Timer: mNav.RenderCurrent: s = s & "Suche inkl. Seitenaufbau" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    st.Search = "Eintrag 123"
    t0 = Timer: rows = mViewTable.ViewRows(vid, n): s = s & "Weitere Suche „Eintrag 123“ (" & n & " Treffer, Suchindex wiederverwendet)" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    st.Search = ""
    st.Conds = "|" & mCfg.FieldIdByName(tA, "Feld 8") & "|OFFEN||"
    t0 = Timer: rows = mViewTable.ViewRows(vid, n): s = s & "Filter Status offen (" & n & " Treffer)" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    st.Conds = ""
    st.SortSpec = mCfg.FieldIdByName(tA, "Feld 1") & ":D"
    t0 = Timer: rows = mViewTable.ViewRows(vid, n): s = s & "Sortierung absteigend über 10.000 DS" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    st.SortSpec = ""

    rid = mData.TData(tA).IdAt(5000)
    Set f = New frmRecord
    t0 = Timer: f.PrepareOnly tA, rid: s = s & "Eintrag öffnen (Maske mit 50 Feldern aufbauen)" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    Unload f
    Set vals = New cMap
    vals.SetVal mCfg.FieldIdByName(tA, "Name"), "Geändert " & Format$(Now, "hh:nn:ss")
    t0 = Timer
    mApp.BeginWork
    mData.TxBegin "Messung"
    mData.UpdateRecord tA, rid, vals
    mLog.Add "Geändert", tA, rid, ""
    mData.TxCommit
    mApp.EndWork
    s = s & "Eintrag speichern (Transaktion, Protokoll, Neuberechnung)" & vbTab & Format$(Timer - t0, "0.00") & vbLf

    did = mDash.CreateDashboard("Leistungsmessung", "6 Standardelemente")
    spec = Array("Kachel|ANZAHL||", "Balken|ANZAHL||Feld 4", "Säulen|ANZAHL||Feld 8", "Zeitverlauf|ANZAHL||Feld 3", "Anteil|ANZAHL||Feld 5", "Tabelle|SUMME|Feld 2|Feld 4")
    For i = 0 To 5
        vb = Split(spec(i), "|")
        Set props = New cMap
        props.SetVal "Titel", "Element " & (i + 1)
        props.SetVal "Typ", vb(0)
        props.SetVal "Tabelle", tA
        props.SetVal "Kennzahl", vb(1)
        If Len(vb(2)) > 0 Then props.SetVal "Feld", mCfg.FieldIdByName(tA, CStr(vb(2)))
        If Len(vb(3)) > 0 Then props.SetVal "Gruppierung", mCfg.FieldIdByName(tA, CStr(vb(3)))
        If vb(0) = "Zeitverlauf" Then props.SetVal "Datumsstufe", "Monat"
        mDash.CreateElement did, props
    Next
    t0 = Timer: mNav.GoDash did: s = s & "Dashboard mit 6 Elementen über 10.000 DS aktualisieren" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    t0 = Timer: mNav.GoDash did: s = s & "Dashboard erneut (warm)" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    Dim ev As Variant
    For Each ev In mCfg.ElementIds(did)
        t0 = Timer: mDash.Compute CStr(ev), did: s = s & "   davon Berechnung " & mDash.ElemGet(CStr(ev), "Typ") & vbTab & Format$(Timer - t0, "0.00") & vbLf
    Next

    For i = 1 To ArrLen(mCfg.ViewIds(tA)) - 1
        vid = mCfg.ViewIds(tA)(i)
        t0 = Timer: mNav.GoTable tA, vid: s = s & "Ansicht „" & mCfg.ViewGet(vid, "Name") & "“ öffnen" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    Next
    t0 = Timer: mCalc.RecalcAll: s = s & "Alle Berechnungen neu (10.000 Rollups über 40.000 Zuordnungen)" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    t0 = Timer: mNav.GoStart: s = s & "Startseite" & vbTab & Format$(Timer - t0, "0.00") & vbLf
    Measure = s
    h = FreeFile
    Open ThisWorkbook.Path & "\Leistung_Ergebnis.txt" For Output As #h
    Print #h, Replace(s, vbLf, vbCrLf)
    Close #h
    Exit Function
crash:
    Measure = s & vbLf & "ABBRUCH: " & Err.Number & " " & Err.Description
End Function

' Seitenwechsel je Seite messen (nur Entwicklung): Seitenaufbau und Navigation getrennt,
' jeweils 3 Durchlaeufe; mit Zwischenzeiten aus mRender.Lap
Public Function ProfilePages() As String
    Dim pages As Variant, p As Variant, k As Long, t0 As Double, t1 As Double, s As String, ids As Variant, h As Integer
    On Error GoTo crash
    mApp.gStarted = True
    ids = mCfg.TableIds()
    pages = Array("Start", "T:" & ids(0), "T:" & ids(1), "Dashboard", "Erfassung", "Einrichten", "Hilfe", "Technik")
    For k = 1 To 3
        For Each p In pages
            mRender.gProf = ""
            t0 = Timer
            ' wie ein Klick in der Navigation (inkl. Wiederverwendung unveraenderter Seiten)
            If Left$(p, 2) = "T:" Then
                mNav.NavClick CStr(p)
            ElseIf p = "Dashboard" Then
                mNav.GoDash ""
            Else
                mNav.GoPage CStr(p)
            End If
            DoEvents
            s = s & p & vbTab & Format$((Timer - t0) * 1000, "0") & " ms gesamt" & vbTab & IIf(Len(mRender.gProf) = 0, "(wiederverwendet)", mRender.gProf) & vbLf
        Next
    Next
    ProfilePages = s
    h = FreeFile
    Open ThisWorkbook.Path & "\Seiten_Profil.txt" For Output As #h
    Print #h, Replace(s, vbLf, vbCrLf)
    Close #h
    Exit Function
crash:
    ProfilePages = s & vbLf & "ABBRUCH: " & Err.Number & " " & Err.Description
End Function

' Kosten einzelner Formen (nur Entwicklung): je 40 Stueck auf dem Suchblatt
Public Function BenchShapes() As String
    Dim ws As Worksheet, i As Long, t0 As Double, s As String, sh As Shape
    Set ws = mRender.Sheet(SH_SEARCH)
    mRender.BeginPage ws
    t0 = Timer
    For i = 1 To 40: Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, 200, 100 + i * 2, 100, 24): Next
    s = s & "AddShape nackt: " & Format$((Timer - t0) * 25, "0.0") & " ms/Stk" & vbLf
    t0 = Timer
    For i = 1 To 40: mRender.Btn ws, "", "Schaltfläche " & i, 200, 100, "mCmd.Cmd_GoStart", "primary", ChrW(&HE710): Next
    s = s & "Btn primary (mit Schatten): " & Format$((Timer - t0) * 25, "0.0") & " ms/Stk" & vbLf
    t0 = Timer
    For i = 1 To 40: mRender.Btn ws, "", "Schaltfläche " & i, 200, 100, "mCmd.Cmd_GoStart", "ghost", ChrW(&HE710): Next
    s = s & "Btn ghost (ohne Schatten): " & Format$((Timer - t0) * 25, "0.0") & " ms/Stk" & vbLf
    t0 = Timer
    For i = 1 To 40: mRender.TextShape ws, "tx" & i, "Text " & i, 200, 100, 200, 20, 10, CText(), False, 1: Next
    s = s & "TextShape: " & Format$((Timer - t0) * 25, "0.0") & " ms/Stk" & vbLf
    t0 = Timer
    For i = 1 To 40
        Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, 200, 100, 100, 24)
        sh.TextFrame.Characters.Text = "Legacy " & i
    Next
    s = s & "AddShape + TextFrame (alt): " & Format$((Timer - t0) * 25, "0.0") & " ms/Stk" & vbLf
    t0 = Timer
    For i = 1 To 40
        Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, 200, 100, 100, 24)
        sh.TextFrame2.TextRange.Text = "TF2 " & i
    Next
    s = s & "+ TextFrame2.Text: " & Format$((Timer - t0) * 25, "0.0") & " ms/Stk" & vbLf
    t0 = Timer
    For i = 1 To 40
        Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, 200, 100, 100, 24)
        With sh.TextFrame2
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse
        End With
    Next
    s = s & "+ TF2 Raender/Anker/Umbruch: " & Format$((Timer - t0) * 25, "0.0") & " ms/Stk" & vbLf
    t0 = Timer
    For i = 1 To 40
        Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, 200, 100, 100, 24)
        sh.TextFrame2.TextRange.Text = "TF2 " & i
        With sh.TextFrame2.TextRange.Font
            .Name = UI_FONT: .Size = 9.5: .Fill.ForeColor.RGB = CText()
        End With
    Next
    s = s & "+ TF2 Text+Schrift: " & Format$((Timer - t0) * 25, "0.0") & " ms/Stk" & vbLf
    t0 = Timer
    For i = 1 To 40
        Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, 200, 100, 100, 24)
        With sh.TextFrame
            .Characters.Text = "TF " & i
            .Characters.Font.Name = UI_FONT: .Characters.Font.Size = 9.5: .Characters.Font.Color = CText()
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .VerticalAlignment = xlVAlignCenter: .HorizontalAlignment = xlHAlignCenter
        End With
    Next
    s = s & "Legacy TextFrame komplett: " & Format$((Timer - t0) * 25, "0.0") & " ms/Stk" & vbLf
    t0 = Timer
    For i = 1 To 40
        Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, 200, 100, 100, 24)
        With sh.Shadow
            .Visible = msoTrue: .ForeColor.RGB = RGB(0, 0, 0): .Transparency = 0.9: .OffsetX = 0: .OffsetY = 1: .Blur = 2: .Size = 100
        End With
    Next
    s = s & "+ Schatten: " & Format$((Timer - t0) * 25, "0.0") & " ms/Stk" & vbLf
    t0 = Timer
    For i = 1 To 40
        Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, 200, 100, 100, 24)
        sh.TextFrame2.TextRange.Text = "Autosize " & i
        sh.TextFrame2.WordWrap = msoFalse
        sh.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        sh.TextFrame2.AutoSize = msoAutoSizeNone
    Next
    s = s & "+ AutoSize: " & Format$((Timer - t0) * 25, "0.0") & " ms/Stk" & vbLf
    mRender.EndPage ws
    BenchShapes = s
End Function

' Laufzeitprofil der Neuberechnung (nur Entwicklung): n Datensaetze mit je 4 Zuordnungen
Public Function ProfileRecalc(ByVal n As Long) As String
    Dim tid As String, tR As String, linkF As String, rollF As String, props As cMap, a() As Variant, t As cTableData, r As Long, j As Long
    Dim rels() As Variant, ids() As String, nR As Long, t0 As Double, s As String, x As String, f As cField, v As Variant, k As Long
    tR = mCfg.TableIdByName("Ressourcen")
    tid = mStruct.CreateTable("Profil", "", &HE8FD&)
    mStruct.SetLabelField tid, mStruct.AddField(tid, "Name", FT_TEXT)
    Set props = New cMap
    props.SetVal "Zieltabelle", tR: props.SetVal "Mehrfach", "Ja": props.SetVal "Gegenname", "Profil"
    linkF = mStruct.AddField(tid, "Link", FT_LINK, props)
    Set props = New cMap
    props.SetVal "Formel", "ROLLUP|F:" & linkF & "|K:ANZAHL": props.SetVal "Ergebnistyp", FT_ZAHL
    rollF = mStruct.AddField(tid, "Anzahl", FT_FORMEL, props)
    Set t = mData.TData(tid)
    t0 = Timer
    ReDim a(1 To n, 1 To t.ColCount)
    For r = 1 To n
        v = mData.NewRowArray(tid, "")
        For j = 1 To t.ColCount: a(r, j) = v(j): Next
        a(r, 1) = "R" & Format$(mCfg.ReserveIds("R", 1), "0000000")
        a(r, t.ColOf(mCfg.FieldIdByName(tid, "Name"))) = "Eintrag " & r
    Next
    mData.InsertRecordsBulk tid, a
    s = s & "Einfügen " & n & ": " & Format$(Timer - t0, "0.0") & " s" & vbLf
    Set t = mData.TData(tid)
    Set f = mCfg.Field(linkF)
    ReDim ids(1 To mData.TData(tR).RowCount)
    For r = 1 To mData.TData(tR).RowCount: ids(r) = mData.TData(tR).IdAt(r): Next
    nR = UBound(ids)
    ReDim rels(1 To n * 4, 1 To 4)
    k = 0
    For r = 1 To n
        For j = 0 To 3
            k = k + 1
            rels(k, 1) = linkF: rels(k, 2) = t.IdAt(r): rels(k, 3) = ids(1 + ((r * 7 + j * 5) Mod nR)): rels(k, 4) = Now
        Next
    Next
    t0 = Timer
    mData.TxBegin "Profil"
    mData.RelAddBulk rels
    mData.TxCommit
    mData.ClearUndo
    s = s & "Zuordnungen " & (n * 4) & ": " & Format$(Timer - t0, "0.0") & " s" & vbLf
    mData.InvalidateRel
    t0 = Timer
    x = mData.LinkTargets(linkF, t.IdAt(1))
    s = s & "Zuordnungsindex aufbauen: " & Format$(Timer - t0, "0.00") & " s" & vbLf
    t0 = Timer
    For r = 1 To n
        x = mData.LinkTargets(linkF, t.IdAt(r))
    Next
    s = s & "LinkTargets x" & n & ": " & Format$(Timer - t0, "0.00") & " s" & vbLf
    t0 = Timer
    For r = 1 To n
        v = mCalc.EvalDef(mCfg.Field(rollF).Formula, t, r)
    Next
    s = s & "EvalDef ROLLUP x" & n & ": " & Format$(Timer - t0, "0.00") & " s" & vbLf
    t0 = Timer
    mCalc.RecalcAll
    s = s & "RecalcAll: " & Format$(Timer - t0, "0.00") & " s" & vbLf
    Dim fv As Variant, tt As cTableData, vals() As Variant
    t0 = Timer
    Set props = New cMap
    For Each fv In mCalc.TopoOrder()
        props.SetVal CStr(fv), 1
    Next
    s = s & "TopoOrder: " & Format$(Timer - t0, "0.00") & " s (" & props.Count & " Formeln)" & vbLf
    For Each fv In props.Keys
        Set f = mCfg.Field(CStr(fv))
        Set tt = mData.TData(f.TableID)
        t0 = Timer
        ReDim vals(1 To Application.Max(1, tt.RowCount))
        For r = 1 To tt.RowCount
            vals(r) = mCalc.EvalDef(f.Formula, tt, r)
        Next
        x = Format$(Timer - t0, "0.00")
        t0 = Timer
        If tt.RowCount > 0 Then mData.WriteColumnRaw f.TableID, f.ID, vals
        s = s & "  " & mCfg.TableName(f.TableID) & "." & f.Name & " (" & tt.RowCount & "): rechnen " & x & " s, schreiben " & Format$(Timer - t0, "0.00") & " s" & vbLf
    Next
    ProfileRecalc = s
    Dim h As Integer
    h = FreeFile
    Open ThisWorkbook.Path & "\Profil_Ergebnis_" & n & ".txt" For Output As #h
    Print #h, Replace(s, vbLf, vbCrLf)
    Close #h
End Function

' Zustandsabfrage fuer T01/T18 (keine Datenaenderung)
Public Function Probe() As String
    Dim w As Boolean
    mUI.gQuiet = True: mUI.gQuietLog = ""
    w = mApp.RequireWritable("Die Prüfung")
    mUI.gQuiet = False
    Probe = "gestartet=" & mApp.gStarted & "; lesemodus=" & mApp.gReadOnly & "; schreiben erlaubt=" & w & "; datei=" & ThisWorkbook.Name & _
            "; seite=" & ActiveSheet.Name & "; tabellen=" & ArrLen(mCfg.TableIds()) & "; navigation=" & (mNav.gPage <> "")
End Function

' Gezielter Laufzeitfehler mitten in einer Transaktion (fuer T22)
Private Sub ForcedError()
    Dim tA As String, vals As New cMap, rid As String
    On Error GoTo f
    tA = mCfg.TableIdByName("Aufgaben")
    rid = mData.TData(tA).IdAt(1)
    mApp.BeginWork
    Application.Calculation = xlCalculationManual
    mData.TxBegin "Selbsttest Fehler"
    vals.SetVal mCfg.FieldIdByName(tA, "Aufgabe"), "WIRD ZURÜCKGENOMMEN"
    mData.UpdateRecord tA, rid, vals
    Err.Raise 11
    Exit Sub
f: mApp.Boundary "Selbsttest"
End Sub

' Fachlicher Fingerabdruck: Anzahl Tabellen/Felder/Ansichten/Dashboards/Elemente, Datensaetze
' je Tabelle, Zuordnungen und Pruefsumme ueber alle Werte (fuer Speichern/Wiederoeffnen und Sicherung)
Public Function Snapshot() As String
    Dim ids As Variant, i As Long, t As cTableData, r As Long, c As Long, s As String, sum As Double, txt As String, k As Long
    ids = mCfg.TableIds()
    s = "Tabellen=" & ArrLen(ids)
    For i = 0 To ArrLen(ids) - 1
        Set t = mData.TData(CStr(ids(i)))
        s = s & " | " & mCfg.TableName(CStr(ids(i))) & ": " & t.RowCount & " DS, " & mCfg.Fields(CStr(ids(i)), False).Count & " Felder, " & ArrLen(mCfg.ViewIds(CStr(ids(i)))) & " Ansichten"
        For r = 1 To t.RowCount
            For c = 1 To t.ColCount
                txt = CStr(t.Cell(r, c))
                For k = 1 To Len(txt)
                    sum = sum + AscW(Mid$(txt, k, 1)) * ((k Mod 7) + 1)
                Next
            Next
        Next
    Next
    Dim tv As cTableData, vsum As Double
    Set tv = mCfg.TD(SH_VIEW)
    For r = 1 To tv.RowCount
        For c = 1 To tv.ColCount
            txt = CStr(tv.Cell(r, c))
            For k = 1 To Len(txt)
                vsum = vsum + AscW(Mid$(txt, k, 1)) * ((k Mod 5) + 1)
            Next
        Next
    Next
    s = s & " | Ansichten=" & tv.RowCount & " (Prüfsumme " & Format$(vsum, "0") & ")"
    s = s & " | Zuordnungen=" & mCfg.TD(SH_REL).RowCount & " | Dashboards=" & mCfg.TD(SH_DSB).RowCount & " | Elemente=" & mCfg.TD(SH_ELM).RowCount
    s = s & " | Pruefsumme=" & Format$(sum, "0")
    Snapshot = s
End Function

' T19: Sicherung erstellen, Daten aendern, Sicherung wiederherstellen, Zustand vergleichen
' Tabelle anlegen (wie der Dialog „Neue Tabelle“), löschen, erneut anlegen. Liefert "OK" oder Schritt + Fehler
Public Function CreateDeleteCreateTest() As String
    Dim stp As String, tid As String, k As Long, props As cMap, fid As String
    On Error GoTo f
    For k = 1 To 3
        stp = "anlegen " & k
        mApp.BeginWork
        tid = mStruct.CreateTable("Prueftabelle " & k, "", &HE8FD&)
        Set props = New cMap
        props.SetVal "Pflicht", "Ja"
        stp = "Feld " & k
        fid = mStruct.AddField(tid, "Name", FT_TEXT, props)
        mStruct.SetLabelField tid, fid
        stp = "Ansicht " & k
        mStruct.CreateDefaultView tid
        mApp.EndWork
        stp = "Navigation " & k
        mNav.RefreshNav
        stp = "Tabelle öffnen " & k
        mNav.GoTable tid
        If k < 3 Then
            stp = "löschen " & k
            mApp.BeginWork
            mStruct.DeleteTable tid
            mApp.EndWork
            mNav.gTid = "": mNav.gVid = "": mNav.gPage = "Start"
            mNav.RefreshNav
            mNav.RenderCurrent
        End If
    Next
    CreateDeleteCreateTest = "OK: dreimal angelegt, zweimal gelöscht"
    Exit Function
f:
    CreateDeleteCreateTest = "FEHLER bei „" & stp & "“: " & Err.Number & " " & Err.Description & " (" & Err.Source & ")"
    mApp.ResetWork
End Function

' Sicherungsname ohne Tag und Uhrzeit ("Name_2025-06-03_143000_v1.7.0.xlsm" -> "Name_JJJJ-MM-TT_hhmmss_v1.7.0.xlsm")
Private Function StampMasked(ByVal nm As String) As String
    Dim p As Long
    p = InStrRev(nm, "_v")
    If p > 19 Then
        If Mid$(nm, p - 18, 1) = "_" And Mid$(nm, p - 7, 1) = "_" Then nm = Left$(nm, p - 19) & "_JJJJ-MM-TT_hhmmss" & Mid$(nm, p)
    End If
    StampMasked = nm
End Function

Public Function BackupRestoreTest() As String
    Dim a As String, b As String, c As String, ok As Boolean, l As Collection, path As String, tA As String, rid As String, vals As New cMap
    On Error GoTo crash
    mUI.gQuiet = True: mUI.gQuietLog = ""
    tA = mCfg.TableIdByName("Aufgaben")
    a = Snapshot()
    ok = mBackup.BackupNow("Selbsttest T19", True)
    Set l = mBackup.ListBackups()
    path = mBackup.BackupFolder() & "\" & l(1)(0)
    ' Aenderung nach der Sicherung: Eintrag umbenennen und einen archivieren
    rid = FirstActive(tA, 0)
    vals.SetVal mCfg.FieldIdByName(tA, "Aufgabe"), "GEÄNDERT NACH SICHERUNG"
    mData.TxBegin "Selbsttest T19"
    mData.UpdateRecord tA, rid, vals
    mData.TxCommit
    mForms.SetArchived tA, FirstActive(tA, 1), True
    b = Snapshot()
    mBackup.RestoreFrom path
    c = Snapshot()
    BackupRestoreTest = "T19" & vbTab & IIf(ok And a = c And a <> b, "BESTANDEN", "NICHT BESTANDEN") & vbTab & _
                        "Sicherung " & StampMasked(CStr(l(1)(0))) & " (" & Format$(l(1)(2) / 1024, "0") & " KB) wiederhergestellt; Zustand vorher/nachher identisch: " & (a = c) & _
                        "; Zwischenstand abweichend: " & (a <> b) & vbLf & "vorher:  " & a & vbLf & "nachher: " & c
    mUI.gQuiet = False
    Exit Function
crash:
    mUI.gQuiet = False
    BackupRestoreTest = "T19" & vbTab & "ABBRUCH" & vbTab & Err.Description & vbLf & mUI.gQuietLog
End Function

' T17: Fingerabdruck in Datei schreiben (vor Speichern bzw. nach Wiederoeffnen)
Public Function SnapshotToFile(ByVal fileName As String) As String
    Dim h As Integer
    SnapshotToFile = Snapshot()
    h = FreeFile
    Open ThisWorkbook.Path & "\" & fileName For Output As #h
    Print #h, SnapshotToFile
    Close #h
End Function

Public Function SaveForTest() As String
    mUI.gQuiet = True: mUI.gQuietLog = ""
    SaveForTest = IIf(mBackup.ControlledSave(False), "GESPEICHERT", "NICHT GESPEICHERT") & IIf(Len(mUI.gQuietLog) > 0, " " & mUI.gQuietLog, "")
    mUI.gQuiet = False
End Function

' Ergebnis neben der Datei ablegen (Selbsttest_Ergebnis.txt, Windows-1252)
Private Sub WriteResult()
    Dim h As Integer
    On Error Resume Next
    h = FreeFile
    Open ThisWorkbook.Path & "\Selbsttest_Ergebnis.txt" For Output As #h
    ' Entwicklungsstand nur mit Monat und Jahr (das Ergebnis liegt dem Lieferpaket bei)
    Print #h, "Selbsttest " & Format$(Now, "mmmm yyyy") & " · " & ThisWorkbook.Name & " · Excel " & Application.Version & " (" & IIf(InStr(Application.OperatingSystem, "64") > 0, "64", "32") & " Bit)"
    Print #h, Replace(mOut, vbLf, vbCrLf)
    Close #h
End Sub

'==============================================================================
' L: Liste übernehmen (1.7.0): Tabellensuche, Kopfzeilen, Rückfragen, Typen,
' Umwandlung, 2ask-Codeplan und eine echte Übernahme in eine neue Tabelle
'==============================================================================
Private Sub ListImportChecks()
    Call L01: Call L02: Call L03: Call L04: Call L05: Call L06: Call L07: Call L08: Call L09
    Call L10: Call L11: Call L12: Call L13: Call L14: Call L15: Call L16: Call L17
    Call L18: Call L19: Call L20
    mListImport.LiReset
End Sub

' 1.8.0: Hinweiszeile unter dem Kopf, Gruppenüberschrift vor allgemeinen Unterbegriffen, Fußnote
Private Sub L18()
    Dim plan As cMap, c As cLiCol
    On Error GoTo f
    LiLoad LiSh("Gemeinden", Array(Array("Gebiet", "Fläche km2 1)", "Bevölkerung", "", "Reisegebiete", ""), _
        Array("", "", "insgesamt", "männlich", "", ""), _
        Array("", "31.12.2025 (Jahr)", "31.12.2025 (Jahr)", "", "Schlüssel", "Bezeichnung"), _
        Array("Nord", 12.5, 1000, 480, "F01", "Ostsee"), Array("Süd", 20.1, 2000, 990, "F02", "Nordsee"), _
        Array("West", 8.7, 500, 240, "F01", "Ostsee"), Array("Ost", 9.9, 700, 350, "F03", "Harz")))
    mListImport.LiChooseTable 1
    Set c = LiCol("Fläche km2")
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L18", "Hinweiszeile unter dem Kopf (Stichtag, „Schlüssel“/„Bezeichnung“) ist kein Datensatz; „Bevölkerung männlich“, „Reisegebiete Schlüssel“, Fußnote entfernt", _
        mListImport.liTables(1).AnnRow = 3 And mListImport.LiRecordCount() = 4 And Not c Is Nothing And Not LiCol("Bevölkerung männlich") Is Nothing And _
        Not LiCol("Reisegebiete Bezeichnung") Is Nothing And Not LiCol("Reisegebiete Schlüssel") Is Nothing And LiPV(plan, 2, "Fläche km2") = 12.5, _
        "Einträge " & mListImport.LiRecordCount() & ", Hinweiszeile " & mListImport.liTables(1).AnnRow
    Exit Sub
f:  Check "L18", "Hinweiszeile unter dem Kopf", False, "Laufzeitfehler: " & Err.Description
End Sub

' 1.8.0: Beträge wie in Verwaltungslisten, Ja/Nein-Varianten, Platzhalter, "dito", keine unechte Mehrfachauswahl
Private Sub L19()
    Dim plan As cMap
    On Error GoTo f
    LiLoad LiSh("Planung", Array(Array("Maßnahme", "Kosten", "Anmeldung", "Ort", "Anzahl"), _
        Array("Fest", "ca. 1.200 €", "ja!", "Bürgerhaus", 10), Array("Kurs", "500,-", "(ja)", "Bürgerhaus, gr. Saal", "-"), _
        Array("Café", "s.o.", "nein", "Stadtteilbüro", 12), Array("Lesung", "300 €", "x", "Bibliothek", "k.A."), _
        Array("Markt", "1.000 €", "nein", "Schulhof", 8), Array("Kino", "250 €", "Ja", "Park", 20)))
    mListImport.LiChooseTable 1
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L19", "„ca. 1.200 €“, „500,-“ und „s.o.“ (Wert darüber) als Geldbetrag; „ja!“, „(ja)“ als Ja; „-“, „k.A.“ leer; Ort mit Komma bleibt Text", _
        LiCol("Kosten").FType = FT_GELD And LiPV(plan, 2, "Kosten") = 1200 And LiPV(plan, 3, "Kosten") = 500 And LiPV(plan, 4, "Kosten") = 500 And _
        LiCol("Anmeldung").FType = FT_JANEIN And LiPV(plan, 3, "Anmeldung") = True And LiCol("Ort").FType = FT_TEXT And _
        LiCol("Anzahl").FType = FT_ZAHL And IsEmpty(LiPV(plan, 3, "Anzahl")) And Not LiHasQ("typ:S:anzahl"), _
        "Kosten " & LiTxt(LiPV(plan, 2, "Kosten")) & "/" & LiTxt(LiPV(plan, 3, "Kosten")) & "/" & LiTxt(LiPV(plan, 4, "Kosten")) & ", Ort " & LiCol("Ort").FType
    Exit Sub
f:  Check "L19", "Beträge, Ja/Nein, Platzhalter, dito", False, "Laufzeitfehler: " & Err.Description
End Sub

' 1.8.0: unkodierter 2ask-Export mit Codeplan (Antworttexte statt Codes)
Private Sub L20()
    Dim p As String, wb As Workbook, ws As Worksheet, cp As Worksheet, plan As cMap, prevSU As Boolean
    On Error GoTo f
    p = ThisWorkbook.Path & "\_Selbsttest_2ask_text.xlsx"
    If Len(Dir$(p)) > 0 Then Kill p
    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "results"
    ws.Range("A1:B1").Value = Array("ID", "In welchem Bezirk?")
    ws.Range("A3:B3").Value = Array("{}", "{}")
    ws.Range("A4:B4").Value = Array("survey_session_id", "Q1")
    ws.Range("A5:B5").Value = Array("AAA111", "Süd")
    ws.Range("A6:B6").Value = Array("BBB222", "Nord")
    Set cp = wb.Worksheets.Add(After:=ws)
    cp.Name = "code_plan"
    cp.Range("A1:F1").Value = Array("export_key", "type", "sub_type", "data_type", "coding", "de")
    cp.Range("A2:F2").Value = Array("survey_session_id", "standard_field", Empty, "string", Empty, Empty)
    cp.Range("A3:F3").Value = Array("Q1", "question", "multiple_choice", "number", Empty, "In welchem Bezirk?")
    cp.Range("A4:F4").Value = Array("Q1", "answer_option", "radio", "number", 1, "Nord")
    cp.Range("A5:F5").Value = Array("Q1", "answer_option", "radio", "number", 2, "Süd")
    Application.DisplayAlerts = False
    wb.SaveAs p, 51
    wb.Close False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = prevSU
    mListImport.LiReadFile p, "2ask"
    mListImport.LiPlan
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Kill p
    Check "L20", "Unkodierter 2ask-Export mit Codeplan: Antworttexte werden den Optionen zugeordnet (nicht leer)", _
        LiPV(plan, 2, "In welchem Bezirk?") = "Süd" And LiPV(plan, 3, "In welchem Bezirk?") = "Nord" And mListImport.liNotes.Count = 0, _
        "Werte „" & LiTxt(LiPV(plan, 2, "In welchem Bezirk?")) & "“, „" & LiTxt(LiPV(plan, 3, "In welchem Bezirk?")) & "“, Hinweise " & mListImport.liNotes.Count
    Exit Sub
f:  Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Check "L20", "Unkodierter 2ask-Export", False, "Laufzeitfehler: " & Err.Description
End Sub

' Blatt aus Zeilen (Array von Arrays) bauen
Private Function LiSh(ByVal nm As String, ByVal rows As Variant, Optional ByVal csv As Boolean = False) As cLiSheet
    Dim r As Long, c As Long, nc As Long, a() As Variant
    For r = 0 To UBound(rows)
        If UBound(rows(r)) + 1 > nc Then nc = UBound(rows(r)) + 1
    Next
    ReDim a(1 To UBound(rows) + 1, 1 To nc)
    For r = 0 To UBound(rows)
        For c = 0 To UBound(rows(r))
            a(r + 1, c + 1) = rows(r)(c)
        Next
    Next
    Set LiSh = mListImport.LiSheetFromArray(nm, a, csv)
End Function

Private Sub LiLoad(ParamArray sheets() As Variant)
    Dim col As New Collection, i As Long
    For i = 0 To UBound(sheets): col.Add sheets(i): Next
    mListImport.LiLoadSheets col
End Sub

Private Function LiCol(ByVal nm As String) As cLiCol
    Dim c As cLiCol
    For Each c In mListImport.liCols
        If c.Name = nm Then Set LiCol = c: Exit Function
    Next
End Function

Private Function LiHasQ(ByVal prefix As String) As Boolean
    Dim q As cLiQ
    For Each q In mListImport.liQs
        If Left$(q.ID, Len(prefix)) = prefix Then LiHasQ = True: Exit Function
    Next
End Function

' Wert im Importplan: Zeile r (2 = erster Eintrag), Spalte nach Feldname
Private Function LiPV(ByVal plan As cMap, ByVal r As Long, ByVal nm As String) As Variant
    Dim d As Variant, c As Long
    d = plan.GetVal("data")
    LiPV = "#fehlt"
    If r > UBound(d, 1) Then Exit Function
    For c = 1 To UBound(d, 2)
        If d(1, c) = nm Then LiPV = d(r, c): Exit Function
    Next
End Function

Private Function LiTxt(ByVal v As Variant) As String
    If IsEmpty(v) Then LiTxt = "(leer)" Else LiTxt = CStr(v)
End Function

Private Function L02Sheet() As cLiSheet
    Set L02Sheet = LiSh("Anmeldungen", Array(Array("Name", "Teilnahme", "", ""), Array("", "Mo", "Di", "Mi"), _
        Array("A", "x", "", "x"), Array("B", "X", "x", ""), Array("C", "", "x", "x"), Array("D", "x", "x", "x")))
End Function

Private Function L06Sheet() As cLiSheet
    Set L06Sheet = LiSh("Vorgänge", Array(Array("Vorgang", "Status"), Array("A", "erledigt"), Array("B", "offen"), Array("C", "erl."), _
        Array("D", "erledigt"), Array("E", "offen"), Array("F", "Erledigt")))
End Function

Private Sub L01()
    Dim t As cLiTab
    On Error GoTo f
    LiLoad LiSh("Liste", Array(Array("Fortbildung 2025"), Array(), Array("Name", "Ort", "Betrag"), Array("Anna", "Mainz", 10), _
        Array("Ben", "Mainz", 20), Array("Cem", "Worms", 5), Array("Dora", "Bingen", 7), Array("Eva", "Mainz", 8), Array("Summe", "", 50)))
    Set t = mListImport.liTables(1)
    Check "L01", "Liste übernehmen: Tabelle mit Titel, Leerzeile und Summenzeile erkannt (Kopf Zeile 3, 5 Einträge, 1 Summenzeile)", _
        mListImport.liTables.Count = 1 And t.HdrRow = 3 And t.DataRowCount() = 5 And t.CountClass("summe") = 1 And t.Title = "Fortbildung 2025", _
        "Tabellen " & mListImport.liTables.Count & ", Kopf Zeile " & t.HdrRow & ", Einträge " & t.DataRowCount() & ", Titel " & t.Title
    Exit Sub
f:  Check "L01", "Liste übernehmen: Tabelle mit Titel und Summenzeile erkannt", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L02()
    Dim c As cLiCol, plan As cMap
    On Error GoTo f
    LiLoad L02Sheet()
    mListImport.LiChooseTable 1
    Set c = LiCol("Teilnahme")
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L02", "Zweizeiliger Kopf; Kreuzspalten Mo, Di, Mi werden eine Mehrfachauswahl „Teilnahme“", _
        c.FType = FT_MEHRFACH And c.Opts.Count = 3 And LiPV(plan, 2, "Teilnahme") = "Mo;Mi" And mListImport.LiRecordCount() = 4, _
        c.Opts.Count & " Optionen, erster Eintrag „" & LiTxt(LiPV(plan, 2, "Teilnahme")) & "“"
    Exit Sub
f:  Check "L02", "Kreuzspalten werden eine Mehrfachauswahl", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L03()
    Dim c As cLiCol, plan As cMap
    On Error GoTo f
    LiLoad LiSh("Telefon", Array(Array("Name", "Funktion", "Telefon"), Array("Stadtverwaltung"), Array("A", "Leitung", "0123 45-1"), _
        Array("B", "Sachbearbeitung", "0123 45-2"), Array("Freie Träger"), Array("C", "Leitung", "0123 98-3")))
    mListImport.LiChooseTable 1
    Set c = LiCol("Gruppe")
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L03", "Gruppenzeilen werden ein Auswahlfeld an jedem Eintrag darunter", _
        c.FType = FT_AUSWAHL And c.Opts.Count = 2 And LiPV(plan, 4, "Gruppe") = "Freie Träger" And mListImport.LiRecordCount() = 3, _
        c.Opts.Count & " Optionen, dritter Eintrag „" & LiTxt(LiPV(plan, 4, "Gruppe")) & "“"
    Exit Sub
f:  Check "L03", "Gruppenzeilen werden ein Feld", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L04()
    Dim plan As cMap, ok As Boolean, s As String
    On Error GoTo f
    LiLoad LiSh("Typen", Array(Array("Datum", "Betrag", "PLZ", "Telefon", "E-Mail", "Anzahl"), _
        Array(DateSerial(2025, 1, 8), "1.234,50 €", "01067", "0123 45-3456", "a@b.de", 3), _
        Array("12.3.25", "99,00 €", "55116", "0611 98 76 54", "c@d.de", 5), _
        Array(DateSerial(2025, 2, 1), "10 €", "55122", "0123/45-2000", "e@f.de", 7), _
        Array(DateSerial(2025, 3, 1), "5 €", "55124", "+49 123 555012", "g@h.de", 9), _
        Array(DateSerial(2025, 4, 1), "7,5 €", "55126", "0123 44 55 66", "i@j.de", 11)))
    mListImport.LiChooseTable 1
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    ok = LiCol("Datum").FType = FT_DATUM And LiCol("Betrag").FType = FT_GELD And LiCol("PLZ").FType = FT_TEXT And _
         LiCol("Telefon").FType = FT_TELEFON And LiCol("E-Mail").FType = FT_EMAIL And LiCol("Anzahl").FType = FT_ZAHL
    ok = ok And LiPV(plan, 3, "Datum") = DateSerial(2025, 3, 12) And LiPV(plan, 2, "Betrag") = 1234.5 And LiPV(plan, 2, "PLZ") = "01067"
    s = LiCol("Datum").FType & "/" & LiCol("Betrag").FType & "/" & LiCol("PLZ").FType & "/" & LiCol("Telefon").FType & "/" & _
        LiCol("E-Mail").FType & "/" & LiCol("Anzahl").FType & "; „12.3.25“ = " & LiTxt(LiPV(plan, 3, "Datum")) & "; Betrag " & LiTxt(LiPV(plan, 2, "Betrag"))
    Check "L04", "Typvorschläge: Datum (auch „12.3.25“), Geld, PLZ mit Null als Text, Telefon, E-Mail, Nummer", ok, s
    Exit Sub
f:  Check "L04", "Typvorschläge", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L05()
    Dim plan As cMap, rows() As Variant, i As Long
    On Error GoTo f
    ReDim rows(0 To 10)
    rows(0) = Array("Vorgang", "Frist")
    For i = 1 To 9: rows(i) = Array("V" & i, DateSerial(2025, i, 1)): Next
    rows(10) = Array("V10", "KW 14")
    LiLoad LiSh("Fristen", rows)
    mListImport.LiChooseTable 1
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L05", "Nicht passender Wert („KW 14“ im Datumsfeld): Rückfrage, Wert bleibt leer, Hinweis im Bericht", _
        LiCol("Frist").FType = FT_DATUM And LiHasQ("typ:") And IsEmpty(LiPV(plan, 11, "Frist")) And mListImport.liNotes.Count = 1, _
        "Hinweise " & mListImport.liNotes.Count
    Exit Sub
f:  Check "L05", "Nicht passender Wert", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L06()
    Dim plan As cMap
    On Error GoTo f
    LiLoad L06Sheet()
    mListImport.LiChooseTable 1
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L06", "Abkürzung „erl.“ als „erledigt“ vorgeschlagen; Groß-/Kleinschreibung zusammengeführt (2 Optionen)", _
        LiHasQ("abk:") And LiCol("Status").Opts.Count = 2 And LiPV(plan, 4, "Status") = "erledigt" And LiPV(plan, 7, "Status") = "erledigt", _
        LiCol("Status").Opts.Count & " Optionen, „erl.“ wird „" & LiTxt(LiPV(plan, 4, "Status")) & "“"
    Exit Sub
f:  Check "L06", "Abkürzungen", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L07()
    Dim plan As cMap, c As cLiCol
    On Error GoTo f
    LiLoad LiSh("Maßnahmen", Array(Array("Maßnahme", "Themen"), Array("A", "Sport, Gesundheit"), Array("B", "Bildung; Sprache"), _
        Array("C", "Sport"), Array("D", "Gesundheit / Bildung"), Array("E", "Sprache")))
    mListImport.LiChooseTable 1
    Set c = LiCol("Themen")
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L07", "Mehrere Werte je Zelle (Komma, Semikolon, Schrägstrich) werden eine Mehrfachauswahl", _
        c.FType = FT_MEHRFACH And c.Opts.Count = 4 And LiPV(plan, 2, "Themen") = "Sport;Gesundheit" And LiPV(plan, 5, "Themen") = "Gesundheit;Bildung", _
        c.Opts.Count & " Optionen, „" & LiTxt(LiPV(plan, 5, "Themen")) & "“"
    Exit Sub
f:  Check "L07", "Mehrere Werte je Zelle", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L08()
    Dim plan As cMap
    On Error GoTo f
    LiLoad LiSh("Belegung", Array(Array("Belegung 2025"), Array(), Array("Einrichtung", "Jan", "Feb", "Mrz", "Apr"), _
        Array("Kita A", 10, 11, 12, Empty), Array("Kita B", 20, Empty, 22, 23), Array("Summe", 30, 11, 34, 23)))
    mListImport.LiChooseTable 1
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L08", "Monate als Spalten werden je Zeile und Monat ein Eintrag (Jahr aus dem Titel, leere Monate entfallen)", _
        LiHasQ("pivot") And mListImport.LiRecordCount() = 6 And LiPV(plan, 2, "Monat") = DateSerial(2025, 1, 1) And LiPV(plan, 7, "Monat") = DateSerial(2025, 4, 1) _
        And LiCol("Einrichtung").FType = FT_AUSWAHL, "Einträge " & mListImport.LiRecordCount() & ", erster Monat " & LiTxt(LiPV(plan, 2, "Monat"))
    Exit Sub
f:  Check "L08", "Monate als Spalten", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L09()
    Dim plan As cMap
    On Error GoTo f
    LiLoad LiSh("Antrag", Array(Array("Antrag auf Zuschuss"), Array(), Array("Antragsteller:", "Verein X"), Array("Telefon:", "0123 456789"), _
        Array("Teilnehmende:", 24), Array("Zuschuss:", 2400)))
    mListImport.LiChooseTable 1
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L09", "Formular (Bezeichnung: Wert) wird ein Eintrag mit passenden Typen", _
        mListImport.liTables(1).Kind = "formular" And mListImport.LiRecordCount() = 1 And LiCol("Teilnehmende").FType = FT_ZAHL And _
        LiCol("Telefon").FType = FT_TELEFON And LiCol("Zuschuss").FType = FT_GELD And LiPV(plan, 2, "Antragsteller") = "Verein X", _
        "Art " & mListImport.liTables(1).Kind & ", Felder " & mListImport.liCols.Count
    Exit Sub
f:  Check "L09", "Formular", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L10()
    Dim plan As cMap
    On Error GoTo f
    LiLoad LiSh("Kita A", Array(Array("Artikel", "Menge", "Einheit", "Preis"), Array("Kreide", 5, "Packung", 4.2), Array("Knete", 2, "Dose", 3.9))), _
           LiSh("Kita B", Array(Array("Artikel", "Anzahl", "Einheit", "Preis"), Array("Bälle", 3, "Stück", 12)))
    mListImport.LiChooseTable 1
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L10", "Gleich aufgebaute Blätter zusammengeführt, Blattname als Feld, „Anzahl“ und „Menge“ zusammengelegt", _
        mListImport.liTables(1).Kind = "zusammen" And mListImport.LiRecordCount() = 3 And LiCol("Anzahl") Is Nothing And _
        LiPV(plan, 4, "Menge") = 3 And LiPV(plan, 4, "Blatt") = "Kita B" And LiCol("Preis").FType = FT_GELD, _
        "Einträge " & mListImport.LiRecordCount() & ", Menge im dritten Eintrag " & LiTxt(LiPV(plan, 4, "Menge"))
    Exit Sub
f:  Check "L10", "Blätter zusammenführen", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L11()
    Dim plan As cMap
    On Error GoTo f
    LiLoad LiSh("export.csv", Array(Array("FALLNR", "EINGANG", "BETRAG", "AKTIV"), Array("0004711", "2025-01-14", "1450.00", "J"), _
        Array("0004712", "2025-01-20", "2380.50", "N")), True)
    mListImport.LiChooseTable 1
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L11", "CSV: lesbare Feldnamen, Dezimalpunkt (1450.00 = 1.450), ISO-Datum, führende Nullen, J/N", _
        LiPV(plan, 2, "Betrag") = 1450 And LiPV(plan, 3, "Betrag") = 2380.5 And LiPV(plan, 2, "Eingang") = DateSerial(2025, 1, 14) And _
        LiPV(plan, 2, "Fallnummer") = "0004711" And LiPV(plan, 3, "Aktiv") = False And LiCol("Betrag").FType = FT_GELD, _
        "Betrag " & LiTxt(LiPV(plan, 2, "Betrag")) & ", Fallnummer " & LiTxt(LiPV(plan, 2, "Fallnummer"))
    Exit Sub
f:  Check "L11", "CSV-Export", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L12()
    Dim plan As cMap
    On Error GoTo f
    LiLoad LiSh("Inventar", Array(Array("Raum", "Gerät"), Array("1.12", "Laptop"), Array("2.05", "Drucker"), Array("1.14", "Laptop")))
    mListImport.LiChooseTable 1
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L12", "Raum „1.12“ (in Excel als Text gespeichert) bleibt Text", _
        LiCol("Raum").FType = FT_TEXT And LiPV(plan, 2, "Raum") = "1.12", "Wert " & LiTxt(LiPV(plan, 2, "Raum"))
    Exit Sub
f:  Check "L12", "Raum bleibt Text", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L13()
    Dim errNo As Long
    On Error GoTo f
    LiLoad LiSh("KW 20", Array(Array("Uhrzeit", "Montag", "Dienstag", "Mittwoch"), Array("08:00", "Kurs", "", ""), _
        Array("09:00", "", "Chor", ""), Array("10:00", "", "", "Tanz")))
    On Error Resume Next
    mListImport.LiChooseTable 1
    errNo = Err.Number
    Err.Clear
    On Error GoTo f
    Check "L13", "Wochenraster erkannt und mit verständlicher Meldung abgelehnt", _
        mListImport.liTables(1).Kind = "raster" And errNo = vbObjectError + 513, "Art " & mListImport.liTables(1).Kind
    Exit Sub
f:  Check "L13", "Wochenraster", False, "Laufzeitfehler: " & Err.Description
End Sub

Private Sub L14()
    Check "L14", "Optionsnamen ohne Komma und Semikolon („Sonstiges, und zwar“ wird „Sonstiges und zwar“)", _
        mListImport.LiOptName("Sonstiges, und zwar") = "Sonstiges und zwar" And mListImport.LiOptName(" a;b ") = "a b"
End Sub

Private Sub L15()
    Dim plan As cMap, c As cLiCol
    On Error GoTo f
    LiLoad L06Sheet()
    mListImport.LiChooseTable 1
    mListImport.LiSetFieldPart "S:status", "type", FT_TEXT
    mListImport.LiSetFieldPart "S:status", "name", "Bearbeitungsstand"
    mListImport.LiPlan
    Set c = LiCol("Bearbeitungsstand")
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Check "L15", "Änderung durch die Nutzerin oder den Nutzer (Name, Typ Text) bleibt beim Neuplanen erhalten; Abkürzungsfrage entfällt", _
        c.FType = FT_TEXT And Not LiHasQ("abk:") And LiPV(plan, 4, "Bearbeitungsstand") = "erl.", "Wert „" & LiTxt(LiPV(plan, 4, "Bearbeitungsstand")) & "“"
    Exit Sub
f:  Check "L15", "Änderung durch die Nutzerin oder den Nutzer", False, "Laufzeitfehler: " & Err.Description
End Sub

' 2ask: kleiner kodierter Export mit Codeplan wird erzeugt und gelesen
Private Sub L16()
    Dim p As String, wb As Workbook, ws As Worksheet, cp As Worksheet, plan As cMap, prevSU As Boolean, i As Long, r As Long, ok As Boolean, s As String
    On Error GoTo f
    p = ThisWorkbook.Path & "\_Selbsttest_2ask.xlsx"
    If Len(Dir$(p)) > 0 Then Kill p
    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "results"
    ws.Range("A1:F1").Value = Array("ID", "In welchem Bezirk?", "Welche Angebote?", "Welche Angebote?", "Zusammenarbeit", "Zusammenarbeit")
    ws.Range("A2:F2").Value = Array("", "", "Answer-Option: Fachberatung", "Answer-Option: Supervision", "Answer: kompliziert - unkompliziert", "Answer: langsam - schnell")
    ws.Range("A3:F3").Value = Array("{}", "{}", "{}", "{}", "{}", "{}")
    ws.Range("A4:F4").Value = Array("survey_session_id", "Q1", "Q2_01", "Q2_02", "Q3_1", "Q3_2")
    ws.Range("A5:F5").Value = Array("AAA111", 2, 1, Empty, 4, 2)
    ws.Range("A6:F6").Value = Array("BBB222", 1, 1, 2, -99, 5)
    Set cp = wb.Worksheets.Add(After:=ws)
    cp.Name = "code_plan"
    cp.Range("A1:F1").Value = Array("export_key", "type", "sub_type", "data_type", "coding", "de")
    cp.Range("A2:F2").Value = Array("survey_session_id", "standard_field", Empty, "string", Empty, Empty)
    cp.Range("A3:F3").Value = Array("Q1", "question", "multiple_choice", "number", Empty, "In welchem Bezirk?")
    cp.Range("A4:F4").Value = Array("Q1", "answer_option", "radio", "number", 1, "Nord")
    cp.Range("A5:F5").Value = Array("Q1", "answer_option", "radio", "number", 2, "Süd")
    cp.Range("A6:F6").Value = Array("Q2_01", "question", "multiple_choice", Empty, Empty, "Welche Angebote?")
    cp.Range("A7:F7").Value = Array("Q2_01", "answer_option", "checkbox", "number", 1, "Fachberatung")
    cp.Range("A8:F8").Value = Array("Q2_02", "answer_option", "checkbox", "number", 2, "Supervision")
    cp.Range("A9:F9").Value = Array("Q3_1", "question", "matrix_table", Empty, Empty, "Zusammenarbeit")
    r = 10
    cp.Range("A" & r & ":F" & r).Value = Array("Q3_1", "answer", "radio", "number", Empty, "kompliziert"): r = r + 1
    For i = 1 To 5: cp.Range("A" & r & ":F" & r).Value = Array("Q3_1", "answer_option", "radio", "number", i, "[" & i & "]"): r = r + 1: Next
    cp.Range("A" & r & ":F" & r).Value = Array("Q3_2", "answer", "radio", "number", Empty, "langsam"): r = r + 1
    For i = 1 To 5: cp.Range("A" & r & ":F" & r).Value = Array("Q3_2", "answer_option", "radio", "number", i, "[" & i & "]"): r = r + 1: Next
    Application.DisplayAlerts = False
    wb.SaveAs p, 51
    wb.Close False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = prevSU
    mListImport.LiReadFile p, "2ask"
    mListImport.LiPlan
    Set plan = mListImport.LiBuildPlan("x", "", "", "")
    Kill p
    ok = mListImport.liMode = "2ask" And LiCol("In welchem Bezirk?").FType = FT_AUSWAHL And LiCol("In welchem Bezirk?").Opts(1) = "Nord" And _
         LiPV(plan, 2, "In welchem Bezirk?") = "Süd" And LiPV(plan, 3, "Welche Angebote?") = "Fachberatung;Supervision" And _
         LiCol("kompliziert - unkompliziert").FType = FT_ZAHL And LiPV(plan, 2, "kompliziert - unkompliziert") = 4 And _
         IsEmpty(LiPV(plan, 3, "kompliziert - unkompliziert")) And InStr(mListImport.LiAskSummary(), "1 umkodierte fehlende Antwort") > 0
    Check "L16", "2ask-Export mit Codeplan: Codes werden Optionen (Reihenfolge des Codeplans), Mehrfachauswahl aus Einzelspalten, Skala als Nummer, -99 leer", ok, _
        "Bezirk „" & LiTxt(LiPV(plan, 2, "In welchem Bezirk?")) & "“, Angebote „" & LiTxt(LiPV(plan, 3, "Welche Angebote?")) & "“"
    Exit Sub
f:  Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Check "L16", "2ask-Export mit Codeplan", False, "Laufzeitfehler: " & Err.Description
End Sub

' Echte Übernahme in eine neue Tabelle (Optionen in Reihenfolge, Mehrfachauswahl, Bezeichnungsfeld)
Private Sub L17()
    Dim plan As cMap, errs As Collection, valid As Collection, upd As Long, tid As String, fidT As String, ops As String, oid As Variant, t As cTableData
    Dim v As String, a As Variant, i As Long, nm As String
    On Error GoTo f
    mUI.gQuiet = True
    LiLoad L02Sheet()
    mListImport.LiChooseTable 1
    nm = "Selbsttest Liste " & Format$(Now, "hhnnss")
    Set plan = mListImport.LiBuildPlan(nm, "", "", "S:name")
    Set errs = mIO.CheckImport(plan, valid, upd)
    tid = mIO.ExecuteImport(plan, valid)
    mUI.gQuiet = False
    Set t = mData.TData(tid)
    fidT = mCfg.FieldIdByName(tid, "Teilnahme")
    For Each oid In mCfg.OptionIds(fidT)
        ops = ops & IIf(Len(ops) > 0, ",", "") & mCfg.OptName(CStr(oid))
    Next
    a = Split(NzS(t.Cell(1, t.ColOf(fidT))), ";")
    For i = 0 To UBound(a): v = v & IIf(Len(v) > 0, ",", "") & mCfg.OptName(CStr(a(i))): Next
    Check "L17", "Übernahme in eine neue Tabelle: 4 Einträge, Optionen Mo, Di, Mi in dieser Reihenfolge, erster Eintrag Mo und Mi, Bezeichnungsfeld „Name“", _
        errs.Count = 0 And t.RowCount = 4 And ops = "Mo,Di,Mi" And v = "Mo,Mi" And mCfg.FieldName(mCfg.LabelField(tid)) = "Name", _
        "Einträge " & t.RowCount & ", Optionen " & ops & ", erster Eintrag " & v
    Exit Sub
f:  mUI.gQuiet = False
    Check "L17", "Übernahme in eine neue Tabelle", False, "Laufzeitfehler: " & Err.Description
End Sub
