Attribute VB_Name = "mCapture"
'==============================================================================
' mCapture - Erfassungsdateien: Aus Tabellen eigenstaendige, makrofreie
' Excel-Dateien (XLSX) fuer andere Personen erzeugen und Ruecklaeufe einlesen.
'   Zweck "Neu"     : leere Zeilen bzw. Formular, Ruecklauf legt Eintraege an.
'   Zweck "Pruefen" : vorhandene Eintraege vorbelegt (graue Spalten nur zur
'                     Anzeige, gelbe zum Ausfuellen), je Empfaenger aufgeteilt;
'                     der Ruecklauf aktualisiert die Eintraege ueber ihre ID.
' Mehrfachauswahl erscheint als eine Ankreuzspalte je Option (senkrechter
' Titel, Auswahl "Ja" oder leer). Eine Datei kann mehrere Blaetter (Tabellen)
' enthalten. Jede Datei traegt eine eindeutige Datei-ID; doppelte Ruecklaeufe
' werden erkannt, Aenderungen seit dem Versand als Konflikt ausgewiesen.
'==============================================================================
Option Explicit

Private Const INFO_SHEET As String = "_Info"
Private Const LIST_SHEET As String = "Listen"
Private Const CAPT_SHEET As String = "Erfassung"
Private Const MARK As String = "DBK-ERFASSUNG"
Private Const R_BAND As Long = 9         ' Band "Bestehende Einträge" (Prüfdateien)
Private Const R_GROUP As Long = 10       ' Gruppenzeile (Titel ueber Ankreuzspalten)
Private Const R_HEAD As Long = 11        ' Spaltenkoepfe
Private Const HDR_MIN_W As Double = 84   ' Mindestbreite des Kopfbereichs (Excel-Spaltenbreiten)
Private Const MULTI_MAX As Long = 30     ' mehr Optionen: Mehrfachauswahl als Text "A; B"
Public gCapDiag As String               ' letzte technische Auffälligkeit beim Erzeugen (Prüfung/Diagnose)
Private Const ART_NEU As String = "Neue Einträge sammeln (am häufigsten)"
Private Const ART_PRUEF As String = "Vorhandene Einträge prüfen und aktualisieren lassen (erweitert)"
Private Const ABS_LISTE As String = "Aus einer festen Liste wählen (empfohlen)"
Private Const ABS_TAB As String = "Aus einer Tabelle dieser Datei wählen (z. B. Einrichtungen)"
Private Const ABS_FREI As String = "Frei eintragen lassen"
Private Const ABS_NEIN As String = "Nicht abfragen"
Private Const SEP_MARK As String = "§"   ' Spalte A der Trennzeile "Neue Einträge" (wird beim Einlesen übersprungen)

Private mSel As String
Private mRowMap As cMap
Private mRespMap As cMap          ' Zeile der Rückmeldungen -> Array(Erfassung, Name, Zustand)
Private mRespCol As Long          ' erste Spalte der Rückmeldungen
Private Const RS_OPEN As Long = 0
Private Const RS_GOT As Long = 1
Private Const RS_NODATA As Long = 2
Private mRestoreAsked As cMap   ' je Einlesen: Erfassung wiederhergestellt (True) / abgelehnt (False)

'==============================================================================
' Optionen einer Festlegung (Spalte "Optionen", Schluessel=Wert)
'==============================================================================
Public Function COpt(ByVal cid As String, ByVal key As String, Optional ByVal dflt As String = "") As String
    COpt = KVGet(mCfg.GetS(SH_CAP, cid, "Optionen"), key, dflt)
End Function

Private Function IsCheck(ByVal cid As String) As Boolean
    IsCheck = (COpt(cid, "art") = "pruefen")
End Function

Private Function ParentOf(ByVal cid As String) As String
    ParentOf = COpt(cid, "eltern")
End Function

'==============================================================================
' Teilnehmende (wer soll antworten?) und Rückmeldungen
'   absart: frei = Name frei eintragen | liste = feste Liste | tabelle = Einträge
'           einer Tabelle (Bezeichnung) | nein = nicht abfragen
'   Die feste Liste steht in _System "Teilnehmer_<ID>" (eine Zeile je Name).
'==============================================================================
Public Function SenderMode(ByVal cid As String) As String
    SenderMode = COpt(cid, "absart", "")
    If Len(SenderMode) = 0 Then SenderMode = IIf(mCfg.GetS(SH_CAP, cid, "Absender") = "Nein", "nein", "frei")
End Function

' Erwartete Teilnehmende (Anzeige-Namen, Reihenfolge wie eingegeben bzw. alphabetisch)
Public Function ExpectedSenders(ByVal cid As String) As Collection
    Dim res As New Collection, seen As New cMap, v As Variant, tid As String, t As cTableData, r As Long, lf As String, nm As String
    Set ExpectedSenders = res
    Select Case SenderMode(cid)
        Case "tabelle"
            tid = COpt(cid, "abstab")
            If Not mCfg.TableExists(tid) Then Exit Function
            Set t = mData.TData(tid)
            lf = mCfg.LabelField(tid)
            For r = 1 To t.RowCount
                If Not t.IsArchived(r) Then
                    nm = Trim$(mData.RecordLabel(tid, t.IdAt(r)))
                    If Len(nm) > 0 And Not seen.Exists(LCase$(nm)) Then seen.SetVal LCase$(nm), True: res.Add nm
                End If
            Next
        Case Else
            For Each v In Split(Replace(mCfg.SysGet("Teilnehmer_" & cid, ""), vbCr, ""), vbLf)
                nm = Trim$(CStr(v))
                If Len(nm) > 0 And Not seen.Exists(LCase$(nm)) Then seen.SetVal LCase$(nm), True: res.Add nm
            Next
    End Select
End Function

' Eingegangene Rückmeldungen: Absender (klein) -> Array(Anzeigename, Zeitpunkt, Zeilen)
Public Function Responses(ByVal cid As String) As cMap
    Dim res As New cMap, t As cTableData, r As Long, s As String
    Set Responses = res
    If Not mApp.SheetExists(SH_RET) Then Exit Function
    Set t = mCfg.TD(SH_RET)
    For r = 1 To t.RowCount
        If CStr(t.Val_(r, "Erfassung")) = cid Then
            s = Trim$(CStr(t.Val_(r, "Absender")))
            If Len(s) = 0 Then s = Trim$(CStr(t.Val_(r, "Empfaenger")))
            If Len(s) = 0 Then s = "(ohne Absender) " & CStr(t.Val_(r, "Datei"))
            res.SetVal LCase$(s), Array(s, t.Val_(r, "Importiert"), t.Val_(r, "Zeilen"))
        End If
    Next
End Function

' Fehlanzeigen: Teilnehmende, die ohne Datei gemeldet haben, dass es nichts zu melden gibt
' (z. B. per Mail „bei uns keine Kinder mit Förderbedarf“). Name (klein) -> Array(Anzeigename, Zeitpunkt)
Public Function NoDataMarks(ByVal cid As String) As cMap
    Dim res As New cMap, v As Variant, p As Variant
    Set NoDataMarks = res
    For Each v In Split(Replace(mCfg.SysGet("Fehlanzeige_" & cid, ""), vbCr, ""), vbLf)
        p = Split(CStr(v), vbTab)
        If UBound(p) >= 0 Then
            If Len(Trim$(CStr(p(0)))) > 0 Then res.SetVal LCase$(Trim$(CStr(p(0)))), Array(Trim$(CStr(p(0))), IIf(UBound(p) >= 1, InvDate(CStr(p(UBound(p)))), Empty))
        End If
    Next
End Function

Public Sub SetNoData(ByVal cid As String, ByVal nm As String, ByVal on_ As Boolean)
    Dim m As cMap, k As Variant, a As Variant, s As String
    Set m = NoDataMarks(cid)
    If on_ Then
        m.SetVal LCase$(nm), Array(nm, Now)
    ElseIf m.Exists(LCase$(nm)) Then
        m.Remove LCase$(nm)
    End If
    For Each k In m.Keys
        a = m.GetVal(CStr(k))
        s = s & IIf(Len(s) > 0, vbLf, "") & a(0) & vbTab & FmtStamp(a(1))
    Next
    mCfg.SysSet "Fehlanzeige_" & cid, s
End Sub

' Wie viele der erwarteten Teilnehmenden haben geantwortet (Datei eingelesen oder Fehlanzeige)?
Public Function AnsweredCount(ByVal cid As String, Optional ByRef nNoData As Long) As Long
    Dim exp_ As Collection, got As cMap, nod As cMap, k As Variant
    Set exp_ = ExpectedSenders(cid)
    Set got = Responses(cid)
    Set nod = NoDataMarks(cid)
    nNoData = 0
    For Each k In exp_
        If got.Exists(LCase$(CStr(k))) Then
            AnsweredCount = AnsweredCount + 1
        ElseIf nod.Exists(LCase$(CStr(k))) Then
            AnsweredCount = AnsweredCount + 1
            nNoData = nNoData + 1
        End If
    Next
End Function

Public Sub SetSenderList(ByVal cid As String, ByVal txt As String)
    Dim v As Variant, s As String, seen As New cMap, nm As String
    For Each v In Split(Replace(txt, vbCr, ""), vbLf)
        nm = Trim$(CStr(v))
        If Len(nm) > 0 And Not seen.Exists(LCase$(nm)) Then seen.SetVal LCase$(nm), True: s = s & IIf(Len(s) > 0, vbLf, "") & nm
    Next
    mCfg.SysSet "Teilnehmer_" & cid, s
End Sub

Private Function AbsLabel(ByVal key As String) As String
    Select Case key
        Case "liste": AbsLabel = ABS_LISTE
        Case "tabelle": AbsLabel = ABS_TAB
        Case "nein": AbsLabel = ABS_NEIN
        Case Else: AbsLabel = ABS_FREI
    End Select
End Function

Private Function AbsKey(ByVal lbl As String) As String
    Select Case lbl
        Case ABS_LISTE: AbsKey = "liste"
        Case ABS_TAB: AbsKey = "tabelle"
        Case ABS_NEIN: AbsKey = "nein"
        Case Else: AbsKey = "frei"
    End Select
End Function

' Hauptfestlegung (fuer Blaetter: die Datei, zu der sie gehoeren)
Private Function RootOf(ByVal cid As String) As String
    RootOf = ParentOf(cid)
    If Len(RootOf) = 0 Then RootOf = cid
End Function

Private Function ChildIds(ByVal cid As String) As Collection
    Dim res As New Collection, ids As Variant, i As Long
    ids = CaptIds()
    For i = 0 To ArrLen(ids) - 1
        If ParentOf(CStr(ids(i))) = cid Then res.Add CStr(ids(i))
    Next
    Set ChildIds = res
End Function

' Alle Blaetter einer Datei: zuerst die Hauptfestlegung, dann die weiteren Blaetter
Private Function SheetDefs(ByVal root As String) As Collection
    Dim res As New Collection, c As Variant
    res.Add root
    For Each c In ChildIds(root)
        res.Add CStr(c)
    Next
    Set SheetDefs = res
End Function

'==============================================================================
' Seite "Erfassungsdateien"
'==============================================================================
Public Sub Render()
    Dim ws As Worksheet, ids As Variant, rr As Long, i As Long, cid As String, c As Variant, rTop As Long, firstRoot As String
    Dim avail As Double, wName As Double, wArt As Double, panelRight As Boolean
    EnsureSheets
    Set ws = mRender.Sheet(SH_CAPTURE)
    mRender.BeginPage ws
    Set mRowMap = New cMap
    Set mRespMap = Nothing
    ' Breiten passend zum Fenster: Liste links, "Rückmeldungen" rechts, bei wenig Platz darunter
    avail = mRender.PageRight() - ws.Columns(3).Left
    wName = 230: wArt = 150
    If avail < 740 Then wName = Application.Max(150, wName - (740 - avail) * 0.6): wArt = Application.Max(120, wArt - (740 - avail) * 0.4)
    mRender.SetColWidthPt ws, 3, 26
    mRender.SetColWidthPt ws, 4, wName
    mRender.SetColWidthPt ws, 5, 120
    mRender.SetColWidthPt ws, 6, wArt
    mRender.SetColWidthPt ws, 7, 104
    mRender.SetColWidthPt ws, 8, 110
    mRender.SetColWidthPt ws, 9, 24
    mRender.SetColWidthPt ws, 10, 250
    mRender.SetColWidthPt ws, 11, 96
    panelRight = (avail >= ws.Columns(12).Left - ws.Columns(3).Left)
    mRender.DrawTitle ws, ChrW(&HE8A5), "Erfassungsdateien", "Eine Excel-Liste an viele verschicken: alle Antworten landen gesammelt in Ihrer Tabelle"
    mRender.FlowBegin ws, R_TABS, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowBtn "c_new", "Neue Erfassung", "mCapture.Cmd_New", "primary", ChrW(&HE710)
    mRender.FlowBtn "c_imp", "Rücklauf einlesen…", "mCapture.Cmd_Import", "secondary", ChrW(&HE8B5)
    mRender.FlowBtn "c_how", "So funktioniert's", "mCapture.Cmd_Intro", "secondary", ChrW(&HE897)
    mRender.FlowBtn "c_log", "Rücklaufprotokoll", "mCapture.Cmd_Log", "ghost", ChrW(&HE81C)
    mRender.FlowEnd
    ' Ablauf als Bild: vier Schritte
    ws.Rows(R_TOOLS).RowHeight = 8
    ws.Rows(R_CHIPS).RowHeight = 86
    DrawFlowStrip ws, ws.Rows(R_CHIPS).Top + 2, ws.Cells(R_CHIPS, 3).Left, Application.Min(IIf(panelRight, ws.Columns(12).Left, ws.Columns(9).Left) - ws.Columns(3).Left - 8, avail - 8), 80
    ' Liste der Erfassungen
    ids = CaptIds()
    For i = 0 To ArrLen(ids) - 1
        If Len(ParentOf(CStr(ids(i)))) = 0 Then
            If Len(firstRoot) = 0 Then firstRoot = CStr(ids(i))
        End If
    Next
    If Len(mSel) = 0 Or Len(mCfg.GetS(SH_CAP, mSel, "ID")) = 0 Then mSel = firstRoot
    rr = R_DATA
    ws.Cells(rr, 3).Value = "Ihre Erfassungen"
    ws.Cells(rr, 3).Font.Size = 12: ws.Cells(rr, 3).Font.Name = UI_FONT_SEMI
    ws.Rows(rr).RowHeight = 26
    rTop = rr
    rr = rr + 1
    mRender.FlowBegin ws, rr, mRender.ContentLeft(ws), ws.Cells(rr, 9).Left
    mRender.FlowText "c_hint", "Zeile anklicken, dann:", CMuted(), 9, False, 8
    mRender.FlowBtn "c_exp", "Datei erzeugen…", "mCapture.Cmd_Export", "primary", ChrW(&HEDE1)
    mRender.FlowBtn "c_edit", "Bearbeiten…", "mCapture.Cmd_Edit", "secondary", ChrW(&HE70F)
    mRender.FlowBtn "c_people", "Teilnehmende…", "mCapture.Cmd_Senders", "secondary", ChrW(&HE716)
    mRender.FlowBtn "c_more", "Weiteres…", "mCapture.Cmd_MoreMenu", "secondary", ChrW(&HE712)
    mRender.FlowEnd
    rr = rr + 1
    ws.Range(ws.Cells(rr, 3), ws.Cells(rr, 8)).Value = Array("", "Erfassung", "Tabelle", "Art", "Rückmeldungen", "Letzter Rücklauf")
    With ws.Range(ws.Cells(rr, 3), ws.Cells(rr, 8))
        .Interior.Color = CSurface(): .Font.Color = CMuted(): .Font.Size = 9.5
    End With
    For i = 0 To ArrLen(ids) - 1
        cid = CStr(ids(i))
        If Len(ParentOf(cid)) = 0 Then
            rr = rr + 1
            DrawCaptRow ws, rr, cid, False
            For Each c In ChildIds(cid)
                rr = rr + 1
                DrawCaptRow ws, rr, CStr(c), True
            Next
        End If
    Next
    If ArrLen(ids) = 0 Then
        rr = rr + 1
        ws.Cells(rr, 4).Value = "Noch keine Erfassung. Mit „Neue Erfassung“ geht es los; „So funktioniert's“ zeigt den Ablauf."
        ws.Cells(rr, 4).Font.Color = CMuted()
        ws.Rows(rr).RowHeight = 26
    End If
    mRender.GridBorders ws.Range(ws.Cells(R_DATA + 2, 3), ws.Cells(rr, 8))
    ' rechts: Rückmeldungen der gewählten Erfassung
    If Len(mSel) > 0 Then
        If panelRight Then DrawResponses ws, rTop, RootOf(mSel), 10 Else DrawResponses ws, rr + 2, RootOf(mSel), 4
    End If
    mRender.EndPage ws
    ' beim ersten Besuch den Ablauf kurz erklären
    If mCfg.SysGet("ErfassungIntro", "") <> "1" And Not mApp.gReadOnly And Not mUI.gQuiet Then
        mCfg.SysSet "ErfassungIntro", "1"
        mApp.MarkCfgDirty
        Application.OnTime Now, mRender.MacroRef("mCapture.Cmd_Intro")
    End If
End Sub

' Ablauf in vier Schritten (Formen): Liste festlegen -> verschicken -> einlesen -> Gesamtliste
Private Sub DrawFlowStrip(ByVal ws As Worksheet, ByVal y As Single, ByVal x As Single, ByVal w As Single, ByVal h As Single)
    Dim i As Long, cw As Single, gap As Single, s As Shape, xx As Single, icons As Variant, ttl As Variant, txt As Variant, acc As Long
    icons = Array(ChrW(&HE8FD), ChrW(&HE715), ChrW(&HE8B5), ChrW(&HE9D9))
    ttl = Array("Liste festlegen", "Datei verschicken", "Rücklauf einlesen", "Gesamtliste")
    txt = Array("Spalten wählen", "per E-Mail an alle", "Dateien auswählen", "alles in einer Tabelle")
    acc = CAccentFill()
    gap = 26
    cw = (w - gap * 3) / 4
    For i = 0 To 3
        xx = x + i * (cw + gap)
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, xx, y, cw, h)
        s.Name = "fl_c" & i
        s.Adjustments.Item(1) = 0.1
        s.Fill.ForeColor.RGB = Blend(acc, RGB(255, 255, 255), 0.9)
        s.Line.ForeColor.RGB = Blend(acc, RGB(255, 255, 255), 0.65): s.Line.Weight = 0.75
        s.Placement = xlFreeFloating
        ' Nummer im Kreis
        Set s = ws.Shapes.AddShape(msoShapeOval, xx + 10, y + 10, 22, 22)
        s.Name = "fl_n" & i
        s.Fill.ForeColor.RGB = acc: s.Line.Visible = msoFalse
        With s.TextFrame2
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = CStr(i + 1)
            .TextRange.Font.Size = 10: .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Name = UI_FONT
            .TextRange.Font.Fill.ForeColor.RGB = COnAccent()
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
        s.Placement = xlFreeFloating
        mRender.TextShape ws, "fl_i" & i, CStr(icons(i)), xx + cw - 34, y + 10, 24, 24, 15, CAccent(), False, 2
        ws.Shapes("fl_i" & i).TextFrame2.TextRange.Font.Name = ICON_FONT
        mRender.TextShape ws, "fl_t" & i, CStr(ttl(i)), xx + 10, y + 38, cw - 20, 18, 10.5, CText(), True, 1
        mRender.TextShape ws, "fl_x" & i, CStr(txt(i)), xx + 10, y + 56, cw - 20, 18, 8.5, CMuted(), False, 1
        If i < 3 Then
            Set s = ws.Shapes.AddShape(msoShapeRightArrow, xx + cw + 5, y + h / 2 - 7, gap - 10, 14)
            s.Name = "fl_a" & i
            s.Fill.ForeColor.RGB = Blend(acc, RGB(255, 255, 255), 0.45): s.Line.Visible = msoFalse
            s.Placement = xlFreeFloating
        End If
    Next
End Sub

Private Sub DrawCaptRow(ByVal ws As Worksheet, ByVal rr As Long, ByVal cid As String, ByVal isChild As Boolean)
    Dim tid As String, exp_ As Collection, got As cMap, n As Long, k As Variant, last As Variant, a As Variant
    tid = mCfg.GetS(SH_CAP, cid, "Tabelle")
    ws.Cells(rr, 3).Value = IIf(isChild, "", ChrW(&HE8A5))
    ws.Cells(rr, 3).Font.Name = ICON_FONT: ws.Cells(rr, 3).Font.Color = CAccent(): ws.Cells(rr, 3).HorizontalAlignment = xlCenter
    ws.Cells(rr, 4).NumberFormat = "@"
    If isChild Then
        ws.Cells(rr, 4).Value = "   " & ChrW(&H21B3) & "  Blatt „" & SheetTitle(cid) & "“"
        ws.Cells(rr, 4).Font.Color = CMuted()
    Else
        ws.Cells(rr, 4).Value = mCfg.GetS(SH_CAP, cid, "Name")
        ws.Cells(rr, 4).Font.Name = UI_FONT_SEMI
    End If
    ws.Cells(rr, 5).NumberFormat = "@"
    ws.Cells(rr, 5).Value = mCfg.TableName(tid)
    ws.Cells(rr, 6).Value = IIf(IsCheck(cid), "Bestand prüfen (erweitert)", "Neue Einträge sammeln")
    If IsCheck(cid) And Len(COpt(cid, "teilen")) > 0 And mCfg.FieldExists(COpt(cid, "teilen")) Then
        ws.Cells(rr, 6).Value = "Prüfen je „" & mCfg.FieldName(COpt(cid, "teilen")) & "“"
    End If
    If Not isChild Then
        Set got = Responses(cid)
        Set exp_ = ExpectedSenders(cid)
        ws.Cells(rr, 7).NumberFormat = "@"
        If exp_.Count > 0 Then
            ws.Cells(rr, 7).Value = AnsweredCount(cid) & " von " & exp_.Count
        Else
            ws.Cells(rr, 7).Value = CStr(got.Count)
        End If
        For Each k In got.Keys
            a = got.GetVal(CStr(k))
            If IsEmpty(last) Then
                last = a(1)
            ElseIf IsDate(a(1)) Then
                If CDate(a(1)) > CDate(last) Then last = a(1)
            End If
        Next
        If Not IsEmpty(last) Then ws.Cells(rr, 8).Value = FmtStamp(last)
    End If
    ws.Rows(rr).RowHeight = 24
    mRowMap.SetVal CStr(rr), cid
    If cid = mSel Then ws.Range(ws.Cells(rr, 3), ws.Cells(rr, 8)).Interior.Color = CAccentSoft()
End Sub

' Rechts neben der Liste: wer hat schon geantwortet? (Teilnehmende mit Haken, sonst "offen")
Private Sub DrawResponses(ByVal ws As Worksheet, ByVal r0 As Long, ByVal cid As String, ByVal c1 As Long)
    Dim exp_ As Collection, got As cMap, k As Variant, rr As Long, n As Long, done_ As Long, a As Variant, s As Shape, x As Single, w As Single
    Dim seen As New cMap, pct As Double, nod As cMap, nNoData As Long
    Set exp_ = ExpectedSenders(cid)
    Set got = Responses(cid)
    Set nod = NoDataMarks(cid)
    Set mRespMap = New cMap
    mRespCol = c1
    x = ws.Cells(r0, c1).Left
    w = ws.Cells(r0, c1 + 2).Left - x
    ws.Cells(r0, c1).Value = "Rückmeldungen"
    ws.Cells(r0, c1).Font.Size = 12: ws.Cells(r0, c1).Font.Name = UI_FONT_SEMI
    rr = r0 + 1
    done_ = AnsweredCount(cid, nNoData)
    ws.Cells(rr, c1).NumberFormat = "@"
    If exp_.Count > 0 Then
        ws.Cells(rr, c1).Value = done_ & " von " & exp_.Count & " haben geantwortet" & _
                                 IIf(nNoData > 0, " (davon " & nNoData & IIf(nNoData = 1, " Fehlanzeige)", " Fehlanzeigen)"), "")
        pct = done_ / exp_.Count
    Else
        ws.Cells(rr, c1).Value = got.Count & IIf(got.Count = 1, " Rückmeldung", " Rückmeldungen") & " eingelesen"
    End If
    ws.Cells(rr, c1).Font.Color = CMuted()
    ws.Cells(rr, c1).VerticalAlignment = xlTop
    If exp_.Count > 0 Then
        ' Fortschrittsbalken unter der Zahl
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, ws.Rows(rr).Top + 17, w - 8, 6)
        s.Name = "rp_track": s.Line.Visible = msoFalse: s.Fill.ForeColor.RGB = CSurface2(): s.Placement = xlFreeFloating
        If pct > 0 Then
            Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, ws.Rows(rr).Top + 17, Application.Max(6, (w - 8) * pct), 6)
            s.Name = "rp_fill": s.Line.Visible = msoFalse: s.Fill.ForeColor.RGB = CSuccess(): s.Placement = xlFreeFloating
        End If
    End If
    rr = rr + 1
    ws.Range(ws.Cells(rr, c1), ws.Cells(rr, c1 + 1)).Value = Array(IIf(exp_.Count > 0, "Teilnehmende", "Absender"), "Stand")
    With ws.Range(ws.Cells(rr, c1), ws.Cells(rr, c1 + 1))
        .Interior.Color = CSurface(): .Font.Color = CMuted(): .Font.Size = 9.5
    End With
    ' zuerst die offenen (wer fehlt noch?), dann die eingegangenen, dann die Fehlanzeigen
    For Each k In exp_
        If Not got.Exists(LCase$(CStr(k))) And Not nod.Exists(LCase$(CStr(k))) Then
            rr = rr + 1: n = n + 1
            ResponseRow ws, rr, CStr(k), RS_OPEN, Empty, c1
            mRespMap.SetVal CStr(rr), Array(cid, CStr(k), RS_OPEN)
        End If
    Next
    For Each k In exp_
        If got.Exists(LCase$(CStr(k))) Then
            rr = rr + 1: n = n + 1
            a = got.GetVal(LCase$(CStr(k)))
            ResponseRow ws, rr, CStr(k), RS_GOT, a(1), c1
            seen.SetVal LCase$(CStr(k)), True
        End If
    Next
    For Each k In exp_
        If Not got.Exists(LCase$(CStr(k))) And nod.Exists(LCase$(CStr(k))) Then
            rr = rr + 1: n = n + 1
            a = nod.GetVal(LCase$(CStr(k)))
            ResponseRow ws, rr, CStr(k), RS_NODATA, a(1), c1
            mRespMap.SetVal CStr(rr), Array(cid, CStr(k), RS_NODATA)
        End If
    Next
    ' Antworten von Absendern, die (noch) nicht in der Liste stehen
    For Each k In got.Keys
        If Not seen.Exists(CStr(k)) Then
            rr = rr + 1: n = n + 1
            a = got.GetVal(CStr(k))
            ResponseRow ws, rr, CStr(a(0)) & IIf(exp_.Count > 0, "  (nicht in der Liste)", ""), RS_GOT, a(1), c1
        End If
    Next
    If n = 0 Then
        rr = rr + 1
        ws.Cells(rr, c1).Value = "Noch keine Rückmeldung. Unter „Teilnehmende…“ tragen Sie ein, an wen die Datei geht."
        ws.Cells(rr, c1).Font.Color = CMuted(): ws.Cells(rr, c1).Font.Size = 9
        ws.Range(ws.Cells(rr, c1), ws.Cells(rr, c1 + 1)).Merge
        ws.Cells(rr, c1).WrapText = True
        ws.Rows(rr).RowHeight = Application.Max(ws.Rows(rr).RowHeight, 36)
    End If
    mRender.GridBorders ws.Range(ws.Cells(r0 + 2, c1), ws.Cells(rr, c1 + 1))
    If exp_.Count > 0 And n > 0 Then
        ' Hinweis: Antwort ohne Datei (z. B. per Mail) per Klick erfassen
        rr = rr + 1
        ws.Cells(rr, c1).Value = "Antwort ohne Datei, z. B. per Mail „bei uns nichts zu melden“? Auf „offen“ klicken " & ChrW(&H2192) & " Fehlanzeige."
        ws.Cells(rr, c1).Font.Color = CMuted(): ws.Cells(rr, c1).Font.Size = 8.5
        ws.Range(ws.Cells(rr, c1), ws.Cells(rr, c1 + 1)).Merge
        ws.Cells(rr, c1).WrapText = True
        ws.Cells(rr, c1).VerticalAlignment = xlTop
        ws.Rows(rr).RowHeight = 30
    End If
End Sub

' Zeile der Rückmeldungen: offen (anklickbar) / eingegangen mit Datum / Fehlanzeige (anklickbar)
Private Sub ResponseRow(ByVal ws As Worksheet, ByVal rr As Long, ByVal nm As String, ByVal state As Long, ByVal stamp As Variant, ByVal c1 As Long)
    ws.Cells(rr, c1).NumberFormat = "@"
    ws.Cells(rr, c1 + 1).NumberFormat = "@"
    Select Case state
        Case RS_GOT
            ws.Cells(rr, c1).Value = ChrW(&H2713) & "  " & nm
            ws.Cells(rr, c1).Font.Color = CText()
            ws.Cells(rr, c1).Characters(1, 1).Font.Color = CSuccess()
            ws.Cells(rr, c1 + 1).Value = IIf(IsDate(stamp), FmtDate(CDate(stamp)), "eingegangen")
            ws.Cells(rr, c1 + 1).Font.Color = CSuccess()
        Case RS_NODATA
            ws.Cells(rr, c1).Value = ChrW(&H2713) & "  " & nm
            ws.Cells(rr, c1).Font.Color = CText()
            ws.Cells(rr, c1).Characters(1, 1).Font.Color = CMuted()
            ws.Cells(rr, c1 + 1).Value = "Fehlanzeige" & IIf(IsDate(stamp), " · " & Format$(CDate(stamp), "dd.mm."), "")
            ws.Cells(rr, c1 + 1).Font.Color = CMuted()
        Case Else
            ws.Cells(rr, c1).Value = ChrW(&H25CB) & "  " & nm
            ws.Cells(rr, c1).Font.Color = CMuted()
            ws.Cells(rr, c1).Characters(1, 1).Font.Color = CBorderStrong()
            ws.Cells(rr, c1 + 1).Value = "offen " & ChrW(&H25BE)
            ws.Cells(rr, c1 + 1).Font.Color = CWarn()
    End Select
    ws.Cells(rr, c1 + 1).Font.Size = 9
    If ws.Rows(rr).RowHeight < 22 Then ws.Rows(rr).RowHeight = 22
End Sub

' Klick auf eine offene Zeile bzw. eine Fehlanzeige: kleines Menü, kein Formular
Private Sub ResponseMenu(ByVal info As Variant)
    Dim p As String
    p = CStr(info(0)) & Chr$(31) & CStr(info(1))
    If CLng(info(2)) = RS_NODATA Then
        mRender.Popup Array("Fehlanzeige zurücknehmen (wieder „offen“)|mCapture.Cmd_NoDataOff|" & p)
    Else
        mRender.Popup Array("Fehlanzeige eintragen: " & CStr(info(1)) & " hat nichts zu melden|mCapture.Cmd_NoDataOn|" & p)
    End If
End Sub

Public Sub Cmd_NoDataOn()
    NoDataCmd True
End Sub

Public Sub Cmd_NoDataOff()
    NoDataCmd False
End Sub

Private Sub NoDataCmd(ByVal on_ As Boolean)
    Dim p As Variant
    On Error GoTo f
    p = Split(mRender.MenuParam(), Chr$(31))
    If UBound(p) < 1 Then Exit Sub
    If Not mApp.RequireWritable("Das Eintragen einer Fehlanzeige") Then Exit Sub
    SetNoData CStr(p(0)), CStr(p(1)), on_
    mLog.Add "Erfassung", "", "", "", IIf(on_, "Fehlanzeige eingetragen: ", "Fehlanzeige zurückgenommen: ") & p(1) & " (" & mCfg.GetS(SH_CAP, CStr(p(0)), "Name") & ")"
    mApp.MarkCfgDirty
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Fehlanzeige"
End Sub

Private Function SheetTitle(ByVal cid As String) As String
    SheetTitle = mCfg.GetS(SH_CAP, cid, "Titel")
    If Len(SheetTitle) = 0 Then SheetTitle = mCfg.TableName(mCfg.GetS(SH_CAP, cid, "Tabelle"))
End Function

Public Sub Cmd_GoPage()
    On Error GoTo f
    mNav.GoPage "Erfassung"
    Exit Sub
f: mApp.Boundary "Erfassungsdateien"
End Sub

Public Sub OnSelect(ByVal target As Range)
    Dim cid As String, prevEv As Boolean, info As Variant
    ' Klick in die Rückmeldungen: offene Zeile bzw. Fehlanzeige -> kleines Menü
    If Not mRespMap Is Nothing And target.Cells.Count = 1 Then
        If (target.Column = mRespCol Or target.Column = mRespCol + 1) And mRespMap.Exists(CStr(target.Row)) Then
            info = mRespMap.GetVal(CStr(target.Row))
            ResponseMenu info
            ' Auswahl wieder unter die Navigation, damit derselbe Eintrag erneut anklickbar ist
            prevEv = Application.EnableEvents
            Application.EnableEvents = False
            On Error Resume Next
            If ActiveSheet.Name = SH_CAPTURE Then ActiveSheet.Cells(R_TOP, 1).Select
            On Error GoTo 0
            Application.EnableEvents = prevEv
            Exit Sub
        End If
    End If
    If mRowMap Is Nothing Then Exit Sub
    cid = CStr(mRowMap.GetVal(CStr(target.Row), ""))
    If Len(cid) = 0 Or cid = mSel Then Exit Sub
    mSel = cid
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    mNav.RenderCurrent
    Application.EnableEvents = prevEv
End Sub

Public Sub OnDoubleClick(ByVal target As Range, ByRef Cancel As Boolean)
    Dim cid As String
    If mRowMap Is Nothing Then Exit Sub
    cid = CStr(mRowMap.GetVal(CStr(target.Row), ""))
    If Len(cid) = 0 Then Exit Sub
    Cancel = True
    mSel = cid
    Cmd_Export
End Sub

Private Function CaptIds() As Variant
    Dim t As cTableData, r As Long, a() As Variant, n As Long
    Set t = mCfg.TD(SH_CAP)
    If t.RowCount = 0 Then CaptIds = Array(): Exit Function
    ReDim a(0 To t.RowCount - 1)
    For r = 1 To t.RowCount
        If Len(NzS(t.Cell(r, 1))) > 0 Then a(n) = CStr(t.Cell(r, 1)): n = n + 1
    Next
    If n = 0 Then CaptIds = Array(): Exit Function
    ReDim Preserve a(0 To n - 1)
    CaptIds = a
End Function

' Blaetter und Spalten fuer Erfassungsdateien in aelteren Dateien nachruesten
Public Sub EnsureSheets()
    Dim sh As Variant, ws As Worksheet, hdr As Variant, i As Long, lo_ As ListObject, prevEv As Boolean
    If mApp.SheetExists(SH_CAP) And mApp.SheetExists(SH_RET) And mApp.SheetExists(SH_CAPTURE) Then
        EnsureColumns
        Exit Sub
    End If
    If mApp.gReadOnly Then RaiseUser "Erfassungsdateien sind in dieser schreibgeschützten Datei noch nicht eingerichtet."
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    If Not mApp.SheetExists(SH_CAPTURE) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(SH_SEARCH))
        ws.Name = SH_CAPTURE
        ws.Cells.Font.Name = UI_FONT
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If
    For Each sh In Array(SH_CAP, SH_RET)
        If Not mApp.SheetExists(CStr(sh)) Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = CStr(sh)
            hdr = mCfg.CfgHeaders(CStr(sh))
            For i = 0 To UBound(hdr)
                ws.Cells(1, i + 1).Value = hdr(i)
            Next
            ws.Columns.NumberFormat = "@"
            Set lo_ = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(hdr) + 1)), , xlYes)
            lo_.Name = Mid$(CStr(sh), 2)
            lo_.TableStyle = "TableStyleLight1"
            ws.Visible = xlSheetVeryHidden
        End If
    Next
    mCfg.InvalidateAll
    EnsureColumns
    Application.EnableEvents = prevEv
End Sub

Private Sub EnsureColumns()
    If mApp.gReadOnly Then Exit Sub
    mCfg.EnsureColumn SH_CAP, "Optionen"
    mCfg.EnsureColumn SH_RET, "Empfaenger"
End Sub

'==============================================================================
' Festlegen / Bearbeiten (gefuehrt in wenigen Schritten)
'==============================================================================
Public Sub Cmd_New()
    On Error GoTo f
    EditDefinition "", ""
    Exit Sub
f: mApp.Boundary "Neue Erfassungsdatei"
End Sub

' Erklärung mit Ablaufbild (beim ersten Besuch automatisch)
Public Sub Cmd_Intro()
    Dim f As frmCaptureIntro
    On Error GoTo f
    Set f = New frmCaptureIntro
    f.Setup
    f.Show
    If f.StartNew Then
        Unload f
        Cmd_New
    Else
        Unload f
    End If
    Exit Sub
f: mApp.Boundary "So funktioniert's"
End Sub

' Teilnehmende einer Erfassung pflegen: an wen ging die Datei? (für die Rückmeldungsübersicht)
Public Sub Cmd_Senders()
    Dim root As String, fp As frmProps, ids As Variant, names() As Variant, i As Long, mode As String, cur As String, k As Variant, opt As String
    On Error GoTo f
    If Not RequireSel() Then Exit Sub
    If Not mApp.RequireWritable("Das Ändern der Teilnehmenden") Then Exit Sub
    root = RootOf(mSel)
    ids = mCfg.TableIds()
    ReDim names(0 To ArrLen(ids) - 1)
    For i = 0 To ArrLen(ids) - 1: names(i) = mCfg.TableName(CStr(ids(i))): Next
    mode = SenderMode(root)
    For Each k In ExpectedSenders(root)
        If mode <> "tabelle" Then cur = cur & IIf(Len(cur) > 0, vbCrLf, "") & CStr(k)
    Next
    Set fp = New frmProps
    fp.Init "Teilnehmende: " & mCfg.GetS(SH_CAP, root, "Name"), _
            "An wen geht die Erfassungsdatei? Mit dieser Liste sehen Sie rechts unter „Rückmeldungen“, wer schon geantwortet hat. " & _
            "In neu erzeugten Dateien wählen die Empfänger ihren Namen aus dieser Liste.", 560
    fp.AddCombo "absart", "Woher kommen die Namen?", Array(ABS_LISTE, ABS_TAB, ABS_FREI), AbsLabel(IIf(mode = "nein", "frei", mode))
    fp.AddText "abslist", "Teilnehmende (ein Name je Zeile)", cur, "Tipp: Namen aus Excel oder einer E-Mail kopieren und hier einfügen.", True
    fp.ShowWhen "abslist", "absart", ABS_LISTE
    fp.AddCombo "abstab", "Tabelle mit den Teilnehmenden", names, IIf(mCfg.TableExists(COpt(root, "abstab")), mCfg.TableName(COpt(root, "abstab")), CStr(names(0)))
    fp.ShowWhen "abstab", "absart", ABS_TAB
    If Not fp.ShowDialog("Übernehmen") Then Unload fp: Exit Sub
    opt = mCfg.GetS(SH_CAP, root, "Optionen")
    opt = KVSet(opt, "absart", AbsKey(fp.GetText("absart")))
    opt = KVSet(opt, "abstab", IIf(AbsKey(fp.GetText("absart")) = "tabelle", mCfg.TableIdByName(fp.GetText("abstab")), ""))
    mCfg.SetV SH_CAP, root, "Optionen", opt
    If AbsKey(fp.GetText("absart")) = "liste" Then SetSenderList root, fp.GetText("abslist")
    If mCfg.GetS(SH_CAP, root, "Absender") = "Nein" Then
        mCfg.SetV SH_CAP, root, "Absender", "Ja"
        mCfg.SetV SH_CAP, root, "AbsenderFeld", EnsureSenderField(mCfg.GetS(SH_CAP, root, "Tabelle"))
    End If
    Unload fp
    mApp.MarkCfgDirty
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Teilnehmende"
End Sub

' Seltener Gebrauchtes zur gewählten Erfassung
Public Sub Cmd_MoreMenu()
    On Error GoTo f
    If Not RequireSel() Then Exit Sub
    mRender.Popup Array("Rücklauf-Übersicht (Dateien je Empfänger)|mCapture.Cmd_Status", _
                        "Weiteres Blatt hinzufügen…|mCapture.Cmd_AddSheet", _
                        "Rücklaufprotokoll|mCapture.Cmd_Log", _
                        "Erfassung entfernen…|mCapture.Cmd_Delete||1|1")
    Exit Sub
f: mApp.Boundary "Erfassung"
End Sub

Public Sub Cmd_Edit()
    On Error GoTo f
    If Not RequireSel() Then Exit Sub
    EditDefinition mSel, ParentOf(mSel)
    Exit Sub
f: mApp.Boundary "Erfassungsdatei bearbeiten"
End Sub

Public Sub Cmd_AddSheet()
    On Error GoTo f
    If Not RequireSel() Then Exit Sub
    If mCfg.GetS(SH_CAP, RootOf(mSel), "Modus") = "Formular" Then
        mUI.Alert "Formular-Dateien bestehen aus genau einem Blatt. Für mehrere Blätter die Art „Liste“ wählen.", "Blatt hinzufügen"
        Exit Sub
    End If
    EditDefinition "", RootOf(mSel)
    Exit Sub
f: mApp.Boundary "Blatt hinzufügen"
End Sub

Private Function RequireSel() As Boolean
    If Len(mSel) = 0 Or Len(mCfg.GetS(SH_CAP, mSel, "ID")) = 0 Then
        mUI.Alert "Bitte zuerst eine Erfassungsdatei in der Liste anklicken.", "Erfassungsdateien"
        Exit Function
    End If
    RequireSel = True
End Function

' Legt eine Erfassungsdatei bzw. ein Blatt fest; liefert die ID ("" = abgebrochen).
' presetTid: weiteres Blatt für diese Tabelle ohne Schritt 1 (bei Auswahl mehrerer Tabellen)
Private Function EditDefinition(ByVal cid As String, ByVal parent As String, Optional ByVal presetTid As String = "", _
                                Optional ByVal presetCheck As Boolean = False) As String
    Dim extra As String, newTid As String, minStp As Long, sfx As String, xt As Variant, nOk As Long
    Dim fp As frmProps, ids As Variant, names() As Variant, i As Long, tid As String, ok As Boolean, flds As Collection, f As cField
    Dim items() As Variant, fids() As String, n As Long, pre As String, sel As String, a As Variant, m As cMap, picked As String, curTid As String
    Dim rows As Long, nm As String, isChild As Boolean, check As Boolean, opt As String, anz As String, vNames() As Variant, vids As Variant
    Dim sNames() As Variant, sIds() As String, k As Variant, isNew As Boolean, curView As String, curSplit As String
    If Not mApp.RequireWritable("Das Festlegen einer Erfassungsdatei") Then Exit Function
    EnsureSheets
    isChild = (Len(parent) > 0)
    isNew = (Len(cid) = 0)
    ids = mCfg.TableIds()
    If ArrLen(ids) = 0 Then mUI.Alert "Bitte zuerst eine Tabelle anlegen. Die Erfassungsdatei gehört zu einer vorhandenen Tabelle.", "Erfassungsdateien": Exit Function
    ReDim names(0 To ArrLen(ids) - 1)
    For i = 0 To ArrLen(ids) - 1: names(i) = mCfg.TableName(CStr(ids(i))): Next
    If Not isNew Then curTid = mCfg.GetS(SH_CAP, cid, "Tabelle") Else curTid = CStr(ids(0))
    If mNav.gPage = "Tabelle" And isNew And Len(mNav.gTid) > 0 Then curTid = mNav.gTid
    If Not isNew Then opt = mCfg.GetS(SH_CAP, cid, "Optionen")

    ' Assistent als Schrittfolge: Jeder Schritt kann vor (Weiter) und zurueck (Zurueck);
    ' die Angaben bleiben beim Zuruecklaufen erhalten. Gespeichert wird erst am Ende.
    Dim stp As Long, dirn As Long, back As Boolean, nSteps As Long, modeForm As Boolean, viewId As String, splitId As String, newRows As Long
    Dim dTitle As String, dIntro As String, dAbs As Boolean, dAbsLbl As String, dAbsW As String, dFarbe As String, dRows As Long, dHow As Boolean, dProt As Boolean
    Dim dFont As String, dAbsArt As String, dAbsList As String, dAbsTab As String
    tid = curTid
    check = (KVGet(opt, "art") = "pruefen")
    modeForm = (mCfg.GetS(SH_CAP, cid, "Modus") = "Formular")
    If isNew Then nm = "" Else nm = mCfg.GetS(SH_CAP, cid, IIf(isChild, "Titel", "Name"))
    anz = KVGet(opt, "anzeige")
    If isNew Then sel = "" Else sel = mCfg.GetS(SH_CAP, cid, "Felder")
    viewId = KVGet(opt, "ansicht"): splitId = KVGet(opt, "teilen"): newRows = CLng(Val(KVGet(opt, "neuzeilen", "0")))
    dTitle = IIf(isNew, "", mCfg.GetS(SH_CAP, cid, "Titel"))
    dIntro = mCfg.GetS(SH_CAP, cid, "Einleitung")
    dAbs = isNew Or mCfg.GetS(SH_CAP, cid, "Absender") <> "Nein"
    dAbsLbl = KVGet(opt, "abslabel", "Ihr Name / Ihre Stelle")
    dAbsW = KVGet(opt, "absbreite", "Breit")
    dFarbe = KVGet(opt, "farbe", mDocStyle.DefaultStyle())
    dFont = KVGet(opt, "schrift", mDocStyle.DefaultFont())
    dAbsArt = IIf(isNew, "liste", SenderMode(cid))
    dAbsTab = KVGet(opt, "abstab")
    dAbsList = IIf(isNew, "", Replace(mCfg.SysGet("Teilnehmer_" & cid, ""), vbLf, vbCrLf))
    dRows = IIf(isNew, 200, CLng(Val(mCfg.GetS(SH_CAP, cid, "Zeilen"))))
    dHow = (KVGet(opt, "anleitung", "1") = "1")
    dProt = isNew Or mCfg.GetS(SH_CAP, cid, "Schutz") <> "Nein"
    minStp = 1
    If Len(presetTid) > 0 Then
        ' weiteres Blatt aus der Mehrfachauswahl: Tabelle, Zweck und Name stehen fest
        tid = presetTid: check = presetCheck: nm = mCfg.TableName(presetTid): minStp = 2
    End If
    stp = minStp: dirn = 1
    Do While stp <= 6
        If stp < minStp Then stp = minStp
        nSteps = IIf(isChild, 3, IIf(check, 4, 3))
        sfx = IIf(isChild Or Len(extra) > 0, " (Blatt „" & mCfg.TableName(tid) & "“)", "")
        Select Case stp
            Case 1
                ' Schritt 1: Grundlagen
                Set fp = New frmProps
                If isChild Then
                    fp.Init IIf(isNew, "Weiteres Blatt", "Blatt bearbeiten") & " " & StepLabel(1, minStp, check, isChild), _
                            "Ein weiteres Blatt in derselben Datei, z. B. eine zweite Tabelle, die dieselben Empfänger prüfen oder ergänzen.", 560
                    fp.AddText "name", "Name des Blattes *", nm
                Else
                    fp.Init IIf(isNew, "Neue Erfassung", "Erfassung bearbeiten") & " " & StepLabel(1, minStp, check, isChild), _
                            "Sie stellen eine Excel-Liste zusammen, schicken sie an andere (z. B. an alle Kitas) und lesen die ausgefüllten Dateien wieder ein. " & _
                            "Alle Antworten landen gesammelt in Ihrer Tabelle.", 560
                    fp.AddText "name", "Name der Erfassung *", nm, "z. B. „Kinder mit Förderbedarf 2026“, nur für Sie, erscheint nicht in der Datei."
                End If
                fp.AddCombo "tab", "In welche Tabelle sollen die Antworten? *", names, mCfg.TableName(tid), _
                            "Die Spalten der Erfassungsdatei kommen aus dieser Tabelle. Fehlt eine passende Tabelle, legen Sie sie vorher unter „Neue Tabelle“ an."
                fp.AddCombo "art", "Was möchten Sie tun?", Array(ART_NEU, ART_PRUEF), IIf(check, ART_PRUEF, ART_NEU), _
                            "Neue Einträge sammeln: Alle erhalten dieselbe leere Liste (am häufigsten). " & _
                            "Prüfen: Jede Stelle erhält ihre vorhandenen Einträge zum Kontrollieren und Ergänzen (erweitert)."
                If isNew And Not isChild Then
                    ' erweitert: weitere Tabellen als zusätzliche Blätter derselben Datei
                    pre = ""
                    For i = 0 To ArrLen(ids) - 1
                        If ListHas(extra, CStr(ids(i))) Then pre = pre & IIf(Len(pre) > 0, ";", "") & i
                    Next
                    fp.AddCheck "mehrtab", "Erweitert: weitere Tabellen als zusätzliche Blätter in dieselbe Datei", Len(extra) > 0
                    fp.AddMulti "tabs", "Zusätzliche Tabellen (je ein eigenes Blatt)", names, pre
                    fp.ShowWhen "tabs", "mehrtab"
                End If
                If Not isChild Then
                    fp.AddCombo "mode", "Wie sieht die Datei aus?", Array("Liste: jede Stelle trägt mehrere Einträge ein (z. B. alle Kinder einer Kita)", "Formular: genau ein Eintrag je Datei (z. B. Umfrage)"), _
                                IIf(modeForm, "Formular: genau ein Eintrag je Datei (z. B. Umfrage)", "Liste: jede Stelle trägt mehrere Einträge ein (z. B. alle Kinder einer Kita)")
                    fp.ShowWhen "mode", "art", ART_NEU
                End If
                Do
                    If Not fp.ShowDialog("Weiter") Then Unload fp: Exit Function
                    fp.ClearErrors
                    nm = Trim$(fp.GetText("name"))
                    If Len(nm) > 0 Or Not isChild Then
                        Exit Do
                    Else
                        fp.SetError "name", "Bitte einen Namen für das Blatt angeben."
                    End If
                Loop
                newTid = mCfg.TableIdByName(fp.GetText("tab"))
                If isNew And Not isChild Then
                    ' zusätzliche Tabellen (ohne die Haupttabelle) als weitere Blätter
                    extra = ""
                    If fp.GetBool("mehrtab") Then
                        For Each xt In Split(fp.GetText("tabs"), ";")
                            If Len(xt) > 0 Then
                                If CStr(ids(CLng(xt))) <> newTid Then extra = ListAdd(extra, CStr(ids(CLng(xt))))
                            End If
                        Next
                    End If
                End If
                If newTid <> tid Then
                    ' andere Tabelle: Spaltenwahl und Aufteilung gelten nicht mehr
                    tid = newTid
                    anz = "": sel = "": viewId = "": splitId = ""
                End If
                check = (fp.GetIndex("art") = 1)
                If Not isChild Then modeForm = (fp.GetIndex("mode") = 1) And Len(extra) = 0
                Unload fp
            Case 2
                ' Schritt 2: Spalten zur Anzeige (nur Pruefen)
                If check Then
                    Set flds = mCfg.Fields(tid)
                    ReDim items(0 To flds.Count - 1, 0 To 1): ReDim fids(0 To flds.Count - 1)
                    n = 0: pre = ""
                    For Each f In flds
                        If f.Kind <> FT_BUTTON Then
                            items(n, 0) = f.Name
                            items(n, 1) = f.TypeName_
                            fids(n) = f.ID
                            If Len(anz) = 0 And f.ID = mCfg.LabelField(tid) Then pre = pre & IIf(Len(pre) > 0, ";", "") & n
                            If ListHas(anz, f.ID) Then pre = pre & IIf(Len(pre) > 0, ";", "") & n
                            n = n + 1
                        End If
                    Next
                    picked = mUI.ChooseMany("Welche Spalten sollen nur ANGEZEIGT werden? Sie sind in der Datei vorbelegt, grau und nicht änderbar (z. B. Inventarnummer, Bezeichnung, Standort).", _
                                            "Spalten zur Anzeige " & StepLabel(2, minStp, check, isChild) & sfx, TrimItems(items, n), pre, ok, True, back, "Weiter")
                    If back Then stp = stp - 1: dirn = -1: GoTo nextStep
                    If Not ok Then Exit Function
                    anz = ""
                    a = Split(picked, ";")
                    For i = 0 To UBound(a)
                        If Len(a(i)) > 0 Then anz = ListAdd(anz, fids(CLng(a(i))))
                    Next
                Else
                    anz = ""
                End If
            Case 3
                ' Schritt 3: Spalten zum Ausfuellen
                Set flds = mCfg.Fields(tid)
                ReDim items(0 To flds.Count - 1, 0 To 1): ReDim fids(0 To flds.Count - 1)
                n = 0: pre = ""
                For Each f In flds
                    If Capturable(f) And f.Name <> "Erfasst von" And Not ListHas(anz, f.ID) Then
                        items(n, 0) = f.Name
                        items(n, 1) = f.TypeName_ & IIf(f.Required, " · Pflicht", "") & IIf(f.FType = FT_MEHRFACH, " · als Ankreuzspalten", "")
                        fids(n) = f.ID
                        If (Len(sel) = 0 And Not check) Or ListHas(sel, f.ID) Then pre = pre & IIf(Len(pre) > 0, ";", "") & n
                        n = n + 1
                    End If
                Next
                If n = 0 Then mUI.Alert "Die Tabelle hat keine weiteren ausfüllbaren Eigenschaften.", "Erfassungsdateien": Exit Function
                picked = mUI.ChooseMany(IIf(check, "Welche Spalten sollen die Empfänger AUSFÜLLEN oder ändern (gelb)? Vorhandene Werte stehen bereits darin.", _
                                                   "Welche Eigenschaften sollen die Empfänger ausfüllen? Berechnete und automatische Eigenschaften werden nicht abgefragt."), _
                                        "Spalten zum Ausfüllen " & StepLabel(3, minStp, check, isChild) & sfx, TrimItems(items, n), pre, ok, True, back, "Weiter")
                If back Then stp = stp - 1: dirn = -1: GoTo nextStep
                If Not ok Then Exit Function
                If Len(picked) = 0 Then mUI.Alert "Bitte mindestens eine Spalte zum Ausfüllen wählen.", "Erfassungsdateien": GoTo nextStep
                a = Split(picked, ";")
                sel = ""
                For i = 0 To UBound(a)
                    sel = ListAdd(sel, fids(CLng(a(i))))
                Next
                If Not check Then
                    ' Pflichtfelder der Tabelle ohne Standardwert muessen enthalten sein
                    For Each f In flds
                        If f.Required And f.IsEditable And Not f.IsLinkType And Not ListHas(sel, f.ID) And IsBlank(mValid.DefaultValue(f)) Then
                            If Not mUI.Confirm("„" & f.Name & "“ ist in der Tabelle eine Pflichtangabe, wurde aber nicht gewählt. Rückläufe ohne diesen Wert könnten nicht übernommen werden." & vbLf & vbLf & _
                                               "Trotzdem ohne diese Eigenschaft fortfahren?", "Pflichtangabe fehlt", "Fortfahren", "Zurück") Then GoTo nextStep
                        End If
                    Next
                End If
            Case 4
                ' Schritt 4: Eintraege und Aufteilung (nur Pruefen)
                If check Then
                    Set flds = mCfg.Fields(tid)
                    vids = mCfg.ViewIds(tid)
                    ReDim vNames(0 To ArrLen(vids))
                    vNames(0) = "Alle Einträge (ohne archivierte)"
                    curView = vNames(0)
                    For i = 0 To ArrLen(vids) - 1
                        vNames(i + 1) = "Wie Ansicht „" & mCfg.ViewGet(CStr(vids(i)), "Name") & "“"
                        If CStr(vids(i)) = viewId Then curView = vNames(i + 1)
                    Next
                    ReDim sNames(0 To flds.Count): ReDim sIds(0 To flds.Count)
                    sNames(0) = IIf(isChild, "(nicht zuordnen: alle Einträge auf jedem Blatt)", "(nicht aufteilen: eine Datei mit allen Einträgen)")
                    curSplit = sNames(0)
                    n = 1
                    For Each f In flds
                        If CanSplitBy(f) Then
                            sNames(n) = f.Name: sIds(n) = f.ID
                            If f.ID = splitId Then curSplit = f.Name
                            n = n + 1
                        End If
                    Next
                    ReDim Preserve sNames(0 To n - 1)
                    Set fp = New frmProps
                    fp.AllowBack = True
                    fp.Init "Einträge und Empfänger " & StepLabel(4, minStp, check, isChild) & sfx, _
                            IIf(isChild, "Welche Einträge stehen auf diesem Blatt? Über die Zuordnung erhält jeder Empfänger nur seine Einträge (der Wert muss dem Empfängernamen entsprechen).", _
                                         "Welche Einträge werden vorbelegt? Mit „Aufteilen nach“ entsteht je Wert eine eigene Datei mit nur diesen Einträgen, z. B. je Kita."), 560
                    fp.AddCombo "view", "Einträge", vNames, curView
                    fp.AddCombo "split", IIf(isChild, "Dem Empfänger zuordnen über", "Je Empfänger eine eigene Datei, aufgeteilt nach"), sNames, curSplit
                    fp.AddCheck "neu", "Empfänger dürfen zusätzlich neue Einträge ergänzen (leere Zeilen am Ende)", newRows > 0
                    fp.AddText "neuz", "Anzahl leerer Zeilen für neue Einträge", CStr(IIf(newRows > 0, newRows, 10)), "Die Zeilen stehen in einem eigenen Bereich „Neue Einträge“ unter der Liste."
                    fp.ShowWhen "neuz", "neu"
                    If Not fp.ShowDialog(IIf(isChild, IIf(isNew, "Festlegen", "Übernehmen"), "Weiter")) Then
                        back = fp.WentBack
                        Unload fp
                        If back Then stp = stp - 1: dirn = -1: GoTo nextStep
                        Exit Function
                    End If
                    i = fp.GetIndex("view")
                    If i > 0 Then viewId = CStr(vids(i - 1)) Else viewId = ""
                    i = fp.GetIndex("split")
                    If i > 0 Then splitId = sIds(i) Else splitId = ""
                    If fp.GetBool("neu") Then
                        newRows = CLng(Val(fp.GetText("neuz")))
                        If newRows < 1 Then newRows = 10
                        If newRows > 500 Then newRows = 500
                    Else
                        newRows = 0
                    End If
                    Unload fp
                Else
                    viewId = "": splitId = "": newRows = 0
                End If
            Case 5
                ' Schritt 5: Wer füllt aus? (einheitliche Namen -> Übersicht, wer geantwortet hat)
                If Not isChild Then
                    Set fp = New frmProps
                    fp.AllowBack = True
                    fp.Init "Wer füllt aus? " & StepLabel(5, minStp, check, isChild), _
                            "Oben in der Datei tragen die Empfänger ein, wer sie sind (z. B. ihre Kita). Mit einer festen Liste schreiben alle den Namen gleich, " & _
                            "und Sie sehen später auf einen Blick, wer schon geantwortet hat und wer noch fehlt.", 560
                    fp.AddCombo "absart", "Wie wird der Absender angegeben?", Array(ABS_LISTE, ABS_TAB, ABS_FREI, ABS_NEIN), AbsLabel(dAbsArt)
                    fp.AddText "abslist", "Teilnehmende (ein Name je Zeile)", dAbsList, "Zum Beispiel alle 60 Kitas. Die Liste lässt sich später unter „Teilnehmende“ ergänzen.", True
                    fp.ShowWhen "abslist", "absart", ABS_LISTE
                    fp.AddCombo "abstab", "Tabelle mit den Teilnehmenden", names, IIf(mCfg.TableExists(dAbsTab), mCfg.TableName(dAbsTab), CStr(names(0))), _
                                "Zur Auswahl stehen die Bezeichnungen aller aktiven Einträge, z. B. jeder Eintrag der Tabelle „Einrichtungen“."
                    fp.ShowWhen "abstab", "absart", ABS_TAB
                    fp.AddText "abslbl", "Beschriftung des Feldes in der Datei", dAbsLbl
                    fp.ShowWhen "abslbl", "absart", ABS_LISTE & ";" & ABS_TAB & ";" & ABS_FREI
                    Do
                        If Not fp.ShowDialog("Weiter") Then
                            back = fp.WentBack
                            Unload fp
                            If back Then stp = stp - 1: dirn = -1: GoTo nextStep
                            Exit Function
                        End If
                        fp.ClearErrors
                        If fp.GetText("absart") = ABS_LISTE And Len(Trim$(Replace(Replace(fp.GetText("abslist"), vbCr, ""), vbLf, ""))) = 0 Then
                            fp.SetError "abslist", "Bitte mindestens einen Namen eintragen oder „frei eintragen“ wählen."
                        Else
                            Exit Do
                        End If
                    Loop
                    dAbsArt = AbsKey(fp.GetText("absart"))
                    dAbsList = fp.GetText("abslist")
                    dAbsTab = mCfg.TableIdByName(fp.GetText("abstab"))
                    dAbsLbl = Trim$(fp.GetText("abslbl"))
                    If Len(dAbsLbl) = 0 Then dAbsLbl = "Ihr Name / Ihre Stelle"
                    dAbs = (dAbsArt <> "nein")
                    Unload fp
                End If
            Case 6
                ' Schritt 6: Aussehen der Datei (nur Hauptfestlegung): Wichtiges oben, Feinheiten aufklappbar
                If Not isChild Then
                    Set fp = New frmProps
                    fp.AllowBack = True
                    fp.Init "Text und Aussehen " & StepLabel(6, minStp, check, isChild), "So sieht die Datei für die Empfänger aus. Alles lässt sich später ändern.", 560
                    fp.AddText "title", "Überschrift in der Datei", IIf(Len(dTitle) = 0 And isNew, nm, dTitle)
                    fp.AddText "intro", "Kurzer Hinweis für die Empfänger", IIf(isNew And Len(dIntro) = 0, IIf(check, "Bitte prüfen Sie Ihre Einträge, ergänzen Sie die gelben Spalten, speichern Sie die Datei und senden Sie sie zurück.", _
                                                                                             "Bitte die gelb markierten Felder ausfüllen, speichern und die Datei zurücksenden."), dIntro), "", True
                    fp.AddCheck "mehr", "Weitere Einstellungen (Farbstil, Schrift, Zeilen, Blattschutz)", False
                    fp.AddStyle "farbe", "Farbstil der Datei", dFarbe, "Leitfarbe für Kopfzeile und Akzente. Helle Stile erhalten automatisch dunkle Schrift."
                    fp.ShowWhen "farbe", "mehr"
                    fp.AddCombo "schrift", "Schriftart", mDocStyle.DocFonts(), dFont
                    fp.ShowWhen "schrift", "mehr"
                    If Not check Then fp.AddText "rows", "Vorbereitete leere Zeilen", CStr(dRows), "Höchstens 2.000.": fp.ShowWhen "rows", "mehr"
                    If dAbs Then fp.AddCombo "absw", "Breite des Absenderfeldes", Array("Schmal", "Mittel", "Breit", "Sehr breit"), dAbsW: fp.ShowWhen "absw", "mehr"
                    fp.AddCheck "howto", "Kurzanleitung oben in der Datei anzeigen", dHow
                    fp.ShowWhen "howto", "mehr"
                    fp.AddCheck "prot", "Blattschutz (Aufbau kann nicht versehentlich verändert werden)", dProt
                    fp.ShowWhen "prot", "mehr"
                    If Not fp.ShowDialog(IIf(isNew, "Fertig", "Übernehmen")) Then
                        back = fp.WentBack
                        Unload fp
                        If back Then stp = stp - 1: dirn = -1: GoTo nextStep
                        Exit Function
                    End If
                    dTitle = Trim$(fp.GetText("title"))
                    dIntro = Trim$(fp.GetText("intro"))
                    If Not check Then dRows = CLng(Val(fp.GetText("rows")))
                    dProt = fp.GetBool("prot")
                    If dAbs Then dAbsW = fp.GetText("absw")
                    dFarbe = fp.GetText("farbe")
                    dFont = fp.GetText("schrift")
                    dHow = fp.GetBool("howto")
                    Unload fp
                End If
        End Select
        stp = stp + 1
nextStep:
        ' zurueck: nicht benoetigte Schritte ueberspringen
        If dirn = -1 Then
            Do While stp > 1 And Not StepNeeded(stp, check, isChild)
                stp = stp - 1
            Loop
            dirn = 1
        End If
    Loop

    ' Angaben uebernehmen
    Set m = New cMap
    If isChild Then
        m.SetVal "Name", mCfg.GetS(SH_CAP, parent, "Name") & " / " & nm
        m.SetVal "Titel", nm
        m.SetVal "Modus", "Liste"
    Else
        m.SetVal "Name", IIf(Len(nm) = 0, "Erfassung " & mCfg.TableName(tid), nm)
        m.SetVal "Modus", IIf(check Or Not modeForm, "Liste", "Formular")
    End If
    m.SetVal "Tabelle", tid
    m.SetVal "Felder", sel
    opt = KVSet(opt, "art", IIf(check, "pruefen", ""))
    opt = KVSet(opt, "eltern", parent)
    opt = KVSet(opt, "anzeige", anz)
    opt = KVSet(opt, "ansicht", viewId)
    opt = KVSet(opt, "teilen", splitId)
    opt = KVSet(opt, "neuzeilen", IIf(newRows > 0, CStr(newRows), ""))
    If Not isChild Then
        m.SetVal "Titel", dTitle
        m.SetVal "Einleitung", dIntro
        rows = 200
        If Not check Then rows = dRows
        If rows < 1 Then rows = 200
        If rows > 2000 Then rows = 2000
        m.SetVal "Zeilen", CStr(rows)
        m.SetVal "Absender", IIf(dAbs, "Ja", "Nein")
        m.SetVal "Schutz", IIf(dProt, "Ja", "Nein")
        opt = KVSet(opt, "abslabel", dAbsLbl)
        opt = KVSet(opt, "absbreite", dAbsW)
        opt = KVSet(opt, "absart", dAbsArt)
        opt = KVSet(opt, "abstab", IIf(dAbsArt = "tabelle", dAbsTab, ""))
        opt = KVSet(opt, "farbe", dFarbe)
        opt = KVSet(opt, "schrift", dFont)
        opt = KVSet(opt, "anleitung", IIf(dHow, "1", "0"))
    Else
        m.SetVal "Zeilen", IIf(check, "0", "200")
        m.SetVal "Absender", mCfg.GetS(SH_CAP, parent, "Absender")
        m.SetVal "Schutz", mCfg.GetS(SH_CAP, parent, "Schutz")
    End If
    m.SetVal "Optionen", opt
    ' Feld "Erfasst von" bei Bedarf anlegen (je Tabelle)
    If m.GetVal("Absender") = "Ja" Then
        m.SetVal "AbsenderFeld", EnsureSenderField(tid)
    Else
        m.SetVal "AbsenderFeld", ""
    End If
    If isNew Then
        cid = mCfg.NextId("C", 4)
        m.SetVal "ID", cid
        If Not isChild And dAbsArt = "liste" Then SetSenderList cid, dAbsList
        m.SetVal "Erstellt", Now
        m.SetVal "Version", "1"
        mCfg.AddRow SH_CAP, m
        mLog.Add "Erfassungsdatei festgelegt", tid, cid, "", CStr(m.GetVal("Name"))
    Else
        For Each k In m.Keys
            mCfg.SetV SH_CAP, cid, CStr(k), m.GetVal(CStr(k))
        Next
        mCfg.SetV SH_CAP, cid, "Version", CStr(Val(mCfg.GetS(SH_CAP, cid, "Version")) + 1)
        If Not isChild And dAbsArt = "liste" Then SetSenderList cid, dAbsList
        mLog.Add "Erfassungsdatei geändert", tid, cid, "", CStr(m.GetVal("Name"))
    End If
    ' Absenderabfrage gilt fuer alle Blaetter der Datei
    If Not isChild Then
        For Each k In ChildIds(cid)
            mCfg.SetV SH_CAP, CStr(k), "Absender", m.GetVal("Absender")
            mCfg.SetV SH_CAP, CStr(k), "Schutz", m.GetVal("Schutz")
            If m.GetVal("Absender") = "Ja" Then mCfg.SetV SH_CAP, CStr(k), "AbsenderFeld", EnsureSenderField(mCfg.GetS(SH_CAP, CStr(k), "Tabelle"))
        Next
    End If
    mApp.MarkCfgDirty
    EditDefinition = cid
    ' als weiteres Blatt einer Mehrfachauswahl: Rückmeldung gibt der Hauptablauf
    If Len(presetTid) > 0 Then Exit Function
    ' weitere gewählte Tabellen: je ein Blatt in derselben Datei (Spalten je Tabelle wählen)
    For Each xt In Split(extra, ";")
        If Len(xt) > 0 Then
            If Len(EditDefinition("", cid, CStr(xt), check)) > 0 Then nOk = nOk + 1
        End If
    Next
    mSel = cid
    mNav.GoPage "Erfassung"
    If isChild Then
        mUI.Alert "Das Blatt „" & m.GetVal("Titel") & "“ gehört jetzt zur Datei „" & mCfg.GetS(SH_CAP, parent, "Name") & "“.", "Erfassungsdatei"
    ElseIf mUI.Confirm("Die Erfassung „" & m.GetVal("Name") & "“ ist fertig eingerichtet" & IIf(nOk > 0, " (" & nOk + 1 & " Blätter)", "") & _
                       ". Jetzt die Excel-Datei zum Verschicken erzeugen?", "Erfassung", "Datei erzeugen…", "Später") Then
        Cmd_Export
    End If
End Function

' Schrittanzeige "(2/4)": Position unter den tatsaechlich benoetigten Schritten
Private Function StepLabel(ByVal stp As Long, ByVal minStp As Long, ByVal check As Boolean, ByVal isChild As Boolean) As String
    Dim s As Long, k As Long, n As Long
    For s = minStp To 6
        If StepNeeded(s, check, isChild) Then
            n = n + 1
            If s <= stp Then k = n
        End If
    Next
    StepLabel = "(" & k & "/" & n & ")"
End Function

' Wird ein Schritt des Assistenten fuer diese Festlegung gebraucht? (fuer "Zurueck")
Private Function StepNeeded(ByVal stp As Long, ByVal check As Boolean, ByVal isChild As Boolean) As Boolean
    Select Case stp
        Case 2, 4: StepNeeded = check
        Case 5, 6: StepNeeded = Not isChild
        Case Else: StepNeeded = True
    End Select
End Function

Private Function EnsureSenderField(ByVal tid As String) As String
    Dim props As New cMap
    EnsureSenderField = mCfg.FieldIdByName(tid, "Erfasst von")
    If Len(EnsureSenderField) > 0 Then Exit Function
    props.SetVal "Hilfetext", "Name der Person/Stelle aus der Erfassungsdatei (Rücklauf)."
    props.SetVal "Gruppe", "Rücklauf"
    mApp.BeginWork
    EnsureSenderField = mStruct.AddField(tid, "Erfasst von", FT_TEXT, props)
    mApp.EndWork
End Function

Private Function TrimItems(ByRef items() As Variant, ByVal n As Long) As Variant
    Dim r() As Variant, i As Long
    ReDim r(0 To n - 1, 0 To 1)
    For i = 0 To n - 1
        r(i, 0) = items(i, 0): r(i, 1) = items(i, 1)
    Next
    TrimItems = r
End Function

' Welche Eigenschaften koennen in einer Erfassungsdatei ausgefuellt werden?
Private Function Capturable(ByVal f As cField) As Boolean
    If Not f.IsEditable Then Exit Function
    If f.FType = FT_BACKLINK Then Exit Function
    If f.FType = FT_LINK Then Capturable = Not f.Multi: Exit Function
    Capturable = True
End Function

' Felder, nach denen Empfaenger-Dateien gebildet werden koennen (ein Wert je Eintrag)
Private Function CanSplitBy(ByVal f As cField) As Boolean
    Select Case f.FType
        Case FT_AUSWAHL, FT_STATUS, FT_TEXT, FT_JANEIN: CanSplitBy = (f.Kind <> FT_BUTTON)
        Case FT_LINK: CanSplitBy = Not f.Multi
    End Select
End Function

Public Sub Cmd_Delete()
    Dim c As Variant, nm As String
    On Error GoTo f
    If Not RequireSel() Then Exit Sub
    If Not mApp.RequireWritable("Das Entfernen") Then Exit Sub
    If Len(ParentOf(mSel)) > 0 Then
        nm = "das Blatt „" & SheetTitle(mSel) & "“"
    Else
        nm = "die Erfassungsdatei „" & mCfg.GetS(SH_CAP, mSel, "Name") & "“" & IIf(ChildIds(mSel).Count > 0, " mit allen Blättern", "")
    End If
    If Not mUI.Confirm("Soll " & nm & " entfernt werden?" & vbLf & vbLf & _
                       "Bereits übernommene Einträge bleiben erhalten. Noch ausstehende Rückläufe können danach nicht mehr eingelesen werden.", _
                       "Erfassungsdatei entfernen", "Entfernen", "Abbrechen", True) Then Exit Sub
    mLog.Add "Erfassungsdatei entfernt", mCfg.GetS(SH_CAP, mSel, "Tabelle"), mSel, "", mCfg.GetS(SH_CAP, mSel, "Name")
    For Each c In ChildIds(mSel)
        mCfg.DeleteRow SH_CAP, CStr(c)
    Next
    mCfg.DeleteRow SH_CAP, mSel
    mSel = ""
    mApp.MarkCfgDirty
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Erfassungsdatei entfernen"
End Sub

'==============================================================================
' Eintraege fuer vorbelegte Dateien
'==============================================================================
' Bezeichnung des Aufteilungswertes eines Eintrags (z. B. Name der Kita)
Private Function SplitLabel(ByVal sf As cField, ByVal rid As String) As String
    Dim v As Variant
    If sf.IsLinkType Then
        SplitLabel = mFmt.LinkLabels(sf, mData.LinkTargets(sf.ID, rid), 99)
    Else
        v = mData.GetValue(sf.TableID, rid, sf.ID)
        If IsBlank(v) Then Exit Function
        SplitLabel = mFmt.DisplayText(sf, v, rid)
    End If
    SplitLabel = Trim$(SplitLabel)
End Function

' Datensatz-IDs eines Blattes, optional nur die eines Empfaengers
Private Function SourceRids(ByVal cid As String, ByVal recipient As String) As Collection
    Dim res As New Collection, tid As String, vid As String, conds As Collection, rows() As Long, n As Long, k As Long, t As cTableData
    Dim sf As cField, rid As String
    tid = mCfg.GetS(SH_CAP, cid, "Tabelle")
    Set SourceRids = res
    If Not mCfg.TableExists(tid) Then Exit Function
    vid = COpt(cid, "ansicht")
    If Len(vid) > 0 And mCfg.ViewExists(vid) Then
        Set conds = mQuery.ParseConds(mCfg.ViewGet(vid, "Filter"))
    Else
        Set conds = New Collection
    End If
    rows = mQuery.FilterRows(tid, conds, "", New Collection, mQuery.ARCH_ACTIVE, n)
    If Len(vid) > 0 And mCfg.ViewExists(vid) Then
        If Len(mCfg.ViewGet(vid, "Sortierung")) > 0 Then mQuery.SortRows tid, rows, n, mCfg.ViewGet(vid, "Sortierung")
    End If
    If mCfg.FieldExists(COpt(cid, "teilen")) Then Set sf = mCfg.Field(COpt(cid, "teilen"))
    Set t = mData.TData(tid)
    For k = 1 To n
        rid = t.IdAt(rows(k))
        If sf Is Nothing Or Len(recipient) = 0 Then
            res.Add rid
        ElseIf StrComp(SplitLabel(sf, rid), recipient, vbTextCompare) = 0 Then
            res.Add rid
        End If
    Next
End Function

' Alle Empfaenger (Aufteilungswerte) einer Pruefdatei, alphabetisch
Private Function RecipientValues(ByVal cid As String, ByRef noValue As Long) As Collection
    Dim res As New Collection, seen As New cMap, sf As cField, rid As Variant, lbl As String, arr() As String, n As Long, i As Long, j As Long, tmp As String
    Set RecipientValues = res
    If Not mCfg.FieldExists(COpt(cid, "teilen")) Then Exit Function
    Set sf = mCfg.Field(COpt(cid, "teilen"))
    For Each rid In SourceRids(cid, "")
        lbl = SplitLabel(sf, CStr(rid))
        If Len(lbl) = 0 Then
            noValue = noValue + 1
        ElseIf Not seen.Exists(LCase$(lbl)) Then
            seen.SetVal LCase$(lbl), True
            n = n + 1
            ReDim Preserve arr(1 To n)
            arr(n) = lbl
        End If
    Next
    For i = 1 To n - 1
        For j = i + 1 To n
            If StrComp(arr(j), arr(i), vbTextCompare) < 0 Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next
    Next
    For i = 1 To n
        res.Add arr(i)
    Next
End Function

'==============================================================================
' Dateien erzeugen
'==============================================================================
Public Sub Cmd_Export()
    Dim root As String, path As String, names As String, ok As Boolean, a As Variant, i As Long, basePath As String, made As String, cnt As Long, p As String
    Dim recips As Collection, items() As Variant, picked As String, noVal As Long, v As Variant, versand As String, fileId As String, rcpt As Collection
    Dim saved As Boolean
    On Error GoTo f
    If Not RequireSel() Then Exit Sub
    root = RootOf(mSel)
    Set rcpt = New Collection
    If IsCheck(root) And Len(COpt(root, "teilen")) > 0 Then
        ' Empfaenger aus den Daten (z. B. alle Kitas mit Einträgen)
        Set recips = RecipientValues(root, noVal)
        If recips.Count = 0 Then mUI.Alert "Für die gewählten Einträge gibt es keine Werte in „" & mCfg.FieldName(COpt(root, "teilen")) & "“. Es wurde nichts erzeugt.", "Dateien erzeugen": Exit Sub
        ReDim items(0 To recips.Count - 1, 0 To 1)
        i = 0
        For Each v In recips
            items(i, 0) = v
            items(i, 1) = Plural(SourceRids(root, CStr(v)).Count, "Eintrag", "Einträge")
            picked = picked & IIf(i > 0, ";", "") & i
            i = i + 1
        Next
        picked = mUI.ChooseMany("Für welche Empfänger soll je eine Datei erzeugt werden? Jede Datei enthält nur deren Einträge." & _
                                IIf(noVal > 0, vbLf & noVal & " Einträge ohne Angabe in „" & mCfg.FieldName(COpt(root, "teilen")) & "“ werden nicht verschickt.", ""), _
                                "Empfänger wählen", items, picked, ok)
        If Not ok Or Len(picked) = 0 Then Exit Sub
        For Each v In Split(picked, ";")
            rcpt.Add CStr(items(CLng(v), 0))
        Next
    End If
    ' Neue Einträge sammeln: immer genau eine Datei für alle (Absender wählen sich selbst aus)
    path = mIO.AskExportPath(SafeFile(IIf(Len(mCfg.GetS(SH_CAP, root, "Titel")) > 0, mCfg.GetS(SH_CAP, root, "Titel"), mCfg.GetS(SH_CAP, root, "Name"))), ".xlsx")
    If Len(path) = 0 Then Exit Sub
    mApp.BeginWork "Erfassungsdateien werden erzeugt …"
    If rcpt.Count = 0 Then
        fileId = BuildFile(root, path, "")
        versand = Chr$(31) & fileId
        made = path: cnt = 1
    Else
        basePath = Left$(path, Len(path) - 5)
        For Each v In rcpt
            p = basePath & "_" & SafeFile(CStr(v)) & ".xlsx"
            Application.StatusBar = "Erfassungsdatei für " & v & " wird erzeugt …"
            fileId = BuildFile(root, p, CStr(v))
            versand = versand & IIf(Len(versand) > 0, Chr$(30), "") & CStr(v) & Chr$(31) & fileId
            If cnt < 15 Then made = made & vbLf & "• " & Mid$(p, InStrRev(p, "\") + 1)
            cnt = cnt + 1
        Next
        If cnt > 15 Then made = made & vbLf & "• … und " & (cnt - 15) & " weitere"
    End If
    mApp.EndWork
    ' Versand merken (fuer die Ruecklauf-Uebersicht)
    mCfg.SysSet "Versand_" & root, DateInv(Now) & Chr$(29) & versand
    mLog.Add "Erfassungsdatei erzeugt", mCfg.GetS(SH_CAP, root, "Tabelle"), root, "", cnt & " Datei(en)"
    mApp.MarkCfgDirty
    ' Hauptdatei sofort speichern: nur so lassen sich die Rückläufe später dieser Erfassung zuordnen
    saved = False
    If Not mApp.gReadOnly And Len(ThisWorkbook.path) > 0 Then saved = mBackup.ControlledSave(False)
    mNav.RenderCurrent
    mUI.Alert IIf(cnt = 1, "Die Erfassungsdatei ist fertig:" & vbLf & made, cnt & " Erfassungsdateien sind fertig:" & made) & vbLf & vbLf & _
              "So geht es weiter:" & vbLf & _
              "1. Datei per E-Mail an alle Teilnehmenden schicken (als Anhang)." & vbLf & _
              "2. Ausgefüllte Dateien zurückbekommen und in einem Ordner sammeln." & vbLf & _
              "3. Hier „Rücklauf einlesen“ wählen und die Dateien markieren; mehrere auf einmal sind möglich." & vbLf & vbLf & _
              "Unter „Rückmeldungen“ sehen Sie jederzeit, wer schon geantwortet hat und wer noch fehlt." & _
              IIf(saved, vbLf & "Diese Datei wurde dabei gespeichert, damit die Rückläufe später zugeordnet werden können.", _
                         vbLf & "Bitte diese Datei jetzt speichern, sonst lassen sich die Rückläufe später nicht zuordnen."), "Erfassungsdatei erzeugt"
    Exit Sub
f:  Application.DisplayAlerts = True
    mApp.ResetWork
    mApp.Boundary "Erfassungsdatei erzeugen"
End Sub

Private Function SafeFile(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, CStr(ch), "_")
    Next
    SafeFile = Trim$(s)
End Function

Private Function SafeSheet(ByVal s As String, ByVal wb As Workbook) As String
    Dim ch As Variant, baseNm As String, n As Long, ws As Worksheet, taken As Boolean
    For Each ch In Array("\", "/", ":", "*", "?", "[", "]", "'")
        s = Replace(s, CStr(ch), "_")
    Next
    s = Trim$(Left$(s, 28))
    If Len(s) = 0 Or s = INFO_SHEET Or s = LIST_SHEET Then s = "Blatt"
    baseNm = s
    Do
        taken = False
        For Each ws In wb.Worksheets
            If StrComp(ws.Name, s, vbTextCompare) = 0 Then taken = True
        Next
        If Not taken Then Exit Do
        n = n + 1
        s = baseNm & " " & (n + 1)
    Loop
    SafeSheet = s
End Function

' Kennung dieser Datenbankdatei (entsteht beim ersten Erzeugen einer Erfassungsdatei)
Private Function DbId() As String
    DbId = mCfg.SysGet("DatenbankID", "")
    If Len(DbId) = 0 Then
        DbId = "DB" & Mid$(NewFileId(), 2)
        mCfg.SysSet "DatenbankID", DbId
    End If
End Function

Private Function NewFileId() As String
    Static seq As Long
    Randomize
    seq = seq + 1
    NewFileId = "E" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(Int(Rnd() * 1000000), "000000") & Format$(seq Mod 100, "00")
End Function

' Englische Formel in die Formelsprache dieser Excel-Installation uebersetzen (fuer Bedingte Formatierung)
Private Function LocalFormula(ByVal scratch As Range, ByVal f As String) As String
    scratch.Formula = f
    LocalFormula = scratch.FormulaLocal
    scratch.ClearContents
End Function

Private Function HeadColor(ByVal root As String) As Long
    Dim nm As String
    nm = COpt(root, "farbe", mDocStyle.DefaultStyle())
    HeadColor = mDocStyle.DocColor(nm)
End Function

Private Function BandOld() As Long: BandOld = RGB(100, 116, 139): End Function   ' Band/Streifen "Bestehende Einträge"

' Höhe (Punkt) eines umbrochenen Textes in einer Breite (Excel-Spaltenbreiten-Einheiten), geschätzt
Private Function TextHeight(ByVal s As String, ByVal widthUnits As Double, ByVal fontSize As Double) As Single
    Dim perLine As Double, lines As Long, part As Variant
    If widthUnits < 1 Then widthUnits = 1
    perLine = widthUnits * 11 / fontSize * 0.92        ' Zeichen je Zeile (fette Schrift etwas breiter)
    For Each part In Split(s, vbLf)
        lines = lines + Application.Max(1, -Int(-(Len(CStr(part)) / perLine)))
    Next
    If lines < 1 Then lines = 1
    TextHeight = lines * fontSize * 1.4 + 8
End Function

Private Function SpanWidth(ByVal ws As Worksheet, ByVal c0 As Long, ByVal c1 As Long) As Double
    Dim c As Long
    For c = c0 To c1
        If Not ws.Columns(c).Hidden Then SpanWidth = SpanWidth + ws.Columns(c).ColumnWidth
    Next
End Function

' Abschnittsband über die Tabellenbreite (bricht bei schmalen Tabellen um)
Private Sub SectionBand(ByVal rng As Range, ByVal txt As String, ByVal fill As Long, ByVal ink As Long)
    rng.UnMerge
    rng.Merge
    rng.Value = txt
    rng.Font.Bold = True
    rng.Font.Size = 10
    rng.Font.Color = ink
    rng.Interior.Color = fill
    rng.HorizontalAlignment = xlLeft
    rng.VerticalAlignment = xlCenter
    rng.IndentLevel = 1
    rng.WrapText = True
    rng.Borders.LineStyle = xlNone
    rng.Locked = True
    rng.EntireRow.RowHeight = Application.Max(26, TextHeight(txt, SpanWidth(rng.Worksheet, rng.Column, rng.Column + rng.Columns.Count - 1) - 2, 10) + 4)
End Sub

Private Function FillEdit() As Long: FillEdit = RGB(255, 253, 231): End Function
Private Function FillShow() As Long: FillShow = RGB(243, 244, 246): End Function

' Erzeugt eine Datei mit allen Blaettern; liefert die Datei-ID
Private Function BuildFile(ByVal root As String, ByVal path As String, ByVal recipient As String) As String
    Dim wb As Workbook, ws As Worksheet, wl As Worksheet, wi As Worksheet, prevAlerts As Boolean, infoR As Long, listCol As Long
    Dim defs As Collection, d As Variant, first As Boolean, isForm As Boolean, title As String, fileId As String, nm As String, i As Long
    Set defs = SheetDefs(root)
    For Each d In defs
        If Not mCfg.TableExists(mCfg.GetS(SH_CAP, CStr(d), "Tabelle")) Then RaiseUser "Die Tabelle des Blattes „" & SheetTitle(CStr(d)) & "“ existiert nicht mehr."
    Next
    isForm = (mCfg.GetS(SH_CAP, root, "Modus") = "Formular") And Not IsCheck(root) And defs.Count = 1
    title = mCfg.GetS(SH_CAP, root, "Titel")
    If Len(title) = 0 Then title = mCfg.GetS(SH_CAP, root, "Name")
    fileId = NewFileId()
    If FileExists(path) Then Kill path
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set wl = wb.Worksheets(1): wl.Name = LIST_SHEET
    Set wi = wb.Worksheets.Add(After:=wl): wi.Name = INFO_SHEET
    wi.Columns("A:F").NumberFormat = "@"
    infoR = 0
    InfoSet wi, infoR, "Kennung", MARK
    InfoSet wi, infoR, "ErfassungID", root
    InfoSet wi, infoR, "Erfassung", mCfg.GetS(SH_CAP, root, "Name")
    InfoSet wi, infoR, "Version", mCfg.GetS(SH_CAP, root, "Version")
    InfoSet wi, infoR, "TabelleID", mCfg.GetS(SH_CAP, root, "Tabelle")
    InfoSet wi, infoR, "Tabelle", mCfg.TableName(mCfg.GetS(SH_CAP, root, "Tabelle"))
    InfoSet wi, infoR, "Modus", IIf(isForm, "Formular", "Liste")
    InfoSet wi, infoR, "DateiID", fileId
    InfoSet wi, infoR, "Erstellt", DateInv(Now)
    InfoSet wi, infoR, "Quelle", ThisWorkbook.Name
    InfoSet wi, infoR, "DatenbankID", DbId()
    InfoSet wi, infoR, "Programmversion", PROG_VERSION
    InfoSet wi, infoR, "Empfaenger", recipient
    InfoSet wi, infoR, "Aufbau", "2"
    first = True
    i = 0
    For Each d In defs
        i = i + 1
        If defs.Count = 1 Then nm = CAPT_SHEET Else nm = SafeSheet(IIf(i = 1, SheetTitleRoot(root), SheetTitle(CStr(d))), wb)
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(i))
        ws.Name = nm
        If isForm Then
            BuildFormSheet wb, ws, wl, wi, infoR, listCol, CStr(d), recipient, title
        Else
            BuildListSheet wb, ws, wl, wi, infoR, listCol, CStr(d), root, recipient, IIf(first, title, SheetTitle(CStr(d))), first
        End If
        first = False
    Next
    ' Schutz und Ausblenden
    wl.Visible = xlSheetVeryHidden
    wi.Visible = xlSheetVeryHidden
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.Zoom = 100
            ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
            If mCfg.GetS(SH_CAP, root, "Schutz") <> "Nein" Then
                ws.Protect Password:="", DrawingObjects:=True, Contents:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
            End If
        End If
    Next
    wb.Worksheets(1).Activate
    If mCfg.GetS(SH_CAP, root, "Schutz") <> "Nein" Then wb.Protect Structure:=True
    wb.BuiltinDocumentProperties("Title") = title
    wb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = prevAlerts
    ThisWorkbook.Activate
    BuildFile = fileId
End Function

Private Function SheetTitleRoot(ByVal root As String) As String
    SheetTitleRoot = mCfg.TableName(mCfg.GetS(SH_CAP, root, "Tabelle"))
End Function

Private Sub InfoSet(ByVal wi As Worksheet, ByRef r As Long, ByVal key As String, ByVal v As String)
    r = r + 1
    wi.Cells(r, 1).Value = key
    wi.Cells(r, 2).Value = v
End Sub

Private Sub InfoRow(ByVal wi As Worksheet, ByRef r As Long, ByVal a As Variant)
    Dim i As Long
    r = r + 1
    For i = 0 To UBound(a)
        wi.Cells(r, i + 1).Value = CStr(a(i))
    Next
End Sub

Private Sub StyleInput(ByVal rng As Range)
    rng.Interior.Color = FillEdit()
    With rng.Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(209, 213, 219)
    End With
    rng.IndentLevel = 0
End Sub

' Text im Kopfbereich: über die ganze Breite zusammengefasst, umgebrochen, Höhe nach Länge
Private Sub HeadText(ByVal rng As Range, ByVal txt As String, ByVal size As Double, ByVal bold As Boolean, _
                     ByVal widthUnits As Double, ByVal minH As Single)
    rng.Merge
    rng.Value = txt
    rng.Font.Size = size
    rng.Font.Bold = bold
    rng.WrapText = True
    rng.HorizontalAlignment = xlLeft
    rng.VerticalAlignment = xlCenter
    If Len(txt) = 0 Then
        rng.EntireRow.RowHeight = 6
    Else
        rng.EntireRow.RowHeight = Application.Min(400, Application.Max(minH, TextHeight(txt, widthUnits - 1, size)))
    End If
End Sub

' Kopfbereich (Titel, Untertitel, Einleitung, Anleitung, Absender), nur erstes Blatt mit Absender
Private Sub DrawHeader(ByVal ws As Worksheet, ByVal wi As Worksheet, ByRef infoR As Long, ByVal root As String, ByVal recipient As String, _
                       ByVal title As String, ByVal lastCol As Long, ByVal isFirst As Boolean, ByVal tid As String, ByVal check As Boolean)
    Dim accent As Long, rng As Range, c As Long, w As Double, target As Double, howto As String, hw As Double
    Dim senders As Collection, sv As Variant, wl As Worksheet, i As Long
    accent = HeadColor(root)
    hw = SpanWidth(ws, 2, lastCol)
    ws.Rows(1).RowHeight = 8
    ' Titelband: genau so breit wie der Kopfbereich; Titel und Untertitel brechen bei Bedarf um
    With ws.Range(ws.Cells(2, 2), ws.Cells(3, lastCol))
        .Interior.Color = accent
        .Font.Color = mDocStyle.DocOnColor(accent)
    End With
    HeadText ws.Range(ws.Cells(2, 2), ws.Cells(2, lastCol)), title, 17, True, hw, 36
    ws.Cells(2, 2).VerticalAlignment = xlBottom
    HeadText ws.Range(ws.Cells(3, 2), ws.Cells(3, lastCol)), _
             mCfg.TableName(tid) & " · erstellt am " & FmtDate(Date) & IIf(Len(recipient) > 0, " · für " & recipient, ""), 9, False, hw, 22
    ws.Rows(4).RowHeight = 10
    If isFirst Then
        HeadText ws.Range(ws.Cells(5, 2), ws.Cells(5, lastCol)), mCfg.GetS(SH_CAP, root, "Einleitung"), 10.5, False, hw, 22
        If COpt(root, "anleitung", "1") = "1" Then
            If check Then
                howto = "So geht's: Oben stehen Ihre bestehenden Einträge: Graue Spalten sind vorbelegt, gelbe bitte prüfen und bei Bedarf ändern. " & _
                        "Darunter können Sie neue Einträge ergänzen. Der Pfeil in einer Zelle öffnet die Auswahl; Ankreuzspalten: „Ja“ wählen oder leer lassen. Danach speichern und zurücksenden."
            Else
                howto = "So geht's: gelbe Felder ausfüllen · Pfeil in einer Zelle öffnet die Auswahl · Ankreuzspalten: „Ja“ wählen oder leer lassen · Datum als TT.MM.JJJJ · Pflichtangaben (*) werden rot, solange sie fehlen · danach speichern und zurücksenden."
            End If
            HeadText ws.Range(ws.Cells(6, 2), ws.Cells(6, lastCol)), howto, 9, False, hw, 18
            ws.Cells(6, 2).Font.Color = RGB(107, 114, 128)
        Else
            ws.Rows(6).RowHeight = 6
        End If
        If mCfg.GetS(SH_CAP, root, "Absender") = "Ja" Then
            ws.Cells(7, 2).Value = COpt(root, "abslabel", "Ihr Name / Ihre Stelle") & " *"
            ws.Cells(7, 2).Font.Bold = True
            ws.Cells(7, 2).Font.Size = 10.5
            ws.Cells(7, 2).HorizontalAlignment = xlLeft
            ' Beschriftung vollstaendig sichtbar
            If ws.Columns(2).ColumnWidth < Application.Min(44, Len(ws.Cells(7, 2).Value) * 1.15 + 3) Then
                ws.Columns(2).ColumnWidth = Application.Min(44, Len(ws.Cells(7, 2).Value) * 1.15 + 3)
            End If
            Select Case COpt(root, "absbreite", "Breit")
                Case "Schmal": target = 28
                Case "Mittel": target = 42
                Case "Sehr breit": target = 90
                Case Else: target = 60
            End Select
            ' Eingabefeld ueber so viele Spalten zusammenfassen, bis die gewuenschte Breite erreicht ist
            ' (höchstens bis zum rechten Rand des Kopfbereichs)
            If lastCol > 3 And SpanWidth(ws, 3, lastCol) < target Then target = Application.Max(24, SpanWidth(ws, 3, lastCol))
            c = 3: w = 0
            Do
                If Not ws.Columns(c).Hidden Then w = w + ws.Columns(c).ColumnWidth
                If w >= target Or c >= 40 Then Exit Do
                c = c + 1
            Loop
            Set rng = ws.Range(ws.Cells(7, 3), ws.Cells(7, c))
            rng.Merge
            StyleInput rng
            rng.Borders.Color = RGB(156, 163, 175)
            rng.Font.Size = 11
            rng.HorizontalAlignment = xlLeft
            rng.Cells(1, 1).Value = recipient
            rng.Locked = False
            rng.Validation.Delete
            Set senders = ExpectedSenders(root)
            If (SenderMode(root) = "liste" Or SenderMode(root) = "tabelle") And senders.Count > 0 Then
                ' einheitliche Namen: Auswahl aus der Teilnehmerliste (steht auf dem ausgeblendeten Blatt "Listen")
                Set wl = ws.Parent.Worksheets(LIST_SHEET)
                i = 0
                For Each sv In senders
                    i = i + 1
                    wl.Cells(i, 250).NumberFormat = "@"
                    wl.Cells(i, 250).Value = CStr(sv)
                Next
                On Error Resume Next
                ws.Parent.Names.Add Name:="L_ABS", RefersTo:="='" & LIST_SHEET & "'!" & wl.Range(wl.Cells(1, 250), wl.Cells(i, 250)).Address
                On Error GoTo 0
                rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=L_ABS"
                rng.Validation.InCellDropdown = True
                rng.Validation.InputTitle = Left$(COpt(root, "abslabel", "Ihr Name / Ihre Stelle"), 32)
                rng.Validation.InputMessage = "Bitte aus der Liste wählen (Pfeil rechts in der Zelle). Pflichtangabe."
                rng.Validation.ErrorTitle = Left$(COpt(root, "abslabel", "Ihr Name / Ihre Stelle"), 32)
                rng.Validation.ErrorMessage = "Bitte einen Eintrag aus der Liste wählen."
            Else
                rng.Validation.Add Type:=xlValidateInputOnly
                rng.Validation.InputTitle = Left$(COpt(root, "abslabel", "Ihr Name / Ihre Stelle"), 32)
                rng.Validation.InputMessage = "Pflichtangabe (wird beim Einlesen gespeichert)."
            End If
            ws.Rows(7).RowHeight = 28
            InfoSet wi, infoR, "AbsenderZelle", rng.Cells(1, 1).Address(False, False)
            InfoSet wi, infoR, "AbsenderBlatt", ws.Name
        Else
            ws.Rows(7).RowHeight = 6
        End If
    Else
        HeadText ws.Range(ws.Cells(5, 2), ws.Cells(5, lastCol)), _
                 IIf(check, "Graue Spalten sind vorbelegt, gelbe bitte prüfen bzw. ausfüllen.", "Bitte die gelben Felder ausfüllen."), 9.5, False, hw, 20
        ws.Cells(5, 2).Font.Color = RGB(107, 114, 128)
        ws.Rows(6).RowHeight = 6: ws.Rows(7).RowHeight = 6
    End If
    ws.Rows(8).RowHeight = 12
    ws.Rows(R_BAND).RowHeight = 4
End Sub

' Spaltenplan eines Listenblattes: Array(Art, FeldID, OptionsID, Spalte); Art: ANZ | FELD | OPT | ID | STAND
Private Function ColumnPlan(ByVal cid As String, ByRef nVisible As Long) As Collection
    Dim res As New Collection, tid As String, f As cField, c As Long, check As Boolean, anz As String, sel As String, oid As Variant, n As Long
    tid = mCfg.GetS(SH_CAP, cid, "Tabelle")
    check = IsCheck(cid)
    anz = COpt(cid, "anzeige")
    sel = mCfg.GetS(SH_CAP, cid, "Felder")
    c = 1
    For Each f In mCfg.Fields(tid)
        If ListHas(anz, f.ID) Then
            If check And IsOptionColumns(f) And f.ID <> COpt(cid, "teilen") Then
                ' Mehrfachauswahl auch beim Prüfen als Ankreuzspalten: vorhandene Zeilen
                ' zeigen die Haken grau, neue Zeilen werden angekreuzt
                For Each oid In mCfg.OptionIds(f.ID)
                    If mCfg.OptActive(CStr(oid)) Then
                        c = c + 1
                        res.Add Array("OPTA", f.ID, CStr(oid), c)
                    End If
                Next
            Else
                c = c + 1
                res.Add Array("ANZ", f.ID, "", c)
            End If
        End If
    Next
    For Each f In mCfg.Fields(tid)
        If ListHas(sel, f.ID) And Not ListHas(anz, f.ID) Then
            If IsOptionColumns(f) Then
                For Each oid In mCfg.OptionIds(f.ID)
                    If mCfg.OptActive(CStr(oid)) Then
                        c = c + 1
                        res.Add Array("OPT", f.ID, CStr(oid), c)
                    End If
                Next
            Else
                c = c + 1
                res.Add Array("FELD", f.ID, "", c)
            End If
        End If
    Next
    nVisible = c - 1
    If check Then
        res.Add Array("ID", "", "", c + 1)
        res.Add Array("STAND", "", "", c + 2)
    End If
    Set ColumnPlan = res
End Function

' Mehrfachauswahl als Ankreuzspalten (bis MULTI_MAX aktive Optionen)
Private Function IsOptionColumns(ByVal f As cField) As Boolean
    Dim oid As Variant, n As Long
    If f.FType <> FT_MEHRFACH Then Exit Function
    For Each oid In mCfg.OptionIds(f.ID)
        If mCfg.OptActive(CStr(oid)) Then n = n + 1
    Next
    IsOptionColumns = (n > 0 And n <= MULTI_MAX)
End Function

Private Sub BuildListSheet(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal wl As Worksheet, ByVal wi As Worksheet, ByRef infoR As Long, _
                           ByRef listCol As Long, ByVal cid As String, ByVal root As String, ByVal recipient As String, ByVal title As String, ByVal isFirst As Boolean)
    Dim plan As Collection, p As Variant, tid As String, f As cField, check As Boolean, rids As Collection, nRec As Long, nEmpty As Long, nRows As Long
    Dim r0 As Long, r1 As Long, lastCol As Long, nVis As Long, data() As Variant, i As Long, j As Long, rid As String, v As Variant, accent As Long
    Dim rng As Range, grpStart As Long, grpFid As String, maxOpt As Long, c As Long, idCol As Long, standCol As Long, rowRef As String, lists As New cMap
    Dim t As cTableData, tr As Long, curMulti As String, maxLen As Long, sepRow As Long, rN As Long, splitFid As String, splitVal As String
    Dim gc0 As Long, gc1 As Long, sepN As Long, optLines As Long, hdrLast As Long, tblW As Double, w As Double
    Dim hGroup As Single, hHead As Single, hNeed As Single, preN As Long
    tid = mCfg.GetS(SH_CAP, cid, "Tabelle")
    check = IsCheck(cid)
    accent = HeadColor(root)
    splitFid = COpt(cid, "teilen")
    Set plan = ColumnPlan(cid, nVis)
    lastCol = 1 + nVis
    ws.Cells.Font.Name = mDocStyle.DocFont(COpt(root, "schrift", mDocStyle.DefaultFont()))
    ws.Cells.Font.Size = mDocStyle.DocFontSize(COpt(root, "schrift", mDocStyle.DefaultFont()), 10)
    ws.Cells.VerticalAlignment = xlCenter
    ws.Columns(1).ColumnWidth = 2
    ' Auswahllisten
    For Each p In plan
        If p(0) = "FELD" Then
            Set f = mCfg.Field(CStr(p(1)))
            If HasList(f) And Not lists.Exists(f.ID) Then
                listCol = listCol + 1
                WriteList wb, wl, f, listCol
                lists.SetVal f.ID, True
            End If
        End If
    Next
    If Not lists.Exists("#JA") Then
        listCol = listCol + 1
        wl.Cells(1, listCol).Value = "Ja"
        On Error Resume Next
        wb.Names.Add Name:="L_JA", RefersTo:="='" & LIST_SHEET & "'!" & wl.Cells(1, listCol).Address
        ' leere Liste: Ankreuzzelle, wenn die Höchstzahl erreicht ist
        wb.Names.Add Name:="L_LEER", RefersTo:="='" & LIST_SHEET & "'!" & wl.Cells(2, listCol).Address
        On Error GoTo 0
    End If
    ' Zeilen
    If check Then
        Set rids = SourceRids(cid, recipient)
        nRec = rids.Count
        nEmpty = CLng(Val(COpt(cid, "neuzeilen", "0")))
    Else
        Set rids = New Collection
        nEmpty = CLng(Val(mCfg.GetS(SH_CAP, cid, "Zeilen")))
        If nEmpty < 1 Then nEmpty = 200
    End If
    nRows = nRec + nEmpty
    If nRows < 1 Then nRows = 1
    r0 = R_HEAD + 1
    ' Prüfdatei mit Zeilen für Neues: eigener Trennblock (Abstand, Band "Neue Einträge", bei langen
    ' Listen wiederholte Spaltenköpfe) zwischen beiden Bereichen; alle Zeilen mit SEP_MARK
    If nRec > 0 And nEmpty > 0 Then
        sepN = IIf(nRec > 8, 3, 2)
        sepRow = r0 + nRec: rN = sepRow + sepN
    Else
        rN = r0 + nRec
    End If
    r1 = r0 + nRows - 1 + sepN
    ' Spaltenbreiten zuerst (der Kopfbereich richtet sich danach)
    For Each p In plan
        Select Case p(0)
            Case "ANZ": ws.Columns(p(3)).ColumnWidth = ColWidth(mCfg.Field(CStr(p(1))), True)
            Case "FELD": ws.Columns(p(3)).ColumnWidth = ColWidth(mCfg.Field(CStr(p(1))), False)
            Case "OPT", "OPTA"
                ws.Columns(p(3)).ColumnWidth = 4.6
                If Len(mCfg.OptName(CStr(p(2)))) > maxOpt Then maxOpt = Len(mCfg.OptName(CStr(p(2))))
            Case Else: ws.Columns(p(3)).ColumnWidth = 12
        End Select
    Next
    ' lange Optionsnamen: senkrechter Text in zwei oder drei Zeilen statt abgeschnitten
    optLines = Application.Min(3, Application.Max(1, -Int(-(maxOpt * 5.6 / 140))))
    If optLines > 1 Then
        For Each p In plan
            If p(0) = "OPT" Or p(0) = "OPTA" Then ws.Columns(p(3)).ColumnWidth = 4.6 + (optLines - 1) * 2.6
        Next
    End If
    ' Daten der vorhandenen Einträge vorab lesen: Textspalten so breit wie ihr Inhalt (höchstens 48)
    Set t = mData.TData(tid)
    If nRec > 0 Then
        ReDim data(1 To nRec, 1 To lastCol + 2)
        i = 0
        For Each v In rids
            i = i + 1
            rid = CStr(v)
            tr = t.RowOf(rid)
            For Each p In plan
                c = p(3) - 1
                Select Case p(0)
                    Case "ANZ": data(i, c) = ShowValue(mCfg.Field(CStr(p(1))), rid)
                    Case "FELD": data(i, c) = EditValue(mCfg.Field(CStr(p(1))), rid)
                    Case "OPT", "OPTA"
                        curMulti = NzS(mData.GetValue(tid, rid, CStr(p(1))))
                        If ListHas(curMulti, CStr(p(2))) Then data(i, c) = "Ja"
                    Case "ID": data(i, c) = rid
                    Case "STAND"
                        v = mData.GetValue(tid, rid, C_CHANGED)
                        If IsDate(v) Or IsNumeric(v) Then
                            If Not IsBlank(v) Then data(i, c) = CStr(CDbl(CDate(v)))
                        End If
                End Select
            Next
        Next
        For Each p In plan
            If p(0) = "ANZ" Or p(0) = "FELD" Then
                maxLen = 0
                For i = 1 To nRec
                    If VarType(data(i, p(3) - 1)) = vbString Then
                        If Len(data(i, p(3) - 1)) > maxLen Then maxLen = Len(data(i, p(3) - 1))
                    End If
                Next
                If maxLen + 2 > ws.Columns(p(3)).ColumnWidth Then ws.Columns(p(3)).ColumnWidth = Application.Min(48, maxLen + 2)
            End If
        Next
    End If
    ' Kopfbereich mindestens so breit wie ein gut lesbarer Absatz; bei schmalen Tabellen über
    ' weitere (leere) Spalten hinweg, damit Titel und Hinweise nicht abgeschnitten oder ausgefranst wirken
    hdrLast = lastCol
    tblW = 0
    For c = 2 To lastCol
        tblW = tblW + ws.Columns(c).ColumnWidth
    Next
    If tblW < HDR_MIN_W Then
        c = lastCol + IIf(check, 3, 1)
        w = tblW
        Do While w < HDR_MIN_W And c < 60
            w = w + ws.Columns(c).ColumnWidth
            hdrLast = c
            c = c + 1
        Loop
    End If
    DrawHeader ws, wi, infoR, root, recipient, title, hdrLast, isFirst, tid, check
    If check And nRec > 0 Then
        SectionBand ws.Range(ws.Cells(R_BAND, 2), ws.Cells(R_BAND, lastCol)), _
                    "Bestehende Einträge (" & nRec & "): bitte prüfen, gelbe Felder bei Bedarf ändern", BandOld(), RGB(255, 255, 255)
        ws.Cells(R_BAND, 1).Interior.Color = BandOld()
    End If
    ' Blattbeschreibung fuer das Einlesen
    For Each p In plan
        If p(0) = "ID" Then idCol = p(3)
        If p(0) = "STAND" Then standCol = p(3)
    Next
    InfoRow wi, infoR, Array("Blatt", ws.Name, cid, tid, CStr(r0), CStr(r1))
    InfoRow wi, infoR, Array("Spalten", CStr(idCol), CStr(standCol), IIf(check, "pruefen", "neu"))
    If sepRow > 0 Then InfoRow wi, infoR, Array("Trenner", CStr(sepRow), CStr(sepN))
    For Each p In plan
        InfoRow wi, infoR, Array("Spalte", p(0), p(1), p(2), CStr(p(3)))
    Next
    ' Kopfzeilen: Gruppenzeile (Titel ueber Ankreuzspalten) und Spaltenkoepfe
    ws.Rows(R_GROUP).RowHeight = 20
    ws.Rows(R_HEAD).RowHeight = 30
    For Each p In plan
        c = p(3)
        Select Case p(0)
            Case "ANZ", "FELD"
                Set f = mCfg.Field(CStr(p(1)))
                Set rng = ws.Range(ws.Cells(R_GROUP, c), ws.Cells(R_HEAD, c))
                rng.Merge
                rng.Value = f.Name & IIf(p(0) = "FELD" And f.Required, " *", "") & IIf(f.FType = FT_PROZENT And p(0) = "FELD", " (%)", "")
                rng.WrapText = True
                rng.VerticalAlignment = xlBottom
                rng.Font.Bold = True
                rng.Interior.Color = IIf(p(0) = "ANZ", RGB(229, 231, 235), Blend(accent, RGB(255, 255, 255), 0.82))
                rng.Font.Color = IIf(p(0) = "ANZ", RGB(75, 85, 99), RGB(31, 41, 55))
            Case "OPT", "OPTA"
                ws.Cells(R_HEAD, c).Value = mCfg.OptName(CStr(p(2)))
                ws.Cells(R_HEAD, c).Orientation = xlUpward
                ws.Cells(R_HEAD, c).HorizontalAlignment = xlCenter
                ws.Cells(R_HEAD, c).VerticalAlignment = xlBottom
                ws.Cells(R_HEAD, c).Font.Size = 9
                ws.Cells(R_HEAD, c).WrapText = (optLines > 1)
                ws.Cells(R_HEAD, c).Interior.Color = IIf(p(0) = "OPTA", RGB(229, 231, 235), Blend(accent, RGB(255, 255, 255), 0.88))
                If CStr(p(1)) <> grpFid Then
                    If Len(grpFid) > 0 Then MergeGroup ws, grpStart, c - 1, grpFid, accent
                    grpFid = CStr(p(1)): grpStart = c
                End If
        End Select
        If p(0) <> "OPT" And p(0) <> "OPTA" And Len(grpFid) > 0 Then MergeGroup ws, grpStart, c - 1, grpFid, accent: grpFid = ""
    Next
    If Len(grpFid) > 0 Then MergeGroup ws, grpStart, lastCol, grpFid, accent
    ' Kopfhöhen aus dem Platzbedarf: Gruppentitel umbrechen, Spaltenköpfe ganz lesbar
    hGroup = 20
    hHead = 30
    If maxOpt > 0 Then hHead = Application.Min(150, Application.Max(30, -Int(-(maxOpt / optLines)) * 5.8 + 14))
    For Each p In plan
        If p(0) = "OPT" Or p(0) = "OPTA" Then
            GroupSpan plan, CStr(p(1)), CStr(p(0)), gc0, gc1
            If gc0 = p(3) Then
                hNeed = TextHeight(ws.Cells(R_GROUP, gc0).MergeArea.Cells(1, 1).Value, SpanWidth(ws, gc0, gc1), 9)
                If hNeed > hGroup Then hGroup = hNeed
            End If
        End If
    Next
    For Each p In plan
        If p(0) = "ANZ" Or p(0) = "FELD" Then
            hNeed = TextHeight(ws.Cells(R_GROUP, p(3)).Value, ws.Columns(p(3)).ColumnWidth, 10) + 4
            If hNeed > hGroup + hHead Then hHead = hNeed - hGroup
        End If
    Next
    ws.Rows(R_GROUP).RowHeight = hGroup
    ws.Rows(R_HEAD).RowHeight = hHead
    With ws.Range(ws.Cells(R_GROUP, 2), ws.Cells(R_HEAD, lastCol)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(209, 213, 219)
    End With
    ' Formate je Spalte
    For Each p In plan
        Set rng = ws.Range(ws.Cells(r0, p(3)), ws.Cells(r1, p(3)))
        Select Case p(0)
            Case "ANZ"
                ' vorbelegte Zeilen: grau und gesperrt; leere Zeilen fuer Neues: ausfuellbar
                Set f = mCfg.Field(CStr(p(1)))
                If nRec > 0 Then
                    With ws.Range(ws.Cells(r0, p(3)), ws.Cells(r0 + nRec - 1, p(3)))
                        .NumberFormat = "@"
                        .Interior.Color = FillShow()
                        .Font.Color = RGB(55, 65, 81)
                        .Locked = True
                        .HorizontalAlignment = xlLeft
                    End With
                End If
                If nEmpty > 0 Then
                    Set rng = ws.Range(ws.Cells(rN, p(3)), ws.Cells(r1, p(3)))
                    ' die Aufteilungsspalte (z. B. Standort) wird beim Einlesen automatisch gesetzt
                    If Capturable(f) And f.ID <> splitFid Then
                        If HasList(f) And Not lists.Exists(f.ID) Then
                            listCol = listCol + 1
                            WriteList wb, wl, f, listCol
                            lists.SetVal f.ID, True
                        End If
                        rng.Interior.Color = FillEdit()
                        rng.Locked = False
                        ApplyInput wb, rng, f
                    Else
                        ' fest vorgegeben: ruhig hell statt grauer Balken bis zum Ende
                        rng.Interior.Color = RGB(250, 250, 251)
                        rng.Font.Color = RGB(107, 114, 128)
                        rng.Locked = True
                        rng.NumberFormat = "@"
                        ' Aufteilungswert des Empfängers (z. B. "Workshop") steht auch in neuen Zeilen
                        If f.ID = splitFid And Len(recipient) > 0 Then rng.Value = recipient
                    End If
                End If
            Case "OPTA"
                ' vorhandene Zeilen: Haken nur zur Ansicht; neue Zeilen: ankreuzen
                Set f = mCfg.Field(CStr(p(1)))
                rng.HorizontalAlignment = xlCenter
                rng.NumberFormat = "@"
                If nRec > 0 Then
                    With ws.Range(ws.Cells(r0, p(3)), ws.Cells(r0 + nRec - 1, p(3)))
                        .Interior.Color = FillShow()
                        .Font.Color = RGB(55, 65, 81)
                        .Locked = True
                    End With
                End If
                If nEmpty > 0 Then
                    Set rng = ws.Range(ws.Cells(rN, p(3)), ws.Cells(r1, p(3)))
                    rng.Interior.Color = FillEdit()
                    rng.Locked = False
                    GroupSpan plan, f.ID, "OPTA", gc0, gc1
                    ApplyOptionInput rng, f, mCfg.OptName(CStr(p(2))), _
                                     ws.Range(ws.Cells(rN, gc0), ws.Cells(rN, gc1)).Address(RowAbsolute:=False, ColumnAbsolute:=True), wl.Range("ZZ1")
                    If gc0 = p(3) Then MarkTooManyPicks ws.Range(ws.Cells(rN, gc0), ws.Cells(r1, gc1)), _
                                     ws.Range(ws.Cells(rN, gc0), ws.Cells(rN, gc1)).Address(RowAbsolute:=False, ColumnAbsolute:=True), mValid.MaxPicks(f), wl.Range("ZZ1")
                End If
            Case "FELD"
                Set f = mCfg.Field(CStr(p(1)))
                rng.Interior.Color = FillEdit()
                rng.Locked = False
                ApplyInput wb, rng, f
            Case "OPT"
                rng.Interior.Color = FillEdit()
                rng.Locked = False
                rng.HorizontalAlignment = xlCenter
                rng.NumberFormat = "@"
                GroupSpan plan, CStr(p(1)), "OPT", gc0, gc1
                ApplyOptionInput rng, mCfg.Field(CStr(p(1))), mCfg.OptName(CStr(p(2))), _
                                 ws.Range(ws.Cells(r0, gc0), ws.Cells(r0, gc1)).Address(RowAbsolute:=False, ColumnAbsolute:=True), wl.Range("ZZ1")
                If gc0 = p(3) Then MarkTooManyPicks ws.Range(ws.Cells(r0, gc0), ws.Cells(r1, gc1)), _
                                 ws.Range(ws.Cells(r0, gc0), ws.Cells(r0, gc1)).Address(RowAbsolute:=False, ColumnAbsolute:=True), mValid.MaxPicks(mCfg.Field(CStr(p(1)))), wl.Range("ZZ1")
            Case "ID", "STAND"
                rng.NumberFormat = "@"
                rng.Locked = True
                ws.Columns(p(3)).Hidden = True
        End Select
    Next
    If nRec > 0 Then
        ws.Range(ws.Cells(r0, 2), ws.Cells(r0 + nRec - 1, lastCol + 2)).Value = SliceData(data, nRec, lastCol + 2 - 1)
    End If
    Set rng = ws.Range(ws.Cells(r0, 2), ws.Cells(r1, lastCol))
    rng.RowHeight = 20
    With rng.Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(229, 231, 235)
    End With
    ' vorhandene Einträge: lange Texte umbrechen statt abschneiden (Zeilenhöhe passt sich an)
    If nRec > 0 Then
        For Each p In plan
            If p(0) = "ANZ" Or p(0) = "FELD" Then
                If Not mFmt.IsRightAligned(mCfg.Field(CStr(p(1)))) Then
                    ws.Range(ws.Cells(r0, p(3)), ws.Cells(r0 + nRec - 1, p(3))).WrapText = True
                End If
            End If
        Next
        ws.Range(ws.Cells(r0, 2), ws.Cells(r0 + nRec - 1, lastCol)).Rows.AutoFit
        For i = r0 To r0 + nRec - 1
            If ws.Rows(i).RowHeight < 20 Then ws.Rows(i).RowHeight = 20 Else ws.Rows(i).RowHeight = ws.Rows(i).RowHeight + 4
        Next
        ' farbiger Streifen links: grau = bestehend
        ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + nRec - 1, 1)).Interior.Color = BandOld()
    End If
    If nEmpty > 0 And nRec > 0 Then ws.Range(ws.Cells(rN, 1), ws.Cells(r1, 1)).Interior.Color = accent
    ' klare Trennung: Abstand, Band "Neue Einträge" und bei langen Listen wiederholte Spaltenköpfe
    If sepRow > 0 Then
        Set rng = ws.Range(ws.Cells(sepRow, 1), ws.Cells(sepRow + sepN - 1, lastCol + 2))
        rng.Validation.Delete
        rng.FormatConditions.Delete
        rng.ClearContents
        rng.Locked = True
        rng.Interior.Pattern = xlNone
        rng.Borders.LineStyle = xlNone
        ' Abstandszeile
        ws.Rows(sepRow).RowHeight = 18
        ' Band
        SectionBand ws.Range(ws.Cells(sepRow + 1, 2), ws.Cells(sepRow + 1, lastCol)), _
                    "Neue Einträge: hier ergänzen, was oben noch fehlt", accent, mDocStyle.DocOnColor(accent)
        ws.Cells(sepRow + 1, 1).Interior.Color = accent
        ' wiederholte Spaltenköpfe
        If sepN = 3 Then
            For Each p In plan
                c = p(3)
                Select Case p(0)
                    Case "ANZ", "FELD"
                        ws.Cells(sepRow + 2, c).Value = mCfg.FieldName(CStr(p(1)))
                        ws.Cells(sepRow + 2, c).WrapText = True
                    Case "OPT", "OPTA"
                        ws.Cells(sepRow + 2, c).Value = mCfg.OptName(CStr(p(2)))
                        ws.Cells(sepRow + 2, c).Orientation = xlUpward
                        ws.Cells(sepRow + 2, c).WrapText = (optLines > 1)
                        ws.Cells(sepRow + 2, c).HorizontalAlignment = xlCenter
                End Select
            Next
            With ws.Range(ws.Cells(sepRow + 2, 2), ws.Cells(sepRow + 2, lastCol))
                .Font.Bold = True: .Font.Size = 8.5: .Font.Color = RGB(55, 65, 81)
                .Interior.Color = Blend(accent, RGB(255, 255, 255), 0.88)
                .VerticalAlignment = xlBottom
                .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(209, 213, 219)
            End With
            ws.Rows(sepRow + 2).RowHeight = IIf(maxOpt > 0, hHead, 30)
            ws.Cells(sepRow + 2, 1).Interior.Color = accent
        End If
        For i = sepRow To sepRow + sepN - 1
            ws.Cells(i, 1).Value = SEP_MARK
            ws.Cells(i, 1).Font.Size = 6
            ws.Cells(i, 1).Font.Color = IIf(i = sepRow, RGB(255, 255, 255), accent)
        Next
        ' letzte bestehende Zeile mit kräftigerer Unterkante
        With ws.Range(ws.Cells(sepRow - 1, 2), ws.Cells(sepRow - 1, lastCol)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlMedium: .Color = BandOld()
        End With
    End If
    ' Pflichtmarkierung (nur begonnene Zeilen bzw. vorbelegte)
    rowRef = ws.Range(ws.Cells(r0, 2), ws.Cells(r0, lastCol)).Address(False, True)
    ' in neuen Zeilen ist der Aufteilungswert vorbelegt: er zählt nicht als "begonnen"
    preN = 0
    For Each p In plan
        If p(0) = "ANZ" And CStr(p(1)) = splitFid And Len(recipient) > 0 Then preN = 1
    Next
    For Each p In plan
        If p(0) = "FELD" Then
            Set f = mCfg.Field(CStr(p(1)))
            If f.Required Then
                ' ohne den Trennblock (dessen Band würde sonst als "begonnene Zeile" gelten)
                If sepRow > 0 Then
                    AddRequiredFormat wl, ws.Range(ws.Cells(r0, p(3)), ws.Cells(sepRow - 1, p(3))), rowRef
                    AddRequiredFormat wl, ws.Range(ws.Cells(rN, p(3)), ws.Cells(r1, p(3))), _
                                      ws.Range(ws.Cells(rN, 2), ws.Cells(rN, lastCol)).Address(False, True), preN
                ElseIf nRec = 0 Then
                    AddRequiredFormat wl, ws.Range(ws.Cells(r0, p(3)), ws.Cells(r1, p(3))), rowRef, preN
                Else
                    AddRequiredFormat wl, ws.Range(ws.Cells(r0, p(3)), ws.Cells(r1, p(3))), rowRef
                End If
            End If
        End If
    Next
    ' Seitenlayout ohne Rueckfrage beim Drucker: Mit einem haengenden Standarddrucker (z. B. "OneNote (Desktop) - Protected")
    ' blieb das Erzeugen der Erfassungsdatei sonst stehen (November 2026)
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = ws.Rows(R_GROUP & ":" & R_HEAD).Address
        .CenterFooter = "Erfassungsdatei · " & mCfg.GetS(SH_CAP, root, "Name") & IIf(Len(recipient) > 0, " · " & recipient, "")
    End With
    Application.PrintCommunication = True
    ws.Activate
    ws.Cells(r0, 2).Select
End Sub

Private Function SliceData(ByRef a() As Variant, ByVal n As Long, ByVal cols As Long) As Variant
    Dim b() As Variant, r As Long, c As Long
    ReDim b(1 To n, 1 To cols)
    For r = 1 To n
        For c = 1 To cols
            b(r, c) = a(r, c)
        Next
    Next
    SliceData = b
End Function

Private Sub MergeGroup(ByVal ws As Worksheet, ByVal c0 As Long, ByVal c1 As Long, ByVal fid As String, ByVal accent As Long)
    Dim rng As Range, f As cField
    Set f = mCfg.Field(fid)
    Set rng = ws.Range(ws.Cells(R_GROUP, c0), ws.Cells(R_GROUP, c1))
    rng.Merge
    rng.Value = f.Name & IIf(f.Required, " *", "") & IIf(mValid.MaxPicks(f) > 0, "  (höchstens " & mValid.MaxPicks(f) & " ankreuzen)", "  (ankreuzen)")
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    rng.Font.Bold = True
    rng.Font.Size = 9
    rng.WrapText = True              ' Höhe der Gruppenzeile richtet sich nach dem längsten Titel
    rng.Interior.Color = Blend(accent, RGB(255, 255, 255), 0.82)
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(156, 163, 175)
    End With
End Sub

' Anzeigewert einer grauen Spalte (lesbarer Text)
Private Function ShowValue(ByVal f As cField, ByVal rid As String) As Variant
    If f.IsLinkType Then
        ShowValue = mFmt.LinkLabels(f, mData.LinkTargets(f.ID, rid), 99)
    Else
        ShowValue = mFmt.DisplayText(f, mData.GetValue(f.TableID, rid, f.ID), rid)
    End If
End Function

' Vorbelegung einer gelben Spalte (Wert so, dass er beim Einlesen wieder erkannt wird)
Private Function EditValue(ByVal f As cField, ByVal rid As String) As Variant
    Dim v As Variant
    If f.FType = FT_LINK Then
        EditValue = mFmt.LinkLabels(f, mData.LinkTargets(f.ID, rid), 1)
        Exit Function
    End If
    v = mData.GetValue(f.TableID, rid, f.ID)
    If IsBlank(v) Or IsErrVal(v) Then EditValue = Empty: Exit Function
    Select Case f.FType
        Case FT_JANEIN
            If f.Kind = FT_KONTROLL And Not CBool(v) Then EditValue = Empty Else EditValue = IIf(CBool(v), "Ja", "Nein")
        Case FT_AUSWAHL, FT_STATUS: EditValue = mCfg.OptName(CStr(v))
        Case FT_MEHRFACH: EditValue = mFmt.DisplayText(f, v, rid)
        Case FT_DATUM, FT_DATUMZEIT: EditValue = CDate(v)
        Case FT_ZAHL, FT_GELD, FT_PROZENT: EditValue = CDbl(v)
        Case Else: EditValue = CStr(v)
    End Select
End Function

Private Sub BuildFormSheet(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal wl As Worksheet, ByVal wi As Worksheet, ByRef infoR As Long, _
                           ByRef listCol As Long, ByVal cid As String, ByVal recipient As String, ByVal title As String)
    Dim fids As Variant, f As cField, i As Long, r As Long, rng As Range, oid As Variant, tid As String, r0 As Long
    Dim nAct As Long, gAddr As String
    tid = mCfg.GetS(SH_CAP, cid, "Tabelle")
    fids = Split(mCfg.GetS(SH_CAP, cid, "Felder"), ";")
    ws.Cells.Font.Name = mDocStyle.DocFont(COpt(RootOf(cid), "schrift", mDocStyle.DefaultFont()))
    ws.Cells.Font.Size = mDocStyle.DocFontSize(COpt(RootOf(cid), "schrift", mDocStyle.DefaultFont()), 10)
    ws.Cells.VerticalAlignment = xlCenter
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 32
    ws.Columns(3).ColumnWidth = 48
    ws.Columns(4).ColumnWidth = 46
    For i = 0 To UBound(fids)
        Set f = mCfg.Field(CStr(fids(i)))
        If Not f Is Nothing Then
            If HasList(f) And Not IsOptionColumns(f) Then listCol = listCol + 1: WriteList wb, wl, f, listCol
        End If
    Next
    listCol = listCol + 1
    wl.Cells(1, listCol).Value = "Ja"
    wb.Names.Add Name:="L_JA", RefersTo:="='" & LIST_SHEET & "'!" & wl.Cells(1, listCol).Address
    wb.Names.Add Name:="L_LEER", RefersTo:="='" & LIST_SHEET & "'!" & wl.Cells(2, listCol).Address
    DrawHeader ws, wi, infoR, cid, recipient, title, 4, True, tid, False
    r0 = R_HEAD - 1
    InfoRow wi, infoR, Array("Blatt", ws.Name, cid, tid, "0", "0")
    InfoRow wi, infoR, Array("Spalten", "0", "0", "formular")
    r = r0
    For i = 0 To UBound(fids)
        Set f = mCfg.Field(CStr(fids(i)))
        If Not f Is Nothing Then
            ws.Cells(r, 2).Value = f.Name & IIf(f.Required, " *", "")
            ws.Cells(r, 2).Font.Bold = True
            ws.Cells(r, 2).VerticalAlignment = xlTop
            If IsOptionColumns(f) Then
                ' Mehrfachauswahl: je Option eine Zeile mit "Ja" oder leer
                ws.Cells(r, 4).Value = "Zutreffendes mit „Ja“ markieren" & IIf(mValid.MaxPicks(f) > 0, ", höchstens " & mValid.MaxPicks(f), "")
                ws.Cells(r, 4).Font.Size = 8.5: ws.Cells(r, 4).Font.Color = RGB(107, 114, 128)
                ' Gruppe = alle Ankreuzzellen dieses Feldes (untereinander in Spalte C)
                nAct = 0
                For Each oid In mCfg.OptionIds(f.ID)
                    If mCfg.OptActive(CStr(oid)) Then nAct = nAct + 1
                Next
                gAddr = ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + nAct, 3)).Address(True, True)
                MarkTooManyPicks ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + nAct, 3)), gAddr, mValid.MaxPicks(f), wl.Range("ZZ1")
                For Each oid In mCfg.OptionIds(f.ID)
                    If mCfg.OptActive(CStr(oid)) Then
                        r = r + 1
                        ws.Cells(r, 2).Value = "      " & mCfg.OptName(CStr(oid))
                        Set rng = ws.Cells(r, 3)
                        StyleInput rng
                        rng.Locked = False
                        ApplyOptionInput rng, f, mCfg.OptName(CStr(oid)), gAddr, wl.Range("ZZ1")
                        ws.Rows(r).RowHeight = 22
                        InfoRow wi, infoR, Array("Zelle", "OPT", f.ID, CStr(oid), rng.Address(False, False))
                    End If
                Next
            Else
                Set rng = ws.Cells(r, 3)
                StyleInput rng
                ApplyInput wb, rng, f
                rng.Locked = False
                ws.Cells(r, 4).Value = HintFor(f)
                ws.Cells(r, 4).Font.Size = 8.5: ws.Cells(r, 4).Font.Color = RGB(107, 114, 128)
                ws.Cells(r, 4).WrapText = True: ws.Cells(r, 4).VerticalAlignment = xlTop
                ws.Rows(r).RowHeight = IIf(f.FType = FT_LTEXT, 66, IIf(Len(HintFor(f)) > 75, 38, 26))
                If f.FType = FT_LTEXT Then rng.WrapText = True: rng.VerticalAlignment = xlTop
                If f.Required Then AddRequiredFormat wl, rng, ""
                InfoRow wi, infoR, Array("Zelle", "FELD", f.ID, "", rng.Address(False, False))
            End If
            r = r + 1
        End If
    Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterFooter = "Erfassungsdatei · " & mCfg.GetS(SH_CAP, cid, "Name")
    End With
    Application.PrintCommunication = True
    ws.Activate
    ws.Cells(r0, 3).Select
End Sub

Private Function ColWidth(ByVal f As cField, ByVal showOnly As Boolean) As Double
    Select Case f.FType
        Case FT_LTEXT: ColWidth = IIf(showOnly, 30, 40)
        Case FT_DATUM, FT_JANEIN, FT_PROZENT: ColWidth = 13
        Case FT_DATUMZEIT: ColWidth = 17
        Case FT_ZAHL, FT_GELD: ColWidth = 13
        Case FT_AUSWAHL, FT_STATUS, FT_LINK: ColWidth = 20
        Case FT_MEHRFACH, FT_BACKLINK: ColWidth = 26
        Case Else: ColWidth = 22
    End Select
    If Len(f.Name) + 3 > ColWidth Then ColWidth = Application.Min(Len(f.Name) + 3, 36)
End Function

Private Function HasList(ByVal f As cField) As Boolean
    Select Case f.FType
        Case FT_AUSWAHL, FT_STATUS, FT_MEHRFACH, FT_JANEIN, FT_LINK: HasList = True
    End Select
End Function

' Auswahlwerte eines Feldes in die Listenspalte schreiben und als Name "L_<FeldID>" bereitstellen
Private Sub WriteList(ByVal wb As Workbook, ByVal wl As Worksheet, ByVal f As cField, ByVal col As Long)
    Dim v As Variant, n As Long, t As cTableData, r As Long
    wl.Columns(col).NumberFormat = "@"
    Select Case f.FType
        Case FT_JANEIN
            wl.Cells(1, col).Value = "Ja": wl.Cells(2, col).Value = "Nein": n = 2
        Case FT_LINK
            Set t = mData.TData(f.TargetTable)
            For r = 1 To t.RowCount
                If Not t.IsArchived(r) And n < 5000 Then
                    n = n + 1
                    wl.Cells(n, col).Value = mData.RecordLabel(f.TargetTable, t.IdAt(r))
                End If
            Next
        Case Else
            For Each v In mCfg.OptionIds(f.ID)
                If mCfg.OptActive(CStr(v)) Then
                    n = n + 1
                    wl.Cells(n, col).Value = mCfg.OptName(CStr(v))
                End If
            Next
    End Select
    If n = 0 Then wl.Cells(1, col).Value = "": n = 1
    On Error Resume Next
    wb.Names.Add Name:="L_" & f.ID, RefersTo:="='" & LIST_SHEET & "'!" & wl.Range(wl.Cells(1, col), wl.Cells(n, col)).Address
    On Error GoTo 0
End Sub

' Ankreuzzelle: nur "Ja" oder leer. Mit Höchstzahl (groupAddr = alle Ankreuzzellen dieses Feldes
' in derselben Zeile bzw. im Formular): Die Auswahlliste bietet "Ja" nur an, solange in der Gruppe
' noch nicht maxN Haken gesetzt sind (eigener Haken zählt nicht mit); Eintippen wird ebenso geprüft.
Private Sub ApplyOptionInput(ByVal rng As Range, ByVal f As cField, ByVal optName As String, _
                             Optional ByVal groupAddr As String = "", Optional ByVal scratch As Range = Nothing)
    Dim maxN As Long, fml As String, first As String, eng As String, limited As Boolean
    maxN = mValid.MaxPicks(f)
    On Error Resume Next
    rng.Validation.Delete
    If maxN > 0 And Len(groupAddr) > 0 And Not scratch Is Nothing Then
        first = rng.Cells(1, 1).Address(False, False)
        ' relative Bezüge gelten ab der aktiven Zelle: erste Zelle des Bereichs wählen
        eng = "=IF(COUNTIF(" & groupAddr & ",""Ja"")-(" & first & "=""Ja"")>=" & maxN & ",L_LEER,L_JA)"
        fml = LocalFormula(scratch, eng)
        rng.Worksheet.Parent.Activate
        rng.Worksheet.Activate
        rng.Cells(1, 1).Select
        ' aus VBA erwartet Validation.Add die englische Schreibweise; die lokale nur als Rückfall
        Err.Clear
        rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=eng
        If Err.Number <> 0 Then
            Err.Clear
            rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=fml
        End If
        limited = (Err.Number = 0)
        If Not limited Then
            gCapDiag = "Prüfregel: " & Err.Description & " | " & fml
            Err.Clear
            ' Rückfall: einfache Ankreuzliste; zu viele Haken markiert die rote Hervorhebung
            rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=L_JA"
        End If
        Err.Clear
    Else
        rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=L_JA"
    End If
    rng.Validation.InCellDropdown = True
    ' bei Höchstzahl verweist die Liste auf eine leere Zelle: "Leere ignorieren" würde dann jede Eingabe
    ' zulassen. Leeren einer Zelle bleibt trotzdem möglich (Excel prüft Löschen nicht).
    rng.Validation.IgnoreBlank = Not limited
    rng.Validation.InputTitle = Left$(f.Name, 32)
    rng.Validation.InputMessage = Left$("„" & optName & "“ trifft zu? Dann „Ja“ wählen, sonst leer lassen." & _
                                        IIf(maxN > 0, " Höchstens " & maxN & " ankreuzen.", ""), 250)
    rng.Validation.ShowInput = True
    rng.Validation.ErrorTitle = Left$(f.Name, 32)
    If maxN > 0 Then
        rng.Validation.ErrorMessage = "Höchstens " & maxN & " ankreuzen. Wenn schon " & maxN & "-mal „Ja“ gewählt ist, bitte zuerst ein anderes „Ja“ entfernen."
    Else
        rng.Validation.ErrorMessage = "Bitte „Ja“ wählen oder die Zelle leer lassen."
    End If
    On Error GoTo 0
End Sub

' Zu viele Haken (z. B. hineinkopiert): ganze Gruppe rot markieren
Private Sub MarkTooManyPicks(ByVal block As Range, ByVal groupAddr As String, ByVal maxN As Long, ByVal scratch As Range)
    Dim fc As FormatCondition
    If maxN <= 0 Then Exit Sub
    On Error Resume Next
    block.Worksheet.Activate
    block.Cells(1, 1).Select
    Set fc = block.FormatConditions.Add(Type:=xlExpression, Formula1:=LocalFormula(scratch, "=COUNTIF(" & groupAddr & ",""Ja"")>" & maxN))
    fc.Interior.Color = RGB(254, 226, 226)
    fc.Font.Color = RGB(153, 27, 27)
    On Error GoTo 0
End Sub

' Spalten einer Ankreuzgruppe (gleiches Feld, gleiche Art) im Spaltenplan
Private Sub GroupSpan(ByVal plan As Collection, ByVal fid As String, ByVal kind As String, ByRef c0 As Long, ByRef c1 As Long)
    Dim p As Variant
    c0 = 0: c1 = 0
    For Each p In plan
        If CStr(p(1)) = fid And CStr(p(0)) = kind Then
            If c0 = 0 Then c0 = p(3)
            c1 = p(3)
        End If
    Next
End Sub

' Pruefregel, Zahlenformat und Eingabehinweis einer Eingabezelle bzw. Spalte
Private Sub ApplyInput(ByVal wb As Workbook, ByVal rng As Range, ByVal f As cField)
    Dim mn As String, mx As String, hint As String
    hint = HintFor(f)
    On Error Resume Next
    rng.Validation.Delete
    Select Case f.FType
        Case FT_TEXT, FT_PFAD
            rng.NumberFormat = "@"
            If f.MaxLen > 0 Then rng.Validation.Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlLessEqual, Formula1:=CStr(f.MaxLen)
        Case FT_LTEXT
            rng.NumberFormat = "@"
            rng.WrapText = True
        Case FT_ZAHL, FT_GELD, FT_PROZENT
            If f.FType = FT_GELD Then
                rng.NumberFormat = "#,##0.00 €"
            ElseIf f.FType = FT_PROZENT Then
                rng.NumberFormat = "0" & IIf(f.Decimals > 0, "." & String$(f.Decimals, "0"), "") & " %"
            Else
                rng.NumberFormat = IIf(f.Decimals > 0, "#,##0." & String$(f.Decimals, "0"), IIf(mFmt.IsYearLike(f), "0", "#,##0"))
            End If
            rng.HorizontalAlignment = xlLeft
            mn = "-999999999999": mx = "999999999999"
            If Not IsEmpty(f.MinVal) Then mn = CStr(CDbl(f.MinVal))
            If Not IsEmpty(f.MaxVal) Then mx = CStr(CDbl(f.MaxVal))
            rng.Validation.Add Type:=IIf(f.FType = FT_ZAHL And f.Decimals = 0, xlValidateWholeNumber, xlValidateDecimal), _
                               AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=mn, Formula2:=mx
        Case FT_DATUM, FT_DATUMZEIT
            rng.NumberFormat = IIf(f.FType = FT_DATUM, "dd.mm.yyyy", "dd.mm.yyyy hh:mm")
            rng.HorizontalAlignment = xlLeft
            rng.Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="109574"
        Case FT_JANEIN, FT_AUSWAHL, FT_STATUS, FT_LINK
            rng.NumberFormat = "@"
            rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=L_" & f.ID
            rng.Validation.InCellDropdown = True
        Case FT_MEHRFACH
            ' sehr viele Optionen: als Text "A; B" mit Auswahlhilfe
            rng.NumberFormat = "@"
            rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:="=L_" & f.ID
            rng.Validation.InCellDropdown = True
            rng.Validation.ShowError = False
    End Select
    If rng.Validation.Type < 0 Or Err.Number <> 0 Then
        Err.Clear
        rng.Validation.Add Type:=xlValidateInputOnly
    End If
    rng.Validation.IgnoreBlank = True
    rng.Validation.InputTitle = Left$(f.Name, 32)
    rng.Validation.InputMessage = Left$(hint, 250)
    rng.Validation.ShowInput = True
    rng.Validation.ErrorTitle = Left$(f.Name, 32)
    rng.Validation.ErrorMessage = Left$("Ungültige Eingabe. " & hint, 250)
    On Error GoTo 0
End Sub

Private Function HintFor(ByVal f As cField) As String
    Dim s As String
    Select Case f.FType
        Case FT_DATUM: s = "Datum TT.MM.JJJJ"
        Case FT_DATUMZEIT: s = "Datum und Uhrzeit, z. B. 03.06.2025 14:30"
        Case FT_ZAHL: s = "Zahl" & IIf(f.Decimals > 0, " mit bis zu " & f.Decimals & " Nachkommastellen", "")
        Case FT_GELD: s = "Betrag in Euro, z. B. 1.250,00"
        Case FT_PROZENT: s = "Prozentwert, z. B. 25 für 25 %"
        Case FT_JANEIN: s = "Ja oder Nein (Pfeil)"
        Case FT_AUSWAHL, FT_STATUS: s = "Bitte aus der Liste wählen (Pfeil)"
        Case FT_MEHRFACH: s = "Ein Wert aus der Liste oder mehrere, durch Semikolon getrennt (z. B. A; B)"
        Case FT_LINK: s = "Eintrag aus der Liste wählen (" & mCfg.TableName(f.TargetTable) & ")"
        Case FT_LTEXT: s = "Freitext, Zeilenumbruch mit Alt+Enter"
        Case Else
            Select Case f.Kind
                Case FT_EMAIL: s = "E-Mail-Adresse, z. B. name@beispiel.de"
                Case FT_URL: s = "Webadresse, z. B. https://www.beispiel.de"
                Case FT_TELEFON: s = "Telefonnummer"
                Case FT_DATEIEN: s = "Dateipfad(e), mehrere durch | getrennt"
                Case Else: s = "Text" & IIf(f.MaxLen > 0, ", höchstens " & f.MaxLen & " Zeichen", "")
            End Select
    End Select
    If mFmt.IsYearLike(f) Then
        s = mFmt.YearHint(f, True)
    ElseIf Not IsEmpty(f.MinVal) And Not IsEmpty(f.MaxVal) And (f.FType = FT_ZAHL Or f.FType = FT_GELD) Then s = s & " zwischen " & FmtNum(CDbl(f.MinVal), 2, True) & " und " & FmtNum(CDbl(f.MaxVal), 2, True)
    End If
    If f.Required Then s = s & " · Pflichtangabe"
    If Len(f.Help) > 0 Then s = f.Help & " · " & s
    HintFor = s
End Function

' Rote Markierung fehlender Pflichtangaben (in Listen nur fuer begonnene Zeilen)
Private Sub AddRequiredFormat(ByVal scratchSheet As Worksheet, ByVal rng As Range, ByVal rowRef As String, Optional ByVal prefilled As Long = 0)
    Dim fml As String, fc As FormatCondition, first As String
    On Error Resume Next
    first = rng.Cells(1, 1).Address(False, False)
    If Len(rowRef) = 0 Then
        fml = "=ISBLANK(" & first & ")"
    Else
        fml = "=AND(COUNTA(" & rowRef & ")>" & prefilled & ",ISBLANK(" & first & "))"
    End If
    rng.Worksheet.Activate
    rng.Cells(1, 1).Select
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=LocalFormula(scratchSheet.Range("Z1"), fml))
    fc.Interior.Color = RGB(254, 226, 226)
    fc.Font.Color = RGB(153, 27, 27)
    On Error GoTo 0
End Sub

'==============================================================================
' Ruecklauf-Uebersicht: wer hat schon geantwortet?
'==============================================================================
Public Sub Cmd_Status()
    Dim root As String, s As String, parts As Variant, a As Variant, items() As Variant, i As Long, n As Long, t As cTableData, r As Long
    Dim fileId As String, got As Long, stamp As String
    On Error GoTo f
    If Not RequireSel() Then Exit Sub
    root = RootOf(mSel)
    s = mCfg.SysGet("Versand_" & root, "")
    If Len(s) = 0 Then mUI.Alert "Für „" & mCfg.GetS(SH_CAP, root, "Name") & "“ wurden noch keine Dateien erzeugt.", "Rücklauf-Übersicht": Exit Sub
    stamp = FmtStamp(InvDate(Left$(s, InStr(s, Chr$(29)) - 1)))
    parts = Split(Mid$(s, InStr(s, Chr$(29)) + 1), Chr$(30))
    ReDim items(0 To UBound(parts), 0 To 2)
    Set t = mCfg.TD(SH_RET)
    For i = 0 To UBound(parts)
        a = Split(parts(i), Chr$(31))
        If UBound(a) >= 1 Then
            fileId = CStr(a(1))
            items(n, 0) = IIf(Len(a(0)) > 0, a(0), "(eine Datei ohne Empfänger)")
            items(n, 1) = "offen"
            items(n, 2) = ""
            For r = 1 To t.RowCount
                If CStr(t.Cell(r, 1)) = fileId Then
                    items(n, 1) = ChrW(&H2713) & " eingegangen " & FmtStamp(t.Val_(r, "Importiert"))
                    items(n, 2) = NzS(t.Val_(r, "Absender")) & " · " & NzS(t.Val_(r, "Zeilen")) & " Zeilen"
                    got = got + 1
                End If
            Next
            n = n + 1
        End If
    Next
    mUI.ShowReport "Rücklauf-Übersicht", "„" & mCfg.GetS(SH_CAP, root, "Name") & "“, Versand vom " & stamp & ": " & got & " von " & n & " Dateien eingegangen.", _
                   items, 3, "190;190;200", Array("Schließen")
    Exit Sub
f: mApp.Boundary "Rücklauf-Übersicht"
End Sub

'==============================================================================
' Ruecklauf einlesen
'==============================================================================
Public Sub Cmd_Import()
    Dim files As Variant, i As Long, rec As Variant, items() As Variant, n As Long, msg As String, nConf As Long, skipConf As Boolean
    Dim errs As Collection, e As Variant, plan As Collection, okRecs As Long, tidSet As New cMap, res As Variant, ans As Long, k As Variant
    On Error GoTo f
    If Not mApp.RequireWritable("Das Einlesen von Rückläufen") Then Exit Sub
    EnsureSheets
    files = Application.GetOpenFilename("Excel-Dateien (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , "Ausgefüllte Erfassungsdateien auswählen (mehrere möglich)", , True)
    If Not IsArray(files) Then Exit Sub
    Set plan = New Collection
    Set errs = New Collection
    Set mRestoreAsked = Nothing
    mApp.BeginWork "Rückläufe werden geprüft …"
    Dim seenKey As New cMap, rr As Variant, dupKey As String
    For i = LBound(files) To UBound(files)
        rr = ReadReturn(CStr(files(i)), errs)
        ' dieselbe Antwort (gleiche Datei, gleicher Absender) zweimal ausgewählt: nur einmal übernehmen
        If CStr(rr(5)) <> "" Then
            dupKey = CStr(rr(6)) & "|" & IIf(Len(CStr(rr(1))) > 0, LCase$(CStr(rr(1))), LCase$(CStr(rr(0))))
            If seenKey.Exists(dupKey) Then
                rr = Array(rr(0), rr(1), New Collection, 0, "doppelt ausgewählt (dieselbe Antwort wie „" & seenKey.GetVal(dupKey) & "“), übersprungen", _
                           "", rr(6), rr(7), 0, rr(9), 0, 0, 0, New cMap)
            Else
                seenKey.SetVal dupKey, CStr(rr(0))
            End If
        End If
        plan.Add rr
    Next
    mApp.EndWork
    ' Uebersicht: Datei, Absender, neu, geaendert, unveraendert, Fehler, Hinweis
    ReDim items(0 To plan.Count - 1, 0 To 6)
    For Each rec In plan
        items(n, 0) = rec(0)
        items(n, 1) = rec(1)
        items(n, 2) = rec(12)
        items(n, 3) = rec(11)
        items(n, 4) = rec(10)
        items(n, 5) = rec(3)
        items(n, 6) = rec(4)
        okRecs = okRecs + rec(2).Count
        nConf = nConf + rec(8)
        For Each k In rec(13).Keys
            tidSet.SetVal CStr(k), True
        Next
        n = n + 1
    Next
    msg = "Geprüft: " & plan.Count & " Datei(en). Übernommen werden nur Dateien ohne Fehler; bereits eingelesene Dateien werden übersprungen. " & _
          "Vorhandene Einträge werden über ihre ID aktualisiert (nur geänderte Werte)." & _
          IIf(errs.Count > 0, vbLf & vbLf & "Fehlerhafte Zeilen: " & errs.Count, "")
    If okRecs = 0 Then
        mUI.ShowReport "Rücklauf einlesen", msg & vbLf & vbLf & "Es gibt nichts zu übernehmen (keine neuen oder geänderten Einträge).", items, 7, "170;140;40;62;78;46;300", Array("Schließen"), False, Array("Datei", "Absender", "neu", "geändert", "unverändert", "Fehler", "Hinweis")
        If errs.Count > 0 Then ShowErrors errs
        MarkUnchangedFiles plan
        Exit Sub
    End If
    Select Case mUI.ShowReport("Rücklauf einlesen: Vorschau", msg, items, 7, "170;140;40;62;78;46;300", IIf(errs.Count > 0, Array("Gültige übernehmen", "Fehler anzeigen", "Abbrechen"), Array("Übernehmen", "Abbrechen")), False, Array("Datei", "Absender", "neu", "geändert", "unverändert", "Fehler", "Hinweis"))
        Case 1
        Case 2
            If errs.Count > 0 Then ShowErrors errs
            Exit Sub
        Case Else
            Exit Sub
    End Select
    If nConf > 0 Then
        ans = mUI.Ask3(nConf & " Einträge wurden seit dem Versand der Datei in dieser Datenbank geändert. " & _
                       "Sollen die Rückmeldungen diese Einträge trotzdem aktualisieren (nur die zurückgemeldeten Spalten)?", "Konflikte", _
                       "Rückmeldungen übernehmen", "Diese Einträge auslassen", "Abbrechen")
        If ans = 0 Or ans = 3 Then Exit Sub
        skipConf = (ans = 2)
    End If
    If Not mBackup.OfferBackup("vor dem Einlesen der Rückläufe") Then Exit Sub
    res = ExecuteImport(plan, skipConf)
    ' Ergebnis und nächster Schritt in einem Fenster: Gesamtliste ansehen oder bei den Rückmeldungen bleiben
    msg = res(0) & " neue Einträge angelegt, " & res(1) & " Einträge aktualisiert" & IIf(res(2) > 0, ", " & res(2) & " Konflikte ausgelassen", "") & "." & vbLf & _
          "Jede Stelle wird je Datei nur einmal übernommen; das Rücklaufprotokoll vermerkt, was eingelesen wurde." & _
          IIf(errs.Count > 0, vbLf & vbLf & errs.Count & " fehlerhafte Zeile(n) wurden nicht übernommen.", "")
    If errs.Count > 0 Then
        mUI.Alert msg, "Rücklauf einlesen"
        If mUI.Confirm("Fehlerliste jetzt anzeigen?", "Rücklauf einlesen", "Anzeigen", "Später") Then ShowErrors errs
        msg = ""
    End If
    If tidSet.Count = 1 And mNav.gPage = "Erfassung" Then
        For Each e In tidSet.Keys
            If mUI.Confirm(IIf(Len(msg) > 0, msg & vbLf & vbLf, "") & "Die Einträge stehen jetzt in der Tabelle „" & mCfg.TableName(CStr(e)) & "“.", "Rücklauf eingelesen", _
                           "Tabelle anzeigen", "Bei den Rückmeldungen bleiben") Then
                mNav.GoTable CStr(e)
            Else
                mNav.RenderCurrent
            End If
        Next
    Else
        If Len(msg) > 0 Then mUI.Alert msg, "Rücklauf einlesen"
        If tidSet.Count = 1 Then
            For Each e In tidSet.Keys
                mNav.GoTable CStr(e)
            Next
        Else
            mNav.RenderCurrent
        End If
    End If
    Exit Sub
f:  mApp.ResetWork
    mApp.Boundary "Rücklauf einlesen"
End Sub

Private Sub ShowErrors(ByVal errs As Collection)
    Dim items() As Variant, e As Variant, n As Long
    ReDim items(0 To errs.Count - 1, 0 To 3)
    For Each e In errs
        items(n, 0) = e(0): items(n, 1) = e(1): items(n, 2) = e(2): items(n, 3) = e(3)
        n = n + 1
    Next
    mUI.ShowReport "Rücklauf: nicht übernommene Zeilen", "Diese Zeilen wurden nicht übernommen. Bitte in der jeweiligen Datei korrigieren und die Datei erneut einlesen " & _
                   "(eine Datei mit Fehlern gilt erst nach vollständiger Übernahme als eingelesen).", items, 4, "170;90;140;320", Array("Schließen"), False, Array("Datei", "Blatt / Zeile", "Spalte", "Problem")
End Sub

' Pruefdateien ohne jede Aenderung trotzdem als eingegangen vermerken
Private Sub MarkUnchangedFiles(ByVal plan As Collection)
    Dim p As Variant, any_ As Boolean
    For Each p In plan
        If p(3) = 0 And p(10) > 0 And p(2).Count = 0 And Len(p(6)) > 0 Then
            If Not AlreadyImported(CStr(p(6)), CStr(p(1)), CStr(p(0))) Then RecordReturn p, 0: any_ = True
        End If
    Next
    If any_ Then mApp.MarkCfgDirty
End Sub

Private Sub RecordReturn(ByVal p As Variant, ByVal n As Long)
    Dim m As New cMap
    m.SetVal "DateiID", CStr(p(6))
    m.SetVal "Datei", CStr(p(0))
    m.SetVal "Erfassung", CStr(p(7))
    m.SetVal "Absender", CStr(p(1))
    m.SetVal "Zeilen", CStr(n)
    m.SetVal "Importiert", Now
    m.SetVal "Benutzer", UserId()
    m.SetVal "Empfaenger", CStr(p(9))
    mCfg.AddRow SH_RET, m
End Sub

' Erfassung (und ihre Blätter) aus der _Info einer Rücklaufdatei wiederherstellen. Gelingt nur,
' wenn die Datei aus dieser Datenbank stammt und die Tabelle noch existiert; fragt einmal je Erfassung.
Private Function RestoreCapture(ByVal info As cMap, ByVal sheets As Collection, ByVal fname As String, ByRef note As String) As Boolean
    Dim cid As String, tid As String, dbFile As String, sh As Variant, c As Variant, m As cMap, sel As String, first As Boolean, scid As String
    cid = info.GetVal("ErfassungID", "")
    tid = info.GetVal("TabelleID", "")
    dbFile = info.GetVal("DatenbankID", "")
    If mRestoreAsked Is Nothing Then Set mRestoreAsked = New cMap
    If Len(cid) = 0 Or Not mCfg.TableExists(tid) Then
        note = "gehört zu keiner Erfassung dieser Datei (Erfassung entfernt oder Tabelle gelöscht)"
        Exit Function
    End If
    If Len(dbFile) > 0 And Len(mCfg.SysGet("DatenbankID", "")) > 0 And dbFile <> mCfg.SysGet("DatenbankID", "") Then
        note = "stammt aus einer anderen Datenbankdatei, übersprungen"
        Exit Function
    End If
    If mRestoreAsked.Exists(cid) Then
        RestoreCapture = CBool(mRestoreAsked.GetVal(cid))
        If Not RestoreCapture Then note = "Erfassung nicht wiederhergestellt, übersprungen"
        Exit Function
    End If
    If mApp.gReadOnly Then note = "Erfassung fehlt in dieser Datei (Lesemodus, nicht wiederherstellbar)": Exit Function
    If Not mUI.Confirm("Die Datei „" & fname & "“ gehört zur Erfassung „" & info.GetVal("Erfassung", "") & "“, die in dieser Datei nicht gespeichert ist." & vbLf & vbLf & _
                       "Das passiert, wenn die Hauptdatei nach dem Erzeugen der Erfassungsdatei ohne Speichern geschlossen wurde. " & _
                       "Die Erfassung lässt sich aus der Rücklaufdatei wiederherstellen; danach werden die Einträge wie gewohnt eingelesen.", _
                       "Erfassung wiederherstellen", "Wiederherstellen und einlesen", "Überspringen") Then
        mRestoreAsked.SetVal cid, False
        note = "Erfassung nicht wiederhergestellt, übersprungen"
        Exit Function
    End If
    If Len(mCfg.SysGet("DatenbankID", "")) = 0 And Len(dbFile) > 0 Then mCfg.SysSet "DatenbankID", dbFile
    first = True
    For Each sh In sheets
        scid = CStr(sh(1))
        If Len(scid) > 0 And Len(mCfg.GetS(SH_CAP, scid, "ID")) = 0 And mCfg.TableExists(CStr(sh(2))) Then
            sel = ""
            For Each c In sh(5)
                If c(0) = "FELD" Or c(0) = "OPT" Then sel = ListAdd(sel, CStr(c(1)))
            Next
            Set m = New cMap
            m.SetVal "ID", scid
            m.SetVal "Name", IIf(first, info.GetVal("Erfassung", "Erfassung"), info.GetVal("Erfassung", "") & " / " & CStr(sh(0)))
            m.SetVal "Titel", IIf(first, "", CStr(sh(0)))
            m.SetVal "Tabelle", CStr(sh(2))
            m.SetVal "Modus", IIf(CStr(sh(8)) = "formular", "Formular", "Liste")
            m.SetVal "Felder", sel
            m.SetVal "Zeilen", "200"
            m.SetVal "Schutz", "Ja"
            m.SetVal "Absender", IIf(Len(info.GetVal("AbsenderZelle", "")) > 0, "Ja", "Nein")
            m.SetVal "AbsenderFeld", IIf(m.GetVal("Absender") = "Ja", EnsureSenderField(CStr(sh(2))), "")
            m.SetVal "Optionen", KVSet(KVSet("", "art", IIf(CStr(sh(8)) = "pruefen", "pruefen", "")), "eltern", IIf(first, "", cid))
            m.SetVal "Erstellt", Now
            m.SetVal "Version", info.GetVal("Version", "1")
            mCfg.AddRow SH_CAP, m
            first = False
        End If
    Next
    mLog.Add "Erfassung wiederhergestellt", tid, cid, "", "aus " & fname
    mApp.MarkCfgDirty
    mRestoreAsked.SetVal cid, True
    RestoreCapture = True
End Function

' Wert "angekreuzt"?
Private Function IsTicked(ByVal v As Variant) As Boolean
    Dim s As String
    If IsBlank(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbBoolean Then IsTicked = CBool(v): Exit Function
    s = LCase$(Trim$(CStr(v)))
    Select Case s
        Case "ja", "x", "1", "wahr", "true", "j", "yes": IsTicked = True
    End Select
    If s = ChrW(&H2713) Or s = ChrW(&H2714) Then IsTicked = True
End Function

Private Function SameVal(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsBlank(a) And IsBlank(b) Then SameVal = True: Exit Function
    If IsBlank(a) Or IsBlank(b) Then Exit Function
    If VarType(a) = vbBoolean Or VarType(b) = vbBoolean Then
        On Error Resume Next
        SameVal = (CBool(a) = CBool(b))
        Exit Function
    End If
    If (IsNumeric(a) Or IsDate(a)) And (IsNumeric(b) Or IsDate(b)) And VarType(a) <> vbString And VarType(b) <> vbString Then
        SameVal = (Abs(CDbl(a) - CDbl(b)) < 0.0000001)
    Else
        SameVal = (CStr(a) = CStr(b))
    End If
End Function

' Liest eine Datei. Ergebnis:
'   0 Dateiname, 1 Absender, 2 Collection gueltiger Saetze, 3 Fehlerzahl, 4 Hinweis, 5 TabelleID (erste),
'   6 DateiID, 7 ErfassungID, 8 Konflikte, 9 Empfaenger, 10 unveraendert, 11 geaendert, 12 neu, 13 cMap Tabellen
' Ein gueltiger Satz ist Array(cMap Werte, cMap Verknuepfungen, Zeile, TabelleID, DatensatzID, Konflikt)
Private Function ReadReturn(ByVal path As String, ByVal errs As Collection) As Variant
    Dim wb As Workbook, wi As Worksheet, ws As Worksheet, prevSec As MsoAutomationSecurity, info As New cMap, r As Long, fname As String
    Dim cid As String, tid As String, fileId As String, sender As String, recs As New Collection, nErr As Long, note As String
    Dim sheets As New Collection, sh As Variant, cols As Collection, legacy As Boolean, nConf As Long, nSame As Long, nUpd As Long, nNew As Long
    Dim tids As New cMap, a As Variant, cur As Collection, absWs As Worksheet
    fname = Mid$(path, InStrRev(path, "\") + 1)
    ' Datei schon in Excel geöffnet (z. B. zum Nachsehen)? Dann direkt von dort lesen
    Dim openWb As Workbook, keepOpen As Boolean
    For Each openWb In Application.Workbooks
        If StrComp(openWb.Name, fname, vbTextCompare) = 0 Then
            If StrComp(openWb.FullName, path, vbTextCompare) = 0 Then
                Set wb = openWb: keepOpen = True
            Else
                note = "eine andere Datei mit demselben Namen ist in Excel geöffnet; bitte diese zuerst schließen"
                GoTo done
            End If
        End If
    Next
    If Not keepOpen Then
        prevSec = Application.AutomationSecurity
        On Error GoTo openFail
        Application.AutomationSecurity = msoAutomationSecurityForceDisable
        Set wb = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, Notify:=False)
        Application.AutomationSecurity = prevSec
        wb.Windows(1).Visible = False
        On Error GoTo 0
    End If
    On Error Resume Next
    Set wi = wb.Worksheets(INFO_SHEET)
    On Error GoTo 0
    If wi Is Nothing Then note = "keine Erfassungsdatei (aus diesem Programm)": GoTo done
    ' _Info lesen: Schluessel, Blaetter und ihre Spalten
    For r = 1 To 5000
        If Len(CStr(wi.Cells(r, 1).Value)) = 0 Then Exit For
        Select Case CStr(wi.Cells(r, 1).Value)
            Case "Blatt"
                Set cols = New Collection
                sheets.Add Array(CStr(wi.Cells(r, 2).Value), CStr(wi.Cells(r, 3).Value), CStr(wi.Cells(r, 4).Value), _
                                 CLng(Val(wi.Cells(r, 5).Value)), CLng(Val(wi.Cells(r, 6).Value)), cols, 0, 0, "")
            Case "Spalten"
                a = sheets(sheets.Count)
                sheets.Remove sheets.Count
                sheets.Add Array(a(0), a(1), a(2), a(3), a(4), a(5), CLng(Val(wi.Cells(r, 2).Value)), CLng(Val(wi.Cells(r, 3).Value)), CStr(wi.Cells(r, 4).Value))
            Case "Spalte", "Zelle"
                If sheets.Count > 0 Then sheets(sheets.Count)(5).Add Array(CStr(wi.Cells(r, 2).Value), CStr(wi.Cells(r, 3).Value), CStr(wi.Cells(r, 4).Value), CStr(wi.Cells(r, 5).Value))
            Case "Feld"
                ' Aufbau 1 (Version 1.1): Feld | ID | Spalte bzw. Zelladresse
                legacy = True
                If cur Is Nothing Then Set cur = New Collection
                cur.Add Array("FELD", CStr(wi.Cells(r, 2).Value), "", CStr(wi.Cells(r, 3).Value))
            Case Else
                info.SetVal CStr(wi.Cells(r, 1).Value), CStr(wi.Cells(r, 2).Value)
        End Select
    Next
    If info.GetVal("Kennung", "") <> MARK Then note = "keine Erfassungsdatei (aus diesem Programm)": GoTo done
    cid = info.GetVal("ErfassungID", "")
    fileId = info.GetVal("DateiID", "")
    tid = mCfg.GetS(SH_CAP, cid, "Tabelle")
    If Len(mCfg.GetS(SH_CAP, cid, "ID")) = 0 Then
        ' Erfassung hier unbekannt: meist wurde diese Datei nach dem Erzeugen nicht gespeichert.
        ' Aus den Angaben der Rücklaufdatei wiederherstellen (nach Rückfrage)
        If Not RestoreCapture(info, sheets, fname, note) Then tid = "": GoTo done
        tid = mCfg.GetS(SH_CAP, cid, "Tabelle")
    End If
    If tid <> info.GetVal("TabelleID", "") Or Not mCfg.TableExists(tid) Then note = "Zieltabelle passt nicht": tid = "": GoTo done
    If Len(info.GetVal("DatenbankID", "")) > 0 And info.GetVal("DatenbankID", "") <> DbId() Then note = "stammt aus einer anderen Datenbankdatei, übersprungen": tid = "": GoTo done
    If legacy Then
        ' aeltere Dateien: ein Blatt "Erfassung" mit Tabelle (Liste) oder Zelladressen (Formular)
        Set ws = wb.Worksheets(CAPT_SHEET)
        If info.GetVal("Modus", "") = "Formular" Then
            Set cols = New Collection
            For Each a In cur
                cols.Add Array("FELD", a(1), "", a(3))
            Next
            sheets.Add Array(CAPT_SHEET, cid, tid, 0, 0, cols, 0, 0, "formular")
        Else
            Dim lo_ As ListObject
            On Error Resume Next
            Set lo_ = ws.ListObjects("Erfassung")
            On Error GoTo 0
            If lo_ Is Nothing Then note = "Tabelle „Erfassung“ fehlt": tid = "": GoTo done
            sheets.Add Array(CAPT_SHEET, cid, tid, lo_.DataBodyRange.Row, lo_.DataBodyRange.Row + lo_.ListRows.Count - 1, cur, 0, 0, "neu")
        End If
    End If
    ' Absender
    If Len(info.GetVal("AbsenderZelle", "")) > 0 Then
        On Error Resume Next
        If Len(info.GetVal("AbsenderBlatt", "")) > 0 Then Set absWs = wb.Worksheets(info.GetVal("AbsenderBlatt")) Else Set absWs = wb.Worksheets(CAPT_SHEET)
        sender = Trim$(CStr(absWs.Range(info.GetVal("AbsenderZelle")).Value))
        On Error GoTo 0
    End If
    If mCfg.GetS(SH_CAP, cid, "Absender") = "Ja" And Len(sender) = 0 Then
        errs.Add Array(fname, "", COpt(cid, "abslabel", "Ihr Name / Ihre Stelle"), "Pflichtangabe: bitte in der Datei ergänzen.")
        nErr = nErr + 1
        note = "Absender fehlt"
        tid = ""
        GoTo done
    End If
    ' doppelt? Dieselbe Datei geht oft an viele, entscheidend sind Datei UND Absender
    If AlreadyImported(fileId, sender, fname) Then
        note = IIf(Len(sender) > 0, "von „" & sender & "“ bereits eingelesen, übersprungen", "bereits eingelesen, übersprungen")
        tid = ""
        GoTo done
    End If
    For Each sh In sheets
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(sh(0)))
        On Error GoTo 0
        If ws Is Nothing Then
            errs.Add Array(fname, CStr(sh(0)), "", "Blatt fehlt in der Datei.")
            nErr = nErr + 1
        ElseIf Not mCfg.TableExists(CStr(sh(2))) Or Len(mCfg.GetS(SH_CAP, CStr(sh(1)), "ID")) = 0 Then
            errs.Add Array(fname, CStr(sh(0)), "", "Die Tabelle bzw. das Blatt gibt es in dieser Datenbank nicht mehr.")
            nErr = nErr + 1
        Else
            ReadSheet ws, sh, fname, sender, info.GetVal("Empfaenger", ""), errs, recs, nErr, nConf, nSame, nUpd, nNew
            tids.SetVal CStr(sh(2)), True
        End If
    Next
    If nErr > 0 Then
        note = "Fehler in " & nErr & " Zeile(n), Datei bleibt als nicht eingelesen markiert"
    ElseIf recs.Count = 0 And nSame > 0 Then
        note = "geprüft, keine Änderungen"
    ElseIf recs.Count = 0 Then
        note = "keine ausgefüllten Zeilen"
    Else
        note = "bereit" & IIf(nConf > 0, " · " & nConf & " seit Versand geändert", "")
    End If
done:
    On Error Resume Next
    If Not wb Is Nothing And Not keepOpen Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    On Error GoTo 0
    If nErr > 0 Then Set recs = New Collection: nConf = 0: nUpd = 0: nNew = 0
    ReadReturn = Array(fname, sender, recs, nErr, note, tid, fileId, cid, nConf, info.GetVal("Empfaenger", ""), nSame, nUpd, nNew, tids)
    Exit Function
openFail:
    Application.AutomationSecurity = prevSec
    note = "konnte nicht geöffnet werden: " & Err.Description
    Resume done
End Function

' Liest ein Blatt: sh = Array(Name, ErfassungID, TabelleID, ErsteZeile, LetzteZeile, Spalten, IDSpalte, StandSpalte, Art)
Private Sub ReadSheet(ByVal ws As Worksheet, ByVal sh As Variant, ByVal fname As String, ByVal sender As String, ByVal recipient As String, _
                      ByVal errs As Collection, ByVal recs As Collection, _
                      ByRef nErr As Long, ByRef nConf As Long, ByRef nSame As Long, ByRef nUpd As Long, ByRef nNew As Long)
    Dim tid As String, cid As String, isForm As Boolean, r As Long, rFrom As Long, rTo As Long, c As Variant, f As cField, cellV As Variant
    Dim useCol As Boolean, o As Variant
    Dim vals As cMap, links As cMap, multi As cMap, rowErr As Boolean, anyVal As Boolean, rid As String, pos As String, outV As Variant, msg As String
    Dim ve As cMap, k As Variant, absFid As String, cur As Variant, changed As cMap, chLinks As cMap, conflict As Boolean, stand As Variant, t As cTableData
    tid = CStr(sh(2)): cid = CStr(sh(1))
    isForm = (CStr(sh(8)) = "formular")
    absFid = mCfg.GetS(SH_CAP, cid, "AbsenderFeld")
    If Len(absFid) > 0 And Not mCfg.FieldExists(absFid) Then absFid = ""
    Set t = mData.TData(tid)
    If isForm Then
        rFrom = 1: rTo = 1
    Else
        rFrom = CLng(sh(3)): rTo = CLng(sh(4))
    End If
    For r = rFrom To rTo
        Set vals = New cMap: Set links = New cMap: Set multi = New cMap
        rowErr = False: anyVal = False: rid = "": conflict = False
        pos = IIf(isForm, ws.Name, ws.Name & " Zeile " & r)
        ' Trennzeile "Neue Einträge" überspringen
        If Not isForm Then
            If CStr(ws.Cells(r, 1).Value) = SEP_MARK Then GoTo nextRow
        End If
        If CLng(sh(6)) > 0 Then rid = Trim$(CStr(ws.Cells(r, CLng(sh(6))).Value))
        ' Werte sammeln
        For Each c In sh(5)
            useCol = (c(0) = "FELD" Or c(0) = "OPT")
            ' graue Anzeigespalten zaehlen nur in den leeren Zeilen fuer neue Eintraege
            If (c(0) = "ANZ" Or c(0) = "OPTA") And Len(rid) = 0 And mCfg.FieldExists(CStr(c(1))) Then
                useCol = Capturable(mCfg.Field(CStr(c(1)))) And CStr(c(1)) <> COpt(cid, "teilen")
            End If
            If useCol Then
                If mCfg.FieldExists(CStr(c(1))) Then
                    If isForm Then cellV = ws.Range(CStr(c(3))).Value Else cellV = ws.Cells(r, CLng(c(3))).Value
                    If c(0) = "OPT" Or c(0) = "OPTA" Then
                        If Not multi.Exists(CStr(c(1))) Then multi.SetVal CStr(c(1)), ""
                        If IsTicked(cellV) Then
                            multi.SetVal CStr(c(1)), multi.GetVal(CStr(c(1))) & IIf(Len(multi.GetVal(CStr(c(1)))) > 0, "; ", "") & mCfg.OptName(CStr(c(2)))
                            anyVal = True
                        ElseIf Not IsBlank(cellV) Then
                            errs.Add Array(fname, pos, mCfg.FieldName(CStr(c(1))) & ": " & mCfg.OptName(CStr(c(2))), "Bitte „Ja“ wählen oder leer lassen.")
                            rowErr = True
                        End If
                    Else
                        Set f = mCfg.Field(CStr(c(1)))
                        If Not IsBlank(cellV) Then anyVal = True
                        If Len(rid) > 0 Then cur = mData.GetValue(tid, rid, f.ID) Else cur = Empty
                        If f.FType = FT_LINK Then
                            If Not IsBlank(cellV) Then
                                outV = LinkByLabel(f, CStr(cellV), msg)
                                If Len(outV) > 0 Then
                                    links.SetVal f.ID, outV
                                Else
                                    errs.Add Array(fname, pos, f.Name, msg): rowErr = True
                                End If
                            Else
                                links.SetVal f.ID, ""
                            End If
                        ElseIf mValid.ParseCell(f, cellV, outV, msg, NzS(cur)) Then
                            vals.SetVal f.ID, outV
                        Else
                            errs.Add Array(fname, pos, f.Name, msg): rowErr = True
                        End If
                    End If
                End If
            End If
        Next
        ' Ankreuzspalten zur Mehrfachauswahl zusammensetzen (deaktivierte Optionen eines Eintrags bleiben erhalten)
        For Each k In multi.Keys
            Set f = mCfg.Field(CStr(k))
            If Len(rid) > 0 Then cur = mData.GetValue(tid, rid, f.ID) Else cur = Empty
            If Len(rid) > 0 And Not IsBlank(cur) Then
                For Each o In Split(CStr(cur), ";")
                    If Len(o) > 0 Then
                        If Not mCfg.OptActive(CStr(o)) Then multi.SetVal CStr(k), multi.GetVal(CStr(k)) & IIf(Len(multi.GetVal(CStr(k))) > 0, "; ", "") & mCfg.OptName(CStr(o))
                    End If
                Next
            End If
            If mValid.ParseInput(f, CStr(multi.GetVal(CStr(k))), outV, msg, NzS(cur)) Then
                vals.SetVal f.ID, outV
            Else
                errs.Add Array(fname, pos, f.Name, msg): rowErr = True
            End If
        Next
        If Len(rid) > 0 Then
            ' vorhandener Eintrag: nur geaenderte Werte
            If Not mData.RecordExists(tid, rid) Then
                errs.Add Array(fname, pos, "", "Diesen Eintrag gibt es in der Datenbank nicht mehr (gelöscht).")
                nErr = nErr + 1
            ElseIf rowErr Then
                nErr = nErr + 1
            Else
                Set changed = New cMap: Set chLinks = New cMap
                For Each k In vals.Keys
                    If Not SameFieldVal(mCfg.Field(CStr(k)), vals.GetVal(CStr(k)), mData.GetValue(tid, rid, CStr(k))) Then changed.SetVal CStr(k), vals.GetVal(CStr(k))
                Next
                For Each k In links.Keys
                    If CStr(links.GetVal(CStr(k))) <> mData.LinkTargets(CStr(k), rid) Then chLinks.SetVal CStr(k), links.GetVal(CStr(k))
                Next
                If changed.Count = 0 And chLinks.Count = 0 Then
                    nSame = nSame + 1
                Else
                    If Len(absFid) > 0 And Len(sender) > 0 Then changed.SetVal absFid, sender
                    If Not mValid.ValidateRecord(tid, rid, changed, chLinks, ve) Then
                        For Each k In ve.Keys
                            errs.Add Array(fname, pos, mCfg.FieldName(CStr(k)), CStr(ve.GetVal(CStr(k))))
                        Next
                        nErr = nErr + 1
                    Else
                        stand = Empty
                        If CLng(sh(7)) > 0 Then stand = ws.Cells(r, CLng(sh(7))).Value
                        If IsNumeric(stand) And Not IsBlank(stand) Then
                            conflict = (CDbl(CDate(mData.GetValue(tid, rid, C_CHANGED))) > CDbl(stand) + 0.00001)
                        End If
                        If conflict Then nConf = nConf + 1
                        recs.Add Array(changed, chLinks, r, tid, rid, conflict)
                        nUpd = nUpd + 1
                    End If
                End If
            End If
        ElseIf anyVal Then
            ' neuer Eintrag
            For Each k In links.Keys
                If Len(links.GetVal(CStr(k))) = 0 Then links.Remove CStr(k)
            Next
            If Len(absFid) > 0 Then vals.SetVal absFid, sender
            ' neue Zeilen einer aufgeteilten Pruefdatei gehoeren zum Empfaenger (z. B. zur Kita)
            AssignRecipient cid, recipient, vals, links, fname, pos, errs, rowErr
            If rowErr Then
                nErr = nErr + 1
            ElseIf Not mValid.ValidateRecord(tid, "", vals, links, ve) Then
                For Each k In ve.Keys
                    errs.Add Array(fname, pos, mCfg.FieldName(CStr(k)), CStr(ve.GetVal(CStr(k))))
                Next
                nErr = nErr + 1
            Else
                recs.Add Array(vals, links, r, tid, "", False)
                nNew = nNew + 1
            End If
        End If
nextRow:
    Next
End Sub

' Vergleich gespeicherter und zurueckgemeldeter Werte (Kontrollkaestchen leer = Nein, Mehrfachauswahl ohne Reihenfolge)
Private Function SameFieldVal(ByVal f As cField, ByVal newV As Variant, ByVal curV As Variant) As Boolean
    Dim a As Variant, i As Long
    If f.Kind = FT_KONTROLL Then
        If IsBlank(curV) Then curV = False
        If IsBlank(newV) Then newV = False
    End If
    If f.FType = FT_MEHRFACH Then
        If ListCount(NzS(newV)) <> ListCount(NzS(curV)) Then Exit Function
        a = Split(NzS(newV), ";")
        For i = 0 To UBound(a)
            If Not ListHas(NzS(curV), CStr(a(i))) Then Exit Function
        Next
        SameFieldVal = True
        Exit Function
    End If
    SameFieldVal = SameVal(newV, curV)
End Function

' Setzt bei neuen Zeilen das Aufteilungsfeld auf den Empfaenger
Private Sub AssignRecipient(ByVal cid As String, ByVal recipient As String, ByVal vals As cMap, ByVal links As cMap, _
                            ByVal fname As String, ByVal pos As String, ByVal errs As Collection, ByRef rowErr As Boolean)
    Dim sf As cField, msg As String, v As String, outV As Variant
    If Len(recipient) = 0 Or Not mCfg.FieldExists(COpt(cid, "teilen")) Then Exit Sub
    Set sf = mCfg.Field(COpt(cid, "teilen"))
    If sf.FType = FT_LINK Then
        If links.Exists(sf.ID) Then Exit Sub
        v = LinkByLabel(sf, recipient, msg)
        If Len(v) > 0 Then links.SetVal sf.ID, v
    Else
        If vals.Exists(sf.ID) Then
            If Not IsBlank(vals.GetVal(sf.ID)) Then Exit Sub
        End If
        If mValid.ParseInput(sf, recipient, outV, msg) Then vals.SetVal sf.ID, outV
    End If
End Sub

' Schon eingelesen? Gleiche Datei-ID und gleicher Absender (ohne Absender: gleicher Dateiname)
Private Function AlreadyImported(ByVal fileId As String, ByVal sender As String, ByVal fname As String) As Boolean
    Dim t As cTableData, r As Long
    If Len(fileId) = 0 Then Exit Function
    Set t = mCfg.TD(SH_RET)
    For r = 1 To t.RowCount
        If CStr(t.Cell(r, 1)) = fileId Then
            If Len(sender) > 0 Then
                If StrComp(Trim$(CStr(t.Val_(r, "Absender"))), sender, vbTextCompare) = 0 Then AlreadyImported = True: Exit Function
            ElseIf StrComp(CStr(t.Val_(r, "Datei")), fname, vbTextCompare) = 0 Then
                AlreadyImported = True: Exit Function
            End If
        End If
    Next
End Function

Private Function LinkByLabel(ByVal f As cField, ByVal label As String, ByRef msg As String) As String
    Dim t As cTableData, r As Long, hit As String, n As Long
    Set t = mData.TData(f.TargetTable)
    For r = 1 To t.RowCount
        If Not t.IsArchived(r) Then
            If StrComp(mData.RecordLabel(f.TargetTable, t.IdAt(r)), Trim$(label), vbTextCompare) = 0 Then
                n = n + 1: hit = t.IdAt(r)
            End If
        End If
    Next
    If n = 1 Then
        LinkByLabel = hit
    ElseIf n = 0 Then
        msg = """" & label & """ wurde in " & mCfg.TableName(f.TargetTable) & " nicht gefunden."
    Else
        msg = """" & label & """ ist in " & mCfg.TableName(f.TargetTable) & " mehrdeutig."
    End If
End Function

' Uebernimmt alle gueltigen Saetze in einer Transaktion und vermerkt die Dateien.
' Liefert Array(neu, aktualisiert, Konflikte ausgelassen)
Private Function ExecuteImport(ByVal plan As Collection, Optional ByVal skipConflicts As Boolean = False) As Variant
    Dim p As Variant, rec As Variant, rid As String, a As Variant, t As cTableData, k As Variant, c As Long, nNew As Long, nUpd As Long, nSkip As Long
    Dim nFile As Long, tid As String
    mApp.BeginWork "Rückläufe werden übernommen …"
    On Error GoTo failed
    mData.TxBegin "Rücklauf einlesen"
    For Each p In plan
        If p(2).Count > 0 Then
            nFile = 0
            For Each rec In p(2)
                tid = CStr(rec(3))
                If Len(rec(4)) = 0 Then
                    Set t = mData.TData(tid)
                    rid = mData.NewRecordId()
                    a = mData.NewRowArray(tid, rid)
                    For Each k In rec(0).Keys
                        c = t.ColOf(CStr(k))
                        If c > 0 Then a(c) = rec(0).GetVal(CStr(k))
                    Next
                    mData.InsertRecord tid, a
                    For Each k In rec(1).Keys
                        mData.SetLinks CStr(k), rid, CStr(rec(1).GetVal(CStr(k)))
                    Next
                    nNew = nNew + 1: nFile = nFile + 1
                ElseIf rec(5) And skipConflicts Then
                    nSkip = nSkip + 1
                Else
                    rid = CStr(rec(4))
                    If rec(0).Count > 0 Then mData.UpdateRecord tid, rid, rec(0)
                    For Each k In rec(1).Keys
                        mData.SetLinks CStr(k), rid, CStr(rec(1).GetVal(CStr(k)))
                    Next
                    mLog.Add "Geändert", tid, rid, Join(rec(0).Keys, ";"), "Rücklauf " & p(1)
                    nUpd = nUpd + 1: nFile = nFile + 1
                End If
            Next
            mLog.Add "Rücklauf eingelesen", CStr(p(5)), CStr(p(7)), "", p(0) & " (" & nFile & ")"
        End If
    Next
    mData.TxCommit
    ' Dateien vermerken (ausserhalb der Datentransaktion: Konfiguration)
    For Each p In plan
        If (p(2).Count > 0 Or (p(3) = 0 And p(10) > 0)) And Len(p(6)) > 0 Then RecordReturn p, p(2).Count
    Next
    mCalc.RecalcAll
    mApp.EndWork
    mApp.MarkDirty
    ExecuteImport = Array(nNew, nUpd, nSkip)
    Exit Function
failed:
    Dim emsg As String
    emsg = Err.Description
    mData.TxAbortIfOpen
    mApp.ResetWork
    RaiseUser "Die Übernahme ist fehlgeschlagen; es wurde nichts übernommen (" & emsg & ")."
End Function

Public Sub Cmd_Log()
    Dim t As cTableData, r As Long, items() As Variant, n As Long
    On Error GoTo f
    EnsureSheets
    Set t = mCfg.TD(SH_RET)
    If t.RowCount = 0 Then mUI.Alert "Es wurden noch keine Rückläufe eingelesen.", "Rücklaufprotokoll": Exit Sub
    ReDim items(0 To t.RowCount - 1, 0 To 4)
    For r = t.RowCount To 1 Step -1
        items(n, 0) = FmtStamp(t.Val_(r, "Importiert"))
        items(n, 1) = mCfg.GetS(SH_CAP, CStr(t.Val_(r, "Erfassung")), "Name")
        items(n, 2) = NzS(t.Val_(r, "Datei"))
        items(n, 3) = NzS(t.Val_(r, "Absender"))
        items(n, 4) = NzS(t.Val_(r, "Zeilen"))
        n = n + 1
    Next
    mUI.ShowReport "Rücklaufprotokoll", "Eingelesene Erfassungsdateien (neueste zuerst). Jede Datei wird anhand ihrer Datei-ID nur einmal übernommen.", items, 5, "110;150;180;120;50", Array("Schließen")
    Exit Sub
f: mApp.Boundary "Rücklaufprotokoll"
End Sub

'==============================================================================
' Selbsttest (nur Demokopie): Dateien erzeugen, wie Empfaenger ausfuellen, einlesen
'==============================================================================
' Zelle eines Feldes (bzw. einer Ankreuzspalte) in Zeile i (1 = erste Datenzeile) eines Blattes
Private Function TestCell(ByVal wb As Workbook, ByVal sheetNo As Long, ByVal i As Long, ByVal fid As String, Optional ByVal oid As String = "") As Range
    Dim wi As Worksheet, r As Long, blk As Long, ws As Worksheet, r0 As Long, sep As Long
    Set wi = wb.Worksheets(INFO_SHEET)
    For r = 1 To 5000
        If Len(CStr(wi.Cells(r, 1).Value)) = 0 Then Exit For
        If CStr(wi.Cells(r, 1).Value) = "Blatt" Then
            blk = blk + 1
            If blk = sheetNo Then Set ws = wb.Worksheets(CStr(wi.Cells(r, 2).Value)): r0 = CLng(Val(wi.Cells(r, 5).Value))
        ElseIf blk = sheetNo And CStr(wi.Cells(r, 1).Value) = "Trenner" Then
            ' Zeilennummer i zählt nur Datenzeilen: ab dem Trennblock entsprechend weiter
            sep = CLng(Val(wi.Cells(r, 2).Value))
            If r0 + i - 1 >= sep Then i = i + Application.Max(1, CLng(Val(wi.Cells(r, 3).Value)))
        ElseIf blk = sheetNo And (CStr(wi.Cells(r, 1).Value) = "Spalte" Or CStr(wi.Cells(r, 1).Value) = "Zelle") Then
            If CStr(wi.Cells(r, 3).Value) = fid And CStr(wi.Cells(r, 4).Value) = oid Then
                If CStr(wi.Cells(r, 1).Value) = "Zelle" Then
                    Set TestCell = ws.Range(CStr(wi.Cells(r, 5).Value))
                Else
                    Set TestCell = ws.Cells(r0 + i - 1, CLng(Val(wi.Cells(r, 5).Value)))
                End If
                Exit Function
            End If
        End If
    Next
End Function

Private Sub TestUnprotect(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    wb.Unprotect ""
    For Each ws In wb.Worksheets
        ws.Unprotect ""
    Next
End Sub

' Erzeugt eine Erfassungsdatei ohne Dialog (Beispieldateien, automatisierte Abläufe); liefert den Pfad
Public Function BuildFileFor(ByVal cid As String, ByVal path As String, Optional ByVal recipient As String = "") As String
    BuildFile cid, path, recipient
    BuildFileFor = path
End Function

Public Function SelfTestRoundTrip(ByVal folder As String) As String
    Dim tA As String, tP As String, cid As String, m As cMap, sel As String, nm As Variant, path As String, wb As Workbook
    Dim errs As Collection, res As Variant, before As Long, s As String, absFid As String, pname As String, i As Long, props As New cMap
    Dim fid As String, plan As Collection, pathF As String, fZ As String, fS As String, fP As String, oids As Collection, rid1 As String
    Dim ok As Boolean, cnt As Long, fR As String, rng As Range, rec As Variant
    tA = mCfg.TableIdByName("Aufgaben")
    tP = mCfg.TableIdByName("Projekte")
    EnsureSheets
    absFid = EnsureSenderField(tA)
    fZ = mCfg.FieldIdByName(tA, "Zuständig")
    fS = mCfg.FieldIdByName(tA, "Status")
    fP = mCfg.FieldIdByName(tA, "Projekt")
    For Each nm In Array("Aufgabe", "Status", "Priorität", "Frist", "Aufwand (h)", "Wichtig", "Zuständig", "Projekt")
        sel = ListAdd(sel, mCfg.FieldIdByName(tA, CStr(nm)))
    Next
    Set m = New cMap
    cid = mCfg.NextId("C", 4)
    m.SetVal "ID", cid: m.SetVal "Name", "Selbsttest Liste": m.SetVal "Tabelle", tA: m.SetVal "Felder", sel
    m.SetVal "Modus", "Liste": m.SetVal "Titel", "Meldung offener Aufgaben": m.SetVal "Einleitung", "Test"
    m.SetVal "Zeilen", "50": m.SetVal "Absender", "Ja": m.SetVal "AbsenderFeld", absFid: m.SetVal "Schutz", "Ja": m.SetVal "Version", "1"
    m.SetVal "Optionen", "abslabel=Meldende Stelle;absbreite=Sehr breit"
    mCfg.AddRow SH_CAP, m
    path = folder & "\Rücklauf_Liste_Testamt.xlsx"
    BuildFile cid, path, "Testamt"
    s = s & "C01" & vbTab & IIf(FileExists(path), "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Erfassungsdatei (Liste, XLSX ohne Makros) erzeugt: " & Format$(FileLen(path) / 1024, "0") & " KB" & vbLf
    pname = mData.RecordLabel(tP, mData.TData(tP).IdAt(1))
    Set wb = Application.Workbooks.Open(path)
    TestUnprotect wb
    Set oids = mCfg.OptionIds(fZ)
    Set rng = TestCell(wb, 1, 1, fZ, CStr(oids(1)))
    ok = Not rng Is Nothing
    If ok Then ok = (rng.Validation.Type = xlValidateList And rng.Worksheet.Cells(R_HEAD, rng.Column).Orientation = xlUpward And _
                     TestCell(wb, 1, 1, fS).Validation.Type = xlValidateList And TestCell(wb, 1, 1, mCfg.FieldIdByName(tA, "Frist")).Validation.Type = xlValidateDate And _
                     wb.FileFormat = xlOpenXMLWorkbook And InStr(rng.Worksheet.Range("B7").Value, "Meldende Stelle") > 0 And rng.Worksheet.Range("C7").MergeArea.Columns.Count > 1)
    s = s & "C02" & vbTab & IIf(ok, "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Auswahlliste, Datumsprüfung, Ankreuzspalten mit senkrechtem Titel für „Zuständig“ (" & oids.Count & " Optionen), Absenderfeld beschriftet und breit" & vbLf
    TestCell(wb, 1, 1, mCfg.FieldIdByName(tA, "Aufgabe")).Value = "Rücklauf A": TestCell(wb, 1, 1, fS).Value = "Offen"
    TestCell(wb, 1, 1, mCfg.FieldIdByName(tA, "Frist")).Value = DateSerial(2026, 11, 20)
    TestCell(wb, 1, 1, mCfg.FieldIdByName(tA, "Aufwand (h)")).Value = 3.5: TestCell(wb, 1, 1, mCfg.FieldIdByName(tA, "Wichtig")).Value = "Ja"
    TestCell(wb, 1, 1, fZ, CStr(oids(1))).Value = "Ja": TestCell(wb, 1, 1, fZ, CStr(oids(2))).Value = "Ja": TestCell(wb, 1, 1, fP).Value = pname
    TestCell(wb, 1, 2, mCfg.FieldIdByName(tA, "Aufgabe")).Value = "Rücklauf B": TestCell(wb, 1, 2, fS).Value = "In Arbeit"
    TestCell(wb, 1, 2, mCfg.FieldIdByName(tA, "Priorität")).Value = "Hoch"
    TestCell(wb, 1, 3, mCfg.FieldIdByName(tA, "Aufgabe")).Value = "Rücklauf C": TestCell(wb, 1, 3, fS).Value = "Unsinn"
    wb.Save
    wb.Close False
    ThisWorkbook.Activate
    Set errs = New Collection
    res = ReadReturn(path, errs)
    s = s & "C03" & vbTab & IIf(res(3) = 1 And res(2).Count = 0 And errs.Count >= 1, "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Ungültiger Wert („Unsinn“ als Status) wird erkannt; Datei mit Fehlern wird nicht teilweise übernommen: " & res(4) & vbLf
    Set wb = Application.Workbooks.Open(path)
    TestUnprotect wb
    TestCell(wb, 1, 3, fS).Value = "Wartet"
    wb.Save
    wb.Close False
    ThisWorkbook.Activate
    Set errs = New Collection
    res = ReadReturn(path, errs)
    before = mData.TData(tA).RowCount
    Set plan = New Collection
    plan.Add res
    res = ExecuteImport(plan)
    s = s & "C04" & vbTab & IIf(res(0) = 3 And mData.TData(tA).RowCount = before + 3, "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Korrigierte Datei: 3 Einträge übernommen (eine Transaktion)" & vbLf
    Dim lastRid As String, firstRid As String
    firstRid = mData.TData(tA).IdAt(before + 1)
    lastRid = mData.TData(tA).IdAt(before + 3)
    s = s & "C05" & vbTab & IIf(mData.GetValue(tA, firstRid, absFid) = "Testamt" And Len(mData.LinkTargets(fP, firstRid)) > 0 And _
                                  ListCount(NzS(mData.GetValue(tA, firstRid, fZ))) = 2 And _
                                  Abs(CDbl(mData.GetValue(tA, firstRid, mCfg.FieldIdByName(tA, "Aufwand (h)"))) - 3.5) < 0.001, "BESTANDEN", "NICHT BESTANDEN") & _
            vbTab & "Werte korrekt übernommen: Absender, Verknüpfung über Bezeichnung, zwei Ankreuzspalten ergeben Mehrfachauswahl, Zahl; ID " & NzS(mData.GetValue(tA, lastRid, mCfg.FieldIdByName(tA, "Nr."))) & vbLf
    Set errs = New Collection
    res = ReadReturn(path, errs)
    s = s & "C06" & vbTab & IIf(res(2).Count = 0 And InStr(res(4), "bereits") > 0, "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Dieselbe Datei ein zweites Mal: " & res(4) & vbLf
    ' Ab 1.6: Dieselbe Datei geht an viele; eine zweite Stelle mit derselben Datei wird angenommen
    Dim path2 As String, res2 As Variant, restoredOk As Boolean
    path2 = folder & "\Rücklauf_Liste_Zweite_Stelle.xlsx"
    Set wb = Application.Workbooks.Open(path)
    TestUnprotect wb
    wb.Worksheets(1).Range(InfoValue(wb, "AbsenderZelle", 2)).Value = "Jugendamt"
    wb.SaveAs path2
    wb.Close False
    ThisWorkbook.Activate
    Set errs = New Collection
    res2 = ReadReturn(path2, errs)
    s = s & "C16" & vbTab & IIf(res2(2).Count = 3 And InStr(res2(4), "bereit") = 1, "BESTANDEN", "NICHT BESTANDEN") & vbTab & _
            "Dieselbe Erfassungsdatei von einer zweiten Stelle („Jugendamt“) wird angenommen: " & res2(2).Count & " Einträge, " & res2(4) & vbLf
    ' Hauptdatei nach dem Erzeugen nicht gespeichert: Erfassung fehlt -> aus der Rücklaufdatei wiederherstellen
    mCfg.DeleteRow SH_CAP, cid
    Set mRestoreAsked = Nothing
    Set errs = New Collection
    Dim wasQuiet As Boolean
    wasQuiet = mUI.gQuiet
    mUI.gQuiet = True                ' Rückfrage "Wiederherstellen" gilt als bestätigt
    res2 = ReadReturn(path2, errs)
    mUI.gQuiet = wasQuiet
    restoredOk = (Len(mCfg.GetS(SH_CAP, cid, "ID")) > 0 And mCfg.GetS(SH_CAP, cid, "Tabelle") = tA And res2(2).Count = 3)
    s = s & "C17" & vbTab & IIf(restoredOk, "BESTANDEN", "NICHT BESTANDEN") & vbTab & _
            "Erfassung fehlte (nicht gespeichert) und wurde aus der Rücklaufdatei wiederhergestellt; Rücklauf lesbar: " & res2(4) & vbLf
    ' Fehlanzeige: Teilnehmende ohne Datei („nichts zu melden“) zählen als geantwortet, zurücknehmbar
    Dim nA0 As Long, nA1 As Long, nA2 As Long, nNd As Long
    SetSenderList cid, "Testamt" & vbLf & "Schule am Park" & vbLf & "Kita Sonne"
    nA0 = AnsweredCount(cid)
    SetNoData cid, "Schule am Park", True
    nA1 = AnsweredCount(cid, nNd)
    SetNoData cid, "Schule am Park", False
    nA2 = AnsweredCount(cid)
    s = s & "C18" & vbTab & IIf(nA1 = nA0 + 1 And nNd = 1 And nA2 = nA0, "BESTANDEN", "NICHT BESTANDEN") & vbTab & _
            "Fehlanzeige „Schule am Park“: geantwortet " & nA0 & " -> " & nA1 & " (davon " & nNd & " Fehlanzeige), zurückgenommen -> " & nA2 & vbLf
    ' Formular-Variante (Mehrfachauswahl als Ja-Zeilen)
    m.SetVal "ID", mCfg.NextId("C", 4): m.SetVal "Name", "Selbsttest Formular": m.SetVal "Modus", "Formular": m.SetVal "Absender", "Nein": m.SetVal "AbsenderFeld", ""
    m.SetVal "Optionen", ""
    mCfg.AddRow SH_CAP, m
    pathF = folder & "\Rücklauf_Formular.xlsx"
    BuildFile CStr(m.GetVal("ID")), pathF, ""
    Set wb = Application.Workbooks.Open(pathF)
    TestUnprotect wb
    TestCell(wb, 1, 1, mCfg.FieldIdByName(tA, "Aufgabe")).Value = "Umfrageantwort"
    TestCell(wb, 1, 1, fS).Value = "Offen"
    TestCell(wb, 1, 1, fZ, CStr(oids(3))).Value = "Ja"
    wb.Save
    wb.Close False
    ThisWorkbook.Activate
    Set errs = New Collection
    res = ReadReturn(pathF, errs)
    Set plan = New Collection
    plan.Add res
    before = mData.TData(tA).RowCount
    res = ExecuteImport(plan)
    s = s & "C07" & vbTab & IIf(res(0) = 1 And mData.TData(tA).RowCount = before + 1 And _
                                  NzS(mData.GetValue(tA, mData.TData(tA).IdAt(before + 1), fZ)) = CStr(oids(3)), "BESTANDEN", "NICHT BESTANDEN") & _
            vbTab & "Formular-Datei mit Ja-Zeilen für die Mehrfachauswahl erzeugt, ausgefüllt und übernommen" & vbLf
    ' Datei aus einer anderen Datenbankdatei
    Dim wi As Worksheet, rr As Long
    Set wb = Application.Workbooks.Open(path)
    TestUnprotect wb
    Set wi = wb.Worksheets(INFO_SHEET)
    For rr = 1 To 60
        If CStr(wi.Cells(rr, 1).Value) = "DatenbankID" Then wi.Cells(rr, 2).Value = "DBfremd"
        If CStr(wi.Cells(rr, 1).Value) = "DateiID" Then wi.Cells(rr, 2).Value = NewFileId()
    Next
    wb.Save
    wb.Close False
    ThisWorkbook.Activate
    Set errs = New Collection
    res = ReadReturn(path, errs)
    s = s & "C08" & vbTab & IIf(res(2).Count = 0 And InStr(res(4), "anderen Datenbankdatei") > 0, "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Datei aus einer anderen Datenbankdatei wird erkannt: " & res(4) & vbLf
    ' Pruefdatei je Projekt: vorbelegt, nur eigene Eintraege, Aktualisierung ueber ID, Konflikt
    Set m = New cMap
    cid = mCfg.NextId("C", 4)
    m.SetVal "ID", cid: m.SetVal "Name", "Selbsttest Prüfen": m.SetVal "Tabelle", tA
    m.SetVal "Felder", ListAdd(ListAdd(fS, fZ), mCfg.FieldIdByName(tA, "Wichtig"))
    m.SetVal "Modus", "Liste": m.SetVal "Titel", "Bitte Aufgaben prüfen": m.SetVal "Einleitung", "Test"
    m.SetVal "Zeilen", "0": m.SetVal "Absender", "Ja": m.SetVal "AbsenderFeld", absFid: m.SetVal "Schutz", "Ja": m.SetVal "Version", "1"
    m.SetVal "Optionen", KVSet(KVSet(KVSet(KVSet("", "art", "pruefen"), "anzeige", ListAdd(mCfg.FieldIdByName(tA, "Aufgabe"), mCfg.FieldIdByName(tA, "Frist"))), "teilen", fP), "neuzeilen", "5")
    mCfg.AddRow SH_CAP, m
    cnt = SourceRids(cid, pname).Count
    pathF = folder & "\Rücklauf_Prüfen_Projekt.xlsx"
    BuildFile cid, pathF, pname
    Set wb = Application.Workbooks.Open(pathF)
    TestUnprotect wb
    rid1 = CStr(wb.Worksheets(1).Cells(R_HEAD + 1, CLng(Val(InfoValue(wb, "Spalten", 2)))).Value)
    ok = (cnt > 0 And Len(rid1) > 0 And mData.RecordExists(tA, rid1))
    If ok Then ok = (SplitLabel(mCfg.Field(fP), rid1) = pname And CStr(TestCell(wb, 1, 1, mCfg.FieldIdByName(tA, "Aufgabe")).Value) = mData.RecordLabel(tA, rid1) And _
                     TestCell(wb, 1, 1, mCfg.FieldIdByName(tA, "Aufgabe")).Locked And Not TestCell(wb, 1, 1, fS).Locked And _
                     Len(CStr(TestCell(wb, 1, cnt + 1, mCfg.FieldIdByName(tA, "Aufgabe")).Value)) = 0)
    s = s & "C09" & vbTab & IIf(ok, "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Prüfdatei für „" & pname & "“: " & cnt & " eigene Einträge vorbelegt (graue Spalten gesperrt, gelbe offen, versteckte ID), 5 leere Zeilen für Neues" & vbLf
    TestCell(wb, 1, 1, fS).Value = "Erledigt"
    For i = 1 To oids.Count
        TestCell(wb, 1, 1, fZ, CStr(oids(i))).Value = Empty
    Next
    TestCell(wb, 1, 1, fZ, CStr(oids(2))).Value = "Ja"
    ' neue Zeile in der Pruefdatei (graue Spalten sind dort ausfuellbar)
    TestCell(wb, 1, cnt + 1, mCfg.FieldIdByName(tA, "Aufgabe")).Value = "Neu aus Prüfdatei"
    TestCell(wb, 1, cnt + 1, mCfg.FieldIdByName(tA, "Frist")).Value = DateSerial(2026, 12, 1)
    TestCell(wb, 1, cnt + 1, fS).Value = "Offen"
    wb.Save
    wb.Close False
    ThisWorkbook.Activate
    Set errs = New Collection
    res = ReadReturn(pathF, errs)
    before = mData.TData(tA).RowCount
    Set plan = New Collection
    plan.Add res
    rec = ExecuteImport(plan)
    ok = (res(11) = 1 And res(12) = 1 And res(10) = cnt - 1 And errs.Count = 0 And rec(0) = 1 And rec(1) = 1)
    s = s & "C10" & vbTab & IIf(ok, "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Rücklauf der Prüfdatei: " & res(11) & " geändert, " & res(10) & " unverändert, " & res(12) & " neu, " & errs.Count & " Fehler: " & res(4) & vbLf
    fR = mData.TData(tA).IdAt(mData.TData(tA).RowCount)
    s = s & "C11" & vbTab & IIf(mData.GetValue(tA, rid1, fS) = mCfg.OptIdByName(fS, "Erledigt") And NzS(mData.GetValue(tA, rid1, fZ)) = CStr(oids(2)) And _
                                  mData.GetValue(tA, rid1, absFid) = pname And mData.TData(tA).RowCount = before + 1 And _
                                  mFmt.LinkLabels(mCfg.Field(fP), mData.LinkTargets(fP, fR), 9) = pname, "BESTANDEN", "NICHT BESTANDEN") & _
            vbTab & "Eintrag über die ID aktualisiert (Status, Ankreuzspalten, Absender), keine Dublette; neue Zeile automatisch dem Empfänger („" & pname & "“) zugeordnet" & vbLf
    ' zweites Einlesen derselben Pruefdatei wird erkannt; Ruecklauf-Uebersicht kennt die Datei
    Set errs = New Collection
    res = ReadReturn(pathF, errs)
    s = s & "C12" & vbTab & IIf(InStr(res(4), "bereits") > 0, "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Prüfdatei ein zweites Mal: " & res(4) & vbLf
    ' zweites Blatt (Tabelle Projekte) in derselben Datei, dem Empfaenger ueber den Projektnamen zugeordnet
    Dim child As String, fV As String, pRid As String
    fV = mCfg.FieldIdByName(tP, "Verantwortlich")
    Set m = New cMap
    child = mCfg.NextId("C", 4)
    m.SetVal "ID", child: m.SetVal "Name", "Selbsttest Prüfen / Projekt": m.SetVal "Tabelle", tP: m.SetVal "Felder", fV
    m.SetVal "Modus", "Liste": m.SetVal "Titel", "Projektangaben": m.SetVal "Zeilen", "0": m.SetVal "Absender", "Ja"
    m.SetVal "AbsenderFeld", EnsureSenderField(tP): m.SetVal "Schutz", "Ja": m.SetVal "Version", "1"
    m.SetVal "Optionen", KVSet(KVSet(KVSet(KVSet("", "art", "pruefen"), "anzeige", mCfg.FieldIdByName(tP, "Projekt")), "teilen", mCfg.FieldIdByName(tP, "Projekt")), "eltern", cid)
    mCfg.AddRow SH_CAP, m
    pathF = folder & "\Rücklauf_Prüfen_zwei_Blätter.xlsx"
    BuildFile cid, pathF, pname
    Set wb = Application.Workbooks.Open(pathF)
    TestUnprotect wb
    ok = (wb.Worksheets.Count = 4)
    Set rng = TestCell(wb, 2, 1, fV)
    If Not rng Is Nothing Then rng.Value = "Leitung aus Selbsttest" Else ok = False
    wb.Save
    wb.Close False
    ThisWorkbook.Activate
    Set errs = New Collection
    res = ReadReturn(pathF, errs)
    Set plan = New Collection
    plan.Add res
    rec = ExecuteImport(plan)
    pRid = mCfg.TableIdByName("Projekte")
    For i = 1 To mData.TData(tP).RowCount
        If mData.RecordLabel(tP, mData.TData(tP).IdAt(i)) = pname Then pRid = mData.TData(tP).IdAt(i)
    Next
    ok = ok And rec(1) = 1 And errs.Count = 0 And NzS(mData.GetValue(tP, pRid, fV)) = "Leitung aus Selbsttest"
    s = s & "C13" & vbTab & IIf(ok, "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Eine Datei mit zwei Blättern (Aufgaben und Projekte) erzeugt; Projektangabe über das zweite Blatt aktualisiert: " & res(4) & vbLf
    ' Prüfdatei (ab 1.5): Mehrfachauswahl nur zur Anzeige -> Ankreuzspalten, Trennzeile, Empfänger vorbelegt
    Dim c14 As String, sepOk As Boolean, preOk As Boolean, lockOk As Boolean, wsT As Worksheet, nr As Long
    Set m = New cMap
    c14 = mCfg.NextId("C", 4)
    m.SetVal "ID", c14: m.SetVal "Name", "Selbsttest Ankreuzen": m.SetVal "Tabelle", tA: m.SetVal "Felder", fS
    m.SetVal "Modus", "Liste": m.SetVal "Titel", "Aufgaben prüfen": m.SetVal "Zeilen", "0": m.SetVal "Absender", "Nein"
    m.SetVal "AbsenderFeld", "": m.SetVal "Schutz", "Ja": m.SetVal "Version", "1"
    m.SetVal "Optionen", KVSet(KVSet(KVSet(KVSet("", "art", "pruefen"), "anzeige", ListAdd(ListAdd(mCfg.FieldIdByName(tA, "Aufgabe"), fZ), fP)), "teilen", fP), "neuzeilen", "3")
    mCfg.AddRow SH_CAP, m
    cnt = SourceRids(c14, pname).Count
    pathF = folder & "\Rücklauf_Prüfen_Ankreuzen.xlsx"
    BuildFile c14, pathF, pname
    Set wb = Application.Workbooks.Open(pathF)
    TestUnprotect wb
    Set wsT = wb.Worksheets(1)
    nr = R_HEAD + 1 + cnt                      ' Trennzeile direkt nach den vorhandenen Einträgen
    sepOk = (CStr(wsT.Cells(nr, 1).Value) = SEP_MARK And CStr(wsT.Cells(nr + 1, 1).Value) = SEP_MARK And _
             InStr(CStr(wsT.Cells(nr + 1, 2).Value), "Neue Einträge") = 1 And InStr(CStr(wsT.Cells(R_BAND, 2).Value), "Bestehende Einträge") = 1)
    preOk = (CStr(TestCell(wb, 1, cnt + 1, fP).Value) = pname)
    lockOk = TestCell(wb, 1, 1, fZ, CStr(oids(1))).Locked And Not TestCell(wb, 1, cnt + 1, fZ, CStr(oids(1))).Locked
    TestCell(wb, 1, cnt + 1, mCfg.FieldIdByName(tA, "Aufgabe")).Value = "Neu mit Ankreuzen"
    TestCell(wb, 1, cnt + 1, fS).Value = "Offen"
    TestCell(wb, 1, cnt + 1, fZ, CStr(oids(1))).Value = "Ja"
    wb.Save
    wb.Close False
    ThisWorkbook.Activate
    Set errs = New Collection
    res = ReadReturn(pathF, errs)
    before = mData.TData(tA).RowCount
    Set plan = New Collection
    plan.Add res
    rec = ExecuteImport(plan)
    fR = mData.TData(tA).IdAt(mData.TData(tA).RowCount)
    ok = sepOk And preOk And lockOk And errs.Count = 0 And mData.TData(tA).RowCount = before + 1 And _
         NzS(mData.GetValue(tA, fR, fZ)) = CStr(oids(1)) And mFmt.LinkLabels(mCfg.Field(fP), mData.LinkTargets(fP, fR), 9) = pname
    s = s & "C14" & vbTab & IIf(ok, "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Prüfdatei mit Mehrfachauswahl zur Anzeige: Ankreuzspalten (vorhandene Zeilen gesperrt, neue ankreuzbar), Trennzeile „Neue Einträge“, " & _
            "Empfänger in neuen Zeilen vorbelegt; neue Zeile mit Haken übernommen (Trenner " & sepOk & ", Vorbelegung " & preOk & ", Sperren " & lockOk & ", Fehler " & errs.Count & ")" & vbLf
    ' Höchstzahl bei Mehrfachauswahl (ab 1.5): Regel in der Datei, Ablehnung beim Einlesen, Prüfung im Programm
    Dim vf As String, maxMsg As String, ruleOk As Boolean, rejOk As Boolean, apiOk As Boolean
    mCfg.SetV SH_FLD, fZ, "Maximum", "2"
    mCfg.InvalidateAll
    apiOk = Not mValid.CheckValue(mCfg.Field(fZ), CStr(oids(1)) & ";" & CStr(oids(2)) & ";" & CStr(oids(3)), maxMsg) And _
            mValid.CheckValue(mCfg.Field(fZ), CStr(oids(1)) & ";" & CStr(oids(2)), "")
    pathF = folder & "\Rücklauf_Prüfen_Hoechstzahl.xlsx"
    cnt = SourceRids(c14, pname).Count          ' C14 hat einen Eintrag ergänzt
    BuildFile c14, pathF, pname
    Set wb = Application.Workbooks.Open(pathF)
    TestUnprotect wb
    On Error Resume Next
    vf = TestCell(wb, 1, cnt + 1, fZ, CStr(oids(2))).Validation.Formula1
    On Error GoTo 0
    ruleOk = (InStr(vf, "L_LEER") > 0 And InStr(UCase$(wb.Worksheets(1).Cells(R_GROUP, TestCell(wb, 1, 1, fZ, CStr(oids(1))).Column).MergeArea.Cells(1, 1).Value), "HÖCHSTENS 2") > 0)
    TestCell(wb, 1, cnt + 1, mCfg.FieldIdByName(tA, "Aufgabe")).Value = "Zu viele Haken"
    TestCell(wb, 1, cnt + 1, fS).Value = "Offen"
    ' die Regel lässt zwei Haken zu und erkennt den dritten als ungültig
    TestCell(wb, 1, cnt + 1, fZ, CStr(oids(1))).Value = "Ja"
    TestCell(wb, 1, cnt + 1, fZ, CStr(oids(2))).Value = "Ja"
    On Error Resume Next
    ruleOk = ruleOk And TestCell(wb, 1, cnt + 1, fZ, CStr(oids(2))).Validation.Value = True
    TestCell(wb, 1, cnt + 1, fZ, CStr(oids(3))).Value = "Ja"
    ruleOk = ruleOk And TestCell(wb, 1, cnt + 1, fZ, CStr(oids(3))).Validation.Value = False
    On Error GoTo 0
    wb.Save
    wb.Close False
    ThisWorkbook.Activate
    Set errs = New Collection
    res = ReadReturn(pathF, errs)
    rejOk = (res(3) > 0 And res(2).Count = 0)
    mCfg.SetV SH_FLD, fZ, "Maximum", Empty
    mCfg.InvalidateAll
    s = s & "C15" & vbTab & IIf(apiOk And ruleOk And rejOk, "BESTANDEN", "NICHT BESTANDEN") & vbTab & "Mehrfachauswahl „höchstens 2“: Programm lehnt 3 ab (" & maxMsg & "), " & _
            "Erfassungsdatei hat Ankreuzregel und Hinweis im Kopf (" & ruleOk & IIf(Len(gCapDiag) > 0, "; " & gCapDiag, "") & "), Rücklauf mit 3 Haken wird nicht übernommen (" & rejOk & ")" & vbLf
    SelfTestRoundTrip = s
End Function

Private Function InfoValue(ByVal wb As Workbook, ByVal key As String, ByVal col As Long) As String
    Dim wi As Worksheet, r As Long
    Set wi = wb.Worksheets(INFO_SHEET)
    For r = 1 To 5000
        If Len(CStr(wi.Cells(r, 1).Value)) = 0 Then Exit For
        If CStr(wi.Cells(r, 1).Value) = key Then InfoValue = CStr(wi.Cells(r, col).Value): Exit Function
    Next
End Function
