Attribute VB_Name = "mPages"
'==============================================================================
' mPages - Startseite, Einrichten, Suchergebnisse, Hilfe & Sicherung,
' Technische Dokumentation.
'==============================================================================
Option Explicit

Private mRowMap As cMap          ' Seite|Zeile -> Ziel (Tabelle|Datensatz o. a.)
Private mSetupSel As String

Private Function RowMap() As cMap
    If mRowMap Is Nothing Then Set mRowMap = New cMap
    Set RowMap = mRowMap
End Function

Public Function RowTarget(ByVal sheetName As String, ByVal r As Long) As String
    RowTarget = CStr(RowMap.GetVal(sheetName & "|" & r, ""))
End Function

' Zeilenziele nur dieser Seite verwerfen (andere Seiten bleiben gueltig, auch wenn sie
' beim naechsten Besuch ohne Neuaufbau gezeigt werden)
Private Sub ResetRows(ByVal sheetName As String)
    Dim k As Variant
    For Each k In RowMap.Keys
        If Left$(CStr(k), Len(sheetName) + 1) = sheetName & "|" Then mRowMap.Remove CStr(k)
    Next
End Sub

'==============================================================================
' Start
'==============================================================================
Public Sub RenderStart()
    Dim ws As Worksheet, ids As Variant, i As Long, x As Single, y As Single, cw As Single, ch As Single, cols As Long, s As Shape, t As Shape
    Dim tid As String, n As Long, rr As Long, logRows As Variant, k As Long, sr As Range, x0 As Single, W As Single, lbl As String, cnt As Long
    Set ws = mRender.Sheet(SH_START)
    mRender.BeginPage ws
    ResetRows SH_START
    mRender.SetColWidthPt ws, 3, 130
    mRender.SetColWidthPt ws, 4, 160
    mRender.SetColWidthPt ws, 5, 260
    mRender.SetColWidthPt ws, 6, 160
    mRender.DrawTitle ws, ChrW(&HE80F), "Start", mCfg.SysGet("AppName", APP_TITLE) & IIf(mCfg.SysGet("Demo", "") = "Ja", "  ·  DEMODATEI mit fiktiven Beispieldaten", "")
    Set sr = mRender.SearchBox(ws, R_TABS, mNav.gGlobalSearch, "mCmd.Cmd_GlobalSearch", 280)
    mRender.FlowBegin ws, R_TABS, sr.Left + sr.Width + 6, mRender.PageRight()
    mRender.FlowBtn "b_gsearch", "In allen Tabellen suchen", "mCmd.Cmd_GlobalSearch"
    mRender.FlowEnd
    mRender.FlowBegin ws, R_TOOLS, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowBtn "b_newtab", "Neue Tabelle", "mCmd.Cmd_NewTable", "primary", ChrW(&HE710)
    mRender.FlowBtn "b_tpl", "Aus Vorlage", "mCmd.Cmd_FromTemplate", "secondary", ChrW(&HE8A5)
    mRender.FlowBtn "b_imp", "Daten importieren", "mCmd.Cmd_Import", "secondary", ChrW(&HE8B5)
    mRender.FlowBtn "b_list", "Liste übernehmen", "mCmd.Cmd_ListImport", "secondary", ChrW(&HE8FD)
    mRender.FlowBtn "b_dash", "Dashboards", "mCmd.Cmd_GoDash", "secondary", ChrW(&HE9D2)
    mRender.FlowBtn "b_capt", "Erfassungsdateien", "mCapture.Cmd_GoPage", "secondary", ChrW(&HE8A5)
    mRender.FlowBtn "b_help", "Hilfe", "mCmd.Cmd_GoHelp", "secondary", ChrW(&HE897)
    If Not mApp.gReadOnly Then
        If mCfg.SysGet("TitelFestgelegt", "") = "Ja" Then
            mRender.FlowBtn "b_title", "Titel bearbeiten", "mCmd.Cmd_EditTitle", "link", ChrW(&HE70F)
        Else
            mRender.FlowBtn "b_title", "Titel der Datei festlegen", "mCmd.Cmd_EditTitle", "chip", ChrW(&HE70F)
        End If
    End If
    mRender.FlowEnd
    ws.Rows(R_CHIPS).RowHeight = 8

    x0 = mRender.ContentLeft(ws)
    W = mRender.PageRight() - x0
    If W > 1200 Then W = 1200
    ids = mCfg.TableIds()
    y = ws.Rows(R_DATA).Top + 2
    mRender.TextShape ws, "h_tables", "Ihre Tabellen", x0, y, 300, 22, 12, CText(), True, 1
    y = y + 28
    If ArrLen(ids) = 0 Then
        mRender.TextShape ws, "h_empty", "Noch keine Tabellen. Legen Sie eine eigene Tabelle an oder starten Sie mit einer Vorlage (Aufgaben & Projekte, Kontakte, Ressourcen, Dokumentation). " & _
                          "Die Hilfe beschreibt den Weg „In zehn Schritten zur eigenen Datei“.", x0, y, W, 40, 10, CMuted(), False, 1
        y = y + 50
    Else
        cols = Int((W + 14) / 250)
        If cols < 1 Then cols = 1
        If cols > 4 Then cols = 4
        cw = (W - 14 * (cols - 1)) / cols
        ch = 74
        For i = 0 To ArrLen(ids) - 1
            tid = CStr(ids(i))
            x = x0 + (i Mod cols) * (cw + 14)
            If i > 0 And i Mod cols = 0 Then y = y + ch + 12
            Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, cw, ch)
            s.Name = "tcard_" & tid
            s.Adjustments.Item(1) = 0.08
            s.Fill.ForeColor.RGB = CWhite()
            s.Line.ForeColor.RGB = CBorder(): s.Line.Weight = 0.75
            s.Placement = xlFreeFloating
            With s.Shadow
                .Visible = msoTrue: .ForeColor.RGB = RGB(0, 0, 0): .Transparency = 0.92: .OffsetY = 1: .OffsetX = 0: .Blur = 4
            End With
            s.OnAction = mRender.MacroRef("mCmd.Cmd_OpenTableCard")
            Set t = mRender.TextShape(ws, "tci_" & tid, mCfg.TableIcon(tid), x + 14, y + 12, 26, 26, 16, CAccent(), False, 1)
            t.TextFrame2.TextRange.Font.Name = ICON_FONT
            t.OnAction = mRender.MacroRef("mCmd.Cmd_OpenTableCard")
            ' Name, Zaehler und Beschreibung in einer Textform (weniger Formen = schnellerer Aufbau)
            cnt = ActiveRecords(tid)
            lbl = mCfg.TableName(tid) & vbCr & Plural(cnt, "Eintrag", "Einträge") & "  ·  " & Plural(ArrLen(mCfg.ViewIds(tid)), "Ansicht", "Ansichten") & _
                  vbCr & Truncate(mCfg.GetS(SH_TAB, tid, "Beschreibung"), CLng(cw / 5.5))
            Set t = mRender.TextShape(ws, "tcn_" & tid, lbl, x + 46, y + 10, cw - 56, ch - 16, 8.5, CMuted(), False, 1)
            With t.TextFrame2
                .VerticalAnchor = msoAnchorTop
                .WordWrap = msoFalse
                With .TextRange.Paragraphs(1).Font
                    .Name = UI_FONT_SEMI: .Size = 11.5: .Fill.ForeColor.RGB = CText()
                End With
                .TextRange.Paragraphs(1).ParagraphFormat.SpaceAfter = 3
                .TextRange.Paragraphs(2).Font.Size = 9
            End With
            t.OnAction = mRender.MacroRef("mCmd.Cmd_OpenTableCard")
        Next
        y = y + ch + 24
    End If

    ' Zuletzt geaendert
    rr = R_DATA
    Do While ws.Rows(rr).Top < y
        rr = rr + 1
    Loop
    ws.Cells(rr, 3).Value = "Zuletzt geändert"
    ws.Cells(rr, 3).Font.Size = 12: ws.Cells(rr, 3).Font.Name = UI_FONT_SEMI
    ws.Rows(rr).RowHeight = 26
    rr = rr + 1
    ws.Range(ws.Cells(rr, 3), ws.Cells(rr, 6)).Value = Array("Zeitpunkt", "Tabelle", "Eintrag", "Aktion")
    With ws.Range(ws.Cells(rr, 3), ws.Cells(rr, 6))
        .Interior.Color = CSurface(): .Font.Color = CMuted(): .Font.Size = 9.5
    End With
    logRows = mLog.RecentRows(200)
    n = 0
    If Not IsEmpty(logRows) Then
        Dim recent(1 To 12, 1 To 4) As Variant
        For k = 1 To UBound(logRows, 1)
            tid = NzS(logRows(k, 4))
            If mCfg.TableExists(tid) And Left$(NzS(logRows(k, 5)), 1) = "R" Then
                If mData.RecordExists(tid, NzS(logRows(k, 5))) Then
                    n = n + 1
                    recent(n, 1) = FmtStamp(logRows(k, 1))
                    recent(n, 2) = mCfg.TableName(tid)
                    recent(n, 3) = mData.RecordLabel(tid, NzS(logRows(k, 5)))
                    recent(n, 4) = NzS(logRows(k, 3))
                    RowMap.SetVal SH_START & "|" & (rr + n), tid & "|" & NzS(logRows(k, 5))
                    If n >= 12 Then Exit For
                End If
            End If
        Next
        If n > 0 Then
            ws.Range(ws.Cells(rr + 1, 3), ws.Cells(rr + n, 6)).NumberFormat = "@"
            ws.Range(ws.Cells(rr + 1, 3), ws.Cells(rr + 12, 6)).Value = recent
        End If
    End If
    If n = 0 Then
        ws.Cells(rr + 1, 3).Value = "Noch keine Änderungen."
        ws.Cells(rr + 1, 3).Font.Color = CMuted()
        n = 1
    Else
        ws.Range(ws.Cells(rr + 1, 3), ws.Cells(rr + n, 6)).RowHeight = 21
        mRender.GridBorders ws.Range(ws.Cells(rr, 3), ws.Cells(rr + n, 6))
        ws.Cells(rr + n + 1, 3).Value = "Doppelklick öffnet den Eintrag."
        ws.Cells(rr + n + 1, 3).Font.Color = CMuted(): ws.Cells(rr + n + 1, 3).Font.Size = 8.5
    End If
    ' Dateiinfo
    ws.Cells(rr + n + 3, 3).Value = "Datei: " & ThisWorkbook.Name & "   ·   Programmversion " & PROG_VERSION & " (Format " & FORMAT_VERSION & ")   ·   vollständig offline, ohne Cloud und ohne Add-ins"
    ws.Cells(rr + n + 3, 3).Font.Color = CMuted(): ws.Cells(rr + n + 3, 3).Font.Size = 8.5
    mRender.EndPage ws
End Sub

Public Function ActiveRecords(ByVal tid As String) As Long
    Dim t As cTableData, r As Long
    Set t = mData.TData(tid)
    For r = 1 To t.RowCount
        If Not t.IsArchived(r) Then ActiveRecords = ActiveRecords + 1
    Next
End Function

'==============================================================================
' Suche ueber alle Tabellen
'==============================================================================
Public Sub RenderSearch()
    Dim ws As Worksheet, ids As Variant, i As Long, tid As String, rows() As Long, n As Long, k As Long, rr As Long, sf As Collection, f As cField
    Dim total As Long, sr As Range, q As String
    Set ws = mRender.Sheet(SH_SEARCH)
    mRender.BeginPage ws
    ResetRows SH_SEARCH
    q = mNav.gGlobalSearch
    mRender.SetColWidthPt ws, 3, 280
    mRender.SetColWidthPt ws, 4, 360
    mRender.SetColWidthPt ws, 5, 160
    mRender.DrawTitle ws, ChrW(&HE721), "Suche", "Treffer in allen Tabellen (archivierte ausgenommen)"
    Set sr = mRender.SearchBox(ws, R_TABS, q, "mCmd.Cmd_GlobalSearch", 280)
    mRender.FlowBegin ws, R_TABS, sr.Left + sr.Width + 6, mRender.PageRight()
    mRender.FlowBtn "b_gsearch", "Suchen", "mCmd.Cmd_GlobalSearch"
    mRender.FlowBtn "b_back", "Zur Startseite", "mCmd.Cmd_GoStart", "secondary", ChrW(&HE80F)
    mRender.FlowEnd
    rr = R_DATA
    ids = mCfg.TableIds()
    For i = 0 To ArrLen(ids) - 1
        tid = CStr(ids(i))
        Set sf = New Collection
        For Each f In mCfg.Fields(tid)
            If f.FType <> FT_FORMEL Then sf.Add f
        Next
        If Len(q) > 0 Then
            rows = mQuery.FilterRows(tid, New Collection, q, sf, mQuery.ARCH_ACTIVE, n)
        Else
            n = 0
        End If
        If n > 0 Then
            ws.Cells(rr, 3).Value = mCfg.TableName(tid) & "  (" & n & ")"
            ws.Cells(rr, 3).Font.Name = UI_FONT_SEMI: ws.Cells(rr, 3).Font.Size = 11
            ws.Range(ws.Cells(rr, 3), ws.Cells(rr, 5)).Interior.Color = CSurface2()
            rr = rr + 1
            For k = 1 To Application.Min(n, 50)
                ws.Cells(rr, 3).NumberFormat = "@"
                ws.Cells(rr, 3).Value = mData.RecordLabel(tid, mData.TData(tid).IdAt(rows(k)))
                ws.Cells(rr, 4).Value = mData.RecordHint(tid, mData.TData(tid).IdAt(rows(k)), 3)
                ws.Cells(rr, 4).Font.Color = CMuted()
                RowMap.SetVal SH_SEARCH & "|" & rr, tid & "|" & mData.TData(tid).IdAt(rows(k))
                rr = rr + 1
            Next
            If n > 50 Then
                ws.Cells(rr, 3).Value = (n - 50) & " weitere Treffer; in der Tabelle mit derselben Suche eingrenzen"
                ws.Cells(rr, 3).Font.Color = CMuted(): ws.Cells(rr, 3).Font.Size = 8.5
                rr = rr + 1
            End If
            rr = rr + 1
            total = total + n
        End If
    Next
    If total = 0 Then
        mRender.EmptyState ws, R_DATA, IIf(Len(q) = 0, "Suchbegriff eingeben", "Keine Treffer für „" & q & "“"), _
                           "Gesucht wird in allen Feldern aller Tabellen. Mehrere Wörter müssen alle vorkommen.", "", "", ChrW(&HE721)
    Else
        mRender.GridBorders ws.Range(ws.Cells(R_DATA, 3), ws.Cells(rr - 1, 5))
    End If
    mRender.EndPage ws, 0, 0, "C" & R_TABS
End Sub

'==============================================================================
' Einrichten
'==============================================================================
Public Sub RenderSetup()
    Dim ws As Worksheet, ids As Variant, i As Long, tid As String, rr As Long, f As cField, nF As Long, x As Single, did As Variant
    Dim setup As Boolean
    Set ws = mRender.Sheet(SH_SETUP)
    mRender.BeginPage ws
    ResetRows SH_SETUP
    setup = mApp.IsSetupMode
    mRender.SetColWidthPt ws, 3, 30
    mRender.SetColWidthPt ws, 4, 200
    mRender.SetColWidthPt ws, 5, 80
    mRender.SetColWidthPt ws, 6, 70
    mRender.SetColWidthPt ws, 7, 80
    mRender.SetColWidthPt ws, 8, 100
    mRender.SetColWidthPt ws, 9, 300
    mRender.DrawTitle ws, ChrW(&HE713), "Einrichten", "Tabellen, Farben und Einstellungen dieser Datei"
    mRender.FlowBegin ws, R_TABS, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowBtn "b_title", "Titel: " & Truncate(mCfg.SysGet("AppName", APP_TITLE), 40), "mCmd.Cmd_EditTitle", "secondary", ChrW(&HE70F)
    mRender.FlowBtn "b_newtab", "Neue Tabelle", "mCmd.Cmd_NewTable", "primary", ChrW(&HE710)
    mRender.FlowBtn "b_tpl", "Aus Vorlage", "mCmd.Cmd_FromTemplate", "secondary", ChrW(&HE8A5)
    mRender.FlowBtn "b_imp", "Importieren", "mCmd.Cmd_Import", "secondary", ChrW(&HE8B5)
    mRender.FlowBtn "b_list", "Liste übernehmen", "mCmd.Cmd_ListImport", "secondary", ChrW(&HE8FD)
    mRender.FlowBtn "b_impr", "Letzter Importbericht", "mCmd.Cmd_ImportReport", "secondary", ChrW(&HE9F9)
    mRender.FlowEnd
    mRender.FlowBegin ws, R_TOOLS, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowBtn "b_docstyle", "Dokumentstil: " & IIf(Len(mDocStyle.CurrentDesign()) = 0, "Standard", mDocStyle.CurrentDesign()) & _
                    " · " & mDocStyle.DefaultFont(), "mCmd.Cmd_DocStyle", "secondary", ChrW(&HE790)
    mRender.FlowBtn "b_settings", "Allgemeine Einstellungen", "mCmd.Cmd_Settings", "secondary", ChrW(&HE713)
    mRender.FlowBtn "b_protect", "Einrichtungsschutz", "mCmd.Cmd_SetupProtection", "secondary", ChrW(&HE72E)
    mRender.FlowBtn "b_data", IIf(DataSheetsVisible(), "Datenblätter ausblenden", "Datenzugang (Datenblätter anzeigen)"), "mCmd.Cmd_DataAccess", "secondary", ChrW(&HE80A)
    mRender.FlowBtn "b_recalc", "Alles neu berechnen", "mCmd.Cmd_RecalcAll", "secondary", ChrW(&HE72C)
    mRender.FlowBtn "b_check", "Integrität prüfen", "mCmd.Cmd_Integrity", "secondary", ChrW(&HE9D9)
    mRender.FlowEnd
    ' Farbfelder: Navigation und Akzent direkt anklickbar
    DrawSwatches ws, R_CHIPS
    ' Tabellenliste
    rr = R_DATA
    ws.Cells(rr, 3).Value = "Tabellen"
    ws.Cells(rr, 3).Font.Size = 12: ws.Cells(rr, 3).Font.Name = UI_FONT_SEMI
    ws.Rows(rr).RowHeight = 26
    rr = rr + 1
    mRender.FlowBegin ws, rr, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowText "sel_hint", "Zeile anklicken, dann:", CMuted(), 9, False, 8
    mRender.FlowBtn "t_edit", "Eigenschaften bearbeiten", "mCmd.Cmd_SetupEditTable", "secondary", ChrW(&HE70F)
    mRender.FlowBtn "t_dup", "Duplizieren…", "mCmd.Cmd_SetupDuplicate", "secondary", ChrW(&HE8C8)
    mRender.FlowBtn "t_hide", "Aus-/Einblenden", "mCmd.Cmd_SetupHide", "secondary", ChrW(&HE7B3)
    mRender.FlowBtn "t_up", "Nach oben", "mCmd.Cmd_SetupUp", "secondary", ChrW(&HE74A)
    mRender.FlowBtn "t_down", "Nach unten", "mCmd.Cmd_SetupDown", "secondary", ChrW(&HE74B)
    mRender.FlowBtn "t_del", "Tabelle löschen…", "mCmd.Cmd_SetupDelete", "danger", ChrW(&HE74D)
    mRender.FlowEnd
    rr = rr + 1
    ws.Range(ws.Cells(rr, 3), ws.Cells(rr, 9)).Value = Array("", "Tabelle", "Einträge", "Eigenschaften", "Ansichten", "Status", "Beschreibung")
    With ws.Range(ws.Cells(rr, 3), ws.Cells(rr, 9))
        .Interior.Color = CSurface(): .Font.Color = CMuted(): .Font.Size = 9.5
    End With
    ids = mCfg.TableIds(True)
    For i = 0 To ArrLen(ids) - 1
        tid = CStr(ids(i))
        rr = rr + 1
        ws.Cells(rr, 3).Value = mCfg.TableIcon(tid)
        ws.Cells(rr, 3).Font.Name = ICON_FONT: ws.Cells(rr, 3).Font.Color = CAccent(): ws.Cells(rr, 3).HorizontalAlignment = xlCenter
        ws.Cells(rr, 4).NumberFormat = "@"
        ws.Cells(rr, 4).Value = mCfg.TableName(tid)
        ws.Cells(rr, 4).Font.Name = UI_FONT_SEMI
        ws.Cells(rr, 5).Value = ActiveRecords(tid)
        ws.Cells(rr, 6).Value = mCfg.Fields(tid, False).Count
        ws.Cells(rr, 7).Value = ArrLen(mCfg.ViewIds(tid))
        ws.Cells(rr, 8).Value = IIf(mCfg.GetS(SH_TAB, tid, "Status") = "ausgeblendet", "ausgeblendet", "aktiv")
        ws.Cells(rr, 9).Value = mCfg.GetS(SH_TAB, tid, "Beschreibung")
        ws.Cells(rr, 10).Value = " "   ' lange Beschreibung endet am Tabellenrand (wie in Listen)
        ws.Rows(rr).RowHeight = 24
        RowMap.SetVal SH_SETUP & "|" & rr, "T|" & tid
        If tid = mSetupSel Then ws.Range(ws.Cells(rr, 3), ws.Cells(rr, 9)).Interior.Color = CAccentSoft()
    Next
    If ArrLen(ids) = 0 Then
        rr = rr + 1
        ws.Cells(rr, 4).Value = "Noch keine Tabellen. „Neue Tabelle“ oder „Aus Vorlage“ wählen."
        ws.Cells(rr, 4).Font.Color = CMuted()
    End If
    mRender.GridBorders ws.Range(ws.Cells(R_DATA + 2, 3), ws.Cells(rr, 9))
    ' Dashboards
    rr = rr + 2
    ws.Cells(rr, 3).Value = "Dashboards"
    ws.Cells(rr, 3).Font.Size = 12: ws.Cells(rr, 3).Font.Name = UI_FONT_SEMI
    ws.Rows(rr).RowHeight = 26
    For Each did In mCfg.DashIds()
        rr = rr + 1
        ws.Cells(rr, 4).Value = mCfg.GetS(SH_DSB, CStr(did), "Name")
        ws.Cells(rr, 5).Value = ArrLen(mCfg.ElementIds(CStr(did)))
        ws.Cells(rr, 6).Value = "Elemente"
        ws.Cells(rr, 9).Value = mCfg.GetS(SH_DSB, CStr(did), "Beschreibung")
        RowMap.SetVal SH_SETUP & "|" & rr, "D|" & did
    Next
    rr = rr + 2
    ws.Cells(rr, 3).Value = "Doppelklick auf eine Tabelle öffnet die Eigenschaften (Spalten), auf ein Dashboard das Dashboard. Strukturänderungen erzeugen vorher automatisch eine Sicherung, " & _
                            "wenn Daten betroffen sind (Entfernen, Typwechsel, Import)."
    ws.Cells(rr, 3).Font.Color = CMuted(): ws.Cells(rr, 3).Font.Size = 8.5
    mRender.EndPage ws
End Sub

' Zwei Zeilen: Navigation und Akzent, je kraeftige und helle (Pastell-)Toene
Private Sub DrawSwatches(ByVal ws As Worksheet, ByVal row As Long)
    Dim y As Single
    ws.Rows(row).RowHeight = 62
    y = ws.Rows(row).Top + 4
    SwatchLine ws, "nav", "Navigation", mCmd.NavColorNames(), CNavBg(), y
    SwatchLine ws, "acc", "Akzent", mCmd.AccentColorNames(), CAccentRaw(), y + 28
End Sub

Private Sub SwatchLine(ByVal ws As Worksheet, ByVal kind As String, ByVal caption As String, ByVal names As Variant, ByVal cur As Long, ByVal y As Single)
    Dim x As Single, nm As Variant, s As Shape, c As Long, light As Boolean, prevLight As Boolean
    x = mRender.ContentLeft(ws)
    mRender.TextShape ws, "sw_l_" & kind, caption, x, y, 80, 22, 9, CMuted(), False, 1
    x = x + 78
    For Each nm In names
        c = IIf(kind = "nav", mCmd.NavValue(CStr(nm)), mCmd.AccentValue(CStr(nm)))
        light = IsLightColor(c)
        If light And Not prevLight Then
            ' Trennung zu den hellen Toenen
            x = x + 8
            mRender.TextShape ws, "sw_h_" & kind, "hell", x, y, 28, 22, 8, CMuted(), False, 1
            x = x + 28
        End If
        prevLight = light
        Set s = ws.Shapes.AddShape(msoShapeOval, x, y + 1, 20, 20)
        s.Name = "sw_" & kind & "_" & nm
        s.Fill.ForeColor.RGB = c
        If c = cur Then
            s.Line.ForeColor.RGB = CText(): s.Line.Weight = 2.25
        ElseIf light Then
            s.Line.ForeColor.RGB = CBorderStrong(): s.Line.Weight = 0.75
        Else
            s.Line.Visible = msoFalse
        End If
        s.OnAction = mRender.MacroRef("mCmd.Cmd_Swatch")
        s.Placement = xlFreeFloating
        s.AlternativeText = CStr(nm)
        x = x + 26
    Next
End Sub

' Nur die Farbfelder neu zeichnen (Navigationsfarbe geaendert; Rest der Seite bleibt)
Public Sub RefreshSwatches()
    Dim ws As Worksheet, i As Long
    If mNav.gPage <> "Einrichten" Then Exit Sub
    Set ws = mRender.Sheet(SH_SETUP)
    mApp.BeginWork
    ws.Unprotect Password:=""
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "sw_" Then ws.Shapes(i).Delete
    Next
    DrawSwatches ws, R_CHIPS
    mProtect.ProtectUi ws
    mApp.EndWork
    mApp.SyncSavedFlag
End Sub

Public Function DataSheetsVisible() As Boolean
    On Error Resume Next
    DataSheetsVisible = (ThisWorkbook.Worksheets(SH_TAB).Visible = xlSheetVisible)
End Function

Public Function SetupSelected() As String
    SetupSelected = mSetupSel
End Function

Public Sub OnSetupSelect(ByVal target As Range)
    Dim tg As String
    tg = RowTarget(SH_SETUP, target.Row)
    If Left$(tg, 2) = "T|" Then
        If mSetupSel <> Mid$(tg, 3) Then
            mSetupSel = Mid$(tg, 3)
            HighlightSetup
        End If
    End If
End Sub

Private Sub HighlightSetup()
    Dim ws As Worksheet, k As Variant, p As Variant, r As Long, prevEv As Boolean
    Set ws = mRender.Sheet(SH_SETUP)
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    mApp.BeginWork
    For Each k In RowMap.Keys
        p = Split(CStr(k), "|")
        If p(0) = SH_SETUP Then
            r = CLng(p(1))
            If Left$(RowMap.GetVal(CStr(k)), 2) = "T|" Then
                ws.Range(ws.Cells(r, 3), ws.Cells(r, 9)).Interior.Color = IIf(RowMap.GetVal(CStr(k)) = "T|" & mSetupSel, CAccentSoft(), CWhite())
            End If
        End If
    Next
    mApp.EndWork
    Application.EnableEvents = prevEv
End Sub

'==============================================================================
' Hilfe & Sicherung (Inhalt wird beim Build geschrieben, Kopf/Buttons dynamisch)
'==============================================================================
Public Sub RenderHelp()
    Dim ws As Worksheet
    Set ws = mRender.Sheet(SH_HELP)
    mRender.BeginPage ws, True
    If NzS(ws.Cells(1, 1).Value) <> "HILFE " & PROG_VERSION Then WriteHelp ws
    FitDocWidth ws
    mLogo.PlaceSheetLogo ws, 0, ws.Rows(R_TITLE).Top, 26
    mRender.DrawTitle ws, ChrW(&HE897), "Hilfe & Sicherung", "Bedienung, Sicherung und Wiederherstellung, vollständig offline"
    mRender.FlowBegin ws, R_TABS, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowBtn "b_backup", "Jetzt sichern", "mCmd.Cmd_BackupNow", "primary", ChrW(&HE74E)
    mRender.FlowBtn "b_restore", "Sicherungen & Wiederherstellen", "mCmd.Cmd_Restore", "secondary", ChrW(&HE777)
    mRender.FlowBtn "b_folder", "Sicherungsordner öffnen", "mCmd.Cmd_OpenBackupFolder", "secondary", ChrW(&HE838)
    mRender.FlowBtn "b_bset", "Sicherungseinstellungen", "mCmd.Cmd_BackupSettings", "secondary", ChrW(&HE713)
    mRender.FlowEnd
    mRender.FlowBegin ws, R_TOOLS, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowBtn "b_log", "Änderungsprotokoll", "mCmd.Cmd_ShowLog", "secondary", ChrW(&HE81C)
    mRender.FlowBtn "b_copy", "Neue leere Kopie erstellen", "mCmd.Cmd_NewEmptyCopy", "secondary", ChrW(&HE8C8)
    mRender.FlowBtn "b_reset", "Datei zurücksetzen…", "mCmd.Cmd_ResetData", "danger", ChrW(&HE74D)
    mRender.FlowEnd
    mRender.EndPage ws
End Sub

Public Sub RenderTech()
    Dim ws As Worksheet
    Set ws = mRender.Sheet(SH_TECH)
    mRender.BeginPage ws, True
    If NzS(ws.Cells(1, 1).Value) <> "TECHNIK " & PROG_VERSION Then WriteTech ws
    FitDocWidth ws
    mLogo.PlaceSheetLogo ws, 0, ws.Rows(R_TITLE).Top, 26
    mRender.DrawTitle ws, ChrW(&HE943), "Technische Dokumentation", APP_TITLE & " " & PROG_VERSION & " · " & AUTHOR_NOTE
    mRender.FlowBegin ws, R_TABS, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowBtn "b_code", "Quellcode anzeigen", "mCmd.Cmd_ShowCode", "primary", ChrW(&HE943)
    mRender.FlowBtn "b_zip", "Export ZIP", "mCmd.Cmd_CodeZip", "secondary", ChrW(&HEDE1)
    mRender.FlowBtn "b_html", "Export HTML", "mCmd.Cmd_CodeHtml", "secondary", ChrW(&HEDE1)
    mRender.FlowBtn "b_md", "Export Markdown", "mCmd.Cmd_CodeMd", "secondary", ChrW(&HEDE1)
    mRender.FlowBtn "b_txt", "Export Text", "mCmd.Cmd_CodeTxt", "secondary", ChrW(&HEDE1)
    mRender.FlowEnd
    If ws.Shapes.Count > 0 Then ws.Rows(R_TOOLS).RowHeight = 4      ' keine zweite Knopfreihe
    mRender.EndPage ws
End Sub

' Schreibt formatierten Text ab Zeile R_DATA in Spalte C; Zeilen: "H1|..", "H2|..", "P|..", "B|..", "N|.."
Public Sub WriteDoc(ByVal ws As Worksheet, ByVal lines As Variant, ByVal marker As String)
    Dim i As Long, rr As Long, kind As String, txt As String, p As Long, c As Range, W As Double
    ws.Unprotect ""
    ws.Range(ws.Cells(R_DATA, 1), ws.Cells(ws.Rows.Count, 30)).Clear
    W = 760
    mRender.SetColWidthPt ws, 3, W
    rr = R_DATA
    For i = LBound(lines) To UBound(lines)
        p = InStr(lines(i), "|")
        kind = Left$(lines(i), p - 1)
        txt = Mid$(lines(i), p + 1)
        Set c = ws.Cells(rr, 3)
        c.NumberFormat = "@"
        c.WrapText = True
        c.VerticalAlignment = xlTop
        Select Case kind
            Case "H1"
                c.Value = txt
                c.Font.Name = UI_FONT_SEMI: c.Font.Size = 15: c.Font.Color = CText()
                ws.Rows(rr).RowHeight = 34
                c.VerticalAlignment = xlBottom
            Case "H2"
                c.Value = txt
                c.Font.Name = UI_FONT_SEMI: c.Font.Size = 11.5: c.Font.Color = CAccent()
                ws.Rows(rr).RowHeight = 28
                c.VerticalAlignment = xlBottom
            Case "B"
                c.Value = ChrW(&H2022) & "  " & txt
                c.IndentLevel = 1
                c.Font.Size = 10
                ws.Rows(rr).AutoFit
            Case "N"
                c.Value = txt
                c.Font.Size = 9: c.Font.Color = CMuted()
                c.Interior.Color = CSurface()
                ws.Rows(rr).AutoFit
            Case "C"
                c.Value = txt
                c.Font.Name = "Consolas": c.Font.Size = 9
                c.Interior.Color = CSurface()
                ws.Rows(rr).AutoFit
            Case Else
                c.Value = txt
                c.Font.Size = 10
                ws.Rows(rr).AutoFit
        End Select
        If ws.Rows(rr).RowHeight < 16 Then ws.Rows(rr).RowHeight = 16
        rr = rr + 1
    Next
    ws.Cells(1, 1).Value = marker
    ws.Cells(1, 1).Font.Color = CWhite()
    ws.Cells(1, 1).Font.Size = 1
End Sub

' Textbreite von Hilfe/Technik an Fenster und Zoom anpassen (höchstens 760 pt): Zeilen neu
' umbrechen, damit bei Vergrößerung nichts rechts aus dem Fenster läuft
Public Sub FitDocWidth(ByVal ws As Worksheet)
    Dim W As Double, rr As Long, lastR As Long, c As Range
    ' Filter-, Lücken- und Kopfzeile der Listenseiten werden hier nicht gebraucht (sonst leere Fläche)
    ws.Rows(R_CHIPS).RowHeight = 4
    ws.Rows(R_GAP).RowHeight = 4
    ws.Rows(R_HEAD).RowHeight = 4
    W = Application.Min(760, Application.Max(360, mRender.PageRight() - mRender.ContentLeft(ws) - 24))
    If Abs(ws.Columns(3).Width - W) < 3 Then Exit Sub
    mRender.SetColWidthPt ws, 3, W
    lastR = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    For rr = R_DATA To lastR
        Set c = ws.Cells(rr, 3)
        If Len(CStr(c.Value)) > 0 And c.Font.Size <> 15 And c.Font.Size <> 11.5 Then
            ws.Rows(rr).AutoFit
            If ws.Rows(rr).RowHeight < 16 Then ws.Rows(rr).RowHeight = 16
        End If
    Next
End Sub

Public Sub WriteHelp(ByVal ws As Worksheet)
    WriteDoc ws, mHelpText.HelpLines(), "HILFE " & PROG_VERSION
End Sub

Public Sub WriteTech(ByVal ws As Worksheet)
    WriteDoc ws, mHelpText.TechLines(), "TECHNIK " & PROG_VERSION
End Sub

'--- Protokoll anzeigen ------------------------------------------------------------------
Public Sub ShowLog()
    Dim rows As Variant, items() As Variant, i As Long, n As Long, flds As String, a As Variant, j As Long, tn As String
    rows = mLog.RecentRows(500)
    If IsEmpty(rows) Then mUI.Alert "Das Protokoll ist leer.", "Änderungsprotokoll": Exit Sub
    n = UBound(rows, 1)
    ReDim items(0 To n - 1, 0 To 4)
    For i = 1 To n
        items(i - 1, 0) = FmtStamp(rows(i, 1))
        items(i - 1, 1) = NzS(rows(i, 2))
        items(i - 1, 2) = NzS(rows(i, 3))
        tn = ""
        If mCfg.TableExists(NzS(rows(i, 4))) Then tn = mCfg.TableName(NzS(rows(i, 4)))
        items(i - 1, 3) = tn & IIf(Len(NzS(rows(i, 5))) > 0, " · " & rows(i, 5), "")
        flds = ""
        a = Split(NzS(rows(i, 6)), ";")
        For j = 0 To UBound(a)
            If mCfg.FieldExists(CStr(a(j))) Then flds = flds & IIf(Len(flds) > 0, ", ", "") & mCfg.FieldName(CStr(a(j)))
        Next
        items(i - 1, 4) = flds & IIf(Len(NzS(rows(i, 7))) > 0, IIf(Len(flds) > 0, " · ", "") & rows(i, 7), "")
    Next
    mUI.ShowReport "Änderungsprotokoll (neueste " & n & ")", "Protokolliert werden Zeitpunkt, Windows-Kennung (kein Identitätsnachweis), Art der Änderung, Objekt-ID und betroffene Felder, " & _
                   "keine Feldinhalte. Das Protokoll ist Teil der Datei und jeder Sicherung.", items, 5, "95;70;105;150;210", Array("Schließen")
End Sub
