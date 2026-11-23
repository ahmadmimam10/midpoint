Attribute VB_Name = "mViewTable"
'==============================================================================
' mViewTable - Tabellenansicht: Spalten, Filter, Suche, Sortierung, Gruppierung,
' Seiten, Zusammenfassungen, farbige Auswahlwerte, Direktbearbeitung und
' Mehrzellen-Einfuegen mit denselben Pruefregeln wie das Formular.
' Spalten: A Navigation | B Abstand | C Markierung | D Datensatz-ID (ausgeblendet) | E.. Felder
'==============================================================================
Option Explicit

Public Const COL_MARK As Long = 3
Public Const COL_ID As Long = 4
Public Const COL_F1 As Long = 5

Private mRowRid() As String      ' Blattzeile -> Datensatz-ID (Index ab R_DATA)
Private mColFid() As String      ' Feldspalte (Index ab COL_F1)
Private mNRows As Long
Private mNCols As Long
Private mLastRow As Long
Private mSelRow As Long
Private mTid As String
Private mVid As String
Private mHandling As Boolean
Private mWrapSession As cMap     ' Zeilenumbruch im Lesemodus (nicht gespeichert)

'--- Zeilenumbruch je Ansicht (Standard: aus, nur die erste Zeile) ----------------------
Public Function WrapOn(ByVal vid As String) As Boolean
    If mApp.gReadOnly And Not mWrapSession Is Nothing Then
        If mWrapSession.Exists(vid) Then WrapOn = mWrapSession.GetVal(vid): Exit Function
    End If
    WrapOn = (mCfg.ViewOpt(vid, "umbruch", "0") = "1")
End Function

Public Sub SetWrap(ByVal vid As String, ByVal isOn As Boolean)
    If mApp.gReadOnly Then
        If mWrapSession Is Nothing Then Set mWrapSession = New cMap
        mWrapSession.SetVal vid, isOn
    Else
        mCfg.ViewSetOpt vid, "umbruch", IIf(isOn, "1", "0")
        mApp.MarkDirty
    End If
End Sub

'--- Spalten einer Ansicht ------------------------------------------------------------
Public Function DefaultWidth(ByVal f As cField, ByVal isLabel As Boolean) As Single
    If isLabel Then DefaultWidth = 220: Exit Function
    Select Case f.FType
        Case FT_TEXT: DefaultWidth = 150
        Case FT_LTEXT: DefaultWidth = 220
        Case FT_ZAHL, FT_PROZENT: DefaultWidth = 90
        Case FT_GELD: DefaultWidth = 105
        Case FT_DATUM: DefaultWidth = 92
        Case FT_DATUMZEIT: DefaultWidth = 125
        Case FT_JANEIN: DefaultWidth = 70
        Case FT_AUSWAHL, FT_STATUS: DefaultWidth = 125
        Case FT_MEHRFACH, FT_LINK, FT_BACKLINK: DefaultWidth = 170
        Case FT_FORMEL: DefaultWidth = 115
        Case FT_PFAD: DefaultWidth = 180
        Case Else: DefaultWidth = 130
    End Select
End Function

' Collection von Array(FeldID, Breite); Bezeichnungsfeld steht immer vorn
Public Function ViewColumns(ByVal vid As String) As Collection
    Dim tid As String, spec As String, res As New Collection, parts As Variant, i As Long, p As Variant, f As cField
    Dim lf As String, seen As New cMap, w As Single
    tid = mCfg.ViewGet(vid, "Tabelle")
    lf = mCfg.LabelField(tid)
    spec = mCfg.ViewGet(vid, "Spalten")
    If Len(lf) > 0 And mCfg.FieldExists(lf) Then
        w = DefaultWidth(mCfg.Field(lf), True)
        parts = Split(spec, ";")
        For i = 0 To UBound(parts)
            p = Split(parts(i), ":")
            If UBound(p) >= 1 Then If CStr(p(0)) = lf Then w = Val(p(1))
        Next
        res.Add Array(lf, w)
        seen.SetVal lf, True
    End If
    If Len(spec) = 0 Then
        For Each f In mCfg.Fields(tid)
            If Not seen.Exists(f.ID) And f.Visible Then
                res.Add Array(f.ID, DefaultWidth(f, False))
                seen.SetVal f.ID, True
            End If
        Next
    Else
        parts = Split(spec, ";")
        For i = 0 To UBound(parts)
            p = Split(parts(i), ":")
            If UBound(p) >= 0 Then
                If mCfg.FieldExists(CStr(p(0))) And Not seen.Exists(CStr(p(0))) Then
                    Set f = mCfg.Field(CStr(p(0)))
                    If f.IsActive And f.TableID = tid Then
                        If UBound(p) >= 1 Then w = Val(p(1)) Else w = DefaultWidth(f, False)
                        If w < 40 Then w = 40
                        res.Add Array(f.ID, w)
                        seen.SetVal f.ID, True
                    End If
                End If
            End If
        Next
    End If
    Set ViewColumns = res
End Function

Public Function SearchFields(ByVal vid As String) As Collection
    Dim res As New Collection, c As Variant
    For Each c In ViewColumns(vid)
        res.Add mCfg.Field(CStr(c(0)))
    Next
    Set SearchFields = res
End Function

' Gefilterte und sortierte Zeilen einer Ansicht (gemeinsam fuer alle Darstellungen und Export)
Public Function ViewRows(ByVal vid As String, ByRef n As Long, Optional ByVal withGroup As Boolean = True) As Long()
    Dim st As cViewState, tid As String, rows() As Long
    Set st = mNav.State(vid)
    tid = mCfg.ViewGet(vid, "Tabelle")
    rows = mQuery.FilterRows(tid, mQuery.ParseConds(st.Conds), st.Search, SearchFields(vid), st.ArchMode, n)
    If withGroup And Len(st.GroupFid) > 0 And mCfg.FieldExists(st.GroupFid) Then
        mQuery.SortRows tid, rows, n, st.SortSpec, st.GroupFid, st.GroupLevel
    Else
        mQuery.SortRows tid, rows, n, st.SortSpec
    End If
    ViewRows = rows
End Function

'==============================================================================
' Aufbau
'==============================================================================
Public Sub Render()
    Dim ws As Worksheet, tid As String, vid As String, st As cViewState, cols As Collection, c As Variant, i As Long, j As Long
    Dim f As cField, rows() As Long, n As Long, total As Long, pageSize As Long, pages As Long, p0 As Long, p1 As Long
    Dim out() As Variant, t As cTableData, r As Long, outR As Long, lastCol As Long, rng As Range, x As Single
    Dim fids() As String, flds() As cField, gk As String, gl As String, go As Double, prevKey As String, grpRows As Collection
    Dim conds As Collection, k As Long, s As Shape, sr As Range, grpCount As cMap, hasGroup As Boolean, gf As cField
    Dim nOut As Long, footRow As Long, desc As String

    tid = mNav.gTid: vid = mNav.gVid
    mTid = tid: mVid = vid
    Set st = mNav.State(vid)
    Set ws = mRender.Sheet(SH_TABLE)
    mRender.BeginPage ws
    Set cols = ViewColumns(vid)
    mNCols = cols.Count
    ReDim fids(1 To Application.Max(1, mNCols)): ReDim flds(1 To Application.Max(1, mNCols))
    ReDim mColFid(1 To Application.Max(1, mNCols))
    i = 0
    For Each c In cols
        i = i + 1
        fids(i) = CStr(c(0))
        Set flds(i) = mCfg.Field(fids(i))
        mColFid(i) = fids(i)
        mRender.SetColWidthPt ws, COL_F1 + i - 1, CDbl(c(1))
    Next
    mRender.SetColWidthPt ws, COL_MARK, 22
    ws.Columns(COL_ID).ColumnWidth = 0
    lastCol = COL_F1 + mNCols - 1

    ' Kopf
    desc = mCfg.GetS(SH_TAB, tid, "Beschreibung")
    mRender.DrawTitle ws, mCfg.TableIcon(tid), mCfg.TableName(tid), desc
    DrawViewTabs ws, tid, vid, IIf(st.ArchMode = mQuery.ARCH_ONLY, "", "Neuer Eintrag")

    ' Werkzeugleiste
    mRender.FlowBegin ws, R_TOOLS, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowBtn "b_filter", "Filter", "mCmd.Cmd_Filter", "secondary", ChrW(&HE71C)
    mRender.FlowBtn "b_view", "Ansicht", "mCmd.Cmd_ViewMenu", "secondary", ChrW(&HE8FD)
    mRender.FlowBtn "b_export", "Export", "mCmd.Cmd_Export", "secondary", ChrW(&HEDE1)
    ' Schalter: Kästchen mit Haken, wenn langer Text vollständig (umgebrochen) gezeigt wird
    mRender.FlowBtn "b_wrap", "Zeilenumbruch", "mCmd.Cmd_ToggleWrap", "secondary", IIf(WrapOn(vid), ChrW(&HE73A), ChrW(&HE739))
    If mData.UndoAvailable Then mRender.FlowBtn "b_undo", "Rückgängig", "mCmd.Cmd_Undo", "secondary", ChrW(&HE7A7)
    mRender.FlowGap 10
    If st.ArchMode = mQuery.ARCH_ONLY Then
        mRender.FlowBtn "b_restore", "Wiederherstellen", "mCmd.Cmd_RestoreSelected", "secondary", ChrW(&HE777)
        mRender.FlowBtn "b_delete", "Endgültig löschen", "mCmd.Cmd_DeleteSelected", "danger", ChrW(&HE74D)
        mRender.FlowBtn "b_archoff", "Archiv schließen", "mCmd.Cmd_ToggleArchive", "secondary", ChrW(&HE711)
    Else
        mRender.FlowBtn "b_edit", "Bearbeiten", "mCmd.Cmd_EditSelected", "secondary", ChrW(&HE70F)
        mRender.FlowBtn "b_arch", "Archivieren", "mCmd.Cmd_ArchiveSelected", "secondary", ChrW(&HE7B8)
        mRender.FlowBtn "b_bulk", "Sammeländerung", "mCmd.Cmd_BulkEdit", "secondary", ChrW(&HE762)
    End If
    DrawSetupTools
    mRender.FlowEnd
    mRender.Lap "kopf"

    ' Daten bestimmen
    Set t = mData.TData(tid)
    rows = ViewRows(vid, n)
    mRender.Lap "abfrage"
    total = 0
    For r = 1 To t.RowCount
        If st.ArchMode = mQuery.ARCH_ONLY Then
            If t.IsArchived(r) Then total = total + 1
        ElseIf Not t.IsArchived(r) Then
            total = total + 1
        End If
    Next
    DrawChips ws, st, n, total, IIf(st.ArchMode = mQuery.ARCH_ONLY, "archivierte Einträge", "Einträge")

    pageSize = CLng(Val(mCfg.ViewOpt(vid, "seite", CStr(PAGE_SIZE_DEFAULT))))
    If pageSize < 20 Then pageSize = PAGE_SIZE_DEFAULT
    pages = Application.Max(1, -Int(-n / pageSize))
    If st.Page > pages Then st.Page = pages
    If st.Page < 1 Then st.Page = 1
    p0 = (st.Page - 1) * pageSize + 1
    p1 = Application.Min(n, st.Page * pageSize)

    ' Spaltenkoepfe (Texte in einem Zug schreiben, danach nur das Symbolzeichen formatieren)
    ws.Rows(R_HEAD).RowHeight = 26
    ws.Cells(R_HEAD, COL_ID).Value = "ID"
    Dim hdr() As Variant
    ReDim hdr(1 To 1, 1 To Application.Max(1, mNCols))
    For i = 1 To mNCols
        hdr(1, i) = flds(i).Icon & "  " & flds(i).Name & SortMark(st.SortSpec, fids(i)) & IIf(flds(i).Required, " *", "")
    Next
    Set rng = ws.Range(ws.Cells(R_HEAD, COL_F1), ws.Cells(R_HEAD, COL_F1 + Application.Max(1, mNCols) - 1))
    rng.NumberFormat = "@"
    If mNCols > 0 Then rng.Value = hdr
    Set rng = ws.Range(ws.Cells(R_HEAD, COL_MARK), ws.Cells(R_HEAD, lastCol))
    With rng
        .Interior.Color = CSurface()
        .Font.Color = CMuted()
        .Font.Size = 9.5
        .Font.Name = UI_FONT_SEMI
        .HorizontalAlignment = xlLeft
        .IndentLevel = 0
    End With
    ' Typsymbol vor dem Spaltennamen (nach der Bereichsschrift setzen, sonst wird es ueberschrieben)
    For i = 1 To mNCols
        With ws.Cells(R_HEAD, COL_F1 + i - 1).Characters(1, 1).Font
            .Name = ICON_FONT
            .Size = 9
        End With
    Next
    mRender.GridBorders rng
    rng.Borders(xlEdgeBottom).Color = CBorderStrong()
    ' Spalte hinzufuegen (wie in Notion: + am Ende der Kopfzeile); Doppelklick auf Kopf = Spaltenmenue
    If Not mApp.gReadOnly Then
        Set s = mRender.Btn(ws, "hdr_add", "", ws.Cells(R_HEAD, lastCol + 1).Left + 4, ws.Rows(R_HEAD).Top + 2, "mCmd.Cmd_AddProperty", "ghost", ChrW(&HE710), 22, 0, 30)
        s.AlternativeText = "Eigenschaft hinzufügen"
        ' Formelspalte direkt: Formel schreiben, Spalten der Tabelle anklicken
        Set s = mRender.Btn(ws, "hdr_fx", "fx", s.Left + s.Width + 2, ws.Rows(R_HEAD).Top + 2, "mCmd.Cmd_NewFormulaColumn", "ghost", "", 22, 0, 30)
        s.TextFrame2.TextRange.Font.Italic = msoTrue
        s.AlternativeText = "Neue Formelspalte (Excel-Formel)"
    End If

    mRender.Lap "spaltenkopf"
    ' Gruppen zaehlen (fuer Kopfzeilen)
    hasGroup = (Len(st.GroupFid) > 0 And mCfg.FieldExists(st.GroupFid))
    If hasGroup Then
        Set gf = mCfg.Field(st.GroupFid)
        Set grpCount = New cMap
        For k = 1 To n
            mQuery.GroupOf gf, mQuery.GroupValue(t, gf, rows(k)), t.IdAt(rows(k)), st.GroupLevel, gk, gl, go
            grpCount.AddNum "k" & gk, 1
        Next
    End If

    ' Ausgabe vorbereiten
    nOut = 0
    If n > 0 Then
        ReDim out(1 To (p1 - p0 + 1) * 2 + 2, 1 To mNCols + 2)
        ReDim mRowRid(1 To (p1 - p0 + 1) * 2 + 2)
        prevKey = Chr$(1)
        For k = p0 To p1
            r = rows(k)
            If hasGroup Then
                mQuery.GroupOf gf, mQuery.GroupValue(t, gf, r), t.IdAt(r), st.GroupLevel, gk, gl, go
                If gk <> prevKey Then
                    nOut = nOut + 1
                    out(nOut, 1) = ""
                    out(nOut, 2) = ""
                    out(nOut, 3) = ChrW(&H25BE) & "  " & gl & "   (" & grpCount.GetVal("k" & gk, 0) & ")"
                    mRowRid(nOut) = "#G"
                    prevKey = gk
                End If
            End If
            nOut = nOut + 1
            out(nOut, 1) = ChrW(&HE8A7)
            out(nOut, 2) = t.IdAt(r)
            mRowRid(nOut) = t.IdAt(r)
            For i = 1 To mNCols
                If flds(i).IsLinkType Then
                    out(nOut, i + 2) = mFmt.LinkLabels(flds(i), mData.LinkTargets(fids(i), t.IdAt(r)), 4)
                Else
                    out(nOut, i + 2) = mFmt.CellValue(flds(i), t.Val_(r, fids(i)), t.IdAt(r))
                End If
            Next
        Next
    End If
    mNRows = nOut
    mLastRow = R_DATA + nOut - 1
    mRender.Lap "werte"

    If nOut > 0 Then
        ' Zahlenformate vor dem Schreiben setzen (Text bleibt Text)
        For i = 1 To mNCols
            Set rng = ws.Range(ws.Cells(R_DATA, COL_F1 + i - 1), ws.Cells(mLastRow, COL_F1 + i - 1))
            rng.NumberFormat = mFmt.NumberFormatFor(flds(i), AllWhole(out, nOut, i + 2))
            If flds(i).IsLinkType Or flds(i).IsOptionType Or flds(i).FType = FT_JANEIN Then rng.NumberFormat = "@"
            If mFmt.IsRightAligned(flds(i)) Then rng.HorizontalAlignment = xlRight Else rng.HorizontalAlignment = xlLeft
        Next
        ws.Range(ws.Cells(R_DATA, COL_ID), ws.Cells(mLastRow, COL_ID)).NumberFormat = "@"
        ws.Range(ws.Cells(R_DATA, COL_MARK), ws.Cells(mLastRow, lastCol)).Value = SliceOut(out, nOut, mNCols + 2)
        ' Langer Text in der letzten Spalte endet an ihrem Rand: ein Leerzeichen in der
        ' Nachbarspalte stoppt das Überlaufen (Excel schreibt nur in leere Nachbarzellen weiter)
        ws.Range(ws.Cells(R_DATA, lastCol + 1), ws.Cells(mLastRow, lastCol + 1)).Value = " "
        Set rng = ws.Range(ws.Cells(R_DATA, COL_MARK), ws.Cells(mLastRow, lastCol))
        rng.RowHeight = 22
        rng.WrapText = False
        rng.IndentLevel = 0
        mRender.GridBorders rng
        ws.Range(ws.Cells(R_DATA, COL_MARK), ws.Cells(mLastRow, COL_MARK)).Font.Color = CMuted()
        ' Bearbeitbarkeit
        For i = 1 To mNCols
            If flds(i).IsEditable And Not flds(i).IsLinkType And Not mApp.gReadOnly Then
                ws.Range(ws.Cells(R_DATA, COL_F1 + i - 1), ws.Cells(mLastRow, COL_F1 + i - 1)).Locked = False
            Else
                ws.Range(ws.Cells(R_DATA, COL_F1 + i - 1), ws.Cells(mLastRow, COL_F1 + i - 1)).Font.Color = IIf(flds(i).FType = FT_FORMEL, CMuted(), CText())
            End If
        Next
        mRender.Lap "schreiben"
        ' Gruppenzeilen, Optionsfarben, Fehler. Gleich gefaerbte Zellen werden gesammelt und
        ' je Farbe in einem Zug formatiert (statt Zelle fuer Zelle)
        Dim byColor As New Collection, colorKeys As New Collection, ck As Variant, cellR As Range, errR As Range, grpR As Range, v As Variant
        If st.ArchMode = mQuery.ARCH_ONLY Then ws.Range(ws.Cells(R_DATA, COL_F1), ws.Cells(mLastRow, lastCol)).Font.Color = CMuted()
        For j = 1 To nOut
            If mRowRid(j) = "#G" Then
                Set cellR = ws.Range(ws.Cells(R_DATA + j - 1, COL_MARK), ws.Cells(R_DATA + j - 1, lastCol))
                If grpR Is Nothing Then Set grpR = cellR Else Set grpR = Union(grpR, cellR)
            Else
                r = t.RowOf(mRowRid(j))
                For i = 1 To mNCols
                    If flds(i).IsSingleOption Then
                        v = t.Val_(r, fids(i))
                        If Not IsBlank(v) Then CollectCell byColor, colorKeys, mCfg.OptColor(CStr(v)), ws.Cells(R_DATA + j - 1, COL_F1 + i - 1)
                    ElseIf flds(i).FType = FT_FORMEL Then
                        If IsErrVal(t.Val_(r, fids(i))) Then
                            If errR Is Nothing Then Set errR = ws.Cells(R_DATA + j - 1, COL_F1 + i - 1) Else Set errR = Union(errR, ws.Cells(R_DATA + j - 1, COL_F1 + i - 1))
                        End If
                    ElseIf flds(i).FType = FT_MEHRFACH Then
                        ColorMultiCell ws.Cells(R_DATA + j - 1, COL_F1 + i - 1), t.Val_(r, fids(i))
                    End If
                Next
            End If
        Next
        For Each ck In colorKeys
            With byColor(CStr(ck))
                .Interior.Color = PaletteBg(CStr(ck))
                .Font.Color = PaletteFg(CStr(ck))
            End With
        Next
        If Not errR Is Nothing Then errR.Font.Color = CDanger()
        If Not grpR Is Nothing Then
            With grpR
                .Interior.Color = CSurface2()
                .Font.Name = UI_FONT_SEMI
                .Locked = True
                .HorizontalAlignment = xlLeft
                .Borders(xlInsideVertical).LineStyle = xlNone
            End With
        End If
        With ws.Range(ws.Cells(R_DATA, COL_MARK), ws.Cells(mLastRow, COL_MARK))
            .Font.Name = ICON_FONT
            .Font.Size = 8
            .HorizontalAlignment = xlCenter
        End With
        mRender.Lap "farben"
        If WrapOn(vid) Then
            ' ganzer Inhalt sichtbar: umbrechen, Zeilenhöhe nach Inhalt (mindestens wie sonst)
            Set rng = ws.Range(ws.Cells(R_DATA, COL_F1), ws.Cells(mLastRow, lastCol))
            rng.WrapText = True
            If Not grpR Is Nothing Then grpR.WrapText = False
            rng.Rows.AutoFit
            For j = 1 To nOut
                With ws.Rows(R_DATA + j - 1)
                    If .RowHeight < 22 Then
                        .RowHeight = 22
                    ElseIf .RowHeight > 22 Then
                        .RowHeight = Application.Min(409, .RowHeight + 6)
                    End If
                End With
            Next
        End If
        AddValidations ws, flds, fids
        footRow = mLastRow + 1
        DrawSummaries ws, vid, tid, rows, n, flds, fids, footRow
        DrawPager ws, footRow + 1, st, n, p0, p1, pages
    Else
        mRowRid = Split("")
        If total = 0 And st.ArchMode <> mQuery.ARCH_ONLY Then
            mRender.EmptyState ws, R_DATA, "Noch keine Einträge", "Legen Sie den ersten Eintrag an. Alle Felder der Tabelle erscheinen automatisch in der Erfassungsmaske.", _
                               "Ersten Eintrag anlegen", "mCmd.Cmd_NewRecord", ChrW(&HE8A5)
        ElseIf st.ArchMode = mQuery.ARCH_ONLY Then
            mRender.EmptyState ws, R_DATA, "Das Archiv ist leer", "Archivierte Einträge erscheinen hier und können wiederhergestellt werden.", "", "", ChrW(&HE7B8)
        Else
            mRender.EmptyState ws, R_DATA, "Keine Treffer", "Mit der aktuellen Suche und den Filtern wurde nichts gefunden.", "Ansicht zurücksetzen", "mCmd.Cmd_ResetView", ChrW(&HE721)
        End If
    End If

    mRender.EndPage ws, R_HEAD, COL_F1, IIf(nOut > 0, "E" & R_DATA, "")
    mSelRow = 0
    If Len(st.SelRid) > 0 Then HighlightRid st.SelRid
End Sub

' Seit 1.2.0 kein Umbaumodus: Eigenschaften ueber "+" im Kopf, Doppelklick auf den Kopf
' und "Ansicht > Eigenschaften verwalten". Hier nur ein dezenter Einstieg.
Public Sub DrawSetupTools()
    If mApp.gReadOnly Then Exit Sub
    mRender.FlowGap 10
    mRender.FlowBtn "b_props", "Eigenschaften", "mCmd.Cmd_EditTableHere", "ghost", ChrW(&HE713)
End Sub

' Alle Zahlen einer Ausgabespalte ganzzahlig? (Format von Berechnungen ohne feste Kommastellen)
Private Function AllWhole(ByRef a() As Variant, ByVal n As Long, ByVal col As Long) As Boolean
    Dim j As Long, v As Variant
    For j = 1 To n
        v = a(j, col)
        If VarType(v) = vbDouble Then
            If v <> Fix(v) Then Exit Function
        End If
    Next
    AllWhole = True
End Function

Private Function SliceOut(ByRef a() As Variant, ByVal n As Long, ByVal cols As Long) As Variant
    Dim b() As Variant, r As Long, c As Long
    ReDim b(1 To n, 1 To cols)
    For r = 1 To n
        For c = 1 To cols
            b(r, c) = a(r, c)
        Next
    Next
    SliceOut = b
End Function

Private Function SortMark(ByVal spec As String, ByVal fid As String) As String
    If InStr(1, ";" & spec & ";", ";" & fid & ":A;") > 0 Then SortMark = "  " & ChrW(&H2191)
    If InStr(1, ";" & spec & ";", ";" & fid & ":D;") > 0 Then SortMark = "  " & ChrW(&H2193)
End Function

' Zelle der Farbgruppe key hinzufuegen (Union je Farbe)
Private Sub CollectCell(ByVal byColor As Collection, ByVal colorKeys As Collection, ByVal key As String, ByVal cell As Range)
    Dim cur As Range
    On Error Resume Next
    Set cur = byColor(key)
    On Error GoTo 0
    If cur Is Nothing Then
        byColor.Add cell, key
        colorKeys.Add key
    Else
        byColor.Remove key
        byColor.Add Union(cur, cell), key
    End If
End Sub

Public Sub ColorOptionCell(ByVal cell As Range, ByVal oid As Variant)
    Dim key As String
    If IsBlank(oid) Then Exit Sub
    key = mCfg.OptColor(CStr(oid))
    cell.Interior.Color = PaletteBg(key)
    cell.Font.Color = PaletteFg(key)
End Sub

Public Sub ColorMultiCell(ByVal cell As Range, ByVal v As Variant)
    Dim a As Variant, i As Long, pos As Long, nm As String
    If IsBlank(v) Then Exit Sub
    a = Split(CStr(v), ";")
    pos = 1
    For i = 0 To UBound(a)
        nm = mCfg.OptName(CStr(a(i)))
        If pos + Len(nm) - 1 <= Len(cell.Value) Then
            cell.Characters(pos, Len(nm)).Font.Color = PaletteFg(mCfg.OptColor(CStr(a(i))))
            cell.Characters(pos, Len(nm)).Font.Bold = True
        End If
        pos = pos + Len(nm) + 2
    Next
End Sub

Private Sub AddValidations(ByVal ws As Worksheet, flds() As cField, fids() As String)
    Dim i As Long, rng As Range, nm As String
    On Error Resume Next
    For i = 1 To mNCols
        If flds(i).IsSingleOption Or (flds(i).FType = FT_JANEIN And flds(i).Kind <> FT_KONTROLL) Then
            If flds(i).FType = FT_JANEIN Then nm = "L_JANEIN" Else nm = "L_" & fids(i)
            Set rng = ws.Range(ws.Cells(R_DATA, COL_F1 + i - 1), ws.Cells(mLastRow, COL_F1 + i - 1))
            rng.Validation.Delete
            rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="=" & nm
            rng.Validation.IgnoreBlank = True
            rng.Validation.InCellDropdown = True
            rng.Validation.ShowError = False
        End If
    Next
End Sub

Public Sub DrawViewTabs(ByVal ws As Worksheet, ByVal tid As String, ByVal vid As String, Optional ByVal newCaption As String = "", _
                        Optional ByVal newMacro As String = "mCmd.Cmd_NewRecord")
    Dim ids As Variant, i As Long, vt As String, nb As Shape, tb As Shape, pk As String
    mRender.FlowBegin ws, R_TABS, mRender.ContentLeft(ws), mRender.PageRight()
    ids = mCfg.ViewIds(tid)
    For i = 0 To ArrLen(ids) - 1
        vt = mCfg.ViewGet(CStr(ids(i)), "Typ")
        Set tb = mRender.FlowBtn("tab_" & ids(i), mCfg.ViewGet(CStr(ids(i)), "Name"), "mCmd.Cmd_ViewTab", _
                        IIf(CStr(ids(i)) = vid, "tabActive", "tab"), ViewTypeIcon(vt), 2)
        ' farbige Register: Symbol in der Farbe der Ansichtsart, aktives Register farbig hinterlegt
        pk = ViewTypeColor(vt)
        If CStr(ids(i)) = vid Then
            tb.Fill.ForeColor.RGB = PaletteBg(pk)
            tb.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PaletteFg(pk)
        End If
        tb.TextFrame2.TextRange.Characters(1, 1).Font.Fill.ForeColor.RGB = Blend(PaletteFg(pk), RGB(255, 255, 255), 0.1)
    Next
    mRender.FlowGap 6
    mRender.FlowBtn "tab_add", "Ansicht", "mCmd.Cmd_NewView", "tab", ChrW(&HE710)
    If Len(newCaption) > 0 And Not mApp.gReadOnly Then
        Set nb = mRender.FlowBtn("b_new", newCaption, newMacro, "primary", ChrW(&HE710))
        If mRender.PageRight() - nb.Width > nb.Left Then nb.Left = mRender.PageRight() - nb.Width
    End If
    mRender.FlowEnd
    ' feine Linie unter den Registern
    Dim ln As Shape, y As Single
    y = ws.Rows(R_TABS).Top + ws.Rows(R_TABS).Height - 1
    Set ln = ws.Shapes.AddLine(mRender.ContentLeft(ws), y, mRender.PageRight(), y)
    ln.Name = "tab_line"
    ln.Line.ForeColor.RGB = CBorder()
    ln.Line.Weight = 0.75
    ln.Placement = xlFreeFloating
End Sub

Public Function ViewTypeColor(ByVal vt As String) As String
    Select Case vt
        Case VT_CAL: ViewTypeColor = "gruen"
        Case VT_BOARD: ViewTypeColor = "orange"
        Case VT_GANTT: ViewTypeColor = "lila"
        Case Else: ViewTypeColor = "blau"
    End Select
End Function

Public Function ViewTypeIcon(ByVal vt As String) As String
    Select Case vt
        Case VT_CAL: ViewTypeIcon = ChrW(&HE787)
        Case VT_BOARD: ViewTypeIcon = ChrW(&HE8A9)
        Case VT_GANTT: ViewTypeIcon = ChrW(&HE9D5)
        Case Else: ViewTypeIcon = ChrW(&HE80A)
    End Select
End Function

' Filterkacheln, Sortierung, Gruppierung, Suche, Zaehler
Public Sub DrawChips(ByVal ws As Worksheet, ByVal st As cViewState, ByVal n As Long, ByVal total As Long, ByVal unit As String)
    Dim conds As Collection, c As Variant, i As Long, txt As String, p As Variant, parts As Variant
    Dim sr As Range
    Set sr = mRender.SearchBox(ws, R_CHIPS, st.Search, "mCmd.Cmd_Search", 200)
    mRender.FlowBegin ws, R_CHIPS, sr.Left + sr.Width + 6, mRender.PageRight()
    mRender.FlowBtn "b_search", "Suchen", "mCmd.Cmd_Search", "secondary", "", 14
    If n = total Then
        txt = FmtInt(n) & " " & IIf(n = 1 And unit = "Einträge", "Eintrag", unit)
    Else
        txt = FmtInt(n) & " von " & FmtInt(total) & " " & unit
    End If
    mRender.FlowText "cnt", txt, CMuted(), 9.5, True, 12
    If Len(st.Search) > 0 Then mRender.FlowBtn "chip_S", "Suche: " & Truncate(st.Search, 30) & "  " & ChrW(&HD7), "mCmd.Cmd_ChipRemove", "chip"
    Set conds = mQuery.ParseConds(st.Conds)
    i = 0
    For Each c In conds
        i = i + 1
        txt = Truncate(mQuery.CondText(c), 60)
        If i > 1 Then txt = IIf(c(0) = "ODER", "oder ", "und ") & txt
        mRender.FlowBtn "chip_F" & i, txt & "  " & ChrW(&HD7), "mCmd.Cmd_ChipRemove", "chip"
    Next
    If Len(st.SortSpec) > 0 Then
        txt = ""
        parts = Split(st.SortSpec, ";")
        For i = 0 To UBound(parts)
            p = Split(parts(i), ":")
            If mCfg.FieldExists(CStr(p(0))) Then
                txt = txt & IIf(Len(txt) > 0, ", ", "") & mCfg.FieldName(CStr(p(0)))
                If UBound(p) >= 1 Then
                    If p(1) = "D" Then txt = txt & " " & ChrW(&H2193) Else txt = txt & " " & ChrW(&H2191)
                End If
            End If
        Next
        If Len(txt) > 0 Then mRender.FlowBtn "chip_O", "Sortiert: " & txt & "  " & ChrW(&HD7), "mCmd.Cmd_ChipRemove", "chip"
    End If
    If Len(st.GroupFid) > 0 And mCfg.FieldExists(st.GroupFid) Then
        mRender.FlowBtn "chip_G", "Gruppiert: " & mCfg.FieldName(st.GroupFid) & "  " & ChrW(&HD7), "mCmd.Cmd_ChipRemove", "chip"
    End If
    If st.IsModified Then mRender.FlowBtn "b_reset", "Ansicht zurücksetzen", "mCmd.Cmd_ResetView", "link", ChrW(&HE777)
    If st.HasUnsavedSettings Then mRender.FlowBtn "b_saveview", "Als Ansicht speichern", "mCmd.Cmd_SaveView", "link", ChrW(&HE74E)
    mRender.FlowEnd
End Sub

Private Sub DrawSummaries(ByVal ws As Worksheet, ByVal vid As String, ByVal tid As String, rows() As Long, ByVal n As Long, _
                          flds() As cField, fids() As String, ByVal rowNo As Long)
    Dim spec As String, i As Long, agg As String, res As String, any_ As Boolean
    spec = mCfg.ViewOpt(vid, "summen", "")
    ws.Rows(rowNo).RowHeight = 22
    For i = 1 To mNCols
        agg = KVGetList(spec, fids(i))
        If Len(agg) > 0 Then
            res = Summary(tid, rows, n, flds(i), agg)
            With ws.Cells(rowNo, COL_F1 + i - 1)
                .NumberFormat = "@"
                .Value = res
                .Font.Size = 8.5
                .Font.Color = CMuted()
                .HorizontalAlignment = IIf(mFmt.IsRightAligned(flds(i)), xlRight, xlLeft)
            End With
            any_ = True
        End If
    Next
    If any_ Then ws.Range(ws.Cells(rowNo, COL_F1), ws.Cells(rowNo, COL_F1 + mNCols - 1)).Borders(xlEdgeTop).Color = CBorderStrong()
End Sub

Private Function KVGetList(ByVal spec As String, ByVal fid As String) As String
    Dim a As Variant, i As Long, p As Variant
    a = Split(spec, ",")
    For i = 0 To UBound(a)
        p = Split(a(i), ":")
        If UBound(p) >= 1 Then If p(0) = fid Then KVGetList = p(1): Exit Function
    Next
End Function

' Zusammenfassung ueber alle gefilterten Zeilen (nicht nur die aktuelle Seite)
Public Function Summary(ByVal tid As String, rows() As Long, ByVal n As Long, ByVal f As cField, ByVal agg As String) As String
    Dim t As cTableData, k As Long, v As Variant, cnt As Long, s As Double, mn As Variant, mx As Variant, filled As Long, c As Long
    Set t = mData.TData(tid)
    c = t.ColOf(f.ID)
    For k = 1 To n
        If f.IsLinkType Then
            v = mData.LinkTargets(f.ID, t.IdAt(rows(k)))
        ElseIf c > 0 Then
            v = t.Cell(rows(k), c)
        Else
            v = Empty
        End If
        If Not IsBlank(v) And Not IsErrVal(v) Then
            filled = filled + 1
            If f.IsNumeric_ Or f.IsDate_ Then
                If IsNumeric(v) Or IsDate(v) Then
                    cnt = cnt + 1
                    s = s + CDbl(v)
                    If IsEmpty(mn) Then
                        mn = CDbl(v)
                    ElseIf CDbl(v) < mn Then
                        mn = CDbl(v)
                    End If
                    If IsEmpty(mx) Then
                        mx = CDbl(v)
                    ElseIf CDbl(v) > mx Then
                        mx = CDbl(v)
                    End If
                End If
            End If
        End If
    Next
    Select Case agg
        Case "ANZAHL": Summary = "Anzahl " & FmtInt(n)
        Case "GEFUELLT": Summary = "Ausgefüllt " & FmtInt(filled)
        Case "LEER": Summary = "Leer " & FmtInt(n - filled)
        Case "SUMME": If cnt > 0 Then Summary = "Summe " & mFmt.DisplayText(f, s) Else Summary = "Summe (leer)"
        Case "MITTEL": If cnt > 0 Then Summary = "Ø " & mFmt.DisplayText(f, s / cnt) Else Summary = "Ø (leer)"
        Case "MIN": If cnt > 0 Then Summary = "Min " & mFmt.DisplayText(f, mn) Else Summary = "Min (leer)"
        Case "MAX": If cnt > 0 Then Summary = "Max " & mFmt.DisplayText(f, mx) Else Summary = "Max (leer)"
    End Select
End Function

Private Sub DrawPager(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal st As cViewState, ByVal n As Long, ByVal p0 As Long, ByVal p1 As Long, ByVal pages As Long)
    Dim x As Single, y As Single, s As Shape
    ws.Rows(rowNo).RowHeight = 34
    x = ws.Columns(COL_F1).Left
    y = ws.Rows(rowNo).Top + 5
    If Not mApp.gReadOnly And st.ArchMode <> mQuery.ARCH_ONLY Then
        Set s = mRender.Btn(ws, "pg_new", "Neuer Eintrag", x, y, "mCmd.Cmd_NewRecord", "link", ChrW(&HE710))
        x = x + s.Width + 16
    End If
    If pages > 1 Then
        Set s = mRender.TextShape(ws, "pg_info", "Einträge " & FmtInt(p0) & " bis " & FmtInt(p1) & " von " & FmtInt(n) & "  ·  Seite " & st.Page & " von " & pages, _
                                  x, y, 260, BTN_HT, 9, CMuted(), False, 1, True)
        x = x + s.Width + 10
        If st.Page > 1 Then
            Set s = mRender.Btn(ws, "pg_prev", "Zurück", x, y, "mCmd.Cmd_PagePrev", "secondary", ChrW(&HE76B))
            x = x + s.Width + 6
        End If
        If st.Page < pages Then mRender.Btn ws, "pg_next", "Weiter", x, y, "mCmd.Cmd_PageNext", "secondary", ChrW(&HE76C)
    End If
End Sub

'==============================================================================
' Auswahl
'==============================================================================
Public Function RidAtRow(ByVal sheetRow As Long) As String
    Dim i As Long
    i = sheetRow - R_DATA + 1
    If i < 1 Or i > mNRows Then Exit Function
    If mRowRid(i) = "#G" Then Exit Function
    RidAtRow = mRowRid(i)
End Function

Public Function FidAtCol(ByVal sheetCol As Long) As String
    Dim i As Long
    i = sheetCol - COL_F1 + 1
    If i < 1 Or i > mNCols Then Exit Function
    FidAtCol = mColFid(i)
End Function

Public Function SelectedRids() As Collection
    Dim res As New Collection, area As Range, r As Range, rid As String, seen As New cMap
    On Error Resume Next
    If Not ActiveSheet.Name = SH_TABLE Then Set SelectedRids = res: Exit Function
    For Each area In Selection.Areas
        For Each r In area.Rows
            rid = RidAtRow(r.Row)
            If Len(rid) > 0 And Not seen.Exists(rid) Then
                res.Add rid
                seen.SetVal rid, True
            End If
        Next
    Next
    Set SelectedRids = res
End Function

Public Sub OnSelect(ByVal target As Range)
    Dim rid As String
    If mNRows = 0 Then Exit Sub
    rid = RidAtRow(target.Row)
    If Len(rid) = 0 Then Exit Sub
    mNav.State(mVid).SelRid = rid
    If target.Rows.Count = 1 Then HighlightRow target.Row
End Sub

Private Sub HighlightRid(ByVal rid As String)
    Dim i As Long
    For i = 1 To mNRows
        If mRowRid(i) = rid Then HighlightRow R_DATA + i - 1: Exit Sub
    Next
End Sub

Private Sub HighlightRow(ByVal sheetRow As Long)
    Dim ws As Worksheet, lastCol As Long, prevEv As Boolean
    Set ws = mRender.Sheet(SH_TABLE)
    lastCol = COL_F1 + mNCols - 1
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    If mSelRow > 0 And mSelRow <> sheetRow Then
        With ws.Range(ws.Cells(mSelRow, COL_MARK), ws.Cells(mSelRow, lastCol))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = CBorder()
        End With
    End If
    With ws.Range(ws.Cells(sheetRow, COL_MARK), ws.Cells(sheetRow, lastCol))
        .Borders(xlEdgeTop).Color = CAccent(): .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = CAccent(): .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = CAccent(): .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Color = CAccent(): .Borders(xlEdgeRight).Weight = xlMedium
    End With
    mSelRow = sheetRow
    Application.EnableEvents = prevEv
End Sub

'==============================================================================
' Direkte Bearbeitung / Einfuegen
'==============================================================================
Public Function IsSearchCell(ByVal target As Range) As Boolean
    Dim ws As Worksheet
    Set ws = target.Worksheet
    IsSearchCell = (target.Row = R_TOOLS And target.Column >= C_CONTENT And target.Cells.Count >= 1 And target.Column <= COL_F1 + 1)
End Function

Public Sub OnChange(ByVal target As Range)
    Dim ws As Worksheet, cell As Range, rid As String, fid As String, f As cField, outV As Variant, msg As String
    Dim changes As New cMap, recVals As cMap, errs As New Collection, k As Variant, recErrs As cMap, ek As Variant
    Dim tid As String, cur As Variant, n As Long, report() As Variant, i As Long, scrollR As Long, selAddr As String, ok As Boolean
    If mHandling Then Exit Sub
    Set ws = target.Worksheet
    tid = mTid
    If target.Row = R_CHIPS And target.Column <= COL_F1 + 1 And target.Rows.Count = 1 Then
        mNav.State(mVid).Search = Trim$(NzS(ws.Cells(R_CHIPS, C_CONTENT).Value))
        mNav.State(mVid).Page = 1
        mNav.RenderCurrent
        Exit Sub
    End If
    If target.Cells.CountLarge > 5000 Then
        RevertCells target
        mUI.Alert "Es wurden zu viele Zellen auf einmal geändert. Bitte in kleineren Abschnitten einfügen (höchstens 5.000 Zellen).", "Einfügen"
        Exit Sub
    End If
    If Not mApp.RequireWritable("Die Bearbeitung") Then RevertCells target: Exit Sub
    ' 1) Alle Zellen pruefen und umwandeln
    For Each cell In target.Cells
        rid = RidAtRow(cell.Row)
        fid = FidAtCol(cell.Column)
        If Len(rid) = 0 Or Len(fid) = 0 Then
            errs.Add Array("Zelle " & cell.Address(False, False), "", "Dieser Bereich kann nicht bearbeitet werden.")
        Else
            Set f = mCfg.Field(fid)
            cur = mData.GetValue(tid, rid, fid)
            If Not f.IsEditable Or f.IsLinkType Then
                errs.Add Array(mData.RecordLabel(tid, rid), f.Name, "Dieses Feld wird über das Formular bzw. automatisch gepflegt.")
            ElseIf mValid.ParseCell(f, cell.Value, outV, msg, NzS(cur)) Then
                If changes.Exists(rid) Then Set recVals = changes.GetVal(rid) Else Set recVals = New cMap: changes.SetVal rid, recVals
                recVals.SetVal fid, outV
            Else
                errs.Add Array(mData.RecordLabel(tid, rid), f.Name, msg)
            End If
        End If
    Next
    ' 2) Datensatzregeln (Pflicht, Eindeutigkeit, Datumsfolge)
    If errs.Count = 0 Then
        For Each k In changes.Keys
            Set recVals = changes.GetVal(CStr(k))
            If Not mValid.ValidateRecord(tid, CStr(k), recVals, New cMap, recErrs) Then
                For Each ek In recErrs.Keys
                    errs.Add Array(mData.RecordLabel(tid, CStr(k)), mCfg.FieldName(CStr(ek)), CStr(recErrs.GetVal(CStr(ek))))
                Next
            End If
        Next
    End If
    If errs.Count > 0 Then
        RevertCells target
        ReDim report(0 To errs.Count - 1, 0 To 2)
        i = 0
        For Each k In errs
            report(i, 0) = k(0): report(i, 1) = k(1): report(i, 2) = k(2)
            i = i + 1
        Next
        mUI.ShowReport "Änderung nicht übernommen", "Mindestens ein Wert ist ungültig. Es wurde nichts gespeichert; alle Zellen zeigen wieder den bisherigen Stand. " & _
                       "Bitte korrigieren und erneut eingeben.", report, 3, "150;110;300", Array("OK")
        Exit Sub
    End If
    ' 3) Gemeinsam speichern
    scrollR = ActiveWindow.ScrollRow
    selAddr = Selection.Address
    mHandling = True
    On Error GoTo failed
    mData.TxBegin "Bearbeitung in der Tabelle (" & Plural(changes.Count, "Eintrag", "Einträge") & ")"
    For Each k In changes.Keys
        mData.UpdateRecord tid, CStr(k), changes.GetVal(CStr(k))
        mLog.Add "Geändert", tid, CStr(k), Join(changes.GetVal(CStr(k)).Keys, ";"), "Tabellenansicht"
    Next
    mData.TxCommit
    mHandling = False
    If changes.Count = 1 And target.Cells.Count = 1 Then
        Dim ks As Variant
        ks = changes.Keys
        RefreshRow CStr(ks(0))
        mRender.RefreshStatus
    Else
        mNav.RenderCurrent
        On Error Resume Next
        ActiveWindow.ScrollRow = scrollR
        ws.Range(selAddr).Select
    End If
    Exit Sub
failed:
    mHandling = False
    mApp.Boundary "Speichern der Tabellenänderung"
    mNav.RenderCurrent
End Sub

' Stellt die angezeigten Werte wieder her (aus den Daten)
Private Sub RevertCells(ByVal target As Range)
    Dim cell As Range, rid As String, fid As String, f As cField, prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    mHandling = True
    On Error Resume Next
    For Each cell In target.Cells
        rid = RidAtRow(cell.Row)
        fid = FidAtCol(cell.Column)
        If Len(rid) > 0 And Len(fid) > 0 Then
            Set f = mCfg.Field(fid)
            If f.IsLinkType Then
                cell.Value = mFmt.LinkLabels(f, mData.LinkTargets(fid, rid), 4)
            Else
                cell.Value = mFmt.CellValue(f, mData.GetValue(mTid, rid, fid), rid)
            End If
        ElseIf cell.Row >= R_DATA Then
            If cell.Column = COL_ID And Len(RidAtRow(cell.Row)) > 0 Then cell.Value = RidAtRow(cell.Row) Else cell.ClearContents
        End If
    Next
    mHandling = False
    Application.EnableEvents = prevEv
End Sub

' Aktualisiert eine Zeile nach einer Einzelaenderung, ohne die Seite neu aufzubauen
Public Sub RefreshRow(ByVal rid As String)
    Dim ws As Worksheet, j As Long, i As Long, f As cField, cell As Range, v As Variant
    Set ws = mRender.Sheet(SH_TABLE)
    mApp.BeginWork
    mHandling = True
    For j = 1 To mNRows
        If mRowRid(j) = rid Then
            For i = 1 To mNCols
                Set f = mCfg.Field(mColFid(i))
                Set cell = ws.Cells(R_DATA + j - 1, COL_F1 + i - 1)
                v = mData.GetValue(mTid, rid, mColFid(i))
                If f.IsLinkType Then
                    cell.Value = mFmt.LinkLabels(f, mData.LinkTargets(f.ID, rid), 4)
                Else
                    cell.Value = mFmt.CellValue(f, v, rid)
                End If
                If f.IsSingleOption Then
                    cell.Interior.Pattern = xlNone
                    cell.Font.Color = CText()
                    ColorOptionCell cell, v
                ElseIf f.FType = FT_FORMEL Then
                    cell.Font.Color = IIf(IsErrVal(v), CDanger(), CMuted())
                ElseIf f.FType = FT_MEHRFACH Then
                    cell.Font.Color = CText(): cell.Font.Bold = False
                    ColorMultiCell cell, v
                End If
            Next
            Exit For
        End If
    Next
    mHandling = False
    mApp.EndWork
    mApp.SyncSavedFlag
End Sub

' Kontrollkaestchen per Doppelklick umschalten (gemeinsame Pruef- und Speicherwege)
Private Sub ToggleCheck(ByVal rid As String, ByVal fid As String)
    Dim vals As New cMap, errs As cMap, cur As Variant
    If Not mApp.RequireWritable("Das Ändern") Then Exit Sub
    cur = mData.GetValue(mTid, rid, fid)
    If IsBlank(cur) Then cur = False
    vals.SetVal fid, Not CBool(cur)
    If Not mValid.ValidateRecord(mTid, rid, vals, New cMap, errs) Then
        mUI.Alert "Die Änderung ist nicht möglich: " & Join(errs.Items, " "), "Kontrollkästchen", True
        Exit Sub
    End If
    mApp.BeginWork
    mData.TxBegin "„" & mCfg.FieldName(fid) & "“ umgeschaltet"
    mData.UpdateRecord mTid, rid, vals
    mLog.Add "Geändert", mTid, rid, fid, "Kontrollkästchen"
    mData.TxCommit
    mApp.EndWork
    RefreshRow rid
    mRender.RefreshStatus
End Sub

' Spaltenmenue (Doppelklick auf eine Ueberschrift): sortieren, filtern, gruppieren, ausblenden, bearbeiten
Public Sub HeaderMenu(ByVal fid As String)
    Dim f As cField, sx As String, isLbl As Boolean
    Set f = mCfg.Field(fid)
    sx = ""
    isLbl = (fid = mCfg.LabelField(mTid))
    If f.FType = FT_FORMEL And Left$(f.Formula, 6) <> "ROLLUP" Then
        sx = "Formel bearbeiten (Spalten anklicken)…|mCmd.Cmd_HdrFormula|" & fid & "|1|1;" & _
             "Name, Format und Weiteres…|mCmd.Cmd_HdrEdit|" & fid
    Else
        sx = "Eigenschaft bearbeiten…|mCmd.Cmd_HdrEdit|" & fid & "|1|1"
    End If
    mRender.Popup Split("Aufsteigend sortieren|mCmd.Cmd_HdrSortA|" & fid & ";Absteigend sortieren|mCmd.Cmd_HdrSortD|" & fid & ";" & _
                        "Nach „" & Replace(Replace(f.Name, ";", ","), "|", "/") & "“ gruppieren|mCmd.Cmd_HdrGroup|" & fid & "|" & IIf(mQuery.CanGroupBy(f), "1", "0") & ";" & _
                        "Filter…|mCmd.Cmd_Filter||1|1;" & _
                        "Spalte in dieser Ansicht ausblenden|mCmd.Cmd_HdrHide|" & fid & "|" & IIf(isLbl, "0", "1") & ";" & _
                        sx & ";" & _
                        "Neue Eigenschaft rechts davon…|mCmd.Cmd_HdrInsert|" & fid & "|1|1;" & _
                        "Neue Formelspalte rechts davon…|mCmd.Cmd_HdrNewFormula|" & fid & ";" & _
                        "Eigenschaft entfernen…|mCmd.Cmd_HdrDelete|" & fid & "|" & IIf(isLbl, "0", "1") & "|1", ";")
End Sub

Public Sub OnDoubleClick(ByVal target As Range, ByRef Cancel As Boolean)
    Dim rid As String, fid As String, st As cViewState, spec As String
    If target.Row = R_HEAD Then
        fid = FidAtCol(target.Column)
        If Len(fid) > 0 Then
            Cancel = True
            HeaderMenu fid
        End If
        Exit Sub
    End If
    rid = RidAtRow(target.Row)
    fid = FidAtCol(target.Column)
    If Len(rid) > 0 And Len(fid) > 0 Then
        Select Case mCfg.Field(fid).Kind
            Case FT_KONTROLL
                Cancel = True
                ToggleCheck rid, fid
                Exit Sub
            Case FT_BUTTON
                Cancel = True
                mForms.RunButton mTid, rid, fid
                Exit Sub
        End Select
    End If
    If Len(rid) > 0 Then
        Cancel = True
        mNav.State(mVid).SelRid = rid
        mForms.OpenRecord mTid, rid
    End If
End Sub
