Attribute VB_Name = "mCmd"
'==============================================================================
' mCmd - Alle Bedienbefehle (Schaltflaechen, Menues) mit Fehlergrenze sowie
' gefuehrte Ablaeufe (Ansichten, Sammelaenderung, Strukturaenderungen,
' Dashboards, Einstellungen).
'==============================================================================
Option Explicit

'==============================================================================
' Allgemein
'==============================================================================
Public Sub Cmd_Save()
    On Error GoTo f
    mBackup.ControlledSave False
    Exit Sub
f: mApp.Boundary "Speichern"
End Sub

Public Sub Cmd_GoStart()
    On Error GoTo f
    mNav.GoStart
    Exit Sub
f: mApp.Boundary "Start"
End Sub

Public Sub Cmd_GoDash()
    On Error GoTo f
    mNav.GoDash ""
    Exit Sub
f: mApp.Boundary "Dashboards"
End Sub

Public Sub Cmd_GoHelp()
    On Error GoTo f
    mNav.GoPage "Hilfe"
    Exit Sub
f: mApp.Boundary "Hilfe"
End Sub

Public Sub Cmd_ToggleMode()
    On Error GoTo f
    mBackup.ToggleAutoSave
    Exit Sub
f: mApp.Boundary "Automatisch speichern"
End Sub

Public Sub Cmd_Undo()
    Dim lbl As String
    On Error GoTo f
    If Not mData.UndoAvailable Then mUI.Alert "In dieser Sitzung gibt es keine Datenänderung, die zurückgenommen werden kann.", "Rückgängig": Exit Sub
    If Not mApp.RequireWritable("Rückgängig") Then Exit Sub
    If Not mUI.Confirm("Letzte Änderung zurücknehmen?" & vbLf & vbLf & "„" & mData.UndoLabel & "“" & vbLf & vbLf & _
                       "Zurückgenommen werden Einträge und Verknüpfungen dieser Änderung. Strukturänderungen sind ausgenommen.", "Rückgängig", "Zurücknehmen", "Abbrechen") Then Exit Sub
    mApp.BeginWork
    lbl = mData.UndoLast()
    mApp.EndWork
    mNav.RenderCurrent
    mRender.RefreshStatus
    Exit Sub
f: mApp.Boundary "Rückgängig"
    mNav.RenderCurrent
End Sub

'==============================================================================
' Suche
'==============================================================================
Public Sub Cmd_Search()
    Dim ws As Worksheet, st As cViewState
    On Error GoTo f
    Set ws = ActiveSheet
    Set st = mNav.CurState
    If st Is Nothing Then Exit Sub
    st.Search = Trim$(NzS(ws.Cells(R_CHIPS, C_CONTENT).Value))
    st.Page = 1
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Suchen"
End Sub

Public Sub Cmd_GlobalSearch()
    Dim ws As Worksheet
    On Error GoTo f
    Set ws = ActiveSheet
    mNav.gGlobalSearch = Trim$(NzS(ws.Cells(R_TABS, C_CONTENT).Value))
    mNav.GoPage "Suche"
    Exit Sub
f: mApp.Boundary "Suchen"
End Sub

'==============================================================================
' Ansichten
'==============================================================================
Public Sub Cmd_ViewTab()
    Dim nm As String
    On Error GoTo f
    nm = mRender.CallerName()
    If Left$(nm, 4) = "tab_" Then mNav.GoTable mNav.gTid, Mid$(nm, 5)
    Exit Sub
f: mApp.Boundary "Ansicht wechseln"
End Sub

Public Sub Cmd_ViewMenu()
    Dim vt As String, st As cViewState, items As New Collection, arr() As Variant, i As Long, v As Variant
    On Error GoTo f
    vt = mCfg.ViewGet(mNav.gVid, "Typ")
    Set st = mNav.CurState
    If vt = VT_TABLE Or vt = VT_GANTT Then items.Add "Sortieren…|mCmd.Cmd_ViewSort"
    If vt = VT_TABLE Then
        items.Add "Gruppieren…|mCmd.Cmd_ViewGroup"
        items.Add "Spalten auswählen…|mCmd.Cmd_ViewColumns"
        items.Add IIf(st.ArchMode = mQuery.ARCH_ONLY, "Archiv schließen", "Archiv anzeigen") & "|mCmd.Cmd_ToggleArchive||1|1"
    End If
    items.Add "Ansicht bearbeiten…|mCmd.Cmd_EditView||1|1"
    items.Add "Einstellungen in Ansicht speichern|mCmd.Cmd_SaveView||" & IIf(st.HasUnsavedSettings, "1", "0")
    items.Add "Als neue Ansicht speichern…|mCmd.Cmd_ViewSaveAs"
    items.Add "Ansicht zurücksetzen|mCmd.Cmd_ResetView||" & IIf(st.IsModified, "1", "0")
    items.Add "Neue Ansicht…|mCmd.Cmd_NewView||1|1"
    items.Add "Ansicht duplizieren|mCmd.Cmd_ViewDuplicate"
    items.Add "Register nach links|mCmd.Cmd_ViewMoveLeft"
    items.Add "Register nach rechts|mCmd.Cmd_ViewMoveRight"
    items.Add "Ansicht löschen…|mCmd.Cmd_ViewDelete||1|1"
    items.Add "Tabelle „" & Truncate(mCfg.TableName(mNav.gTid), 30) & "“ löschen…|mCmd.Cmd_TableDelete||1|1"
    ReDim arr(0 To items.Count - 1)
    For Each v In items
        arr(i) = v: i = i + 1
    Next
    mRender.Popup arr
    Exit Sub
f: mApp.Boundary "Ansicht"
End Sub

Public Sub Cmd_ViewSort()
    Dim f As frmProps, flds As Collection, names() As Variant, i As Long, st As cViewState, parts As Variant, p As Variant, fld As cField, spec As String, k As Long, fid As String
    On Error GoTo fl
    Set st = mNav.CurState
    Set flds = mCfg.Fields(mNav.gTid)
    ReDim names(0 To flds.Count)
    names(0) = "(keine)"
    i = 1
    For Each fld In flds
        names(i) = fld.Name: i = i + 1
    Next
    Set f = New frmProps
    f.Init "Sortieren", "Bis zu drei Stufen. Leere Werte stehen immer am Ende. Gilt für diese Sitzung; mit „Einstellungen in Ansicht speichern“ dauerhaft übernehmen."
    parts = Split(st.SortSpec, ";")
    For k = 1 To 3
        fid = "": p = Empty
        If k - 1 <= UBound(parts) Then
            p = Split(parts(k - 1), ":")
            fid = p(0)
        End If
        f.AddCombo "s" & k, k & ". Sortieren nach", names, IIf(mCfg.FieldExists(fid), mCfg.FieldName(fid), "(keine)")
        Dim dirTxt As String
        dirTxt = "aufsteigend (A bis Z, 0 bis 9, alt nach neu)"
        If IsArray(p) Then
            If UBound(p) >= 1 Then
                If p(1) = "D" Then dirTxt = "absteigend (Z bis A, 9 bis 0, neu nach alt)"
            End If
        End If
        f.AddCombo "d" & k, "Richtung", Array("aufsteigend (A bis Z, 0 bis 9, alt nach neu)", "absteigend (Z bis A, 9 bis 0, neu nach alt)"), dirTxt
    Next
    If f.ShowDialog() Then
        spec = ""
        For k = 1 To 3
            fid = mCfg.FieldIdByName(mNav.gTid, f.GetText("s" & k))
            If Len(fid) > 0 Then spec = spec & IIf(Len(spec) > 0, ";", "") & fid & ":" & IIf(f.GetIndex("d" & k) = 1, "D", "A")
        Next
        st.SortSpec = spec
        Unload f
        mNav.RenderCurrent
    Else
        Unload f
    End If
    Exit Sub
fl: mApp.Boundary "Sortieren"
End Sub

Public Sub Cmd_ViewGroup()
    Dim f As frmProps, names() As Variant, i As Long, st As cViewState, fld As cField, fid As String, n As Long
    On Error GoTo fl
    Set st = mNav.CurState
    ReDim names(0 To mCfg.Fields(mNav.gTid).Count)
    names(0) = "(keine Gruppierung)"
    n = 1
    For Each fld In mCfg.Fields(mNav.gTid)
        If mQuery.CanGroupBy(fld) Then names(n) = fld.Name: n = n + 1
    Next
    ReDim Preserve names(0 To n - 1)
    Set f = New frmProps
    f.Init "Gruppieren", "Einträge werden unter Zwischenüberschriften mit Anzahl zusammengefasst. Bei Mehrfachverknüpfungen zählt die erste Zuordnung."
    f.AddCombo "g", "Gruppieren nach", names, IIf(mCfg.FieldExists(st.GroupFid), mCfg.FieldName(st.GroupFid), "(keine Gruppierung)")
    f.AddCombo "l", "Bei Datumsfeldern zusammenfassen nach", Array("Tag", "Monat", "Quartal", "Jahr"), IIf(Len(st.GroupLevel) > 0, st.GroupLevel, "Monat")
    If f.ShowDialog() Then
        fid = mCfg.FieldIdByName(mNav.gTid, f.GetText("g"))
        st.GroupFid = fid
        st.GroupLevel = f.GetText("l")
        st.Page = 1
        Unload f
        mNav.RenderCurrent
    Else
        Unload f
    End If
    Exit Sub
fl: mApp.Boundary "Gruppieren"
End Sub

Public Sub Cmd_ViewColumns()
    Dim f As frmColumns
    On Error GoTo fl
    Set f = New frmColumns
    f.Setup mNav.gVid
    f.Show: mNav.FocusGrid
    If f.OK Then
        mCfg.SetV SH_VIEW, mNav.gVid, "Spalten", f.ColSpec
        mCfg.ViewSetOpt mNav.gVid, "summen", f.SumSpec
        mLog.Add "Ansicht geändert", mNav.gTid, mNav.gVid, "", "Spalten"
        Unload f
        mNav.RenderCurrent
    Else
        Unload f
    End If
    Exit Sub
fl: mApp.Boundary "Spalten auswählen"
End Sub

Public Sub Cmd_Filter()
    Dim f As frmFilter, st As cViewState
    On Error GoTo fl
    Set st = mNav.CurState
    Set f = New frmFilter
    f.Setup mNav.gTid, st.Conds, "Filter: " & mCfg.TableName(mNav.gTid)
    f.Show: mNav.FocusGrid
    If f.OK Then
        st.Conds = f.Result
        st.Page = 1
        Unload f
        mNav.RenderCurrent
    Else
        Unload f
    End If
    Exit Sub
fl: mApp.Boundary "Filter"
End Sub

Public Sub Cmd_ChipRemove()
    Dim nm As String, st As cViewState, conds As Collection, nc As New Collection, i As Long, c As Variant, k As Long
    On Error GoTo f
    nm = mRender.CallerName()
    Set st = mNav.CurState
    Select Case True
        Case nm = "chip_S": st.Search = ""
        Case nm = "chip_O": st.SortSpec = ""
        Case nm = "chip_G": st.GroupFid = ""
        Case Left$(nm, 6) = "chip_F"
            k = CLng(Mid$(nm, 7))
            Set conds = mQuery.ParseConds(st.Conds)
            For Each c In conds
                i = i + 1
                If i <> k Then
                    If nc.Count = 0 Then c(0) = ""
                    nc.Add c
                End If
            Next
            st.Conds = mQuery.SerializeConds(nc)
    End Select
    st.Page = 1
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Filter entfernen"
End Sub

Public Sub Cmd_ResetView()
    On Error GoTo f
    mNav.ResetState mNav.gVid
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Ansicht zurücksetzen"
End Sub

Public Sub Cmd_SaveView()
    Dim st As cViewState
    On Error GoTo f
    Set st = mNav.CurState
    mCfg.SetV SH_VIEW, mNav.gVid, "Filter", st.Conds
    mCfg.SetV SH_VIEW, mNav.gVid, "Sortierung", st.SortSpec
    mCfg.SetV SH_VIEW, mNav.gVid, "Gruppierung", st.GroupSpec
    mLog.Add "Ansicht gespeichert", mNav.gTid, mNav.gVid, ""
    mNav.RenderCurrent
    mUI.Alert "Filter, Sortierung und Gruppierung wurden in der Ansicht „" & mCfg.ViewGet(mNav.gVid, "Name") & "“ gespeichert. " & _
              "Sie bleiben nach Speichern und erneutem Öffnen der Datei erhalten.", "Ansicht gespeichert"
    Exit Sub
f: mApp.Boundary "Ansicht speichern"
End Sub

Public Sub Cmd_ViewSaveAs()
    Dim nm As String, ok As Boolean, vid As String, st As cViewState
    On Error GoTo f
    Set st = mNav.CurState
    nm = mUI.AskText("Name der neuen Ansicht (übernimmt die aktuellen Filter, Sortierung, Gruppierung und Spalten):", "Als neue Ansicht speichern", mCfg.ViewGet(mNav.gVid, "Name") & " (Kopie)", False, ok)
    If Not ok Or Len(Trim$(nm)) = 0 Then Exit Sub
    vid = mStruct.CreateView(mNav.gTid, Trim$(nm), mCfg.ViewGet(mNav.gVid, "Typ"), mCfg.ViewGet(mNav.gVid, "Optionen"), mCfg.ViewGet(mNav.gVid, "Spalten"), _
                             st.Conds, st.SortSpec, st.GroupSpec)
    mNav.GoTable mNav.gTid, vid
    Exit Sub
f: mApp.Boundary "Ansicht speichern"
End Sub

Public Sub Cmd_ViewDuplicate()
    Dim vid As String
    On Error GoTo f
    vid = mStruct.DuplicateView(mNav.gVid, mCfg.ViewGet(mNav.gVid, "Name") & " (Kopie)")
    mNav.GoTable mNav.gTid, vid
    Exit Sub
f: mApp.Boundary "Ansicht duplizieren"
End Sub

Public Sub Cmd_ViewMoveLeft()
    On Error GoTo f
    mStruct.MoveInOrder SH_VIEW, "Tabelle", mNav.gTid, mNav.gVid, -1
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Register verschieben"
End Sub

Public Sub Cmd_ViewMoveRight()
    On Error GoTo f
    mStruct.MoveInOrder SH_VIEW, "Tabelle", mNav.gTid, mNav.gVid, 1
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Register verschieben"
End Sub

Public Sub Cmd_ViewDelete()
    On Error GoTo f
    If Not mUI.Confirm("Ansicht „" & mCfg.ViewGet(mNav.gVid, "Name") & "“ löschen? Die Daten bleiben unverändert, nur diese Darstellung wird entfernt.", _
                       "Ansicht löschen", "Löschen", "Abbrechen", True) Then Exit Sub
    mStruct.DeleteView mNav.gVid
    mNav.gVid = ""
    mNav.GoTable mNav.gTid
    Exit Sub
f: mApp.Boundary "Ansicht löschen"
End Sub

Public Sub Cmd_NewView()
    Dim f As frmProps, nm As String, vt As String, vid As String, typName As String, back As Boolean
    On Error GoTo fl
    typName = "Tabelle"
step1:
    Set f = New frmProps
    f.Init "Neue Ansicht", "Ansichten zeigen dieselben Daten unterschiedlich an. Sie sind keine Kopien: Änderungen gelten überall."
    f.AddText "name", "Name der Ansicht", nm
    f.AddCombo "typ", "Darstellung", Array("Tabelle", "Kalender", "Board", "Gantt (Projektleiste)"), typName
    Do
        If Not f.ShowDialog("Weiter") Then Unload f: Exit Sub
        f.ClearErrors
        nm = Trim$(f.GetText("name"))
        If Len(nm) > 0 Then Exit Do
        f.SetError "name", "Bitte einen Namen angeben."
    Loop
    typName = f.GetText("typ")
    vt = Split(typName, " ")(0)
    Unload f
    vid = mStruct.CreateView(mNav.gTid, nm, vt)
    If vt <> VT_TABLE Then
        If Not EditViewSettings(vid, True, back) Then
            mCfg.DeleteRow SH_VIEW, vid
            If back Then GoTo step1
            Exit Sub
        End If
    End If
    mNav.GoTable mNav.gTid, vid
    Exit Sub
fl: mApp.Boundary "Neue Ansicht"
End Sub

Public Sub Cmd_EditView()
    On Error GoTo f
    If EditViewSettings(mNav.gVid, False) Then
        mNav.ResetState mNav.gVid
        mNav.RenderCurrent
    End If
    Exit Sub
f: mApp.Boundary "Ansicht bearbeiten"
End Sub

Private Function FieldNamesWhere(ByVal tid As String, ByVal kinds As String, Optional ByVal withNone As String = "") As Variant
    Dim res() As Variant, n As Long, f As cField
    ReDim res(0 To mCfg.Fields(tid).Count)
    If Len(withNone) > 0 Then res(0) = withNone: n = 1
    For Each f In mCfg.Fields(tid)
        If InStr(1, "," & kinds & ",", "," & f.ValueType & ",") > 0 Then res(n) = f.Name: n = n + 1
    Next
    If n = 0 Then FieldNamesWhere = Array(): Exit Function
    ReDim Preserve res(0 To n - 1)
    FieldNamesWhere = res
End Function

' Einstellungen je Ansichtstyp; liefert True bei Uebernahme
' isNew: zweiter Schritt von "Neue Ansicht", mit "Zurueck" (wentBack) zum ersten Schritt
Public Function EditViewSettings(ByVal vid As String, ByVal isNew As Boolean, Optional ByRef wentBack As Boolean) As Boolean
    Dim f As frmProps, tid As String, vt As String, dates As Variant, opts As Variant, nm As String, fid As String, efid As String, s As String
    Dim cardNames() As Variant, cardIds() As String, fld As cField, n As Long, sel As String, v As Variant, pct As Variant, sfid As String
    tid = mCfg.ViewGet(vid, "Tabelle")
    vt = mCfg.ViewGet(vid, "Typ")
    dates = FieldNamesWhere(tid, FT_DATUM & "," & FT_DATUMZEIT)
    opts = FieldNamesWhere(tid, FT_AUSWAHL & "," & FT_STATUS, "(keine)")
    Set f = New frmProps
    f.Init "Ansicht „" & mCfg.ViewGet(vid, "Name") & "“ (" & vt & ")", ""
    f.AddText "name", "Name", mCfg.ViewGet(vid, "Name")
    Select Case vt
        Case VT_TABLE
            f.AddCombo "seite", "Einträge je Seite", Array("50", "100", "250", "500"), mCfg.ViewOpt(vid, "seite", "100"), "Größere Seiten zeigen mehr auf einmal, bauen aber langsamer auf."
        Case VT_CAL
            If ArrLen(dates) = 0 Then mUI.Alert "Für einen Kalender braucht die Tabelle ein Datumsfeld. Bitte zuerst unter Einrichten ein Feld vom Typ Datum anlegen.", "Kalender": Unload f: Exit Function
            f.AddCombo "datum", "Datumsfeld (Termin bzw. Beginn)", dates, mCfg.FieldName(mCfg.ViewOpt(vid, "datum", mCfg.FieldIdByName(tid, CStr(dates(0)))))
            f.AddCombo "ende", "Enddatum (optional, für mehrtägige Einträge)", FieldNamesWhere(tid, FT_DATUM & "," & FT_DATUMZEIT, "(keins)"), _
                       IIf(Len(mCfg.ViewOpt(vid, "ende")) > 0, mCfg.FieldName(mCfg.ViewOpt(vid, "ende")), "(keins)")
            f.AddCombo "farbe", "Farbe nach (optional)", opts, IIf(Len(mCfg.ViewOpt(vid, "farbe")) > 0, mCfg.FieldName(mCfg.ViewOpt(vid, "farbe")), "(keine)")
        Case VT_BOARD
            opts = FieldNamesWhere(tid, FT_AUSWAHL & "," & FT_STATUS)
            If ArrLen(opts) = 0 Then mUI.Alert "Für ein Board braucht die Tabelle ein Feld vom Typ Status oder Einfachauswahl.", "Board": Unload f: Exit Function
            f.AddCombo "gruppe", "Spalten nach", opts, IIf(Len(mCfg.ViewOpt(vid, "gruppe")) > 0, mCfg.FieldName(mCfg.ViewOpt(vid, "gruppe")), CStr(opts(0)))
            ReDim cardNames(0 To mCfg.Fields(tid).Count): ReDim cardIds(0 To mCfg.Fields(tid).Count)
            For Each fld In mCfg.Fields(tid)
                If fld.ID <> mCfg.LabelField(tid) And fld.FType <> FT_LTEXT Then
                    cardNames(n) = fld.Name: cardIds(n) = fld.ID
                    If ListHas(mCfg.ViewOpt(vid, "karte"), fld.ID, ",") Then sel = sel & IIf(Len(sel) > 0, ";", "") & n
                    n = n + 1
                End If
            Next
            If n > 0 Then
                ReDim Preserve cardNames(0 To n - 1)
                f.AddMulti "karte", "Kurzangaben auf den Karten (bis zu 3 empfohlen)", cardNames, sel
            End If
        Case VT_GANTT
            If ArrLen(dates) < 2 Then mUI.Alert "Für die Projektleiste braucht die Tabelle zwei Datumsfelder (Beginn und Ende).", "Projektleiste": Unload f: Exit Function
            f.AddCombo "start", "Beginn", dates, IIf(Len(mCfg.ViewOpt(vid, "start")) > 0, mCfg.FieldName(mCfg.ViewOpt(vid, "start")), CStr(dates(0)))
            f.AddCombo "ende", "Ende", dates, IIf(Len(mCfg.ViewOpt(vid, "ende")) > 0, mCfg.FieldName(mCfg.ViewOpt(vid, "ende")), CStr(dates(1)))
            pct = FieldNamesWhere(tid, FT_PROZENT, "(keiner)")
            f.AddCombo "fortschritt", "Fortschritt (Prozentfeld, optional)", pct, IIf(Len(mCfg.ViewOpt(vid, "fortschritt")) > 0, mCfg.FieldName(mCfg.ViewOpt(vid, "fortschritt")), "(keiner)")
            f.AddCombo "farbe", "Farbe nach (optional)", opts, IIf(Len(mCfg.ViewOpt(vid, "farbe")) > 0, mCfg.FieldName(mCfg.ViewOpt(vid, "farbe")), "(keine)")
            f.AddCombo "massstab", "Maßstab beim Öffnen", Array("Tag", "Woche", "Monat"), mCfg.ViewOpt(vid, "massstab", "Woche")
            f.AddCheck "namen", "Namen der Einträge dezent auf den Leisten zeigen", mCfg.ViewOpt(vid, "namen", "Nein") = "Ja", _
                       "Passt ein Name nicht in die Leiste, steht er klein rechts daneben (bei Platzmangel gekürzt)."
    End Select
    f.AllowBack = isNew
    wentBack = False
    Do
        If Not f.ShowDialog() Then wentBack = f.WentBack: Unload f: Exit Function
        f.ClearErrors
        nm = Trim$(f.GetText("name"))
        If Len(nm) = 0 Then
            f.SetError "name", "Bitte einen Namen angeben."
        ElseIf vt = VT_GANTT And f.GetText("start") = f.GetText("ende") Then
            f.SetError "ende", "Beginn und Ende müssen verschiedene Felder sein."
        Else
            Exit Do
        End If
    Loop
    mCfg.SetV SH_VIEW, vid, "Name", nm
    Select Case vt
        Case VT_TABLE
            mCfg.ViewSetOpt vid, "seite", f.GetText("seite")
        Case VT_CAL
            fid = mCfg.FieldIdByName(tid, f.GetText("datum"))
            mCfg.ViewSetOpt vid, "datum", fid
            efid = mCfg.FieldIdByName(tid, f.GetText("ende"))
            mCfg.ViewSetOpt vid, "ende", efid
            mCfg.ViewSetOpt vid, "farbe", mCfg.FieldIdByName(tid, f.GetText("farbe"))
            If Len(efid) > 0 And efid <> fid Then EnsureNotBefore efid, fid
        Case VT_BOARD
            mCfg.ViewSetOpt vid, "gruppe", mCfg.FieldIdByName(tid, f.GetText("gruppe"))
            s = ""
            For Each v In Split(f.GetText("karte"), ";")
                If Len(v) > 0 Then s = s & IIf(Len(s) > 0, ",", "") & cardIds(CLng(v))
            Next
            mCfg.ViewSetOpt vid, "karte", s
        Case VT_GANTT
            sfid = mCfg.FieldIdByName(tid, f.GetText("start"))
            efid = mCfg.FieldIdByName(tid, f.GetText("ende"))
            mCfg.ViewSetOpt vid, "start", sfid
            mCfg.ViewSetOpt vid, "ende", efid
            mCfg.ViewSetOpt vid, "fortschritt", mCfg.FieldIdByName(tid, f.GetText("fortschritt"))
            mCfg.ViewSetOpt vid, "farbe", mCfg.FieldIdByName(tid, f.GetText("farbe"))
            mCfg.ViewSetOpt vid, "massstab", f.GetText("massstab")
            mCfg.ViewSetOpt vid, "namen", IIf(f.GetBool("namen"), "Ja", "Nein")
            EnsureNotBefore efid, sfid
    End Select
    Unload f
    mLog.Add IIf(isNew, "Ansicht angelegt", "Ansicht geändert"), tid, vid, ""
    EditViewSettings = True
End Function

' Regel "Ende nicht vor Beginn" sicherstellen (Beginn nach Ende ist ungueltig)
Private Sub EnsureNotBefore(ByVal endFid As String, ByVal startFid As String)
    If Len(endFid) = 0 Or Len(startFid) = 0 Then Exit Sub
    If mCfg.Field(endFid).NotBefore <> startFid Then mCfg.SetV SH_FLD, endFid, "NichtVor", startFid
End Sub

Public Sub Cmd_ToggleArchive()
    Dim st As cViewState
    On Error GoTo f
    Set st = mNav.CurState
    If st.ArchMode = mQuery.ARCH_ONLY Then st.ArchMode = mQuery.ARCH_ACTIVE Else st.ArchMode = mQuery.ARCH_ONLY
    st.Page = 1
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Archiv"
End Sub

' Zeilenumbruch in der Tabellenansicht an/aus (je Ansicht gespeichert; im Lesemodus nur für diese Sitzung)
Public Sub Cmd_ToggleWrap()
    Dim vid As String
    On Error GoTo f
    vid = mNav.gVid
    If Len(vid) = 0 Then Exit Sub
    mViewTable.SetWrap vid, Not mViewTable.WrapOn(vid)
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Zeilenumbruch"
End Sub

Public Sub Cmd_PagePrev()
    On Error GoTo f
    mNav.CurState.Page = mNav.CurState.Page - 1
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Seite"
End Sub

Public Sub Cmd_PageNext()
    On Error GoTo f
    mNav.CurState.Page = mNav.CurState.Page + 1
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Seite"
End Sub

'==============================================================================
' Eintraege
'==============================================================================
Private Function CurrentRid() As String
    Select Case ActiveSheet.Name
        Case SH_TABLE
            Dim c As Collection
            Set c = mViewTable.SelectedRids()
            If c.Count > 0 Then CurrentRid = c(1)
        Case SH_BOARD: CurrentRid = mViewBoard.SelectedRid()
        Case SH_GANTT: CurrentRid = mViewGantt.SelectedRid()
        Case SH_CAL
            CurrentRid = mViewCal.CellTarget(Selection.Cells(1, 1))
            If Left$(CurrentRid, 2) = "D:" Or Left$(CurrentRid, 2) = "M:" Then CurrentRid = ""
    End Select
End Function

Public Sub Cmd_NewRecord()
    On Error GoTo f
    mForms.NewRecord mNav.gTid
    Exit Sub
f: mApp.Boundary "Neuer Eintrag"
End Sub

Public Sub Cmd_EditSelected()
    Dim rid As String
    On Error GoTo f
    rid = CurrentRid()
    If Len(rid) = 0 Then mUI.Alert "Bitte zuerst einen Eintrag anklicken (Zeile bzw. Karte).", "Bearbeiten": Exit Sub
    mForms.OpenRecord mNav.gTid, rid
    Exit Sub
f: mApp.Boundary "Bearbeiten"
End Sub

Public Sub Cmd_ArchiveSelected()
    Dim rids As Collection
    On Error GoTo f
    Set rids = mViewTable.SelectedRids()
    If rids.Count = 0 Then mUI.Alert "Bitte zuerst einen oder mehrere Einträge markieren.", "Archivieren": Exit Sub
    If Not mApp.RequireWritable("Das Archivieren") Then Exit Sub
    If Not mUI.Confirm(Plural(rids.Count, "Eintrag", "Einträge") & " archivieren? Archivierte Einträge verschwinden aus Ansichten und Auswertungen, " & _
                       "bleiben aber mit allen Verknüpfungen erhalten und lassen sich über „Ansicht > Archiv anzeigen“ wiederherstellen.", "Archivieren", "Archivieren", "Abbrechen") Then Exit Sub
    mForms.ArchiveMany mNav.gTid, rids, True
    Exit Sub
f: mApp.Boundary "Archivieren"
End Sub

Public Sub Cmd_RestoreSelected()
    Dim rids As Collection
    On Error GoTo f
    Set rids = mViewTable.SelectedRids()
    If rids.Count = 0 Then mUI.Alert "Bitte zuerst einen oder mehrere archivierte Einträge markieren.", "Wiederherstellen": Exit Sub
    If Not mApp.RequireWritable("Das Wiederherstellen") Then Exit Sub
    mForms.ArchiveMany mNav.gTid, rids, False
    Exit Sub
f: mApp.Boundary "Wiederherstellen"
End Sub

Public Sub Cmd_DeleteSelected()
    Dim rids As Collection
    On Error GoTo f
    Set rids = mViewTable.SelectedRids()
    If rids.Count = 0 Then mUI.Alert "Bitte zuerst einen oder mehrere archivierte Einträge markieren.", "Endgültig löschen": Exit Sub
    If Not mApp.RequireWritable("Das Löschen") Then Exit Sub
    mForms.DeleteForever mNav.gTid, rids
    Exit Sub
f: mApp.Boundary "Endgültig löschen"
End Sub

Public Sub Cmd_Export()
    On Error GoTo f
    mRender.Popup Array("Als Excel-Datei (XLSX)…|mCmd.Cmd_ExportXlsx", "Als CSV-Datei (UTF-8, Semikolon)…|mCmd.Cmd_ExportCsv")
    Exit Sub
f: mApp.Boundary "Export"
End Sub

Public Sub Cmd_ExportXlsx()
    On Error GoTo f
    mIO.ExportFlow mNav.gVid
    Exit Sub
f: mApp.Boundary "Export"
End Sub

Public Sub Cmd_ExportCsv()
    On Error GoTo f
    mIO.ExportView mNav.gVid, "CSV"
    Exit Sub
f: mApp.Boundary "Export"
End Sub

'--- Sammelaenderung ------------------------------------------------------------------
Public Sub Cmd_BulkEdit()
    Dim rids As Collection, f As frmProps, names() As Variant, ids() As String, fld As cField, n As Long, fid As String, sel As Long
    Dim target As cField, newV As Variant, msg As String, mode As String, items() As Variant, i As Long, v As Variant, vals As cMap, errs As cMap
    Dim ek As Variant, report As New Collection, oldTxt As String, newTxt As String, cur As Variant, rv As cMap, ok As Boolean, linksVal As String
    On Error GoTo fl
    Set rids = mViewTable.SelectedRids()
    If rids.Count = 0 Then mUI.Alert "Bitte zuerst mehrere Zeilen markieren (Umschalt- oder Strg-Taste beim Klicken).", "Sammeländerung": Exit Sub
    If Not mApp.RequireWritable("Die Sammeländerung") Then Exit Sub
    ReDim names(0 To mCfg.Fields(mNav.gTid).Count): ReDim ids(0 To mCfg.Fields(mNav.gTid).Count)
    For Each fld In mCfg.Fields(mNav.gTid)
        If fld.IsEditable And fld.FType <> FT_BACKLINK Then names(n) = fld.Name: ids(n) = fld.ID: n = n + 1
    Next
    ReDim Preserve names(0 To n - 1)
    sel = mUI.ChooseOne("Welches Feld soll bei " & Plural(rids.Count, "markiertem Eintrag", "markierten Einträgen") & " geändert werden?", "Sammeländerung", names, "Weiter")
    If sel < 0 Then Exit Sub
    fid = ids(sel)
    Set target = mCfg.Field(fid)
    ' Wert erfassen
    Set f = New frmProps
    f.Init "Sammeländerung: " & target.Name, "Neuer Wert für " & Plural(rids.Count, "Eintrag", "Einträge") & ". Leer lassen = Wert entfernen. Vor der Übernahme erscheint eine Vorschau."
    Select Case target.FType
        Case FT_AUSWAHL, FT_STATUS
            ReDim items(0 To mCfg.OptionIds(fid, True).Count)
            items(0) = "(leer)": i = 1
            For Each v In mCfg.OptionIds(fid, True)
                items(i) = mCfg.OptName(CStr(v)): i = i + 1
            Next
            f.AddCombo "v", "Neuer Wert", items, "(leer)"
        Case FT_MEHRFACH
            f.AddCombo "mode", "Art der Änderung", Array("Optionen hinzufügen", "Optionen entfernen", "Auswahl ersetzen"), "Optionen hinzufügen"
            ReDim items(0 To mCfg.OptionIds(fid, True).Count - 1)
            i = 0
            For Each v In mCfg.OptionIds(fid, True)
                items(i) = mCfg.OptName(CStr(v)): i = i + 1
            Next
            f.AddMulti "v", "Optionen", items, ""
        Case FT_JANEIN
            f.AddCombo "v", "Neuer Wert", Array("Ja", "Nein", "(leer)"), "Ja"
        Case FT_LINK
            f.AddInfo "Die Verknüpfung wird im nächsten Schritt über die Auswahlliste festgelegt und ersetzt die bisherigen Zuordnungen."
        Case FT_DATUM, FT_DATUMZEIT
            f.AddDate "v", "Neues Datum", ""
        Case Else
            f.AddText "v", "Neuer Wert", "", mFmt.InputHint(target), target.FType = FT_LTEXT
    End Select
    Do
        If Not f.ShowDialog("Vorschau") Then Unload f: Exit Sub
        f.ClearErrors
        msg = ""
        Select Case target.FType
            Case FT_AUSWAHL, FT_STATUS
                If f.GetText("v") = "(leer)" Then newV = Empty Else newV = mCfg.OptIdByName(fid, f.GetText("v"))
            Case FT_MEHRFACH
                mode = f.GetText("mode")
                newV = ""
                For Each v In Split(f.GetText("v"), ";")
                    If Len(v) > 0 Then newV = ListAdd(CStr(newV), mCfg.OptIdByName(fid, CStr(items(CLng(v)))))
                Next
            Case FT_JANEIN
                If f.GetText("v") = "(leer)" Then newV = Empty Else newV = (f.GetText("v") = "Ja")
            Case FT_LINK
                newV = Empty
            Case Else
                If Not mValid.ParseInput(target, f.GetText("v"), newV, msg) Then f.SetError "v", msg
        End Select
        If Len(msg) = 0 Then Exit Do
    Loop
    Unload f
    If target.FType = FT_LINK Then
        linksVal = mForms.PickRecords(target.TargetTable, "", target.Multi, "Neue Verknüpfung für alle markierten Einträge", ok)
        If Not ok Then Exit Sub
    End If
    ' Vorschau und Pruefung
    ReDim items(0 To rids.Count - 1, 0 To 2)
    i = 0
    For Each v In rids
        Set vals = New cMap
        If target.FType = FT_LINK Then
            oldTxt = mFmt.DisplayText(target, Empty, CStr(v))
            newTxt = mFmt.LinkLabels(target, linksVal)
        Else
            cur = mData.GetValue(mNav.gTid, CStr(v), fid)
            If target.FType = FT_MEHRFACH Then
                Select Case mode
                    Case "Optionen hinzufügen"
                        cur = NzS(cur)
                        For Each ek In Split(CStr(newV), ";")
                            cur = ListAdd(CStr(cur), CStr(ek))
                        Next
                    Case "Optionen entfernen"
                        cur = NzS(cur)
                        For Each ek In Split(CStr(newV), ";")
                            cur = ListRemove(CStr(cur), CStr(ek))
                        Next
                    Case Else
                        cur = newV
                End Select
                If Len(NzS(cur)) = 0 Then cur = Empty
                vals.SetVal fid, cur
            Else
                vals.SetVal fid, newV
            End If
            oldTxt = mFmt.DisplayText(target, mData.GetValue(mNav.gTid, CStr(v), fid))
            newTxt = mFmt.DisplayText(target, vals.GetVal(fid))
            If Not mValid.ValidateRecord(mNav.gTid, CStr(v), vals, New cMap, errs) Then
                For Each ek In errs.Keys
                    report.Add Array(mData.RecordLabel(mNav.gTid, CStr(v)), mCfg.FieldName(CStr(ek)), CStr(errs.GetVal(CStr(ek))))
                Next
            End If
        End If
        items(i, 0) = mData.RecordLabel(mNav.gTid, CStr(v))
        items(i, 1) = oldTxt
        items(i, 2) = newTxt
        i = i + 1
    Next
    If report.Count > 0 Then
        Dim rep() As Variant
        ReDim rep(0 To report.Count - 1, 0 To 2)
        i = 0
        For Each v In report
            rep(i, 0) = v(0): rep(i, 1) = v(1): rep(i, 2) = v(2): i = i + 1
        Next
        mUI.ShowReport "Sammeländerung nicht möglich", "Für einige Einträge wäre der neue Wert ungültig. Es wurde nichts geändert.", rep, 3, "160;110;280", Array("OK")
        Exit Sub
    End If
    If mUI.ShowReport("Vorschau der Sammeländerung", Plural(rids.Count, "Eintrag wird", "Einträge werden") & " im Feld „" & target.Name & "“ geändert. Die Änderung lässt sich mit „Rückgängig“ zurücknehmen.", _
                      items, 3, "200;160;160", Array("Übernehmen", "Abbrechen")) <> 1 Then Exit Sub
    ' Ausfuehren
    mApp.BeginWork
    mData.TxBegin "Sammeländerung „" & target.Name & "“ (" & Plural(rids.Count, "Eintrag", "Einträge") & ")"
    For Each v In rids
        If target.FType = FT_LINK Then
            mData.SetLinks fid, CStr(v), linksVal
        Else
            Set vals = New cMap
            cur = mData.GetValue(mNav.gTid, CStr(v), fid)
            If target.FType = FT_MEHRFACH Then
                Select Case mode
                    Case "Optionen hinzufügen"
                        cur = NzS(cur)
                        For Each ek In Split(CStr(newV), ";"): cur = ListAdd(CStr(cur), CStr(ek)): Next
                    Case "Optionen entfernen"
                        cur = NzS(cur)
                        For Each ek In Split(CStr(newV), ";"): cur = ListRemove(CStr(cur), CStr(ek)): Next
                    Case Else
                        cur = newV
                End Select
                If Len(NzS(cur)) = 0 Then cur = Empty
                vals.SetVal fid, cur
            Else
                vals.SetVal fid, newV
            End If
            mData.UpdateRecord mNav.gTid, CStr(v), vals
        End If
        mLog.Add "Geändert", mNav.gTid, CStr(v), fid, "Sammeländerung"
    Next
    mData.TxCommit
    mApp.EndWork
    mNav.RenderCurrent
    Exit Sub
fl: mApp.Boundary "Sammeländerung"
End Sub

'==============================================================================
' Kalender, Board, Gantt
'==============================================================================
Public Sub Cmd_CalPrev()
    On Error GoTo f
    mNav.CurState.CalMonth = DateSerial(Year(mNav.CurState.CalMonth), Month(mNav.CurState.CalMonth) - 1, 1)
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Kalender"
End Sub

Public Sub Cmd_CalNext()
    On Error GoTo f
    mNav.CurState.CalMonth = DateSerial(Year(mNav.CurState.CalMonth), Month(mNav.CurState.CalMonth) + 1, 1)
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Kalender"
End Sub

Public Sub Cmd_CalToday()
    On Error GoTo f
    mNav.CurState.CalMonth = DateSerial(Year(Date), Month(Date), 1)
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Kalender"
End Sub

Public Sub Cmd_CalMode()
    On Error GoTo f
    mNav.CurState.CalMode = IIf(mNav.CurState.CalMode = "Liste", "Monat", "Liste")
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Kalender"
End Sub

Public Sub Cmd_CalUndated()
    On Error GoTo f
    mViewCal.ShowDay Empty
    Exit Sub
f: mApp.Boundary "Kalender"
End Sub

Public Sub Cmd_BoardLeft()
    On Error GoTo f
    mViewBoard.MoveSelected -1
    Exit Sub
f: mApp.Boundary "Board"
End Sub

Public Sub Cmd_BoardRight()
    On Error GoTo f
    mViewBoard.MoveSelected 1
    Exit Sub
f: mApp.Boundary "Board"
End Sub

Public Sub Cmd_BoardMove()
    On Error GoTo f
    mViewBoard.MoveSelected 0
    Exit Sub
f: mApp.Boundary "Board"
End Sub

Public Sub Cmd_BoardNewHere()
    On Error GoTo f
    mViewBoard.NewInColumn mRender.CallerName()
    Exit Sub
f: mApp.Boundary "Board"
End Sub

Public Sub Cmd_GanttPrev()
    On Error GoTo f
    mViewGantt.Shift -1
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Projektleiste"
End Sub

Public Sub Cmd_GanttNext()
    On Error GoTo f
    mViewGantt.Shift 1
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Projektleiste"
End Sub

Public Sub Cmd_GanttToday()
    Dim st As cViewState
    On Error GoTo f
    Set st = mNav.CurState
    st.GanttStart = st.DefaultGanttStart(st.GanttScale)
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Projektleiste"
End Sub

Public Sub Cmd_GanttScale()
    Dim st As cViewState
    On Error GoTo f
    Set st = mNav.CurState
    st.GanttScale = Mid$(mRender.CallerName(), 4)
    st.GanttStart = st.DefaultGanttStart(st.GanttScale)
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Projektleiste"
End Sub

' Namen dezent auf den Leisten zeigen (wird in der Ansicht gespeichert)
Public Sub Cmd_GanttNames()
    Dim vid As String
    On Error GoTo f
    If Not mApp.RequireWritable("Das Ändern der Ansicht") Then Exit Sub
    vid = mNav.gVid
    mApp.BeginWork
    mCfg.ViewSetOpt vid, "namen", IIf(mCfg.ViewOpt(vid, "namen", "Nein") = "Ja", "Nein", "Ja")
    mApp.EndWork
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Projektleiste"
End Sub

Public Sub Cmd_GanttBar()
    On Error GoTo f
    mViewGantt.OpenFromShape mRender.CallerName()
    Exit Sub
f: mApp.Boundary "Projektleiste"
End Sub

Public Sub Cmd_GanttIncomplete()
    On Error GoTo f
    mViewGantt.ShowIncomplete
    Exit Sub
f: mApp.Boundary "Projektleiste"
End Sub

'==============================================================================
' Start / Einrichten
'==============================================================================
Public Sub Cmd_OpenTableCard()
    Dim nm As String
    On Error GoTo f
    nm = mRender.CallerName()
    mNav.GoTable Mid$(nm, InStr(nm, "_") + 1)
    Exit Sub
f: mApp.Boundary "Tabelle öffnen"
End Sub

Public Sub Cmd_NewTable()
    Dim f As frmTableEditor, tid As String
    On Error GoTo fl
    If Not mApp.RequireSetup("Das Anlegen einer Tabelle") Then Exit Sub
    Set f = New frmTableEditor
    f.Setup ""
    f.Show: mNav.FocusGrid
    tid = f.TableId
    Unload f
    mStruct.RebuildLists
    mNav.RefreshNav
    If Len(tid) > 0 Then mNav.GoTable tid Else mNav.RenderCurrent
    Exit Sub
fl: mApp.Boundary "Neue Tabelle"
End Sub

Public Sub Cmd_FromTemplate()
    Dim lst As Variant, items() As Variant, i As Long, sel As Long, tid As String
    On Error GoTo f
    If Not mApp.RequireSetup("Das Anlegen aus einer Vorlage") Then Exit Sub
    lst = mStruct.TemplateList()
    If IsEmpty(lst) Then mUI.Alert "Es sind keine Vorlagen vorhanden.", "Vorlagen": Exit Sub
    ReDim items(0 To UBound(lst, 1), 0 To 1)
    For i = 0 To UBound(lst, 1)
        items(i, 0) = lst(i, 1): items(i, 1) = lst(i, 2)
    Next
    Dim fm As frmMsg
    Set fm = New frmMsg
    fm.SetupList "Aus Vorlage starten", "Vorlagen legen fertig eingerichtete Tabellen mit Feldern, Auswahlwerten, Berechnungen und Ansichten an, mit derselben Technik wie selbst angelegte Tabellen. " & _
                 "Alles lässt sich anschließend frei anpassen. Es werden keine Beispieldaten angelegt.", items, False, "Anlegen", 2, "180;300", 0
    fm.Show: mNav.FocusGrid
    sel = -1
    If fm.Result = 1 Then sel = fm.SelectedIndex
    Unload fm
    If sel < 0 Then Exit Sub
    mApp.BeginWork
    tid = mStruct.ApplyTemplate(CStr(lst(sel, 0)))
    mLog.Add "Vorlage angewendet", tid, "", "", CStr(lst(sel, 1))
    mApp.EndWork
    mNav.RefreshNav
    If Len(tid) > 0 Then mNav.GoTable tid
    Exit Sub
f: mApp.Boundary "Vorlage"
End Sub

Public Sub Cmd_Import()
    Dim f As frmImport, tid As String
    On Error GoTo fl
    If Not mApp.RequireWritable("Der Import") Then Exit Sub
    Set f = New frmImport
    f.Setup IIf(mNav.gPage = "Tabelle", mNav.gTid, "")
    f.Show: mNav.FocusGrid
    tid = f.ResultTid
    Unload f
    If Len(tid) > 0 Then
        mNav.RefreshNav
        mNav.GoTable tid
    End If
    Exit Sub
fl: mApp.Boundary "Import"
End Sub

' Vorhandene Liste (beliebige Excel-/CSV-Liste oder 2ask-Befragung) uebernehmen
Public Sub Cmd_ListImport()
    Dim f As frmListImport, tid As String, msg As String
    On Error GoTo fl
    If Not mApp.RequireWritable("Die Übernahme einer Liste") Then Exit Sub
    Set f = New frmListImport
    f.Setup
    f.Show: mNav.FocusGrid
    tid = f.ResultTid
    msg = f.ResultMsg
    Unload f
    mListImport.LiReset
    If Len(tid) > 0 Then
        mNav.RefreshNav
        mNav.GoTable tid
        ' Erfolgsmeldung erst nach dem Schließen des Assistenten und dem Seitenaufbau: Nach einer langen Übernahme
        ' (Windows hält Excel kurz für „reagiert nicht“) blieb die Meldung sonst unsichtbar hinter dem Fenster
        Application.ScreenUpdating = True
        DoEvents
        On Error Resume Next
        AppActivate Application.Caption
        On Error GoTo fl
        If Len(msg) > 0 Then mUI.Alert msg, "Liste übernehmen"
    End If
    Exit Sub
fl: mApp.Boundary "Liste übernehmen"
End Sub

Public Sub Cmd_ImportReport()
    Dim rows As Variant, items() As Variant, i As Long, j As Long
    On Error GoTo f
    rows = mIO.ErrorReportRows()
    If IsEmpty(rows) Then mUI.Alert "Es liegt kein Importbericht mit Fehlern vor.", "Importbericht": Exit Sub
    ReDim items(0 To UBound(rows, 1) - 1, 0 To 3)
    For i = 1 To UBound(rows, 1)
        For j = 1 To 4: items(i - 1, j - 1) = NzS(rows(i, j)): Next
    Next
    If mUI.ShowReport(mIO.ErrorReportTitle(), "Zeilen, die beim letzten Import nicht übernommen wurden.", items, 4, "70;140;150;260", Array("Als CSV speichern…", "Schließen")) = 1 Then
        mIO.ExportErrorReport
    End If
    Exit Sub
f: mApp.Boundary "Importbericht"
End Sub

Public Sub Cmd_Settings()
    Dim f As frmProps, accents As Variant, navs As Variant, i As Long, curA As String, curN As String
    On Error GoTo fl
    If Not mApp.RequireSetup("Das Ändern der Einstellungen") Then Exit Sub
    accents = AccentColorNames()
    navs = NavColorNames()
    curA = ColorName(CAccentRaw(), accents, True)
    curN = ColorName(CNavBg(), navs, False)
    Set f = New frmProps
    f.Init "Allgemeine Einstellungen", "Name und Farben dieser Datei. Die Akzentfarbe gilt für Schaltflächen und Markierungen; Statusfarben bleiben je Option einstellbar."
    f.AddText "app", "Anwendungsname (oben in der Navigation)", mCfg.SysGet("AppName", APP_TITLE)
    f.AddText "sub", "Untertitel / Organisation", mCfg.SysGet("Untertitel", "Offline-Datenbank in Excel")
    f.AddCombo "acc", "Akzentfarbe", accents, curA
    f.AddCombo "nav", "Navigationsfarbe", navs, curN
    If f.ShowDialog() Then
        mCfg.SysSet "AppName", Trim$(f.GetText("app"))
        mCfg.SysSet "Untertitel", Trim$(f.GetText("sub"))
        mCfg.SysSet "Akzentfarbe", CStr(AccentValue(f.GetText("acc")))
        mCfg.SysSet "Navigationsfarbe", CStr(NavValue(f.GetText("nav")))
        If Len(mDocStyle.CurrentDesign()) > 0 And (curA <> f.GetText("acc") Or curN <> f.GetText("nav")) Then mCfg.SysSet "Dokumentstil", "Eigene Auswahl"
        mApp.MarkCfgDirty
        mLog.Add "Einstellungen", "", "", "", "Allgemein"
        Unload f
        mNav.RenderCurrent
        mNav.RefreshNav True
    Else
        Unload f
    End If
    Exit Sub
fl: mApp.Boundary "Einstellungen"
End Sub

' Dokumentstil der ganzen Datei: Vorlage (Farben + Schrift) oder eigene Zusammenstellung
Public Sub Cmd_DocStyle()
    Dim pk As frmDesignPick, f As frmProps, choice As String, fonts As Variant
    On Error GoTo fl
    If Not mApp.RequireSetup("Das Ändern des Dokumentstils") Then Exit Sub
    Set pk = New frmDesignPick
    pk.Setup IIf(Len(mDocStyle.CurrentDesign()) = 0, "Eigene Auswahl", mDocStyle.CurrentDesign())
    pk.Show
    If Not pk.OK Then Unload pk: mNav.FocusGrid: Exit Sub
    choice = pk.Result
    Unload pk
    If choice = "Eigene Auswahl" Then
        Set f = New frmProps
        f.Init "Eigener Dokumentstil", "Farben der Oberfläche, Farbstil für Dashboards, Erfassungsdateien und Export sowie die Schrift der ganzen Datei."
        f.AddCombo "nav", "Navigationsfarbe", NavColorNames(), ColorName(CNavBg(), NavColorNames(), False)
        f.AddCombo "acc", "Akzentfarbe (Schaltflächen, Markierungen)", AccentColorNames(), ColorName(CAccentRaw(), AccentColorNames(), True)
        f.AddStyle "farbe", "Farbstil (Dashboards, Erfassungsdateien, Export)", mDocStyle.DefaultStyle()
        fonts = mDocStyle.DocFonts()
        f.AddCombo "schrift", "Schrift", fonts, mDocStyle.DefaultFont(), _
                   "Gilt für Oberfläche, Dashboards und erzeugte Dateien. Garamond, Georgia und Cambria sind Serifenschriften."
        If Not f.ShowDialog("Übernehmen") Then Unload f: mNav.FocusGrid: Exit Sub
        mDocStyle.SetDocStyle "Eigene Auswahl", f.GetText("nav"), f.GetText("acc"), f.GetText("farbe"), f.GetText("schrift")
        Unload f
    Else
        mDocStyle.ApplyDesign choice
    End If
    mApp.MarkCfgDirty
    mLog.Add "Einstellungen", "", "", "", "Dokumentstil: " & choice
    mDocStyle.ResetFontCache
    mNav.RefreshNav True
    mNav.RenderCurrent
    mNav.FocusGrid
    Exit Sub
fl: mApp.Boundary "Dokumentstil"
End Sub

' Titel/Untertitel der Datei (Klick oben links in der Navigation, Startseite, erster Start)
Public Sub EditTitle()
    Dim f As frmProps
    On Error GoTo fl
    If Not mApp.RequireWritable("Das Ändern des Titels") Then Exit Sub
    Set f = New frmProps
    f.Init "Titel der Datei", "Der Titel steht oben links in der Navigation und auf der Startseite, z. B. „Inventar Kitas“ oder „Projektliste Team Nord“."
    f.AddText "app", "Titel", mCfg.SysGet("AppName", APP_TITLE)
    f.AddText "sub", "Untertitel / Organisation (optional)", mCfg.SysGet("Untertitel", "Offline-Datenbank in Excel")
    Do
        If Not f.ShowDialog("Übernehmen") Then Unload f: mNav.FocusGrid: Exit Sub
        f.ClearErrors
        If Len(Trim$(f.GetText("app"))) > 0 Then Exit Do
        f.SetError "app", "Bitte einen Titel angeben."
    Loop
    mCfg.SysSet "AppName", Trim$(f.GetText("app"))
    mCfg.SysSet "Untertitel", Trim$(f.GetText("sub"))
    mCfg.SysSet "TitelFestgelegt", "Ja"
    Unload f
    mApp.MarkCfgDirty
    mLog.Add "Einstellungen", "", "", "", "Titel"
    mNav.RefreshNav
    mNav.RenderCurrent
    mNav.FocusGrid
    Exit Sub
fl: mApp.Boundary "Titel ändern"
End Sub

Public Sub Cmd_EditTitle()
    EditTitle
End Sub

' Farbfelder auf der Seite Einrichten: Name der Form "sw_nav_<Name>" bzw. "sw_acc_<Name>"
Public Sub Cmd_Swatch()
    Dim nm As String, p As Variant
    On Error GoTo fl
    If Not mApp.RequireWritable("Das Ändern der Farben") Then Exit Sub
    nm = mRender.CallerName()
    p = Split(nm, "_")
    If UBound(p) < 2 Then Exit Sub
    ' Ohne Schliessen/Neuoeffnen der Navigation (das flackerte): Navigation wird an Ort und
    ' Stelle neu gezeichnet; bei der Navigationsfarbe nur die Farbfelder der Seite
    If p(1) = "nav" Then
        If NavValue(CStr(p(2))) = CNavBg() Then Exit Sub
        mCfg.SysSet "Navigationsfarbe", CStr(NavValue(CStr(p(2))))
        If Len(mDocStyle.CurrentDesign()) > 0 Then mCfg.SysSet "Dokumentstil", "Eigene Auswahl"
        mApp.MarkCfgDirty
        mPages.RefreshSwatches
        mNav.RefreshNav True
    Else
        If AccentValue(CStr(p(2))) = CAccentRaw() Then Exit Sub
        mCfg.SysSet "Akzentfarbe", CStr(AccentValue(CStr(p(2))))
        If Len(mDocStyle.CurrentDesign()) > 0 Then mCfg.SysSet "Dokumentstil", "Eigene Auswahl"
        mApp.MarkCfgDirty
        mNav.RenderCurrent
        mNav.RefreshNav True
    End If
    Exit Sub
fl: mApp.Boundary "Farbe"
End Sub

' Farben: kraeftige Toene und ruhige helle Pastelltoene (dann mit dunkler Schrift)
Public Function NavColorNames() As Variant
    NavColorNames = Array("Schiefer", "Nachtblau", "Anthrazit", "Petrol", "Aubergine", "Waldgrün", "Bordeaux", "Ozean", "Mokka", "Stahl", _
                          "Flieder", "Salbei", "Himmel", "Sand", "Rosé", "Nebel", "Mainz hell")
End Function

Public Function AccentColorNames() As Variant
    AccentColorNames = Array("Blau", "Petrol", "Grün", "Violett", "Rot", "Orange", "Graphit", "Weinrot", "Terrakotta", _
                             "Flieder hell", "Mint", "Himmel hell", "Pfirsich", "Rosé hell", "Mainz Gelb")
End Function

Public Function AccentValue(ByVal nm As String) As Long
    Select Case nm
        Case "Petrol": AccentValue = RGB(13, 116, 128)
        Case "Grün": AccentValue = RGB(22, 128, 61)
        Case "Violett": AccentValue = RGB(109, 40, 217)
        Case "Rot": AccentValue = RGB(185, 28, 28)
        Case "Orange": AccentValue = RGB(194, 65, 12)
        Case "Graphit": AccentValue = RGB(55, 65, 81)
        Case "Weinrot": AccentValue = RGB(136, 42, 58)
        Case "Terrakotta": AccentValue = RGB(160, 82, 45)
        Case "Flieder hell": AccentValue = RGB(214, 202, 245)
        Case "Mint": AccentValue = RGB(190, 230, 207)
        Case "Himmel hell": AccentValue = RGB(191, 218, 245)
        Case "Pfirsich": AccentValue = RGB(250, 214, 190)
        Case "Rosé hell": AccentValue = RGB(245, 204, 214)
        Case "Mainz Gelb": AccentValue = RGB(235, 189, 0)       ' kräftiges Gelb (Design „Mainz“)
        Case Else: AccentValue = RGB(37, 99, 235)
    End Select
End Function

Public Function NavValue(ByVal nm As String) As Long
    Select Case nm
        Case "Ozean": NavValue = RGB(12, 74, 110)
        Case "Mokka": NavValue = RGB(68, 52, 44)
        Case "Stahl", "Hell": NavValue = RGB(71, 85, 105)
        Case "Nachtblau": NavValue = RGB(23, 37, 84)
        Case "Anthrazit": NavValue = RGB(38, 38, 38)
        Case "Petrol": NavValue = RGB(19, 78, 74)
        Case "Aubergine": NavValue = RGB(59, 7, 100)
        Case "Waldgrün": NavValue = RGB(20, 83, 45)
        Case "Bordeaux": NavValue = RGB(127, 29, 29)
        Case "Flieder": NavValue = RGB(234, 228, 246)
        Case "Salbei": NavValue = RGB(225, 237, 228)
        Case "Himmel": NavValue = RGB(225, 236, 247)
        Case "Sand": NavValue = RGB(243, 237, 226)
        Case "Rosé": NavValue = RGB(247, 232, 235)
        Case "Nebel": NavValue = RGB(236, 239, 243)
        Case "Mainz hell": NavValue = RGB(251, 242, 204)       ' Gelb, 20 % (Design „Mainz“)
        Case Else: NavValue = RGB(30, 41, 59)
    End Select
End Function

Private Function ColorName(ByVal c As Long, ByVal names As Variant, ByVal accent As Boolean) As String
    Dim i As Long
    For i = 0 To UBound(names)
        If accent Then
            If AccentValue(CStr(names(i))) = c Then ColorName = names(i): Exit Function
        Else
            If NavValue(CStr(names(i))) = c Then ColorName = names(i): Exit Function
        End If
    Next
    ColorName = names(0)
End Function

Public Sub Cmd_SetupProtection()
    Dim f As frmProps, p1 As String, p2 As String
    On Error GoTo fl
    If Not mApp.RequireSetup("Das Ändern des Einrichtungsschutzes") Then Exit Sub
    Set f = New frmProps
    f.Init "Einrichtungsschutz", "Mit einem Schutzwort fragt die Datei vor der ersten Strukturänderung einer Sitzung (Tabellen, Eigenschaften, Dashboards) nach dem Wort. Das schützt vor versehentlichen Änderungen. " & _
           "Es ist KEINE Zugriffskontrolle und keine Verschlüsselung: Wer die Datei öffnen kann, kann die Daten lesen. Leer lassen = Schutz entfernen."
    f.AddText "p1", "Schutzwort", "", "", False, True
    f.AddText "p2", "Schutzwort wiederholen", "", "", False, True
    Do
        If Not f.ShowDialog() Then Unload f: Exit Sub
        f.ClearErrors
        p1 = f.GetText("p1"): p2 = f.GetText("p2")
        If p1 = p2 Then Exit Do
        f.SetError "p2", "Die beiden Eingaben stimmen nicht überein."
    Loop
    Unload f
    mCfg.SysSet "Einrichtungsschutz", p1
    mLog.Add "Einstellungen", "", "", "", IIf(Len(p1) > 0, "Einrichtungsschutz gesetzt", "Einrichtungsschutz entfernt")
    mUI.Alert IIf(Len(p1) > 0, "Der Einrichtungsschutz ist gesetzt.", "Der Einrichtungsschutz wurde entfernt."), "Einrichtungsschutz"
    Exit Sub
fl: mApp.Boundary "Einrichtungsschutz"
End Sub

Public Sub Cmd_DataAccess()
    On Error GoTo f
    If mPages.DataSheetsVisible() Then
        mProtect.ShowDataSheets False
        mNav.RenderCurrent
        mUI.Alert "Die Datenblätter sind wieder ausgeblendet.", "Datenzugang"
    Else
        If Not mUI.Confirm("Die strukturierten Datenblätter (Tabellen, Konfiguration, Zuordnungen, Protokoll) werden schreibgeschützt eingeblendet und sind über die Blattregister erreichbar. " & _
                           "Bearbeitungen sind dort gesperrt, weil sie die Prüfregeln umgehen würden. Zurück zur Anwendung über die linke Navigation; ausblenden über dieselbe Schaltfläche.", _
                           "Datenzugang", "Einblenden", "Abbrechen") Then Exit Sub
        mProtect.ShowDataSheets True
        ActiveWindow.DisplayWorkbookTabs = True
        ThisWorkbook.Worksheets(SH_TAB).Activate
        ActiveWindow.DisplayHeadings = True
        ActiveWindow.DisplayGridlines = True
    End If
    Exit Sub
f: mApp.Boundary "Datenzugang"
End Sub

Public Sub Cmd_RecalcAll()
    On Error GoTo f
    If Not mApp.RequireWritable("Die Neuberechnung") Then Exit Sub
    mCalc.RecalcAll
    mApp.MarkDirty
    mNav.RenderCurrent
    mUI.Alert "Alle berechneten Felder wurden neu berechnet.", "Neu berechnen"
    Exit Sub
f: mApp.Boundary "Neu berechnen"
End Sub

Public Sub Cmd_Integrity()
    On Error GoTo f
    mUI.Alert mStart.IntegrityReport(), "Integritätsprüfung"
    Exit Sub
f: mApp.Boundary "Integritätsprüfung"
End Sub

Public Sub Cmd_SetupEditTable()
    Dim f As frmTableEditor, tid As String
    On Error GoTo fl
    tid = mPages.SetupSelected()
    If Len(tid) = 0 Or Not mCfg.TableExists(tid) Then mUI.Alert "Bitte zuerst eine Tabelle in der Liste anklicken.", "Tabelle bearbeiten": Exit Sub
    If Not mApp.RequireSetup("Das Bearbeiten der Felder") Then Exit Sub
    Set f = New frmTableEditor
    f.Setup tid
    f.Show: mNav.FocusGrid
    Unload f
    mStruct.RebuildLists
    mNav.RefreshNav
    mNav.RenderCurrent
    Exit Sub
fl: mApp.Boundary "Tabelle bearbeiten"
End Sub

Public Sub EditTableFromView()
    Dim f As frmTableEditor
    If Not mApp.RequireSetup("Das Bearbeiten der Felder") Then Exit Sub
    Set f = New frmTableEditor
    f.Setup mNav.gTid
    f.Show: mNav.FocusGrid
    Unload f
    mStruct.RebuildLists
    mNav.RefreshNav
    mNav.RenderCurrent
End Sub

Public Sub Cmd_SetupDuplicate()
    Dim tid As String, f As frmProps, nm As String, withData As Boolean, nt As String
    On Error GoTo fl
    tid = mPages.SetupSelected()
    If Len(tid) = 0 Then mUI.Alert "Bitte zuerst eine Tabelle in der Liste anklicken.", "Duplizieren": Exit Sub
    If Not mApp.RequireSetup("Das Duplizieren") Then Exit Sub
    Set f = New frmProps
    f.Init "Tabelle duplizieren", "Legt eine unabhängige Kopie mit neuen IDs an. Verknüpfungen zu anderen Tabellen werden als neue Verknüpfungsfelder angelegt; " & _
           "Verknüpfungen innerhalb der Tabelle werden auf die Kopien umgesetzt. Das Original bleibt unverändert."
    f.AddText "name", "Name der Kopie", mCfg.TableName(tid) & " (Kopie)"
    f.AddCheck "data", "Mit Daten (alle Einträge und Zuordnungen kopieren)", False
    Do
        If Not f.ShowDialog("Duplizieren") Then Unload f: Exit Sub
        f.ClearErrors
        nm = Trim$(f.GetText("name"))
        If Len(nm) = 0 Then
            f.SetError "name", "Bitte einen Namen angeben."
        ElseIf Len(mCfg.TableIdByName(nm)) > 0 Then
            f.SetError "name", "Diesen Tabellennamen gibt es bereits."
        Else
            Exit Do
        End If
    Loop
    withData = f.GetBool("data")
    Unload f
    mApp.BeginWork "Tabelle wird dupliziert …"
    nt = mStruct.DuplicateTable(tid, nm, withData)
    mApp.EndWork
    mNav.RefreshNav
    mNav.RenderCurrent
    mUI.Alert "Die Tabelle „" & nm & "“ wurde angelegt" & IIf(withData, " (mit Daten)", " (nur Struktur)") & ".", "Duplizieren"
    Exit Sub
fl: mApp.Boundary "Duplizieren"
End Sub

Public Sub Cmd_SetupHide()
    Dim tid As String
    On Error GoTo f
    tid = mPages.SetupSelected()
    If Len(tid) = 0 Then mUI.Alert "Bitte zuerst eine Tabelle in der Liste anklicken.", "Ausblenden": Exit Sub
    If Not mApp.RequireSetup("Das Ausblenden") Then Exit Sub
    mStruct.SetTableHidden tid, mCfg.GetS(SH_TAB, tid, "Status") <> "ausgeblendet"
    mNav.RefreshNav
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Ausblenden"
End Sub

Public Sub Cmd_SetupUp()
    On Error GoTo f
    If Len(mPages.SetupSelected()) = 0 Then Exit Sub
    mStruct.MoveTable mPages.SetupSelected(), -1
    mNav.RefreshNav
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Verschieben"
End Sub

Public Sub Cmd_SetupDown()
    On Error GoTo f
    If Len(mPages.SetupSelected()) = 0 Then Exit Sub
    mStruct.MoveTable mPages.SetupSelected(), 1
    mNav.RefreshNav
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Verschieben"
End Sub

Public Sub Cmd_SetupDelete()
    Dim tid As String
    On Error GoTo f
    tid = mPages.SetupSelected()
    If Len(tid) = 0 Then mUI.Alert "Bitte zuerst eine Tabelle in der Liste anklicken, dann „Tabelle löschen…“.", "Tabelle löschen": Exit Sub
    DeleteTableFlow tid
    Exit Sub
f: mApp.Boundary "Tabelle löschen"
End Sub

' Aus dem Ansicht-Menü einer Tabelle: die gerade gezeigte Tabelle löschen
Public Sub Cmd_TableDelete()
    On Error GoTo f
    If Len(mNav.gTid) = 0 Then Exit Sub
    DeleteTableFlow mNav.gTid
    Exit Sub
f: mApp.Boundary "Tabelle löschen"
End Sub

' Tabelle mit allen Einträgen löschen: Abhängigkeiten prüfen, Auswirkungen zeigen,
' Bestätigung per Eingabe, vorher automatische Sicherung. True = gelöscht
Public Function DeleteTableFlow(ByVal tid As String) As Boolean
    Dim deps As Collection, items() As Variant, i As Long, d As Variant, word As String, ok As Boolean, n As Long
    If Not mCfg.TableExists(tid) Then Exit Function
    If Not mApp.RequireSetup("Das Löschen einer Tabelle") Then Exit Function
    Set deps = mStruct.TableDependencies(tid)
    n = mData.TData(tid).RowCount
    ReDim items(0 To deps.Count, 0 To 1)
    items(0, 0) = "Daten": items(0, 1) = Plural(n, "Eintrag wird", "Einträge werden") & " dauerhaft entfernt (inkl. archivierter)."
    i = 1
    For Each d In deps
        items(i, 0) = d(0): items(i, 1) = d(1) & IIf(d(3) = 1, "  [blockiert]", "")
        i = i + 1
    Next
    If mStruct.HasBlocking(deps) Then
        mUI.ShowReport "Tabelle kann nicht gelöscht werden", "Die Tabelle wird noch verwendet. Bitte die mit [blockiert] markierten Abhängigkeiten zuerst auflösen. Alternative: Tabelle ausblenden (Einrichten > Aus-/Einblenden).", _
                       items, 2, "110;440", Array("OK")
        Exit Function
    End If
    If mUI.ShowReport("Tabelle „" & mCfg.TableName(tid) & "“ löschen", "Folgende Auswirkungen werden kontrolliert behandelt. Vorher wird automatisch eine Sicherung erstellt. " & _
                      "Tipp: Unter Einrichten blendet „Aus-/Einblenden“ die Tabelle aus, ohne Daten zu löschen.", items, 2, "110;440", Array("Weiter", "Abbrechen"), True) <> 1 Then Exit Function
    word = mUI.AskText("Zur Bestätigung bitte LÖSCHEN eingeben.", "Tabelle löschen", "", False, ok)
    If Not ok Then Exit Function
    If UCase$(Trim$(word)) <> "LÖSCHEN" And UCase$(Trim$(word)) <> "LOESCHEN" Then mUI.Alert "Die Bestätigung stimmt nicht. Es wurde nichts gelöscht.", "Tabelle löschen": Exit Function
    If Not mBackup.AutoBackup("vor dem Löschen einer Tabelle") Then Exit Function
    mApp.BeginWork
    mStruct.DeleteTable tid
    mApp.EndWork
    mStruct.RebuildLists
    DeleteTableFlow = True
    If mNav.gTid = tid Then
        mNav.gTid = "": mNav.gVid = ""
        If mNav.gPage = "Tabelle" Then mNav.gPage = "Start"
    End If
    mNav.RefreshNav
    mNav.RenderCurrent
    mUI.Alert "Die Tabelle wurde gelöscht. Eine Sicherung des vorherigen Stands liegt im Sicherungsordner (Hilfe & Sicherung > Wiederherstellen).", "Tabelle löschen"
End Function

'==============================================================================
' Eigenschaften direkt in der Ansicht (Spaltenkoepfe, + Eigenschaft)
'==============================================================================
' "Eigenschaften": oeffnet die Eigenschaftenliste der aktuellen Tabelle
Public Sub Cmd_EditTableHere()
    On Error GoTo fl
    If Not mApp.RequireWritable("Das Bearbeiten der Eigenschaften") Then Exit Sub
    EditTableFromView
    Exit Sub
fl: mApp.Boundary "Tabelle umbauen"
End Sub

' Neue Eigenschaft (Spalte) in der aktuellen Tabelle
Public Sub Cmd_AddProperty()
    On Error GoTo fl
    AddPropertyAfter ""
    Exit Sub
fl: mApp.Boundary "Eigenschaft hinzufügen"
End Sub

Private Sub AddPropertyAfter(ByVal afterFid As String)
    Dim f As frmField, res As String
    If Len(mNav.gTid) = 0 Then Exit Sub
    If Not mApp.RequireSetup("Das Hinzufügen einer Eigenschaft") Then Exit Sub
    Set f = New frmField
    f.Setup mNav.gTid, ""
    f.Show: mNav.FocusGrid
    res = f.ResultFid
    Unload f
    If Len(res) = 0 Then mNav.RenderCurrent: Exit Sub
    AfterNewProperty res, afterFid
End Sub

' Neue Eigenschaft in die Spaltenliste der aktuellen Ansicht aufnehmen und neu aufbauen
Public Sub AfterNewProperty(ByVal res As String, ByVal afterFid As String)
    Dim vid As String, spec As String, c As Variant, s As String, done As Boolean
    vid = mNav.gVid
    If Len(vid) > 0 Then
        spec = mCfg.ViewGet(vid, "Spalten")
        If Len(afterFid) > 0 Or Len(spec) > 0 Then
            ' Spaltenliste der Ansicht explizit fuehren und neue Spalte einfuegen
            For Each c In mViewTable.ViewColumns(vid)
                s = s & IIf(Len(s) > 0, ";", "") & c(0) & ":" & c(1)
                If c(0) = afterFid Then s = s & ";" & res: done = True
            Next
            If Not done Then s = s & IIf(Len(s) > 0, ";", "") & res
            mCfg.SetV SH_VIEW, vid, "Spalten", s
        End If
    End If
    mStruct.RebuildLists
    mNav.RefreshNav
    mNav.RenderCurrent
End Sub

Public Sub Cmd_HdrSortA(): HdrSort "A": End Sub
Public Sub Cmd_HdrSortD(): HdrSort "D": End Sub

Private Sub HdrSort(ByVal dirCode As String)
    Dim fid As String, st As cViewState
    On Error GoTo fl
    fid = mRender.MenuParam()
    If Not mCfg.FieldExists(fid) Then Exit Sub
    Set st = mNav.CurState
    st.SortSpec = fid & ":" & dirCode
    st.Page = 1
    mNav.RenderCurrent
    Exit Sub
fl: mApp.Boundary "Sortieren"
End Sub

Public Sub Cmd_HdrGroup()
    Dim fid As String, st As cViewState, f As cField
    On Error GoTo fl
    fid = mRender.MenuParam()
    If Not mCfg.FieldExists(fid) Then Exit Sub
    Set f = mCfg.Field(fid)
    Set st = mNav.CurState
    st.GroupFid = fid
    If f.FType = FT_DATUM Or f.FType = FT_DATUMZEIT Then st.GroupLevel = "Monat"
    st.Page = 1
    mNav.RenderCurrent
    Exit Sub
fl: mApp.Boundary "Gruppieren"
End Sub

' Blendet die Spalte nur in dieser Ansicht aus (Daten bleiben unberuehrt)
Public Sub Cmd_HdrHide()
    Dim fid As String, vid As String, c As Variant, s As String
    On Error GoTo fl
    fid = mRender.MenuParam()
    vid = mNav.gVid
    If Len(fid) = 0 Or Len(vid) = 0 Then Exit Sub
    For Each c In mViewTable.ViewColumns(vid)
        If c(0) <> fid Then s = s & IIf(Len(s) > 0, ";", "") & c(0) & ":" & c(1)
    Next
    mCfg.SetV SH_VIEW, vid, "Spalten", s
    mLog.Add "Ansicht geändert", mNav.gTid, vid, fid, "Spalte ausgeblendet"
    mNav.RenderCurrent
    Exit Sub
fl: mApp.Boundary "Spalte ausblenden"
End Sub

Public Sub Cmd_HdrEdit()
    Dim fid As String, f As frmField
    On Error GoTo fl
    fid = mRender.MenuParam()
    If Not mCfg.FieldExists(fid) Then Exit Sub
    If Not mApp.RequireSetup("Das Bearbeiten einer Eigenschaft") Then Exit Sub
    Set f = New frmField
    f.Setup mNav.gTid, fid
    f.Show: mNav.FocusGrid
    Unload f
    mStruct.RebuildLists
    mNav.RefreshNav
    mNav.RenderCurrent
    Exit Sub
fl: mApp.Boundary "Eigenschaft bearbeiten"
    mNav.RenderCurrent
End Sub

' Formelspalten: schwebende Formelleiste (Spalten der Tabelle anklicken)
Public Sub Cmd_NewFormulaColumn()
    On Error GoTo fl
    If Len(mNav.gTid) = 0 Then Exit Sub
    If Not mApp.RequireSetup("Das Hinzufügen einer Formelspalte") Then Exit Sub
    mForms.OpenXlBar mNav.gTid, "", ""
    Exit Sub
fl: mApp.Boundary "Formelspalte"
End Sub

Public Sub Cmd_HdrNewFormula()
    On Error GoTo fl
    If Not mApp.RequireSetup("Das Hinzufügen einer Formelspalte") Then Exit Sub
    mForms.OpenXlBar mNav.gTid, "", mRender.MenuParam()
    Exit Sub
fl: mApp.Boundary "Formelspalte"
End Sub

Public Sub Cmd_HdrFormula()
    Dim fid As String
    On Error GoTo fl
    fid = mRender.MenuParam()
    If Not mCfg.FieldExists(fid) Then Exit Sub
    If Not mApp.RequireSetup("Das Bearbeiten einer Formel") Then Exit Sub
    If Len(mCfg.Field(fid).Formula) > 0 And Not mCalc.IsXl(mCfg.Field(fid).Formula) Then
        If Not mUI.Confirm("„" & mCfg.FieldName(fid) & "“ wurde mit dem Assistenten berechnet: " & mCalc.Describe(mCfg.Field(fid).Formula) & vbLf & vbLf & _
                           "Stattdessen eine Excel-Formel schreiben? Die bisherige Berechnung wird beim Übernehmen ersetzt.", "Formel", "Formel schreiben", "Abbrechen") Then Exit Sub
    End If
    mForms.OpenXlBar mNav.gTid, fid, ""
    Exit Sub
fl: mApp.Boundary "Formel bearbeiten"
End Sub

Public Sub Cmd_HdrInsert()
    Dim fid As String
    On Error GoTo fl
    fid = mRender.MenuParam()
    AddPropertyAfter fid
    Exit Sub
fl: mApp.Boundary "Eigenschaft hinzufügen"
    mNav.RenderCurrent
End Sub

Public Sub Cmd_HdrDelete()
    Dim fid As String
    On Error GoTo fl
    fid = mRender.MenuParam()
    If Not mCfg.FieldExists(fid) Then Exit Sub
    If fid = mCfg.LabelField(mNav.gTid) Then mUI.Alert "Das Bezeichnungsfeld kann nicht entfernt werden. Legen Sie zuerst unter „Eigenschaften“ ein anderes Bezeichnungsfeld fest.", "Eigenschaft entfernen": Exit Sub
    If Not mApp.RequireSetup("Das Entfernen einer Eigenschaft") Then Exit Sub
    DeleteFieldFlow fid
    mStruct.RebuildLists
    mNav.RefreshNav
    mNav.RenderCurrent
    Exit Sub
fl: mApp.Boundary "Eigenschaft entfernen"
    mNav.RenderCurrent
End Sub
' Feld entfernen mit Abhaengigkeitspruefung (aus dem Tabelleneditor)
Public Sub DeleteFieldFlow(ByVal fid As String)
    Dim deps As Collection, items() As Variant, i As Long, d As Variant, f As cField
    Set f = mCfg.Field(fid)
    Set deps = mStruct.FieldDependencies(fid)
    ReDim items(0 To deps.Count, 0 To 1)
    items(0, 0) = "Daten": items(0, 1) = "Alle Werte des Feldes „" & f.Name & "“ werden entfernt."
    i = 1
    For Each d In deps
        items(i, 0) = d(0): items(i, 1) = d(1) & IIf(d(3) = 1, "  [blockiert]", "")
        i = i + 1
    Next
    If mStruct.HasBlocking(deps) Then
        mUI.ShowReport "Feld kann nicht entfernt werden", "Das Feld wird noch verwendet. Bitte die mit [blockiert] markierten Abhängigkeiten zuerst anpassen. " & _
                       "Alternative ohne Datenverlust: Feld ausblenden.", items, 2, "110;440", Array("OK")
        Exit Sub
    End If
    If mUI.ShowReport("Feld „" & f.Name & "“ entfernen", "Entfernen löscht die Werte; Ausblenden behält sie. Folgende Auswirkungen werden kontrolliert behandelt; vorher wird automatisch gesichert.", _
                      items, 2, "110;440", Array("Entfernen", "Abbrechen"), True) <> 1 Then Exit Sub
    If Not mBackup.AutoBackup("vor dem Entfernen eines Feldes") Then Exit Sub
    mApp.BeginWork
    mStruct.DeleteField fid
    mCalc.RecalcAll
    mApp.EndWork
End Sub

' Feldtyp wechseln mit Konfliktvorschau; liefert neuen Typ oder ""
Public Function ChangeTypeFlow(ByVal fid As String) As String
    Dim f As cField, codes As Variant, names() As Variant, ids() As String, n As Long, c As Variant, sel As Long, toT As String
    Dim conflicts As Collection, newVals As Variant, createOpts As Collection, items() As Variant, i As Long, k As Variant, ans As Long, t As cTableData, r As Long
    Set f = mCfg.Field(fid)
    codes = FieldTypeCodes()
    ReDim names(0 To UBound(codes)): ReDim ids(0 To UBound(codes))
    For Each c In codes
        If mStruct.CanConvert(f.FType, CStr(c)) Then names(n) = FieldTypeName(CStr(c)): ids(n) = CStr(c): n = n + 1
    Next
    If n = 0 Then mUI.Alert "Für diesen Feldtyp ist kein Wechsel möglich (Verknüpfungen und Berechnungen bitte neu anlegen).", "Typ ändern": Exit Function
    ReDim Preserve names(0 To n - 1)
    sel = mUI.ChooseOne("Neuer Typ für „" & f.Name & "“. Vor der Umwandlung werden alle vorhandenen Werte geprüft.", "Feldtyp ändern", names, "Prüfen")
    If sel < 0 Then Exit Function
    toT = ids(sel)
    Set conflicts = mStruct.ConversionPreview(fid, toT, newVals, createOpts)
    If conflicts.Count = 0 Then
        If Not mUI.Confirm("Alle " & FmtInt(mData.TData(f.TableID).RowCount) & " Werte lassen sich umwandeln" & _
                           IIf(createOpts.Count > 0, "; " & createOpts.Count & " neue Auswahlmöglichkeiten werden angelegt", "") & ". Jetzt umwandeln? (Vorher wird gesichert.)", _
                           "Feldtyp ändern", "Umwandeln", "Abbrechen") Then Exit Function
    Else
        ReDim items(0 To conflicts.Count - 1, 0 To 2)
        For Each k In conflicts
            items(i, 0) = k(0): items(i, 1) = k(1): items(i, 2) = k(2): i = i + 1
        Next
        ans = mUI.ShowReport("Konflikte bei der Umwandlung", conflicts.Count & " Werte passen nicht zum neuen Typ „" & FieldTypeName(toT) & "“. " & _
                             "Empfohlen: abbrechen und die Werte zuerst korrigieren. Alternativ können diese Werte ausdrücklich geleert werden (nach Sicherung).", _
                             items, 3, "180;150;240", Array("Abbrechen", "Konfliktwerte leeren und umwandeln"), True)
        If ans <> 2 Then Exit Function
        If Not mUI.Confirm(conflicts.Count & " Werte werden geleert. Wirklich fortfahren?", "Konfliktwerte leeren", "Leeren und umwandeln", "Abbrechen", True) Then Exit Function
        ' Konfliktzeilen sind in newVals leer geblieben
    End If
    If Not mBackup.AutoBackup("vor dem Feldtypwechsel") Then Exit Function
    mApp.BeginWork
    mStruct.ApplyTypeChange fid, toT, newVals, createOpts
    mCalc.RecalcAll
    mApp.EndWork
    ChangeTypeFlow = toT
End Function

'==============================================================================
' Hilfe, Sicherung, Technik
'==============================================================================
Public Sub Cmd_BackupNow()
    On Error GoTo f
    mBackup.BackupNow "manuell"
    Exit Sub
f: mApp.Boundary "Sicherung"
End Sub

Public Sub Cmd_Restore()
    On Error GoTo f
    mBackup.RestoreFlow
    Exit Sub
f: mApp.Boundary "Wiederherstellen"
End Sub

Public Sub Cmd_OpenBackupFolder()
    Dim p As String
    On Error GoTo f
    p = mBackup.BackupFolder()
    If Len(p) = 0 Or Not DirExists(p) Then mUI.Alert "Es gibt noch keinen Sicherungsordner. Er entsteht mit der ersten Sicherung neben der Datei:" & vbLf & p, "Sicherungsordner": Exit Sub
    ThisWorkbook.FollowHyperlink Address:=p, NewWindow:=True
    Exit Sub
f: mApp.Boundary "Sicherungsordner"
End Sub

Public Sub Cmd_BackupSettings()
    Dim f As frmProps, n As Double, m As Double
    On Error GoTo fl
    If Not mApp.RequireWritable("Das Ändern der Sicherungseinstellungen") Then Exit Sub
    Set f = New frmProps
    f.Init "Sicherungseinstellungen", "Sicherungen liegen im Ordner „" & Mid$(mBackup.BackupFolder(), InStrRev(mBackup.BackupFolder(), "\") + 1) & "“ neben der Datei."
    f.AddCheck "auto", "Beim Speichern automatisch sichern", mCfg.SysGet("SicherungBeimSpeichern", "Ja") = "Ja"
    f.AddText "min", "Mindestabstand automatischer Sicherungen (Minuten)", CStr(mCfg.SysLong("SicherungAbstandMinuten", 60))
    f.AddText "keep", "Anzahl aufbewahrter Sicherungen", CStr(mCfg.SysLong("SicherungenBehalten", 10))
    Do
        If Not f.ShowDialog() Then Unload f: Exit Sub
        f.ClearErrors
        If Not ParseNumberDE(f.GetText("min"), m) Or m < 0 Or m > 10080 Then
            f.SetError "min", "Bitte eine Zahl von 0 bis 10080 angeben."
        ElseIf Not ParseNumberDE(f.GetText("keep"), n) Or n < 1 Or n > 200 Then
            f.SetError "keep", "Bitte eine Zahl von 1 bis 200 angeben."
        Else
            Exit Do
        End If
    Loop
    mCfg.SysSet "SicherungBeimSpeichern", IIf(f.GetBool("auto"), "Ja", "Nein")
    mCfg.SysSet "SicherungAbstandMinuten", CStr(CLng(m))
    mCfg.SysSet "SicherungenBehalten", CStr(CLng(n))
    mApp.MarkCfgDirty
    Unload f
    Exit Sub
fl: mApp.Boundary "Sicherungseinstellungen"
End Sub

Public Sub Cmd_ShowLog()
    On Error GoTo f
    mPages.ShowLog
    Exit Sub
f: mApp.Boundary "Protokoll"
End Sub

Public Sub Cmd_NewEmptyCopy()
    On Error GoTo f
    mBackup.NewEmptyCopy
    Exit Sub
f: mApp.Boundary "Leere Kopie"
End Sub

Public Sub Cmd_ResetData()
    On Error GoTo f
    If Not mApp.RequireWritable("Das Zurücksetzen") Then Exit Sub
    mBackup.ResetData
    Exit Sub
f: mApp.Boundary "Datei zurücksetzen"
End Sub

Public Sub Cmd_ShowCode()
    On Error GoTo f
    mCodeCat.ShowCode
    Exit Sub
f: mApp.Boundary "Quellcode"
End Sub

Public Sub Cmd_CodeZip(): On Error GoTo f: mCodeCat.ExportAs "ZIP": Exit Sub
f: mApp.Boundary "Codeexport"
End Sub

Public Sub Cmd_CodeHtml(): On Error GoTo f: mCodeCat.ExportAs "HTML": Exit Sub
f: mApp.Boundary "Codeexport"
End Sub

Public Sub Cmd_CodeMd(): On Error GoTo f: mCodeCat.ExportAs "MD": Exit Sub
f: mApp.Boundary "Codeexport"
End Sub

Public Sub Cmd_CodeTxt(): On Error GoTo f: mCodeCat.ExportAs "TXT": Exit Sub
f: mApp.Boundary "Codeexport"
End Sub

'==============================================================================
' Dashboards
'==============================================================================
Public Sub Cmd_DashTab()
    Dim nm As String
    On Error GoTo f
    nm = mRender.CallerName()
    mDash.EndArrange
    mNav.GoDash Mid$(nm, 6)
    Exit Sub
f: mApp.Boundary "Dashboard"
End Sub

Public Sub Cmd_DashArrange()
    On Error GoTo f
    If Not mApp.RequireSetup("Das Anordnen eines Dashboards") Then Exit Sub
    mDash.BeginArrange
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Anordnen"
End Sub

' Gezogene Karten sofort einrasten (bleibt im Anordnen-Modus)
Public Sub Cmd_DashArrangeSnap()
    On Error GoTo f
    mDash.SnapArrangement
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Anordnen"
End Sub

Public Sub Cmd_DashArrangeDone()
    On Error GoTo f
    mDash.ApplyArrangement
    mDash.EndArrange
    mNav.RenderCurrent
    Exit Sub
f: mDash.EndArrange
    mApp.Boundary "Anordnen"
End Sub

Public Sub Cmd_DashArrangeCancel()
    On Error GoTo f
    mDash.EndArrange
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Anordnen"
End Sub

Public Sub Cmd_DashMore()
    On Error GoTo f
    mRender.Popup Array("Dashboard umbenennen / löschen…|mCmd.Cmd_DashEdit", "Export Excel…|mCmd.Cmd_DashExportXlsx||1|1", "Export PDF…|mCmd.Cmd_DashExportPdf")
    Exit Sub
f: mApp.Boundary "Dashboard"
End Sub

' Live-Vorschau aus dem Gestaltungsdialog (frmProps.PreviewMacro)
Public Sub DashStylePreview(ByVal f As Object)
    On Error Resume Next
    mCfg.SysSet "DashStil_" & mNav.gDash, f.GetText("stil")
    mCfg.SysSet "DashPalette_" & mNav.gDash, f.GetText("pal")
    mNav.RenderCurrent
    Application.ScreenUpdating = True
End Sub

' Stil und Farbpalette: Auswahl wird sofort im Dashboard dahinter gezeigt (Vorschau)
Public Sub Cmd_DashStyle()
    Dim did As String, oldS As String, oldP As String, f As frmProps, ok As Boolean
    On Error GoTo fl
    did = mNav.gDash
    If Len(did) = 0 Then Exit Sub
    If Not mApp.RequireSetup("Die Gestaltung eines Dashboards") Then Exit Sub
    oldS = mDash.DashStyle(did): oldP = mDash.DashPaletteName(did)
    Set f = New frmProps
    f.Init "Gestaltung", "Stil der Karten und Farbpalette der Diagramme. Jede Auswahl wird sofort im Dashboard dahinter gezeigt. " & _
           "Einzelne Elemente können eine eigene Palette haben (Element bearbeiten).", 420
    f.AddCombo "stil", "Stil", mDash.StyleNames(), oldS
    f.AddStyle "pal", "Farbstil der Diagramme", oldP
    f.PreviewMacro = "mCmd.DashStylePreview"
    ok = f.ShowDialog("Übernehmen")
    If ok Then
        mCfg.SysSet "DashStil_" & did, f.GetText("stil")
        mCfg.SysSet "DashPalette_" & did, f.GetText("pal")
    Else
        mCfg.SysSet "DashStil_" & did, oldS: mCfg.SysSet "DashPalette_" & did, oldP
    End If
    Unload f
    mNav.RenderCurrent
    mNav.FocusGrid
    If Not ok Then Exit Sub
    mApp.MarkCfgDirty
    mLog.Add "Dashboard gestaltet", "", did, "", mDash.DashStyle(did) & " / " & mDash.DashPaletteName(did)
    Exit Sub
fl: mApp.Boundary "Gestaltung"
End Sub

Public Sub Cmd_DashNew()
    Dim f As frmProps, did As String
    On Error GoTo fl
    If Not mApp.RequireSetup("Das Anlegen eines Dashboards") Then Exit Sub
    Set f = New frmProps
    f.Init "Neues Dashboard", "Ein Dashboard bündelt Kennzahlen, Diagramme und Ergebnistabellen."
    f.AddText "name", "Name", "Übersicht"
    f.AddText "desc", "Beschreibung (optional)", ""
    If Not f.ShowDialog("Anlegen") Then Unload f: Exit Sub
    did = mDash.CreateDashboard(IIf(Len(Trim$(f.GetText("name"))) > 0, Trim$(f.GetText("name")), "Dashboard"), f.GetText("desc"))
    Unload f
    mNav.GoDash did
    Exit Sub
fl: mApp.Boundary "Neues Dashboard"
End Sub

Public Sub Cmd_DashEdit()
    Dim f As frmProps, did As String
    On Error GoTo fl
    did = mNav.gDash
    If Len(did) = 0 Then Exit Sub
    If Not mApp.RequireSetup("Das Bearbeiten eines Dashboards") Then Exit Sub
    Set f = New frmProps
    f.Init "Dashboard bearbeiten", ""
    f.AddText "name", "Name", mCfg.GetS(SH_DSB, did, "Name")
    f.AddText "desc", "Beschreibung", mCfg.GetS(SH_DSB, did, "Beschreibung")
    f.AddCheck "del", "Dieses Dashboard löschen (die Daten bleiben unverändert)", False
    If Not f.ShowDialog() Then Unload f: Exit Sub
    If f.GetBool("del") Then
        Unload f
        If mUI.Confirm("Dashboard „" & mCfg.GetS(SH_DSB, did, "Name") & "“ mit allen Elementen löschen?", "Dashboard löschen", "Löschen", "Abbrechen", True) Then
            mDash.DeleteDashboard did
            mNav.gDash = ""
        End If
    Else
        mCfg.SetV SH_DSB, did, "Name", Trim$(f.GetText("name"))
        mCfg.SetV SH_DSB, did, "Beschreibung", f.GetText("desc")
        Unload f
    End If
    mNav.GoDash mNav.gDash
    Exit Sub
fl: mApp.Boundary "Dashboard bearbeiten"
End Sub

Public Sub Cmd_DashRefresh()
    On Error GoTo f
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Aktualisieren"
End Sub

Public Sub Cmd_DashExportXlsx()
    On Error GoTo f
    mIO.ExportDashboard mNav.gDash, "XLSX"
    Exit Sub
f: mApp.Boundary "Export"
End Sub

Public Sub Cmd_DashExportPdf()
    On Error GoTo f
    mIO.ExportDashboard mNav.gDash, "PDF"
    Exit Sub
f: mApp.Boundary "Export"
End Sub

Public Sub Cmd_DashFilterClear()
    On Error GoTo f
    mCfg.SetV SH_DSB, mNav.gDash, "Filter", ""
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Dashboardfilter"
End Sub

Public Sub Cmd_DashFilter()
    Dim f As frmProps, did As String, cur As String, tnames() As Variant, ids As Variant, i As Long, d1 As Date, d2 As Date, ht As Boolean
    Dim von As String, bis As String, tid As String, s As String, ff As frmFilter, msg As String
    On Error GoTo fl
    did = mNav.gDash
    cur = mCfg.GetS(SH_DSB, did, "Filter")
    ids = mCfg.TableIds()
    ReDim tnames(0 To ArrLen(ids))
    tnames(0) = "(kein Tabellenfilter)"
    For i = 0 To ArrLen(ids) - 1: tnames(i + 1) = mCfg.TableName(CStr(ids(i))): Next
    Set f = New frmProps
    f.Init "Dashboardfilter", "Der Zeitraum wirkt auf Elemente, für die ein Zeitraumfeld festgelegt ist. Der Tabellenfilter wirkt nur auf Elemente dieser Tabelle. " & _
           "Nicht betroffene Elemente sind gekennzeichnet. Der Filter wird mit dem Dashboard gespeichert."
    f.AddDate "von", "Zeitraum von", IIf(Len(KVGet(cur, "von")) > 0, FmtDate(InvDate(KVGet(cur, "von"))), "")
    f.AddDate "bis", "Zeitraum bis", IIf(Len(KVGet(cur, "bis")) > 0, FmtDate(InvDate(KVGet(cur, "bis"))), "")
    f.AddCombo "tab", "Tabellenfilter für", tnames, IIf(mCfg.TableExists(KVGet(cur, "tab")), mCfg.TableName(KVGet(cur, "tab")), "(kein Tabellenfilter)")
    Do
        If Not f.ShowDialog("Weiter") Then Unload f: Exit Sub
        f.ClearErrors
        von = "": bis = "": msg = ""
        If Len(Trim$(f.GetText("von"))) > 0 Then
            If ParseDateDE(f.GetText("von"), d1, ht) Then von = DateInv(Int(d1)) Else f.SetError "von", "Bitte ein Datum TT.MM.JJJJ angeben.": msg = "x"
        End If
        If Len(Trim$(f.GetText("bis"))) > 0 And Len(msg) = 0 Then
            If ParseDateDE(f.GetText("bis"), d2, ht) Then bis = DateInv(Int(d2)) Else f.SetError "bis", "Bitte ein Datum TT.MM.JJJJ angeben.": msg = "x"
        End If
        If Len(msg) = 0 And Len(von) > 0 And Len(bis) > 0 Then
            If d2 < d1 Then f.SetError "bis", "Das Ende liegt vor dem Beginn.": msg = "x"
        End If
        If Len(msg) = 0 Then Exit Do
    Loop
    tid = mCfg.TableIdByName(f.GetText("tab"))
    Unload f
    s = KVSet("", "von", von)
    s = KVSet(s, "bis", bis)
    If Len(tid) > 0 Then
        Set ff = New frmFilter
        ff.Setup tid, IIf(KVGet(cur, "tab") = tid, KVGet(cur, "f"), ""), "Tabellenfilter für „" & mCfg.TableName(tid) & "“"
        ff.Show
        If ff.OK Then
            s = KVSet(s, "tab", tid)
            s = KVSet(s, "f", ff.Result)
        End If
        Unload ff
    End If
    mCfg.SetV SH_DSB, did, "Filter", s
    mNav.RenderCurrent
    Exit Sub
fl: mApp.Boundary "Dashboardfilter"
End Sub

Public Sub Cmd_DashAddElement()
    On Error GoTo f
    If Len(mNav.gDash) = 0 Then Cmd_DashNew: Exit Sub
    If Not mApp.RequireSetup("Das Hinzufügen eines Dashboard-Elements") Then Exit Sub
    ElementFlow mNav.gDash, ""
    mNav.RenderCurrent
    Exit Sub
f: mDash.gHighlight = ""
    mApp.Boundary "Element hinzufügen"
End Sub

Public Sub Cmd_DashElementMenu()
    Dim eid As String
    On Error GoTo f
    eid = mDash.ElementFromShape(mRender.CallerName())
    If Len(eid) = 0 Then Exit Sub
    mRender.Popup Array("Datensätze anzeigen…|mCmd.Cmd_ElemDetails|" & eid, CopyCaption(eid) & "|mCmd.Cmd_ElemCopy|" & eid, _
                        "Bearbeiten…|mCmd.Cmd_ElemEdit|" & eid & "|1|1", _
                        "Nach vorne|mCmd.Cmd_ElemUp|" & eid, "Nach hinten|mCmd.Cmd_ElemDown|" & eid, "Entfernen…|mCmd.Cmd_ElemDelete|" & eid & "|1|1")
    Exit Sub
f: mApp.Boundary "Element"
End Sub

Private Function CopyCaption(ByVal eid As String) As String
    Select Case mDash.ElemGet(eid, "Typ")
        Case "Kachel": CopyCaption = "Kachel kopieren (für PowerPoint, Word)"
        Case "Tabelle": CopyCaption = "Tabelle kopieren (für PowerPoint, Word)"
        Case Else: CopyCaption = "Diagramm kopieren (für PowerPoint, Word)"
    End Select
End Function

' Dashboard-Element als Bild kopieren
Public Sub Cmd_ElemCopy()
    Dim eid As String, what As String
    On Error GoTo f
    eid = mRender.MenuParam()
    If Len(eid) = 0 Then Exit Sub
    Select Case mDash.ElemGet(eid, "Typ")
        Case "Kachel": what = "Die Kachel"
        Case "Tabelle": what = "Die Tabelle"
        Case Else: what = "Das Diagramm"
    End Select
    If mDash.CopyElement(eid) Then
        mUI.Alert what & " „" & mDash.ElemGet(eid, "Titel") & "“ ist kopiert." & vbLf & vbLf & _
                  "In PowerPoint oder Word mit Strg+V einfügen. Es kommt als Bild an, genau so gestaltet wie hier, " & _
                  "und bleibt beim Vergrößern scharf.", "Kopiert"
    Else
        mUI.Alert "Das Element ließ sich nicht kopieren. Bitte „Aktualisieren“ wählen und es noch einmal versuchen.", "Kopieren"
    End If
    Exit Sub
f: mApp.Boundary "Element kopieren"
End Sub

Public Sub Cmd_ElemDetails()
    On Error GoTo f
    mDash.Details mRender.MenuParam()
    Exit Sub
f: mApp.Boundary "Details"
End Sub

Public Sub Cmd_DashDetails()
    On Error GoTo f
    mDash.Details mDash.ElementFromShape(mRender.CallerName())
    Exit Sub
f: mApp.Boundary "Details"
End Sub

Public Sub Cmd_DashDrill()
    Dim info As Variant, eid As String, tid As String, ttl As String
    On Error GoTo f
    info = mDash.DrillInfo(mRender.CallerName())
    If IsEmpty(info) Then Exit Sub
    eid = info(0)
    tid = mDash.ElemGet(eid, "Tabelle")
    ttl = mDash.ElemGet(eid, "Titel")
    If Not IsEmpty(info(1)) Then
        If mCfg.FieldExists(mDash.ElemGet(eid, "Gruppierung")) Then ttl = ttl & ": " & mDash.GroupLabel(mCfg.Field(mDash.ElemGet(eid, "Gruppierung")), CStr(info(1)), mDash.ElemGet(eid, "Datumsstufe"))
    End If
    mDash.ShowRecordList tid, mDash.DrillRids(eid, mNav.gDash, info(1)), ttl
    Exit Sub
f: mApp.Boundary "Details"
End Sub

Public Sub Cmd_ElemEdit()
    On Error GoTo f
    If Not mApp.RequireSetup("Das Bearbeiten eines Dashboard-Elements") Then Exit Sub
    ElementFlow mNav.gDash, mRender.MenuParam()
    mNav.RenderCurrent
    Exit Sub
f: mDash.gHighlight = ""
    mApp.Boundary "Element bearbeiten"
End Sub

Public Sub Cmd_ElemUp()
    On Error GoTo f
    mStruct.MoveInOrder SH_ELM, "Dashboard", mNav.gDash, mRender.MenuParam(), -1
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Element verschieben"
End Sub

Public Sub Cmd_ElemDown()
    On Error GoTo f
    mStruct.MoveInOrder SH_ELM, "Dashboard", mNav.gDash, mRender.MenuParam(), 1
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Element verschieben"
End Sub

Public Sub Cmd_ElemDelete()
    Dim eid As String
    On Error GoTo f
    eid = mRender.MenuParam()
    If Not mApp.RequireSetup("Das Entfernen eines Dashboard-Elements") Then Exit Sub
    If Not mUI.Confirm("Element „" & mDash.ElemGet(eid, "Titel") & "“ entfernen? Die Daten bleiben unverändert.", "Element entfernen", "Entfernen", "Abbrechen", True) Then Exit Sub
    mCfg.DeleteRow SH_ELM, eid
    mNav.RenderCurrent
    Exit Sub
f: mApp.Boundary "Element entfernen"
End Sub

' Auswertungsassistent: Datenquelle, Kennzahl, Gruppierung, Zeitraum, Filter, Vorschau
Public Function ElementFlow(ByVal did As String, ByVal eid As String) As Boolean
    Dim f As frmProps, ids As Variant, tnames() As Variant, i As Long, tid As String, typ As String, ttl As String, sz As String
    Dim measures As Variant, mlabels() As Variant, fNames() As Variant, gNames() As Variant, dNames() As Variant, fld As cField, n As Long, g As Long, d As Long
    Dim props As New cMap, measure As String, fid As String, gfid As String, dfid As String, filt As String, ff As frmFilter, res As cMap, msg As String, ans As Long
    Dim isNew As Boolean, palNames As Variant, palItems() As Variant, palSel As String, opt As String
    isNew = (Len(eid) = 0)
    ids = mCfg.TableIds()
    If ArrLen(ids) = 0 Then mUI.Alert "Bitte zuerst eine Tabelle anlegen.", "Auswertung": Exit Function
    ReDim tnames(0 To ArrLen(ids) - 1)
    For i = 0 To ArrLen(ids) - 1: tnames(i) = mCfg.TableName(CStr(ids(i))): Next
    ' Schritt 1 (mit "Zurueck" aus Schritt 2 kehrt man mit den gewaehlten Angaben hierher zurueck)
    Dim szName As String, palName As String
    ttl = IIf(isNew, "", mDash.ElemGet(eid, "Titel"))
    tid = IIf(isNew, CStr(ids(0)), mDash.ElemGet(eid, "Tabelle"))
    typ = IIf(isNew, "Kachel", mDash.ElemGet(eid, "Typ"))
    szName = IIf(isNew, "Mittel (½ Breite)", Choose(Application.Max(1, Application.Min(3, Choose(Application.Max(1, CLng(Val(mDash.ElemGet(eid, "Groesse")))), 1, 2, 2, 3))), "Klein (¼ Breite)", "Mittel (½ Breite)", "Breit (ganze Breite)"))
    palName = IIf(Len(KVGet(mDash.ElemGet(eid, "Optionen"), "pal", "")) > 0, KVGet(mDash.ElemGet(eid, "Optionen"), "pal", ""), "(wie Dashboard)")
step1:
    Set f = New frmProps
    f.Init IIf(isNew, "Neues Element", "Element bearbeiten") & " (1/2)", "Datenquelle und Darstellung wählen."
    f.AddText "titel", "Titel", ttl
    f.AddCombo "tab", "Datenquelle (Tabelle)", tnames, mCfg.TableName(tid)
    f.AddCombo "typ", "Darstellung", mDash.ElemTypes(), typ, _
               "Unsicher? „Welches Diagramm passt?“ zeigt jede Darstellung als Bild mit Erklärung."
    f.SetComboHelper "typ", "Welches Diagramm passt?", "mForms.ChooseChartType"
    f.AddCombo "groesse", "Größe im Raster", Array("Klein (¼ Breite)", "Mittel (½ Breite)", "Breit (ganze Breite)"), szName
    palNames = mDash.PaletteNames()
    ReDim palItems(0 To UBound(palNames) + 1)
    palItems(0) = "(wie Dashboard)"
    For i = 0 To UBound(palNames): palItems(i + 1) = palNames(i): Next
    f.AddStyle "pal", "Farbstil", palName, "", "(wie Dashboard)"
    Do
        If Not f.ShowDialog("Weiter") Then Unload f: Exit Function
        f.ClearErrors
        If Len(Trim$(f.GetText("titel"))) > 0 Then Exit Do
        f.SetError "titel", "Bitte einen Titel angeben."
    Loop
    ttl = Trim$(f.GetText("titel"))
    tid = mCfg.TableIdByName(f.GetText("tab"))
    typ = f.GetText("typ")
    Select Case f.GetIndex("groesse")
        Case 0: sz = "1"
        Case 2: sz = "4"
        Case Else: sz = "2"
    End Select
    szName = f.GetText("groesse")
    palName = f.GetText("pal")
    palSel = palName
    If palSel = "(wie Dashboard)" Then palSel = ""
    Unload f
    ' Schritt 2
    measures = mDash.Measures()
    ReDim mlabels(0 To UBound(measures))
    For i = 0 To UBound(measures): mlabels(i) = mDash.MeasureLabel(CStr(measures(i))): Next
    ReDim fNames(0 To mCfg.Fields(tid).Count): ReDim gNames(0 To mCfg.Fields(tid).Count + 1): ReDim dNames(0 To mCfg.Fields(tid).Count + 1)
    fNames(0) = "(kein Feld)": n = 1
    gNames(0) = "(keine Gruppierung)": g = 1
    dNames(0) = "(kein Zeitraumfeld)": d = 1
    For Each fld In mCfg.Fields(tid)
        fNames(n) = fld.Name: n = n + 1
        If mQuery.CanGroupBy(fld) Or fld.FType = FT_BACKLINK Or fld.FType = FT_LINK Then gNames(g) = fld.Name: g = g + 1
        If fld.IsDate_ Then dNames(d) = fld.Name: d = d + 1
    Next
    ReDim Preserve fNames(0 To n - 1): ReDim Preserve gNames(0 To g - 1): ReDim Preserve dNames(0 To d - 1)
    Set f = New frmProps
    f.Init ttl & " (2/2)", "Kennzahl, Gruppierung und Zeitraum. Ohne Gruppierung entsteht eine einzelne Kennzahl; Diagramme brauchen eine Gruppierung."
    measure = IIf(isNew, "ANZAHL", mDash.ElemGet(eid, "Kennzahl"))
    f.AddCombo "kz", "Kennzahl", mlabels, mDash.MeasureLabel(measure)
    f.AddCombo "feld", "Feld für die Kennzahl", fNames, IIf(isNew, "(kein Feld)", IIf(mCfg.FieldExists(mDash.ElemGet(eid, "Feld")), mCfg.FieldName(mDash.ElemGet(eid, "Feld")), "(kein Feld)")), _
               "Summe, Durchschnitt, Median, Minimum, Maximum und Spannweite brauchen ein Zahlen-, Geld-, Prozent- oder Datumsfeld."
    ' Feld nur zeigen, wenn die gewählte Kennzahl eines braucht
    Dim needF As String, mm As Variant
    For Each mm In measures
        If mDash.MeasureNeedsField(CStr(mm)) Then needF = needF & IIf(Len(needF) > 0, ";", "") & mDash.MeasureLabel(CStr(mm))
    Next
    f.ShowWhen "feld", "kz", needF
    f.AddCombo "grp", IIf(typ = "Sankey", "Von (linke Seite)", "Gruppieren nach"), gNames, IIf(isNew, "(keine Gruppierung)", IIf(mCfg.FieldExists(mDash.ElemGet(eid, "Gruppierung")), mCfg.FieldName(mDash.ElemGet(eid, "Gruppierung")), "(keine Gruppierung)"))
    Dim zielF As String
    If typ = "Sankey" Then
        zielF = KVGet(mDash.ElemGet(eid, "Optionen"), "ziel", "")
        f.AddCombo "ziel", "Nach (rechte Seite)", gNames, IIf(mCfg.FieldExists(zielF), mCfg.FieldName(zielF), "(keine Gruppierung)"), _
                   "Beispiel: Von „Status“ nach „Priorität“. Die Bänder zeigen, wie viele Einträge (bzw. welche Summe) von links nach rechts fließen."
    End If
    f.AddCheck "filt", "Zusätzlich einen Filter für dieses Element festlegen", Len(mDash.ElemGet(eid, "Filter")) > 0
    f.AddCheck "mehr", "Weitere Einstellungen (Datumsstufe, Zeitraumfeld, Reihenfolge, Anzahl Gruppen)", False
    f.AllowBack = True
    Do
        If Not f.ShowDialog("Weiter") Then
            If f.WentBack Then Unload f: GoTo step1
            Unload f: Exit Function
        End If
        f.ClearErrors
        measure = measures(f.GetIndex("kz"))
        fid = mCfg.FieldIdByName(tid, f.GetText("feld"))
        gfid = mCfg.FieldIdByName(tid, f.GetText("grp"))
        msg = ""
        If mDash.MeasureNeedsField(measure) And Len(fid) = 0 Then
            f.SetError "feld", "Für „" & f.GetText("kz") & "“ bitte ein Feld wählen.": msg = "x"
        ElseIf mDash.MeasureNumeric(measure) And Len(fid) > 0 Then
            If Not (mCfg.Field(fid).IsNumeric_ Or mCfg.Field(fid).IsDate_) Then f.SetError "feld", "Diese Kennzahl braucht ein Zahlen-, Geld-, Prozent- oder Datumsfeld.": msg = "x"
        End If
        If Len(msg) = 0 And typ <> "Kachel" And typ <> "Tabelle" And Len(gfid) = 0 Then f.SetError "grp", "Diagramme brauchen eine Gruppierung.": msg = "x"
        If Len(msg) = 0 And typ = "Sankey" Then
            zielF = mCfg.FieldIdByName(tid, f.GetText("ziel"))
            If Len(zielF) = 0 Then
                f.SetError "ziel", "Bitte die rechte Seite wählen.": msg = "x"
            ElseIf zielF = gfid Then
                f.SetError "ziel", "Links und rechts bitte verschiedene Felder wählen.": msg = "x"
            ElseIf measure <> "ANZAHL" And measure <> "SUMME" And measure <> "GEFUELLT" Then
                f.SetError "kz", "Sankey braucht eine Kennzahl, die sich addieren lässt: Anzahl oder Summe.": msg = "x"
            End If
        End If
        If Len(msg) = 0 Then Exit Do
    Loop
    props.SetVal "Titel", ttl
    props.SetVal "Typ", typ
    props.SetVal "Tabelle", tid
    props.SetVal "Groesse", sz
    props.SetVal "Kennzahl", measure
    props.SetVal "Feld", fid
    props.SetVal "Gruppierung", gfid
    ' Selteneres: bisherige Werte bzw. sinnvolle Voreinstellungen, Anpassung nur auf Wunsch
    Dim stufe As String, zeitF As String, basis As String, sortM As String, topN As String, fm As frmProps
    stufe = IIf(isNew Or Len(mDash.ElemGet(eid, "Datumsstufe")) = 0, "Monat", mDash.ElemGet(eid, "Datumsstufe"))
    zeitF = IIf(isNew, "", mDash.ElemGet(eid, "Datumsfeld"))
    If mDash.ElemGet(eid, "Tabelle") <> tid Then zeitF = ""
    If Len(zeitF) = 0 Then
        ' Voreinstellung: erstes Datumsfeld der Tabelle (damit der Dashboardfilter "Zeitraum" wirkt)
        If d > 1 Then zeitF = mCfg.FieldIdByName(tid, CStr(dNames(1)))
    End If
    basis = IIf(mDash.ElemGet(eid, "Basis") = "GEFUELLT", "GEFUELLT", "ALLE")
    sortM = KVGet(mDash.ElemGet(eid, "Optionen"), "sort", "wert")
    topN = KVGet(mDash.ElemGet(eid, "Optionen"), "top", "12")
    If f.GetBool("mehr") Then
        Set fm = New frmProps
        fm.Init ttl & ": weitere Einstellungen", "Nur bei Bedarf anpassen."
        fm.AddCombo "stufe", "Datumsgruppierung", Array("Monat", "Quartal", "Jahr", "Tag"), stufe
        fm.AddCombo "zeit", "Zeitraumfeld (für Dashboardfilter „Zeitraum“)", dNames, IIf(Len(zeitF) > 0, mCfg.FieldName(zeitF), "(kein Zeitraumfeld)")
        fm.AddCombo "basis", "Grundgesamtheit bei Anteilen", Array("Alle gefilterten Datensätze", "Nur Datensätze mit Gruppenangabe"), IIf(basis = "GEFUELLT", "Nur Datensätze mit Gruppenangabe", "Alle gefilterten Datensätze")
        fm.AddCombo "sort", "Reihenfolge der Gruppen", Array("Nach Wert (absteigend)", "Nach Bezeichnung / Reihenfolge der Optionen"), IIf(sortM = "name", "Nach Bezeichnung / Reihenfolge der Optionen", "Nach Wert (absteigend)")
        fm.AddText "top", "Höchstens so viele Gruppen anzeigen", topN
        If fm.ShowDialog("Weiter") Then
            stufe = fm.GetText("stufe")
            zeitF = mCfg.FieldIdByName(tid, fm.GetText("zeit"))
            basis = IIf(fm.GetIndex("basis") = 1, "GEFUELLT", "ALLE")
            sortM = IIf(fm.GetIndex("sort") = 1, "name", "wert")
            topN = CStr(Application.Max(1, CLng(Val(fm.GetText("top")))))
        End If
        Unload fm
    End If
    props.SetVal "Datumsstufe", stufe
    props.SetVal "Datumsfeld", zeitF
    props.SetVal "Basis", basis
    opt = KVSet(KVSet("", "sort", sortM), "top", topN)
    opt = KVSet(opt, "pal", palSel)
    If typ = "Sankey" Then opt = KVSet(opt, "ziel", zielF)
    If Not isNew Then opt = KVSet(opt, "hoch", KVGet(mDash.ElemGet(eid, "Optionen"), "hoch", ""))
    props.SetVal "Optionen", opt
    filt = IIf(isNew, "", mDash.ElemGet(eid, "Filter"))
    If mDash.ElemGet(eid, "Tabelle") <> tid Then filt = ""
    If f.GetBool("filt") Then
        Unload f
        Set ff = New frmFilter
        ff.Setup tid, filt, "Filter für „" & ttl & "“"
        ff.Show
        If ff.OK Then filt = ff.Result
        Unload ff
    Else
        filt = ""
        Unload f
    End If
    props.SetVal "Filter", filt
    ' speichern und Vorschau
    If isNew Then
        eid = mDash.CreateElement(did, props)
    Else
        Dim k As Variant
        For Each k In props.Keys
            mCfg.SetV SH_ELM, eid, CStr(k), props.GetVal(CStr(k))
        Next
    End If
    Set res = mDash.Compute(eid, did)
    If res.Exists("error") Then
        msg = "Hinweis: " & res.GetVal("error")
    Else
        msg = "Gesamt " & mDash.FormatMeasure(res, res.GetVal("totalValue")) & IIf(CLng(res.GetVal("n", 0)) > 0, " in " & res.GetVal("n") & " Gruppen", "")
    End If
    ' echte Vorschau: Dashboard mit dem (markierten) Element zeigen, Dialog rechts unten
    mDash.gHighlight = eid
    mNav.RenderCurrent
    ScrollToElement eid
    Application.ScreenUpdating = True
    ans = mUI.Ask3Corner("Vorschau: Das blau umrandete Element zeigt das Ergebnis." & vbLf & msg, "Element „" & ttl & "“", "Übernehmen", "Nochmals bearbeiten", IIf(isNew, "Verwerfen", "Schließen"))
    mDash.gHighlight = ""
    If ans = 2 Then
        ElementFlow = ElementFlow(did, eid)
        Exit Function
    ElseIf ans = 3 And isNew Then
        mCfg.DeleteRow SH_ELM, eid
        Exit Function
    End If
    mLog.Add IIf(isNew, "Element angelegt", "Element geändert"), tid, eid, "", ttl
    mApp.MarkCfgDirty
    ElementFlow = True
End Function

Private Sub ScrollToElement(ByVal eid As String)
    Dim ws As Worksheet, s As Shape, r As Long
    On Error Resume Next
    Set ws = mRender.Sheet(SH_DASH)
    Set s = ws.Shapes("card_" & eid)
    If s Is Nothing Then Exit Sub
    r = s.TopLeftCell.Row
    If s.Top + s.Height > ws.Rows(ActiveWindow.ScrollRow).Top + ActiveWindow.UsableHeight - 20 Then ActiveWindow.ScrollRow = Application.Max(1, r - 1)
End Sub
