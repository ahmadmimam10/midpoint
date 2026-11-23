Attribute VB_Name = "mViewBoard"
'==============================================================================
' mViewBoard - Board nach Status oder Einfachauswahl. Optionen der Ansicht:
' gruppe=Feld, karte=Feld,Feld (Kurzangaben). Karten oeffnen (Doppelklick),
' neu anlegen je Spalte, einer anderen Gruppe zuordnen (Schaltflaechen/Dialog).
' Eintraege ohne Gruppe erhalten eine eigene Spalte.
'==============================================================================
Option Explicit

Private Const CARD_W As Single = 210
Private Const GAP_W As Single = 12
Private Const MAX_CARDS As Long = 40
Private mCell As cMap            ' "Zeile,Spalte" -> Datensatz-ID | "G:OptionID" (Spaltenkopf)
Private mGroups As Collection    ' Options-IDs in Spaltenreihenfolge ("" = ohne Zuordnung)
Private mTid As String
Private mVid As String
Private mSelRid As String

Public Sub Render()
    Dim ws As Worksheet, tid As String, vid As String, st As cViewState, gfid As String, gf As cField, cardF As Variant
    Dim t As cTableData, rows() As Long, n As Long, k As Long, r As Long, sr As Range, oid As Variant, gi As Long, col As Long
    Dim byGroup As cMap, lst As String, a As Variant, i As Long, j As Long, rr As Long, rid As String, cell As Range, maxR As Long
    Dim cnt As Long, s As Shape, f As cField, txt As String, groupsShown As New Collection, hasNone As Boolean, hdr As Range

    tid = mNav.gTid: vid = mNav.gVid: mTid = tid: mVid = vid
    Set st = mNav.State(vid)
    gfid = mCfg.ViewOpt(vid, "gruppe")
    Set ws = mRender.Sheet(SH_BOARD)
    mRender.BeginPage ws
    Set mCell = New cMap
    mRender.DrawTitle ws, mCfg.TableIcon(tid), mCfg.TableName(tid), mCfg.GetS(SH_TAB, tid, "Beschreibung")
    mViewTable.DrawViewTabs ws, tid, vid, "Neuer Eintrag"
    If Not mCfg.FieldExists(gfid) Then
        mRender.EmptyState ws, R_DATA, "Board noch nicht eingerichtet", "Bitte unter „Ansicht > Ansicht bearbeiten“ ein Status- oder Auswahlfeld für die Spalten wählen.", _
                           "Ansicht bearbeiten", "mCmd.Cmd_EditView", ChrW(&HE8A9)
        mRender.EndPage ws
        Exit Sub
    End If
    Set gf = mCfg.Field(gfid)

    ' Spalten: Optionen in Reihenfolge, "ohne Zuordnung" zuerst, wenn vorhanden
    Set t = mData.TData(tid)
    rows = mViewTable.ViewRows(vid, n, False)
    Set byGroup = New cMap
    For k = 1 To n
        r = rows(k)
        byGroup.SetVal "k" & NzS(t.Val_(r, gfid)), ListAdd(CStr(byGroup.GetVal("k" & NzS(t.Val_(r, gfid)), "")), t.IdAt(r))
    Next
    Set mGroups = New Collection
    If byGroup.Exists("k") Then mGroups.Add ""
    For Each oid In mCfg.OptionIds(gfid)
        If mCfg.OptActive(CStr(oid)) Or byGroup.Exists("k" & oid) Then mGroups.Add CStr(oid)
    Next

    ' Spaltenbreiten: Karte, Abstand, Karte, ...
    For gi = 1 To mGroups.Count
        mRender.SetColWidthPt ws, 1 + gi * 2, CARD_W
        mRender.SetColWidthPt ws, 2 + gi * 2, GAP_W
    Next

    ' Werkzeuge
    mRender.FlowBegin ws, R_TOOLS, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowBtn "b_filter", "Filter", "mCmd.Cmd_Filter", "secondary", ChrW(&HE71C)
    mRender.FlowBtn "b_view", "Ansicht", "mCmd.Cmd_ViewMenu", "secondary", ChrW(&HE8FD)
    If mData.UndoAvailable Then mRender.FlowBtn "b_undo", "Rückgängig", "mCmd.Cmd_Undo", "secondary", ChrW(&HE7A7)
    mRender.FlowGap 10
    mRender.FlowBtn "b_left", "Nach links", "mCmd.Cmd_BoardLeft", "secondary", ChrW(&HE76B)
    mRender.FlowBtn "b_right", "Nach rechts", "mCmd.Cmd_BoardRight", "secondary", ChrW(&HE76C)
    mRender.FlowBtn "b_group", "Gruppe ändern…", "mCmd.Cmd_BoardMove", "secondary", ChrW(&HE8AB)
    mRender.FlowBtn "b_edit", "Öffnen", "mCmd.Cmd_EditSelected", "secondary", ChrW(&HE70F)
    mViewTable.DrawSetupTools
    mRender.FlowEnd
    mViewTable.DrawChips ws, st, n, ActiveCount(t), "Einträge"

    cardF = Split(mCfg.ViewOpt(vid, "karte"), ",")
    ' einheitliches Kartenraster ueber alle Spalten (Zeilen werden geteilt)
    Dim nDet As Long, slot As Long, maxCards As Long
    For k = 0 To UBound(cardF)
        If mCfg.FieldExists(CStr(cardF(k))) Then nDet = nDet + 1
    Next
    slot = nDet + 2
    For gi = 1 To mGroups.Count
        cnt = ListCount(CStr(byGroup.GetVal("k" & mGroups(gi), "")))
        If cnt > maxCards Then maxCards = cnt
    Next
    If maxCards > MAX_CARDS Then maxCards = MAX_CARDS + 1
    For i = 0 To maxCards
        ws.Rows(R_DATA + i * slot).RowHeight = 22
        For j = 1 To nDet
            ws.Rows(R_DATA + i * slot + j).RowHeight = 17
        Next
        ws.Rows(R_DATA + i * slot + nDet + 1).RowHeight = 8
    Next
    ws.Rows(R_HEAD).RowHeight = 30
    maxR = R_DATA
    For gi = 1 To mGroups.Count
        col = 1 + gi * 2
        oid = mGroups(gi)
        lst = CStr(byGroup.GetVal("k" & oid, ""))
        cnt = ListCount(lst)
        ' Spaltenkopf
        Set hdr = ws.Cells(R_HEAD, col)
        hdr.NumberFormat = "@"
        If Len(oid) = 0 Then
            hdr.Value = "Ohne " & gf.Name & "   " & cnt
            hdr.Interior.Color = CSurface2(): hdr.Font.Color = CMuted()
        Else
            hdr.Value = mCfg.OptName(CStr(oid)) & IIf(mCfg.OptDone(CStr(oid)), "  " & ChrW(&H2713), "") & "   " & cnt
            hdr.Interior.Color = PaletteBg(mCfg.OptColor(CStr(oid))): hdr.Font.Color = PaletteFg(mCfg.OptColor(CStr(oid)))
        End If
        hdr.Font.Name = UI_FONT_SEMI: hdr.IndentLevel = 1
        mCell.SetVal R_HEAD & "," & col, "G:" & oid
        ' Karten
        a = Split(lst, ";")
        rr = R_DATA
        For i = 0 To UBound(a)
            If i >= MAX_CARDS Then
                Set cell = ws.Cells(rr, col)
                cell.Value = "+ " & (UBound(a) + 1 - MAX_CARDS) & " weitere (Filter nutzen)"
                cell.Font.Size = 8.5: cell.Font.Color = CMuted(): cell.IndentLevel = 1
                rr = rr + slot
                Exit For
            End If
            rid = CStr(a(i))
            ' Titel
            Set cell = ws.Cells(rr, col)
            cell.NumberFormat = "@"
            cell.Value = Truncate(mData.RecordLabel(tid, rid), CLng(CARD_W / 6.2))
            cell.Font.Name = UI_FONT_SEMI: cell.IndentLevel = 1
            mCell.SetVal rr & "," & col, rid
            j = 0
            For k = 0 To UBound(cardF)
                If mCfg.FieldExists(CStr(cardF(k))) Then
                    Set f = mCfg.Field(CStr(cardF(k)))
                    j = j + 1
                    Set cell = ws.Cells(rr + j, col)
                    txt = mFmt.DisplayText(f, mData.GetValue(tid, rid, f.ID), rid)
                    cell.NumberFormat = "@"
                    cell.Value = Truncate(f.Name & ": " & IIf(Len(txt) = 0, "(leer)", txt), CLng(CARD_W / 5.2))
                    cell.Font.Size = 8.5: cell.Font.Color = CMuted(): cell.IndentLevel = 1
                    mCell.SetVal (rr + j) & "," & col, rid
                End If
            Next
            With ws.Range(ws.Cells(rr, col), ws.Cells(rr + nDet, col))
                .Interior.Color = IIf(rid = mSelRid, CAccentSoft(), CWhite())
                .Borders(xlEdgeLeft).Color = CBorderStrong(): .Borders(xlEdgeRight).Color = CBorderStrong()
                .Borders(xlEdgeTop).Color = CBorderStrong(): .Borders(xlEdgeBottom).Color = CBorderStrong()
            End With
            rr = rr + slot
        Next
        ' "Neu" in dieser Spalte
        Set s = mRender.Btn(ws, "new_" & gi, "Neu hier", ws.Cells(rr, col).Left + 4, ws.Rows(rr).Top + 2, "mCmd.Cmd_BoardNewHere", "link", ChrW(&HE710))
        If rr > maxR Then maxR = rr
    Next
    If mGroups.Count = 0 Then
        mRender.EmptyState ws, R_DATA, "Keine Spalten", "Das Gruppenfeld hat noch keine Auswahlmöglichkeiten. Bitte unter Einrichten Optionen anlegen."
    Else
        mViewCal.TextNote ws, maxR + 2, "Karte anklicken und mit „Nach links/rechts“ oder „Gruppe ändern…“ verschieben. Doppelklick öffnet den Eintrag. " & _
                                       "Spaltenkopf doppelklicken legt einen neuen Eintrag in dieser Gruppe an. Drag-and-drop wird nicht angeboten."
    End If
    mRender.EndPage ws, R_HEAD, 0
End Sub

Private Function ActiveCount(ByVal t As cTableData) As Long
    Dim r As Long
    For r = 1 To t.RowCount
        If Not t.IsArchived(r) Then ActiveCount = ActiveCount + 1
    Next
End Function

Public Function CellTarget(ByVal target As Range) As String
    If mCell Is Nothing Then Exit Function
    CellTarget = CStr(mCell.GetVal(target.Row & "," & target.Column, ""))
End Function

Public Function SelectedRid() As String
    SelectedRid = mSelRid
End Function

Public Sub OnSelect(ByVal target As Range)
    Dim tg As String
    tg = CellTarget(target)
    If Len(tg) > 0 And Left$(tg, 2) <> "G:" Then
        If tg <> mSelRid Then
            mSelRid = tg
            mNav.State(mVid).SelRid = tg
            HighlightSel
        End If
    End If
End Sub

Private Sub HighlightSel()
    Dim ws As Worksheet, k As Variant, p As Variant, prevEv As Boolean
    Set ws = mRender.Sheet(SH_BOARD)
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    mApp.BeginWork
    For Each k In mCell.Keys
        p = Split(CStr(k), ",")
        If Left$(mCell.GetVal(CStr(k)), 2) <> "G:" And CLng(p(0)) >= R_DATA Then
            If ws.Cells(CLng(p(0)), CLng(p(1))).Interior.Color = CAccentSoft() Or ws.Cells(CLng(p(0)), CLng(p(1))).Interior.Color = CWhite() Then
                ws.Cells(CLng(p(0)), CLng(p(1))).Interior.Color = IIf(mCell.GetVal(CStr(k)) = mSelRid, CAccentSoft(), CWhite())
            End If
        End If
    Next
    mApp.EndWork
    Application.EnableEvents = prevEv
End Sub

Public Sub OnDoubleClick(ByVal target As Range, ByRef Cancel As Boolean)
    Dim tg As String, pre As cMap
    tg = CellTarget(target)
    If Len(tg) = 0 Then Exit Sub
    Cancel = True
    If Left$(tg, 2) = "G:" Then
        NewInGroup Mid$(tg, 3)
    Else
        mSelRid = tg
        mForms.OpenRecord mTid, tg
    End If
End Sub

Public Sub NewInGroup(ByVal oid As String)
    Dim pre As New cMap
    If Len(oid) > 0 Then pre.SetVal mCfg.ViewOpt(mVid, "gruppe"), oid
    mForms.NewRecord mTid, pre
End Sub

Public Sub NewInColumn(ByVal shapeName As String)
    Dim gi As Long
    gi = CLng(Val(Mid$(shapeName, 5)))
    If gi >= 1 And gi <= mGroups.Count Then NewInGroup CStr(mGroups(gi))
End Sub

' Karte in benachbarte Gruppe (delta) oder gewaehlte Gruppe verschieben
Public Sub MoveSelected(ByVal delta As Long)
    Dim gfid As String, cur As String, gi As Long, i As Long, target As String
    If Len(mSelRid) = 0 Or Not mData.RecordExists(mTid, mSelRid) Then
        mUI.Alert "Bitte zuerst eine Karte anklicken.", "Verschieben": Exit Sub
    End If
    gfid = mCfg.ViewOpt(mVid, "gruppe")
    cur = NzS(mData.GetValue(mTid, mSelRid, gfid))
    For i = 1 To mGroups.Count
        If mGroups(i) = cur Then gi = i
    Next
    If delta = 0 Then
        Dim names() As Variant, ids As New Collection, oid As Variant, n As Long, sel As Long
        ReDim names(0 To mCfg.OptionIds(gfid, True).Count)
        names(0) = "(ohne Zuordnung)": ids.Add ""
        n = 1
        For Each oid In mCfg.OptionIds(gfid, True)
            names(n) = mCfg.OptName(CStr(oid)): ids.Add CStr(oid): n = n + 1
        Next
        sel = mUI.ChooseOne("„" & mData.RecordLabel(mTid, mSelRid) & "“ einer anderen Gruppe zuordnen:", "Gruppe ändern", names, "Zuordnen")
        If sel < 0 Then Exit Sub
        target = ids(sel + 1)
    Else
        If gi + delta < 1 Or gi + delta > mGroups.Count Then Exit Sub
        target = mGroups(gi + delta)
    End If
    If target = cur Then Exit Sub
    SetGroup mSelRid, target
End Sub

Private Sub SetGroup(ByVal rid As String, ByVal oid As String)
    Dim gfid As String, vals As New cMap, errs As cMap, ek As Variant, msg As String
    If Not mApp.RequireWritable("Das Verschieben") Then Exit Sub
    gfid = mCfg.ViewOpt(mVid, "gruppe")
    If Len(oid) > 0 Then vals.SetVal gfid, oid Else vals.SetVal gfid, Empty
    If Not mValid.ValidateRecord(mTid, rid, vals, New cMap, errs) Then
        For Each ek In errs.Keys
            msg = msg & mCfg.FieldName(CStr(ek)) & ": " & errs.GetVal(CStr(ek)) & vbLf
        Next
        mUI.Alert "Die Karte kann nicht verschoben werden, weil der Eintrag sonst ungültig wäre:" & vbLf & vbLf & msg, "Verschieben", True
        Exit Sub
    End If
    mApp.BeginWork
    On Error GoTo failed
    mData.TxBegin "Gruppe von „" & mData.RecordLabel(mTid, rid) & "“ geändert"
    mData.UpdateRecord mTid, rid, vals
    mLog.Add "Geändert", mTid, rid, gfid, "Board"
    mData.TxCommit
    mApp.EndWork
    mNav.RenderCurrent
    Exit Sub
failed:
    mData.TxAbortIfOpen
    mApp.ResetWork
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
