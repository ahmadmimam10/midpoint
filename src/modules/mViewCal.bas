Attribute VB_Name = "mViewCal"
'==============================================================================
' mViewCal - Kalender (Monatsansicht und Terminliste). Einstellungen der Ansicht:
' datum=Feld, ende=Feld (optional), farbe=Auswahl-/Statusfeld (optional).
' Mehrere Termine je Tag: bis zu 4 sichtbar, weitere ueber "+n weitere".
'==============================================================================
Option Explicit

Private Const LINES_PER_DAY As Long = 4
Private mCell As cMap            ' "Zeile,Spalte" -> Datensatz-ID | "D:Datum" (Tag) | "M:Datum" (weitere)
Private mTid As String
Private mVid As String

Public Sub Render()
    Dim ws As Worksheet, tid As String, vid As String, st As cViewState, dfid As String, efid As String, cfid As String
    Dim t As cTableData, rows() As Long, n As Long, k As Long, r As Long, d0 As Double, d1 As Double, sr As Range
    Dim first As Date, start As Date, i As Long, w As Long, dcol As Long, wd As Long, dd As Date, rr As Long
    Dim byDay As cMap, lst As String, a As Variant, j As Long, cnt As Long, colW As Single, x As Single, undated As Long
    Dim lf As String, lbl As String, rid As String, oid As String, cell As Range, tf As cField, total As Long

    tid = mNav.gTid: vid = mNav.gVid: mTid = tid: mVid = vid
    Set st = mNav.State(vid)
    dfid = mCfg.ViewOpt(vid, "datum")
    efid = mCfg.ViewOpt(vid, "ende")
    cfid = mCfg.ViewOpt(vid, "farbe")
    lf = mCfg.LabelField(tid)
    Set ws = mRender.Sheet(SH_CAL)
    mRender.BeginPage ws
    Set mCell = New cMap
    colW = (mRender.PageRight() - mRender.ContentLeft(ws) - 8) / 7
    If colW < 96 Then colW = 96
    For i = 0 To 6
        mRender.SetColWidthPt ws, 3 + i, colW
    Next
    ' Terminliste: Spaltenbreiten vor dem Suchfeld setzen. Das Suchfeld besteht aus Zellen und wuerde sonst nachtraeglich
    ' mitwachsen und unter "Suchen" und die Anzahl der Eintraege laufen.
    If st.CalMode = "Liste" Then ListColWidths ws
    mRender.DrawTitle ws, mCfg.TableIcon(tid), mCfg.TableName(tid), mCfg.GetS(SH_TAB, tid, "Beschreibung")
    mViewTable.DrawViewTabs ws, tid, vid, "Neuer Termin"

    If Not mCfg.FieldExists(dfid) Then
        mRender.EmptyState ws, R_DATA, "Kalender noch nicht eingerichtet", "Bitte unter „Ansicht > Ansicht bearbeiten“ ein Datumsfeld für den Kalender wählen.", _
                           "Ansicht bearbeiten", "mCmd.Cmd_EditView", ChrW(&HE787)
        mRender.EndPage ws
        Exit Sub
    End If

    ' Werkzeuge
    mRender.FlowBegin ws, R_TOOLS, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowBtn "b_prev", "", "mCmd.Cmd_CalPrev", "secondary", ChrW(&HE76B), 2
    mRender.FlowBtn "b_today", "Heute", "mCmd.Cmd_CalToday", "secondary", "", 2
    mRender.FlowBtn "b_next", "", "mCmd.Cmd_CalNext", "secondary", ChrW(&HE76C), 8
    mRender.FlowText "b_month", MonthName_DE(Month(st.CalMonth)) & " " & Year(st.CalMonth), CText(), 12, True, 12
    mRender.FlowBtn "b_mode", IIf(st.CalMode = "Liste", "Monatsansicht", "Terminliste"), "mCmd.Cmd_CalMode", "secondary", IIf(st.CalMode = "Liste", ChrW(&HE787), ChrW(&HE8FD))
    mRender.FlowBtn "b_filter", "Filter", "mCmd.Cmd_Filter", "secondary", ChrW(&HE71C)
    mRender.FlowBtn "b_view", "Ansicht", "mCmd.Cmd_ViewMenu", "secondary", ChrW(&HE8FD)
    If mData.UndoAvailable Then mRender.FlowBtn "b_undo", "Rückgängig", "mCmd.Cmd_Undo", "secondary", ChrW(&HE7A7)

    ' Daten: Eintraege im sichtbaren Zeitraum (6 Wochen) mit Beginn und Ende
    Set t = mData.TData(tid)
    rows = mViewTable.ViewRows(vid, n, False)
    first = st.CalMonth
    start = first - Weekday(first, vbMonday) + 1
    Dim evRid() As String, evA() As Long, evB() As Long, nEv As Long
    ReDim evRid(1 To Application.Max(1, n)): ReDim evA(1 To Application.Max(1, n)): ReDim evB(1 To Application.Max(1, n))
    For k = 1 To n
        r = rows(k)
        If IsBlank(t.Val_(r, dfid)) Or Not (IsDate(t.Val_(r, dfid)) Or IsNumeric(t.Val_(r, dfid))) Then
            undated = undated + 1
        Else
            d0 = Int(CDbl(t.Val_(r, dfid)))
            d1 = d0
            If Len(efid) > 0 Then
                If Not IsBlank(t.Val_(r, efid)) And (IsDate(t.Val_(r, efid)) Or IsNumeric(t.Val_(r, efid))) Then d1 = Int(CDbl(t.Val_(r, efid)))
            End If
            If d1 < d0 Then d1 = d0
            If d1 - d0 > 366 Then d1 = d0 + 366
            If d1 >= CLng(start) And d0 < CLng(start) + 42 Then
                nEv = nEv + 1
                evRid(nEv) = t.IdAt(r): evA(nEv) = CLng(d0): evB(nEv) = CLng(d1)
            End If
        End If
    Next
    mRender.FlowBtn "b_undated", "Ohne Datum (" & undated & ")", "mCmd.Cmd_CalUndated", "secondary", ChrW(&HE946)
    mViewTable.DrawSetupTools
    mRender.FlowEnd
    total = n
    mViewTable.DrawChips ws, st, n, ActiveCount(t), "Einträge"

    If st.CalMode = "Liste" Then
        RenderList ws, t, rows, n, dfid, efid, cfid, st
        mRender.EndPage ws, R_HEAD, 0
        Exit Sub
    End If

    ' Wochentage
    ws.Rows(R_HEAD).RowHeight = 22
    For i = 1 To 7
        With ws.Cells(R_HEAD, 2 + i)
            .Value = Choose(i, "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag")
            .Font.Color = CMuted(): .Font.Size = 9: .HorizontalAlignment = xlLeft: .IndentLevel = 1
        End With
    Next
    ws.Range(ws.Cells(R_HEAD, 3), ws.Cells(R_HEAD, 9)).Borders(xlEdgeBottom).Color = CBorderStrong()

    ' Wochen: Tageskaesten, dann je Woche feste Spuren fuer die Eintraege. Ein mehrtaegiger
    ' Eintrag liegt die ganze Woche in derselben Spur und erscheint als durchgehende Leiste;
    ' ueberlappende Eintraege bekommen getrennte Spuren und vermischen sich nicht.
    Dim todayBox As Range
    rr = R_DATA
    For w = 0 To 5
        ws.Rows(rr).RowHeight = 18
        For j = 1 To LINES_PER_DAY
            ws.Rows(rr + j).RowHeight = 17
        Next
        ws.Rows(rr + LINES_PER_DAY + 1).RowHeight = 15
        For wd = 0 To 6
            dd = start + w * 7 + wd
            dcol = 3 + wd
            Set cell = ws.Cells(rr, dcol)
            cell.Value = IIf(Day(dd) = 1, Day(dd) & ". " & Left$(MonthName_DE(Month(dd)), 3), CStr(Day(dd)))
            cell.HorizontalAlignment = xlRight
            cell.IndentLevel = 1
            cell.Font.Size = 9
            If Month(dd) <> Month(first) Then
                cell.Font.Color = RGB(170, 170, 170)
                ws.Range(ws.Cells(rr, dcol), ws.Cells(rr + LINES_PER_DAY + 1, dcol)).Interior.Color = CSurface()
            Else
                cell.Font.Color = CText()
            End If
            For j = 0 To LINES_PER_DAY + 1
                mCell.SetVal (rr + j) & "," & dcol, "D:" & CLng(dd)
            Next
            With ws.Range(ws.Cells(rr, dcol), ws.Cells(rr + LINES_PER_DAY + 1, dcol))
                .Borders(xlEdgeLeft).Color = CBorder(): .Borders(xlEdgeRight).Color = CBorder()
                .Borders(xlEdgeTop).Color = CBorder(): .Borders(xlEdgeBottom).Color = CBorder()
            End With
            ' Heute: Rahmen und fette Tageszahl statt Flaechenfarbe (sonst wie ein Eintrag)
            If dd = Date Then
                cell.Value = "Heute  " & Day(dd)
                cell.Font.Color = CAccent(): cell.Font.Bold = True
                Set todayBox = ws.Range(ws.Cells(rr, dcol), ws.Cells(rr + LINES_PER_DAY + 1, dcol))
            End If
        Next
        DrawWeekBars ws, rr, CLng(start) + w * 7, evRid, evA, evB, nEv, cfid, colW
        rr = rr + LINES_PER_DAY + 2
    Next
    If Not todayBox Is Nothing Then
        With todayBox
            .Borders(xlEdgeLeft).Color = CAccent(): .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Color = CAccent(): .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeTop).Color = CAccent(): .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeBottom).Color = CAccent(): .Borders(xlEdgeBottom).Weight = xlMedium
        End With
    End If
    TextNote ws, rr + 1, "Doppelklick auf einen Termin öffnet ihn; Doppelklick auf einen freien Tag legt dort einen neuen Eintrag an. " & _
                         "Termine verschieben: Eintrag öffnen und Datum ändern." & IIf(undated > 0, "  " & undated & " Einträge ohne Datum sind über „Ohne Datum“ erreichbar.", "")
    mRender.EndPage ws, R_HEAD, 0
End Sub

' Leisten einer Woche (Tag 0 = Montag in Spalte 3). Spurvergabe: frueher Beginn zuerst,
' bei gleichem Beginn der laengere Eintrag; jeder bekommt die niedrigste Spur, die an allen
' seinen Tagen dieser Woche frei ist. Tage mit mehr Eintraegen als Zeilen zeigen in der
' letzten Zeile "+ n weitere".
Private Sub DrawWeekBars(ByVal ws As Worksheet, ByVal rr As Long, ByVal wk0 As Long, evRid() As String, evA() As Long, evB() As Long, _
                         ByVal nEv As Long, ByVal cfid As String, ByVal colW As Single)
    Dim cand() As Long, m As Long, e As Long, i As Long, j As Long, tmp As Long, a As Long, b As Long, L As Long, free As Boolean
    Dim lane() As Long, occ() As Boolean, cnt(0 To 6) As Long, hid(0 To 6) As Long, vis(0 To 6) As Long, d As Long, s0 As Long
    If nEv = 0 Then Exit Sub
    ReDim cand(1 To nEv)
    For e = 1 To nEv
        If evB(e) >= wk0 And evA(e) <= wk0 + 6 Then m = m + 1: cand(m) = e
    Next
    If m = 0 Then Exit Sub
    ' sortieren (wenige Eintraege je Woche: Einfuegesortierung)
    For i = 2 To m
        tmp = cand(i): j = i - 1
        Do While j >= 1
            If Not Before(tmp, cand(j), evA, evB, wk0) Then Exit Do
            cand(j + 1) = cand(j): j = j - 1
        Loop
        cand(j + 1) = tmp
    Next
    ReDim lane(1 To m): ReDim occ(0 To 6, 0 To m)
    For i = 1 To m
        a = Application.Max(evA(cand(i)), wk0) - wk0: b = Application.Min(evB(cand(i)), wk0 + 6) - wk0
        For L = 0 To m
            free = True
            For d = a To b
                If occ(d, L) Then free = False: Exit For
            Next
            If free Then Exit For
        Next
        lane(i) = L
        For d = a To b: occ(d, L) = True: cnt(d) = cnt(d) + 1: Next
    Next
    For d = 0 To 6
        vis(d) = IIf(cnt(d) > LINES_PER_DAY, LINES_PER_DAY - 1, LINES_PER_DAY)
    Next
    For i = 1 To m
        a = Application.Max(evA(cand(i)), wk0) - wk0: b = Application.Min(evB(cand(i)), wk0 + 6) - wk0
        ' zusammenhaengende sichtbare Abschnitte zeichnen
        s0 = -1
        For d = a To b + 1
            If d <= b Then
                If lane(i) < vis(d) Then
                    If s0 < 0 Then s0 = d
                Else
                    hid(d) = hid(d) + 1
                End If
            End If
            If (d > b Or (d <= b And lane(i) >= vis(Application.Min(d, 6)))) And s0 >= 0 Then
                DrawBar ws, rr + 1 + lane(i), 3 + s0, 3 + d - 1, evRid(cand(i)), (evA(cand(i)) < wk0 + s0), cfid, colW
                s0 = -1
            End If
        Next
    Next
    For d = 0 To 6
        If hid(d) > 0 Then
            With ws.Cells(rr + LINES_PER_DAY, 3 + d)
                .Value = "+ " & hid(d) & " weitere"
                .Font.Size = 8.5: .Font.Color = CAccent(): .Font.Bold = True
                .IndentLevel = 1
            End With
            mCell.SetVal (rr + LINES_PER_DAY) & "," & (3 + d), "M:" & (wk0 + d)
        End If
    Next
End Sub

Private Function Before(ByVal x As Long, ByVal y As Long, evA() As Long, evB() As Long, ByVal wk0 As Long) As Boolean
    Dim ax As Long, ay As Long
    ax = Application.Max(evA(x), wk0): ay = Application.Max(evA(y), wk0)
    If ax <> ay Then Before = (ax < ay): Exit Function
    Before = (evB(x) - ax > evB(y) - ay)
End Function

' Eine Leiste ueber die Spalten c0..c1 einer Zeile: gleiche Flaeche ohne innere Linien,
' Beschriftung nur im ersten Tag (laeuft ueber die leeren Folgezellen der Leiste)
Private Sub DrawBar(ByVal ws As Worksheet, ByVal row As Long, ByVal c0 As Long, ByVal c1 As Long, ByVal rid As String, _
                    ByVal cont As Boolean, ByVal cfid As String, ByVal colW As Single)
    Dim rng As Range, oid As String, bg As Long, fg As Long, c As Long
    Set rng = ws.Range(ws.Cells(row, c0), ws.Cells(row, c1))
    If Len(cfid) > 0 Then oid = NzS(mData.GetValue(mTid, rid, cfid))
    If Len(oid) > 0 Then
        bg = PaletteBg(mCfg.OptColor(oid)): fg = PaletteFg(mCfg.OptColor(oid))
    Else
        bg = CAccentSoft(): fg = Blend(CAccent(), RGB(0, 0, 0), 0.3)
    End If
    rng.Interior.Color = bg
    rng.Font.Color = fg
    rng.Font.Size = 8.5
    rng.NumberFormat = "@"
    rng.IndentLevel = 1
    rng.WrapText = False
    If c1 > c0 Then rng.Borders(xlInsideVertical).Color = bg
    ' Abstand zur naechsten Spur und zur Nachbarleiste
    rng.Borders(xlEdgeBottom).Color = CWhite(): rng.Borders(xlEdgeBottom).Weight = xlMedium
    rng.Borders(xlEdgeTop).Color = CWhite(): rng.Borders(xlEdgeTop).Weight = xlThin
    ws.Cells(row, c0).Value = IIf(cont, ChrW(&H2192) & " ", "") & Truncate(mData.RecordLabel(mTid, rid), CLng((c1 - c0 + 1) * colW / 5.4) - 3)
    For c = c0 To c1
        mCell.SetVal row & "," & c, rid
    Next
End Sub

Private Function ActiveCount(ByVal t As cTableData) As Long
    Dim r As Long
    For r = 1 To t.RowCount
        If Not t.IsArchived(r) Then ActiveCount = ActiveCount + 1
    Next
End Function

Public Sub TextNote(ByVal ws As Worksheet, ByVal row As Long, ByVal txt As String)
    mRender.TextShape ws, "note", txt, mRender.ContentLeft(ws), ws.Rows(row).Top, mRender.PageRight() - mRender.ContentLeft(ws), 30, 8.5, CMuted(), False, 1
End Sub

' Spalten der Terminliste: Datum ausgeschrieben, z. B. "Donnerstag, 10. Dezember 2026" (fett; 150 pt reichten nicht),
' Bezeichnung, Ende, Farbe
Private Sub ListColWidths(ByVal ws As Worksheet)
    mRender.SetColWidthPt ws, 3, 180
    mRender.SetColWidthPt ws, 4, 260
    mRender.SetColWidthPt ws, 5, 120
    mRender.SetColWidthPt ws, 6, 140
End Sub

' Terminliste des Monats (chronologisch)
Private Sub RenderList(ByVal ws As Worksheet, ByVal t As cTableData, rows() As Long, ByVal n As Long, ByVal dfid As String, ByVal efid As String, _
                       ByVal cfid As String, ByVal st As cViewState)
    Dim items As New Collection, k As Long, r As Long, d As Double, m0 As Double, m1 As Double, idx() As Long, keys(1 To 1) As Variant, dirs(1 To 1) As Long
    Dim kv() As Variant, cnt As Long, i As Long, rr As Long, rid As String, prevDay As Double, f As cField, oid As String
    Dim e As Double, hit As Boolean
    m0 = CDbl(st.CalMonth): m1 = CDbl(DateSerial(Year(st.CalMonth), Month(st.CalMonth) + 1, 1))
    ReDim idx(1 To Application.Max(1, n)): ReDim kv(1 To Application.Max(1, t.RowCount))
    For k = 1 To n
        r = rows(k)
        If Not IsBlank(t.Val_(r, dfid)) And (IsDate(t.Val_(r, dfid)) Or IsNumeric(t.Val_(r, dfid))) Then
            d = CDbl(t.Val_(r, dfid))
            hit = (d >= m0 And d < m1)
            ' mehrtägige Einträge, die vor dem Monat beginnen und in ihn hineinreichen, wie in der Monatsansicht
            If Not hit And d < m0 And Len(efid) > 0 Then
                If Not IsBlank(t.Val_(r, efid)) Then
                    If IsDate(t.Val_(r, efid)) Or IsNumeric(t.Val_(r, efid)) Then
                        e = CDbl(t.Val_(r, efid))
                        hit = (Int(e) >= m0)
                    End If
                End If
            End If
            If hit Then
                cnt = cnt + 1
                idx(cnt) = r
                kv(r) = d
            End If
        End If
    Next
    ws.Rows(R_HEAD).RowHeight = 24
    ws.Cells(R_HEAD, 3).Value = "Datum": ws.Cells(R_HEAD, 4).Value = "Bezeichnung"
    ws.Cells(R_HEAD, 5).Value = IIf(Len(efid) > 0, mCfg.FieldName(efid), "")
    ws.Cells(R_HEAD, 6).Value = IIf(Len(cfid) > 0, mCfg.FieldName(cfid), "")
    With ws.Range(ws.Cells(R_HEAD, 3), ws.Cells(R_HEAD, 6))
        .Interior.Color = CSurface(): .Font.Color = CMuted(): .Font.Size = 9.5
    End With
    If cnt = 0 Then
        mRender.EmptyState ws, R_DATA, "Keine Termine in " & MonthName_DE(Month(st.CalMonth)) & " " & Year(st.CalMonth), "Mit den Pfeilen zu einem anderen Monat wechseln oder einen neuen Termin anlegen."
        Exit Sub
    End If
    ReDim Preserve idx(1 To cnt)
    keys(1) = kv: dirs(1) = 1
    MergeSortIdx idx, keys, 1, dirs
    rr = R_DATA
    prevDay = -1
    For i = 1 To cnt
        r = idx(i)
        rid = t.IdAt(r)
        d = CDbl(t.Val_(r, dfid))
        If Int(d) <> prevDay Then
            ws.Cells(rr, 3).Value = DateLong_DE(Int(d))
            ws.Cells(rr, 3).Font.Bold = True
            prevDay = Int(d)
        End If
        ws.Cells(rr, 4).NumberFormat = "@"
        ws.Cells(rr, 4).Value = mData.RecordLabel(mTid, rid)
        If Len(efid) > 0 Then ws.Cells(rr, 5).Value = mFmt.DisplayText(mCfg.Field(efid), t.Val_(r, efid))
        If Len(cfid) > 0 Then
            oid = NzS(t.Val_(r, cfid))
            ws.Cells(rr, 6).Value = mCfg.OptName(oid)
            mViewTable.ColorOptionCell ws.Cells(rr, 6), oid
        End If
        mCell.SetVal rr & ",3", rid: mCell.SetVal rr & ",4", rid: mCell.SetVal rr & ",5", rid: mCell.SetVal rr & ",6", rid
        rr = rr + 1
    Next
    mRender.GridBorders ws.Range(ws.Cells(R_DATA, 3), ws.Cells(rr - 1, 6))
End Sub

Public Function CellTarget(ByVal target As Range) As String
    If mCell Is Nothing Then Exit Function
    CellTarget = CStr(mCell.GetVal(target.Row & "," & target.Column, ""))
End Function

Public Sub OnDoubleClick(ByVal target As Range, ByRef Cancel As Boolean)
    Dim tg As String, d As Date, pre As cMap
    tg = CellTarget(target)
    If Len(tg) = 0 Then Exit Sub
    Cancel = True
    If Left$(tg, 2) = "D:" Then
        d = CDate(CLng(Mid$(tg, 3)))
        If Not mApp.RequireWritable("Das Anlegen") Then Exit Sub
        Set pre = New cMap
        pre.SetVal mCfg.ViewOpt(mVid, "datum"), d
        mForms.NewRecord mTid, pre
    ElseIf Left$(tg, 2) = "M:" Then
        ShowDay CDate(CLng(Mid$(tg, 3)))
    Else
        mForms.OpenRecord mTid, tg
    End If
End Sub

' Alle Eintraege eines Tages bzw. ohne Datum in einer Liste; Doppelklick oeffnet
Public Sub ShowDay(ByVal d As Variant)
    Dim t As cTableData, rows() As Long, n As Long, k As Long, r As Long, dfid As String, efid As String, items() As Variant, cnt As Long
    Dim ids() As String, v As Variant, d0 As Double, d1 As Double, hit As Boolean, sel As Long, title As String
    dfid = mCfg.ViewOpt(mVid, "datum"): efid = mCfg.ViewOpt(mVid, "ende")
    Set t = mData.TData(mTid)
    rows = mViewTable.ViewRows(mVid, n, False)
    ReDim items(0 To Application.Max(0, n - 1), 0 To 1)
    ReDim ids(0 To Application.Max(0, n - 1))
    For k = 1 To n
        r = rows(k)
        v = t.Val_(r, dfid)
        hit = False
        If IsEmpty(d) Then
            hit = IsBlank(v)
        ElseIf Not IsBlank(v) Then
            d0 = Int(CDbl(v)): d1 = d0
            If Len(efid) > 0 Then If Not IsBlank(t.Val_(r, efid)) Then d1 = Int(CDbl(t.Val_(r, efid)))
            hit = (CDbl(d) >= d0 And CDbl(d) <= Application.Max(d0, d1))
        End If
        If hit Then
            items(cnt, 0) = mData.RecordLabel(mTid, t.IdAt(r))
            items(cnt, 1) = mData.RecordHint(mTid, t.IdAt(r))
            ids(cnt) = t.IdAt(r)
            cnt = cnt + 1
        End If
    Next
    If IsEmpty(d) Then title = "Einträge ohne Datum" Else title = "Einträge am " & DateLong_DE(CDate(d))
    If cnt = 0 Then mUI.Alert "Keine Einträge vorhanden.", title: Exit Sub
    Dim out() As Variant, i As Long
    ReDim out(0 To cnt - 1, 0 To 1)
    For i = 0 To cnt - 1: out(i, 0) = items(i, 0): out(i, 1) = items(i, 1): Next
    Dim f As frmMsg
    Set f = New frmMsg
    f.SetupList title, Plural(cnt, "Eintrag", "Einträge") & ". Doppelklick oder „Öffnen“ zeigt den Eintrag.", out, False, "Öffnen", 2, "220;200", 0
    f.Show: mNav.FocusGrid
    sel = -1
    If f.Result = 1 Then sel = f.SelectedIndex
    Unload f
    If sel >= 0 Then mForms.OpenRecord mTid, ids(sel)
End Sub
