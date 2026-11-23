Attribute VB_Name = "mViewGantt"
'==============================================================================
' mViewGantt - Projektleiste mit horizontaler Zeitachse (Tag/Woche/Monat),
' Heute-Markierung und Zeitraumwahl. Optionen: start=Feld, ende=Feld,
' fortschritt=Feld (Prozent, optional), farbe=Auswahl-/Statusfeld (optional).
' Eintraege ohne vollstaendige oder mit ungueltigen Zeitangaben separat erreichbar.
'==============================================================================
Option Explicit

Private Const COL_T As Long = 3      ' Bezeichnung
Private Const COL_S As Long = 4      ' Beginn
Private Const COL_E As Long = 5      ' Ende
Private Const COL_0 As Long = 6      ' erste Zeitspalte
Private mRowRid As cMap
Private mTid As String
Private mVid As String
Private mIncomplete As Collection
Private mSc As String, mP0 As Date, mPEnd As Date, mNu As Long, mWs As Worksheet

Private Function UnitCount(ByVal sc As String) As Long
    Select Case sc
        Case "Tag": UnitCount = 42
        Case "Monat": UnitCount = 18
        Case Else: UnitCount = 26
    End Select
End Function

Private Function UnitWidth(ByVal sc As String) As Single
    Select Case sc
        Case "Tag": UnitWidth = 22
        Case "Monat": UnitWidth = 56
        Case Else: UnitWidth = 34
    End Select
End Function

' Beginn der Einheit i (0-basiert)
Private Function UnitStart(ByVal sc As String, ByVal p0 As Date, ByVal i As Long) As Date
    Select Case sc
        Case "Tag": UnitStart = p0 + i
        Case "Monat": UnitStart = DateSerial(Year(p0), Month(p0) + i, 1)
        Case Else: UnitStart = p0 + i * 7
    End Select
End Function

Public Sub Render()
    Dim ws As Worksheet, tid As String, vid As String, st As cViewState, sfid As String, efid As String, pfid As String, cfid As String
    Dim t As cTableData, rows() As Long, n As Long, k As Long, r As Long, sr As Range, sc As String, nu As Long, uw As Single, i As Long
    Dim p0 As Date, pEnd As Date, d0 As Double, d1 As Double, rr As Long, u0 As Long, u1 As Long, rid As String, cell As Range
    Dim oid As String, bg As Long, fg As Long, prog As Double, nProg As Long, outside As Long, todayCol As Long, lastCol As Long
    Dim us As Date, ue As Date, shown As Long, maxRows As Long, sortKeys() As Variant, idx() As Long, cnt As Long, keys(1 To 1) As Variant, dirs(1 To 1) As Long
    Dim showNames As Boolean

    tid = mNav.gTid: vid = mNav.gVid: mTid = tid: mVid = vid
    Set st = mNav.State(vid)
    sfid = mCfg.ViewOpt(vid, "start"): efid = mCfg.ViewOpt(vid, "ende")
    pfid = mCfg.ViewOpt(vid, "fortschritt"): cfid = mCfg.ViewOpt(vid, "farbe")
    sc = st.GanttScale
    If sc <> "Tag" And sc <> "Monat" Then sc = "Woche"
    If sc = "Woche" Then st.GanttStart = st.GanttStart - Weekday(st.GanttStart, vbMonday) + 1
    If sc = "Monat" Then st.GanttStart = DateSerial(Year(st.GanttStart), Month(st.GanttStart), 1)
    p0 = st.GanttStart
    nu = UnitCount(sc): uw = UnitWidth(sc)
    pEnd = UnitStart(sc, p0, nu) - 1
    Set ws = mRender.Sheet(SH_GANTT)
    mRender.BeginPage ws
    Set mRowRid = New cMap
    Set mIncomplete = New Collection
    mRender.SetColWidthPt ws, COL_T, 210
    mRender.SetColWidthPt ws, COL_S, 74
    mRender.SetColWidthPt ws, COL_E, 74
    For i = 0 To nu - 1
        mRender.SetColWidthPt ws, COL_0 + i, uw
    Next
    lastCol = COL_0 + nu - 1
    mRender.DrawTitle ws, mCfg.TableIcon(tid), mCfg.TableName(tid), mCfg.GetS(SH_TAB, tid, "Beschreibung")
    mViewTable.DrawViewTabs ws, tid, vid, "Neuer Eintrag"
    If Not mCfg.FieldExists(sfid) Or Not mCfg.FieldExists(efid) Then
        mRender.EmptyState ws, R_DATA, "Projektleiste noch nicht eingerichtet", "Bitte unter „Ansicht > Ansicht bearbeiten“ Beginn- und Endfeld wählen.", _
                           "Ansicht bearbeiten", "mCmd.Cmd_EditView", ChrW(&HE9D5)
        mRender.EndPage ws
        Exit Sub
    End If

    mRender.FlowBegin ws, R_TOOLS, mRender.ContentLeft(ws), mRender.PageRight()
    mRender.FlowBtn "b_prev", "", "mCmd.Cmd_GanttPrev", "secondary", ChrW(&HE76B), 2
    mRender.FlowBtn "b_today", "Heute", "mCmd.Cmd_GanttToday", "secondary", "", 2
    mRender.FlowBtn "b_next", "", "mCmd.Cmd_GanttNext", "secondary", ChrW(&HE76C), 8
    mRender.FlowBtn "sc_Tag", "Tage", "mCmd.Cmd_GanttScale", IIf(sc = "Tag", "tabActive", "tab"), "", 2
    mRender.FlowBtn "sc_Woche", "Wochen", "mCmd.Cmd_GanttScale", IIf(sc = "Woche", "tabActive", "tab"), "", 2
    mRender.FlowBtn "sc_Monat", "Monate", "mCmd.Cmd_GanttScale", IIf(sc = "Monat", "tabActive", "tab"), "", 10
    mRender.FlowBtn "b_names", "Namen auf Leisten", "mCmd.Cmd_GanttNames", IIf(mCfg.ViewOpt(vid, "namen", "Nein") = "Ja", "tabActive", "tab"), _
                    IIf(mCfg.ViewOpt(vid, "namen", "Nein") = "Ja", ChrW(&HE73E), ""), 10
    mRender.FlowBtn "b_filter", "Filter", "mCmd.Cmd_Filter", "secondary", ChrW(&HE71C)
    mRender.FlowBtn "b_view", "Ansicht", "mCmd.Cmd_ViewMenu", "secondary", ChrW(&HE8FD)

    ' Daten
    Set t = mData.TData(tid)
    rows = mViewTable.ViewRows(vid, n, False)
    ReDim idx(1 To Application.Max(1, n)): ReDim sortKeys(1 To Application.Max(1, t.RowCount))
    For k = 1 To n
        r = rows(k)
        If IsBlank(t.Val_(r, sfid)) Or IsBlank(t.Val_(r, efid)) Then
            mIncomplete.Add Array(t.IdAt(r), "Beginn oder Ende fehlt")
        ElseIf Not (IsNumeric(t.Val_(r, sfid)) Or IsDate(t.Val_(r, sfid))) Or Not (IsNumeric(t.Val_(r, efid)) Or IsDate(t.Val_(r, efid))) Then
            mIncomplete.Add Array(t.IdAt(r), "Ungültige Datumsangabe")
        ElseIf CDbl(t.Val_(r, efid)) < CDbl(t.Val_(r, sfid)) Then
            mIncomplete.Add Array(t.IdAt(r), "Ungültig: Beginn liegt nach dem Ende")
        Else
            d0 = Int(CDbl(t.Val_(r, sfid))): d1 = Int(CDbl(t.Val_(r, efid)))
            If d1 < CDbl(p0) Or d0 > CDbl(pEnd) Then
                outside = outside + 1
            Else
                cnt = cnt + 1
                idx(cnt) = r
                sortKeys(r) = d0
            End If
        End If
    Next
    mRender.FlowBtn "b_incomplete", "Ohne vollständigen Zeitraum (" & mIncomplete.Count & ")", "mCmd.Cmd_GanttIncomplete", "secondary", ChrW(&HE946)
    mViewTable.DrawSetupTools
    mRender.FlowEnd
    mViewTable.DrawChips ws, st, n, n, "Einträge im Filter"

    ' Zeitachse (zwei Kopfzeilen)
    ws.Rows(R_GAP).RowHeight = 18
    ws.Rows(R_HEAD).RowHeight = 22
    ws.Cells(R_HEAD, COL_T).Value = "Bezeichnung"
    ws.Cells(R_HEAD, COL_S).Value = mCfg.FieldName(sfid)
    ws.Cells(R_HEAD, COL_E).Value = mCfg.FieldName(efid)
    todayCol = 0
    For i = 0 To nu - 1
        us = UnitStart(sc, p0, i): ue = UnitStart(sc, p0, i + 1) - 1
        Set cell = ws.Cells(R_HEAD, COL_0 + i)
        Select Case sc
            Case "Tag"
                cell.Value = CStr(Day(us))
                If Day(us) = 1 Or i = 0 Then ws.Cells(R_GAP, COL_0 + i).Value = MonthName_DE(Month(us)) & " " & Year(us)
                If Weekday(us, vbMonday) >= 6 Then ws.Range(ws.Cells(R_HEAD, COL_0 + i), ws.Cells(R_HEAD, COL_0 + i)).Interior.Color = CSurface2()
            Case "Monat"
                cell.Value = Left$(MonthName_DE(Month(us)), 3)
                If Month(us) = 1 Or i = 0 Then ws.Cells(R_GAP, COL_0 + i).Value = CStr(Year(us))
            Case Else
                cell.Value = "KW" & IsoWeek(us)
                If i = 0 Or Month(us) <> Month(us - 7) Then ws.Cells(R_GAP, COL_0 + i).Value = Left$(MonthName_DE(Month(us)), 3) & " " & Year(us)
        End Select
        cell.HorizontalAlignment = xlCenter
        cell.Font.Size = 8
        If Date >= us And Date <= ue Then todayCol = COL_0 + i
    Next
    With ws.Range(ws.Cells(R_HEAD, COL_T), ws.Cells(R_HEAD, lastCol))
        .Font.Color = CMuted(): .Interior.Color = CSurface()
        .Borders(xlEdgeBottom).Color = CBorderStrong()
    End With
    ws.Range(ws.Cells(R_HEAD, COL_T), ws.Cells(R_HEAD, COL_E)).Font.Size = 9
    ws.Range(ws.Cells(R_GAP, COL_0), ws.Cells(R_GAP, lastCol)).Font.Size = 8.5
    ws.Range(ws.Cells(R_GAP, COL_0), ws.Cells(R_GAP, lastCol)).Font.Color = CMuted()

    ' Balken
    mSc = sc: mP0 = p0: mPEnd = pEnd: mNu = nu: Set mWs = ws
    showNames = (mCfg.ViewOpt(vid, "namen", "Nein") = "Ja")
    maxRows = 300
    If cnt > 0 Then
        ReDim Preserve idx(1 To cnt)
        keys(1) = sortKeys: dirs(1) = 1
        MergeSortIdx idx, keys, 1, dirs
    End If
    rr = R_DATA
    For k = 1 To cnt
        If shown >= maxRows Then Exit For
        r = idx(k)
        rid = t.IdAt(r)
        d0 = Int(CDbl(t.Val_(r, sfid))): d1 = Int(CDbl(t.Val_(r, efid)))
        ws.Rows(rr).RowHeight = 22
        ws.Cells(rr, COL_T).NumberFormat = "@"
        ws.Cells(rr, COL_T).Value = Truncate(mData.RecordLabel(tid, rid), 36)
        ws.Cells(rr, COL_T).IndentLevel = 1
        ws.Cells(rr, COL_S).Value = FmtDate(d0)
        ws.Cells(rr, COL_E).Value = FmtDate(d1)
        ws.Range(ws.Cells(rr, COL_S), ws.Cells(rr, COL_E)).Font.Size = 8.5
        ws.Range(ws.Cells(rr, COL_S), ws.Cells(rr, COL_E)).Font.Color = CMuted()
        If cfid <> "" Then oid = NzS(t.Val_(r, cfid)) Else oid = ""
        If Len(oid) > 0 Then
            bg = PaletteBg(mCfg.OptColor(oid)): fg = PaletteFg(mCfg.OptColor(oid))
        Else
            bg = CAccentSoft(): fg = Blend(CAccent(), RGB(0, 0, 0), 0.3)
        End If
        prog = -1
        If Len(pfid) > 0 Then
            If IsNumeric(t.Val_(r, pfid)) And Not IsBlank(t.Val_(r, pfid)) Then
                prog = CDbl(t.Val_(r, pfid))
                If prog > 1 Then prog = 1
                If prog < 0 Then prog = 0
            End If
        End If
        ' Leiste tagesgenau als Form (auch in Wochen- und Monatsansicht nur der belegte Teil)
        DrawBar ws, rr, rid, d0, d1, bg, fg, prog, showNames
        mRowRid.SetVal CStr(rr), rid
        rr = rr + 1
        shown = shown + 1
    Next
    If rr > R_DATA Then
        mRender.GridBorders ws.Range(ws.Cells(R_DATA, COL_T), ws.Cells(rr - 1, COL_E))
        With ws.Range(ws.Cells(R_DATA, COL_0), ws.Cells(rr - 1, lastCol))
            .Borders(xlInsideVertical).Color = CBorder()
            .Borders(xlEdgeLeft).Color = CBorder()
            .Borders(xlEdgeRight).Color = CBorder()
        End With
        ws.Range(ws.Cells(R_DATA, COL_0), ws.Cells(rr - 1, lastCol)).HorizontalAlignment = xlCenter
    Else
        mRender.EmptyState ws, R_DATA, "Keine Einträge im gewählten Zeitraum", "Mit den Pfeilen den Zeitraum verschieben oder den Maßstab ändern.", "", "", ChrW(&HE9D5)
    End If
    If todayCol > 0 Then
        ' Heute-Linie genau am heutigen Tag (Mitte des Tages), auch in Wochen/Monaten
        Dim ln As Shape, tx As Single
        tx = XOf(CDbl(Date) + 0.5)
        Set ln = ws.Shapes.AddLine(tx, ws.Rows(R_GAP).Top + 2, tx, ws.Rows(Application.Max(rr - 1, R_DATA)).Top + ws.Rows(Application.Max(rr - 1, R_DATA)).Height)
        ln.Name = "g_today"
        ln.Line.ForeColor.RGB = CDanger(): ln.Line.Weight = 1.5
        ln.Placement = xlFreeFloating
        ln.Locked = True
        ws.Cells(R_HEAD, todayCol).Font.Color = CDanger(): ws.Cells(R_HEAD, todayCol).Font.Bold = True
    End If
    mViewCal.TextNote ws, rr + 1, "Zeitraum " & FmtDate(p0) & " bis " & FmtDate(pEnd) & ".  " & _
                      IIf(outside > 0, outside & " Einträge liegen außerhalb dieses Zeitraums (Pfeile nutzen).  ", "") & _
                      IIf(shown >= maxRows And cnt > maxRows, "Angezeigt werden die ersten " & maxRows & " Einträge; bitte filtern.  ", "") & _
                      "Klick auf eine Leiste oder Doppelklick auf die Zeile öffnet den Eintrag. Rote Linie = heute."
    mRender.EndPage ws, R_HEAD, COL_E
End Sub

' Waagerechte Position eines Zeitpunkts (Tagesanfang = ganze Zahl) in der Zeitachse.
' Innerhalb einer Woche/eines Monats anteilig, damit Leisten tagesgenau sind.
Private Function XOf(ByVal d As Double) As Single
    Dim i As Long, us As Double, ue As Double, c As Range
    If d <= CDbl(mP0) Then XOf = mWs.Columns(COL_0).Left: Exit Function
    If d >= CDbl(mPEnd) + 1 Then XOf = mWs.Columns(COL_0 + mNu - 1).Left + mWs.Columns(COL_0 + mNu - 1).Width: Exit Function
    Select Case mSc
        Case "Tag": i = Int(d - CDbl(mP0))
        Case "Monat": i = (Year(d) - Year(mP0)) * 12 + Month(d) - Month(mP0)
        Case Else: i = Int((d - CDbl(mP0)) / 7)
    End Select
    us = CDbl(UnitStart(mSc, mP0, i)): ue = CDbl(UnitStart(mSc, mP0, i + 1))
    Set c = mWs.Columns(COL_0 + i)
    XOf = c.Left + c.Width * (d - us) / (ue - us)
End Function

' Leiste einer Zeile: abgerundete Form vom Beginn bis Ende des letzten Tages; optional
' Fortschritt (dunklerer Anteil) und Name. Passt der Name nicht in die Leiste, steht er
' dezent rechts daneben, bei Platzmangel gekuerzt mit "…".
Private Sub DrawBar(ByVal ws As Worksheet, ByVal row As Long, ByVal rid As String, ByVal d0 As Double, ByVal d1 As Double, _
                    ByVal bg As Long, ByVal fg As Long, ByVal prog As Double, ByVal showNames As Boolean)
    Const BAR_H As Single = 13
    Const CW As Single = 4.1          ' mittlere Zeichenbreite bei 7,5 pt
    Dim x0 As Single, x1 As Single, y As Single, s As Shape, p As Shape, txt As String, need As Single, room As Single, right_ As Single
    x0 = XOf(d0): x1 = XOf(d1 + 1)
    If x1 - x0 < 3 Then x1 = x0 + 3
    y = ws.Rows(row).Top + (ws.Rows(row).Height - BAR_H) / 2
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x0 + 0.5, y, x1 - x0 - 1, BAR_H)
    s.Name = "gb_" & rid
    s.Placement = xlFreeFloating
    On Error Resume Next
    s.Adjustments.Item(1) = 0.35
    On Error GoTo 0
    s.Fill.ForeColor.RGB = bg
    s.Line.Visible = msoFalse
    s.OnAction = "mCmd.Cmd_GanttBar"
    If prog > 0 Then
        Set p = ws.Shapes.AddShape(msoShapeRoundedRectangle, x0 + 0.5, y, Application.Max(3, (x1 - x0 - 1) * prog), BAR_H)
        p.Name = "gp_" & rid
        p.Placement = xlFreeFloating
        On Error Resume Next
        p.Adjustments.Item(1) = 0.35
        On Error GoTo 0
        p.Fill.ForeColor.RGB = Blend(bg, fg, 0.35)
        p.Line.Visible = msoFalse
        p.OnAction = "mCmd.Cmd_GanttBar"
    End If
    If showNames Then txt = mData.RecordLabel(mTid, rid)
    If prog >= 0 Then txt = txt & IIf(Len(txt) > 0, "  · ", "") & FmtNum(prog * 100, 0) & " %"
    If Len(txt) = 0 Then Exit Sub
    need = Len(txt) * CW + 10
    right_ = ws.Columns(COL_0 + mNu - 1).Left + ws.Columns(COL_0 + mNu - 1).Width
    If need <= x1 - x0 Or right_ - x1 < 40 Then
        ' in der Leiste (gekuerzt, wenn rechts kein Platz ist)
        If need > x1 - x0 Then txt = FitText(txt, (x1 - x0 - 10) / CW)
        If Len(txt) = 0 Then Exit Sub
        BarText ws, "gt_" & rid, txt, x0 + 5, y, x1 - x0 - 8, fg
    Else
        ' rechts neben der Leiste, dezent
        room = right_ - x1 - 6
        If need - 10 > room Then txt = FitText(txt, room / CW)
        BarText ws, "gt_" & rid, txt, x1 + 4, y, room, CMuted()
    End If
End Sub

Private Sub BarText(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, ByVal x As Single, ByVal y As Single, _
                    ByVal w As Single, ByVal color As Long)
    Dim s As Shape
    Set s = mRender.TextShape(ws, nm, txt, x, y - 1, w, 15, 7.5, color, False, 1)
    s.TextFrame2.WordWrap = msoFalse
    s.OnAction = "mCmd.Cmd_GanttBar"
End Sub

' Kuerzt auf hoechstens n Zeichen mit "…"
Private Function FitText(ByVal s As String, ByVal n As Single) As String
    If n < 3 Then Exit Function
    If Len(s) <= Int(n) Then FitText = s Else FitText = Left$(s, Int(n) - 1) & ChrW(&H2026)
End Function

' Klick auf Leiste, Fortschritt oder Namen
Public Sub OpenFromShape(ByVal nm As String)
    If Len(nm) < 4 Then Exit Sub
    mForms.OpenRecord mTid, Mid$(nm, 4)
End Sub

Public Sub OnDoubleClick(ByVal target As Range, ByRef Cancel As Boolean)
    Dim rid As String
    If mRowRid Is Nothing Then Exit Sub
    rid = CStr(mRowRid.GetVal(CStr(target.Row), ""))
    If Len(rid) = 0 Then Exit Sub
    Cancel = True
    mForms.OpenRecord mTid, rid
End Sub

Public Function SelectedRid() As String
    On Error Resume Next
    If mRowRid Is Nothing Then Exit Function
    SelectedRid = CStr(mRowRid.GetVal(CStr(Selection.Row), ""))
End Function

Public Sub ShowIncomplete()
    Dim items() As Variant, i As Long, v As Variant, sel As Long, ids() As String, f As frmMsg
    If mIncomplete Is Nothing Then Exit Sub
    If mIncomplete.Count = 0 Then mUI.Alert "Alle gefilterten Einträge haben einen vollständigen, gültigen Zeitraum.", "Projektleiste": Exit Sub
    ReDim items(0 To mIncomplete.Count - 1, 0 To 1)
    ReDim ids(0 To mIncomplete.Count - 1)
    For Each v In mIncomplete
        items(i, 0) = mData.RecordLabel(mTid, CStr(v(0)))
        items(i, 1) = v(1)
        ids(i) = v(0)
        i = i + 1
    Next
    Set f = New frmMsg
    f.SetupList "Einträge ohne vollständigen Zeitraum", "Diese Einträge erscheinen nicht in der Zeitachse. Doppelklick oder „Öffnen“, um Beginn und Ende zu ergänzen.", _
                items, False, "Öffnen", 2, "240;220", 0
    f.Show: mNav.FocusGrid
    sel = -1
    If f.Result = 1 Then sel = f.SelectedIndex
    Unload f
    If sel >= 0 Then mForms.OpenRecord mTid, ids(sel)
End Sub

Public Sub Shift(ByVal dir_ As Long)
    Dim st As cViewState
    Set st = mNav.State(mNav.gVid)
    Select Case st.GanttScale
        Case "Tag": st.GanttStart = st.GanttStart + dir_ * 28
        Case "Monat": st.GanttStart = DateSerial(Year(st.GanttStart), Month(st.GanttStart) + dir_ * 6, 1)
        Case Else: st.GanttStart = st.GanttStart + dir_ * 7 * 12
    End Select
End Sub
