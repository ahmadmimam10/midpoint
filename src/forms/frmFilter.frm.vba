'==============================================================================
' frmFilter - Filter mit fachlich passenden Operatoren je Feldtyp, kombinierbar
' mit UND/ODER (UND bindet staerker). Sichtbare Zusammenfassung, "Alle Filter
' zuruecksetzen". Werte werden typgerecht geprueft.
'==============================================================================
Option Explicit

Public OK As Boolean
Public Result As String
Private Const MAXR As Long = 10
Private mTid As String
Private mFlds As Collection
Private mN As Long
Private mIds(1 To MAXR) As String      ' gewaehlte Options-/Datensatz-IDs je Zeile
Private mEvts As Collection
Private mLoading As Boolean
Private lblSum As MSForms.Label
Private lblErr As MSForms.Label
Private WithEvents btnAdd As MSForms.CommandButton
Private WithEvents btnOk As MSForms.CommandButton
Private WithEvents btnClear As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private mY0 As Single

Private Const XC As Single = 0, WC As Single = 64
Private Const XF As Single = 70, WF As Single = 170
Private Const XO As Single = 246, WO As Single = 160
Private Const XV As Single = 412, WV As Single = 130
Private Const XV2 As Single = 548, WV2 As Single = 100
Private Const XP As Single = 654, WP As Single = 64
Private Const XD As Single = 724, WD As Single = 26

Public Sub Setup(ByVal tid As String, ByVal conds As String, ByVal title As String)
    Dim c As Variant, i As Long, y As Single, f As cField, e As cEvt
    mTid = tid
    Set mEvts = New Collection
    Set mFlds = New Collection
    For Each f In mCfg.Fields(tid)
        mFlds.Add f
    Next
    mUI.StyleForm Me
    Me.Width = PAD * 2 + XD + WD + 16
    If Me.Width > mUI.MaxFormWidth Then Me.Width = mUI.MaxFormWidth
    Me.caption = title
    mUI.AddLabel Me, "t", title, PAD, PAD - 2, 500, 20, True, 12
    mUI.AddLabel Me, "t2", "Mehrere Bedingungen: „und“ bindet stärker als „oder“, z. B. (A und B) oder C. Archivierte Einträge sind immer ausgeblendet (Archiv separat).", _
                 PAD, PAD + 20, Me.InsideWidth - PAD * 2, 28, False, 8.5, CMuted()
    mY0 = PAD + 54
    For i = 1 To MAXR
        BuildRow i
    Next
    y = mY0 + MAXR * 30 + 4
    Set btnAdd = mUI.AddButton(Me, "add", "Bedingung hinzufügen", PAD, y, 160, , ROW_H)
    y = y + ROW_H + 10
    Set lblSum = mUI.AddLabel(Me, "sum", "", PAD, y, Me.InsideWidth - PAD * 2, 30, False, 9, CText())
    y = y + 32
    Set lblErr = mUI.AddLabel(Me, "err", "", PAD, y, Me.InsideWidth - PAD * 2, 15, False, 8.5, CDanger())
    y = y + 22
    Set btnOk = mUI.AddButton(Me, "ok", "Übernehmen", PAD, y, 120, "primary")
    Set btnClear = mUI.AddButton(Me, "clr", "Alle Filter zurücksetzen", PAD + 128, y, 170)
    Set btnCancel = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 104, y, 104)
    btnCancel.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    ' vorhandene Bedingungen laden
    mLoading = True
    mN = 0
    For Each c In mQuery.ParseConds(conds)
        If mCfg.FieldExists(CStr(c(1))) Then
            mN = mN + 1
            If mN > MAXR Then Exit For
            LoadRow mN, c
        End If
    Next
    If mN = 0 Then mN = 1
    mLoading = False
    ShowRows
    UpdateSummary
End Sub

Private Sub BuildRow(ByVal i As Long)
    Dim y As Single, cb As MSForms.ComboBox, tb As MSForms.TextBox, b As MSForms.CommandButton, f As cField
    y = mY0 + (i - 1) * 30
    Set cb = mUI.AddCombo(Me, "c" & i, PAD + XC, y, WC)
    cb.AddItem "und": cb.AddItem "oder": cb.ListIndex = 0
    Bind cb, "c:" & i
    Set cb = mUI.AddCombo(Me, "f" & i, PAD + XF, y, WF)
    For Each f In mFlds
        cb.AddItem f.Name
    Next
    Bind cb, "f:" & i
    Set cb = mUI.AddCombo(Me, "o" & i, PAD + XO, y, WO)
    Bind cb, "o:" & i
    Set tb = mUI.AddText(Me, "v" & i, PAD + XV, y, WV)
    Bind tb, "v:" & i
    Set tb = mUI.AddText(Me, "w" & i, PAD + XV2, y, WV2)
    Bind tb, "v:" & i
    Set b = mUI.AddButton(Me, "p" & i, "Wählen…", PAD + XP, y, WP, , ROW_H)
    Bind b, "p:" & i
    Set b = mUI.AddButton(Me, "x" & i, ChrW(&H2715), PAD + XD, y, WD, , ROW_H)
    b.ControlTipText = "Bedingung entfernen"
    Bind b, "x:" & i
End Sub

Private Sub Bind(ByVal c As Object, ByVal key As String)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, c, key
    mEvts.Add e
End Sub

Private Function FieldAt(ByVal i As Long) As cField
    Dim idx As Long
    idx = Me.Controls("f" & i).ListIndex
    If idx >= 0 Then Set FieldAt = mFlds(idx + 1)
End Function

Private Function OpAt(ByVal i As Long) As String
    Dim f As cField, idx As Long, ops As Variant
    Set f = FieldAt(i)
    If f Is Nothing Then Exit Function
    idx = Me.Controls("o" & i).ListIndex
    ops = mQuery.OpsFor(mQuery.OpKind(f))
    If idx >= 0 And idx <= UBound(ops) Then OpAt = ops(idx)
End Function

Private Sub FillOps(ByVal i As Long)
    Dim f As cField, cb As MSForms.ComboBox, v As Variant
    Set f = FieldAt(i)
    Set cb = Me.Controls("o" & i)
    cb.Clear
    If f Is Nothing Then Exit Sub
    For Each v In mQuery.OpsFor(mQuery.OpKind(f))
        cb.AddItem mQuery.OpLabel(CStr(v))
    Next
    cb.ListIndex = 0
End Sub

Private Sub LoadRow(ByVal i As Long, ByVal c As Variant)
    Dim f As cField, k As Long, ops As Variant, j As Long, kind As String
    For k = 1 To mFlds.Count
        If mFlds(k).ID = CStr(c(1)) Then Me.Controls("f" & i).ListIndex = k - 1: Exit For
    Next
    Me.Controls("c" & i).ListIndex = IIf(c(0) = "ODER", 1, 0)
    FillOps i
    Set f = FieldAt(i)
    kind = mQuery.OpKind(f)
    ops = mQuery.OpsFor(kind)
    For j = 0 To UBound(ops)
        If ops(j) = c(2) Then Me.Controls("o" & i).ListIndex = j
    Next
    Select Case kind
        Case "opt", "status", "multi", "link"
            mIds(i) = CStr(c(3))
            Me.Controls("v" & i).Text = IdNames(f, mIds(i))
        Case "num"
            If Len(c(3)) > 0 Then Me.Controls("v" & i).Text = FmtNum(InvNum(CStr(c(3))), 6, True)
            If Len(c(4)) > 0 Then Me.Controls("w" & i).Text = FmtNum(InvNum(CStr(c(4))), 6, True)
        Case "date"
            If Len(c(3)) > 0 Then Me.Controls("v" & i).Text = FmtDate(InvDate(CStr(c(3))))
            If Len(c(4)) > 0 Then Me.Controls("w" & i).Text = FmtDate(InvDate(CStr(c(4))))
        Case Else
            Me.Controls("v" & i).Text = CStr(c(3))
    End Select
End Sub

Private Function IdNames(ByVal f As cField, ByVal ids As String) As String
    Dim a As Variant, i As Long, s As String, tt As String
    If Len(ids) = 0 Then Exit Function
    a = Split(ids, ",")
    If f.IsLinkType Then
        If f.FType = FT_LINK Then tt = f.TargetTable Else tt = mCfg.Field(f.PairField).TableID
    End If
    For i = 0 To UBound(a)
        If f.IsLinkType Then
            s = s & IIf(Len(s) > 0, ", ", "") & mData.RecordLabel(tt, CStr(a(i)))
        Else
            s = s & IIf(Len(s) > 0, ", ", "") & mCfg.OptName(CStr(a(i)))
        End If
    Next
    IdNames = s
End Function

Private Sub ShowRows()
    Dim i As Long, vis As Boolean, f As cField, kind As String, ar As Long, op As String
    For i = 1 To MAXR
        vis = (i <= mN)
        Me.Controls("c" & i).Visible = vis And i > 1
        Me.Controls("f" & i).Visible = vis
        Me.Controls("o" & i).Visible = vis
        Me.Controls("x" & i).Visible = vis
        Set f = Nothing
        If vis Then Set f = FieldAt(i)
        op = ""
        If Not f Is Nothing Then op = OpAt(i)
        ar = mQuery.OpArity(op)
        kind = ""
        If Not f Is Nothing Then kind = mQuery.OpKind(f)
        Me.Controls("v" & i).Visible = vis And Not f Is Nothing And ar >= 1
        Me.Controls("w" & i).Visible = vis And Not f Is Nothing And ar = 2
        Me.Controls("p" & i).Visible = vis And Not f Is Nothing And ar >= 1 And (kind = "opt" Or kind = "status" Or kind = "multi" Or kind = "link" Or kind = "date")
        Me.Controls("p" & i).caption = IIf(kind = "date", "Datum…", "Wählen…")
        Me.Controls("v" & i).Locked = (kind = "opt" Or kind = "status" Or kind = "multi" Or kind = "link")
        Me.Controls("v" & i).BackColor = IIf(Me.Controls("v" & i).Locked, CSurface(), CWhite())
    Next
    btnAdd.Enabled = (mN < MAXR)
End Sub

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    Dim kind As String, i As Long, k As Long
    kind = Left$(key, 1)
    i = CLng(Mid$(key, 3))
    If mLoading Then Exit Sub
    Select Case kind
        Case "f"
            If evt = "Change" Then
                FillOps i
                mIds(i) = ""
                Me.Controls("v" & i).Text = "": Me.Controls("w" & i).Text = ""
                ShowRows
            End If
        Case "o", "c"
            If evt = "Change" Then ShowRows
        Case "p"
            If evt = "Click" Then PickValue i
        Case "x"
            If evt = "Click" Then RemoveRow i
    End Select
    UpdateSummary
End Sub

Private Sub PickValue(ByVal i As Long)
    Dim f As cField, kind As String, items() As Variant, ids As Collection, v As Variant, n As Long, pre As String, res As String, ok As Boolean
    Dim a As Variant, j As Long, tt As String
    Set f = FieldAt(i)
    If f Is Nothing Then Exit Sub
    kind = mQuery.OpKind(f)
    Select Case kind
        Case "date"
            res = mForms.PickDate(Me.Controls("v" & i).Text, False, f.Name, ok)
            If ok Then Me.Controls("v" & i).Text = res
            If OpAt(i) = "ZWISCHEN" And ok Then
                res = mForms.PickDate(Me.Controls("w" & i).Text, False, f.Name & " (bis)", ok)
                If ok Then Me.Controls("w" & i).Text = res
            End If
        Case "link"
            If f.FType = FT_LINK Then tt = f.TargetTable Else tt = mCfg.Field(f.PairField).TableID
            res = mForms.PickRecords(tt, Replace(mIds(i), ",", ";"), True, "Verknüpfte Einträge für den Filter", ok)
            If ok Then
                mIds(i) = Replace(res, ";", ",")
                Me.Controls("v" & i).Text = IdNames(f, mIds(i))
            End If
        Case Else
            Set ids = mCfg.OptionIds(f.ID)
            If ids.Count = 0 Then mUI.Alert "Dieses Feld hat noch keine Auswahlmöglichkeiten.", "Filter": Exit Sub
            ReDim items(0 To ids.Count - 1)
            n = 0
            For Each v In ids
                items(n) = mCfg.OptName(CStr(v))
                If ListHas(mIds(i), CStr(v), ",") Then pre = pre & IIf(Len(pre) > 0, ";", "") & n
                n = n + 1
            Next
            res = mUI.ChooseMany("Werte für „" & f.Name & "“ wählen (eine oder mehrere):", "Filter", items, pre, ok)
            If ok Then
                mIds(i) = ""
                For Each a In Split(res, ";")
                    If Len(a) > 0 Then mIds(i) = mIds(i) & IIf(Len(mIds(i)) > 0, ",", "") & ids(CLng(a) + 1)
                Next
                Me.Controls("v" & i).Text = IdNames(f, mIds(i))
            End If
    End Select
End Sub

Private Sub RemoveRow(ByVal i As Long)
    Dim j As Long
    mLoading = True
    For j = i To mN - 1
        Me.Controls("c" & j).ListIndex = Me.Controls("c" & (j + 1)).ListIndex
        Me.Controls("f" & j).ListIndex = Me.Controls("f" & (j + 1)).ListIndex
        FillOps j
        Me.Controls("o" & j).ListIndex = Me.Controls("o" & (j + 1)).ListIndex
        Me.Controls("v" & j).Text = Me.Controls("v" & (j + 1)).Text
        Me.Controls("w" & j).Text = Me.Controls("w" & (j + 1)).Text
        mIds(j) = mIds(j + 1)
    Next
    Me.Controls("f" & mN).ListIndex = -1
    Me.Controls("o" & mN).Clear
    Me.Controls("v" & mN).Text = "": Me.Controls("w" & mN).Text = ""
    mIds(mN) = ""
    mN = mN - 1
    If mN < 1 Then mN = 1
    mLoading = False
    ShowRows
    UpdateSummary
End Sub

' Baut die Bedingungen; liefert False mit Meldung bei ungueltigen Werten
Private Function BuildConds(ByRef conds As Collection, ByRef msg As String) As Boolean
    Dim i As Long, f As cField, op As String, kind As String, v1 As String, v2 As String, d As Double, dt As Date, ht As Boolean
    Set conds = New Collection
    For i = 1 To mN
        Set f = FieldAt(i)
        If Not f Is Nothing Then
            op = OpAt(i)
            kind = mQuery.OpKind(f)
            v1 = "": v2 = ""
            If mQuery.OpArity(op) >= 1 Then
                Select Case kind
                    Case "opt", "status", "multi", "link"
                        v1 = mIds(i)
                        If Len(v1) = 0 Then msg = "Zeile " & i & ": Bitte über „Wählen…“ mindestens einen Wert auswählen.": Exit Function
                    Case "num"
                        If Not ParseNumberDE(Me.Controls("v" & i).Text, d) Then msg = "Zeile " & i & ": Bitte eine Zahl eingeben.": Exit Function
                        v1 = NumInv(d)
                        If mQuery.OpArity(op) = 2 Then
                            If Not ParseNumberDE(Me.Controls("w" & i).Text, d) Then msg = "Zeile " & i & ": Bitte eine zweite Zahl (bis) eingeben.": Exit Function
                            v2 = NumInv(d)
                        End If
                    Case "date"
                        If Not ParseDateDE(Me.Controls("v" & i).Text, dt, ht) Then msg = "Zeile " & i & ": Bitte ein Datum TT.MM.JJJJ eingeben.": Exit Function
                        v1 = DateInv(Int(dt))
                        If mQuery.OpArity(op) = 2 Then
                            If Not ParseDateDE(Me.Controls("w" & i).Text, dt, ht) Then msg = "Zeile " & i & ": Bitte ein Enddatum (bis) eingeben.": Exit Function
                            v2 = DateInv(Int(dt))
                        End If
                    Case Else
                        v1 = Me.Controls("v" & i).Text
                        If Len(Trim$(v1)) = 0 Then msg = "Zeile " & i & ": Bitte einen Suchtext eingeben.": Exit Function
                End Select
            End If
            conds.Add Array(IIf(conds.Count = 0, "", IIf(Me.Controls("c" & i).ListIndex = 1, "ODER", "UND")), f.ID, op, v1, v2)
        End If
    Next
    BuildConds = True
End Function

Private Sub UpdateSummary()
    Dim conds As Collection, msg As String
    If BuildConds(conds, msg) Then
        If conds.Count = 0 Then
            lblSum.caption = "Kein Filter: alle aktiven Einträge werden gezeigt."
        Else
            lblSum.caption = "Zusammenfassung: " & mQuery.CondsText(conds)
        End If
        lblErr.caption = ""
    Else
        lblSum.caption = "Zusammenfassung: (unvollständig)"
    End If
End Sub

Private Sub btnAdd_Click()
    If mN < MAXR Then
        mN = mN + 1
        ShowRows
    End If
End Sub

Private Sub btnOk_Click()
    Dim conds As Collection, msg As String
    If Not BuildConds(conds, msg) Then
        lblErr.caption = ChrW(&H26A0) & "  " & msg
        Exit Sub
    End If
    Result = mQuery.SerializeConds(conds)
    OK = True
    Me.Hide
End Sub

Private Sub btnClear_Click()
    Result = ""
    OK = True
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    OK = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True: btnCancel_Click
End Sub

Private Sub UserForm_Terminate()
    Dim e As cEvt
    If Not mEvts Is Nothing Then
        For Each e In mEvts: e.Unbind: Next
    End If
End Sub
