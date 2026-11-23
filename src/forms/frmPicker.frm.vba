'==============================================================================
' frmPicker - Auswahl verknuepfter Eintraege: Suche, Zusatzmerkmale zur
' Unterscheidung gleichnamiger Eintraege, Einfach-/Mehrfachauswahl und
' Schnellanlage eines neuen Zieleintrags (Rueckkehr ohne Verlust der Auswahl).
'==============================================================================
Option Explicit

Public Result As String
Public OK As Boolean
Private mTid As String
Private mMulti As Boolean
Private mSel As String
Private mIds() As String
Private mLabels() As String
Private mN As Long
Private mShown() As String
Private mLoading As Boolean
Private WithEvents txtSearch As MSForms.TextBox
Private WithEvents lst As MSForms.ListBox
Private WithEvents chkArch As MSForms.CheckBox
Private WithEvents btnOk As MSForms.CommandButton
Private WithEvents btnNew As MSForms.CommandButton
Private WithEvents btnClear As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private lblInfo As MSForms.Label
Private lblSel As MSForms.Label

Public Sub Setup(ByVal tid As String, ByVal cur As String, ByVal multi As Boolean, ByVal title As String)
    Dim w As Single, y As Single, h As Single
    mTid = tid: mMulti = multi: mSel = cur
    mUI.StyleForm Me
    w = 620
    If w > mUI.MaxFormWidth Then w = mUI.MaxFormWidth
    Me.Width = w
    Me.caption = title
    mUI.AddIcon Me, "ic", mCfg.TableIcon(tid), PAD - 4, PAD - 3, 16
    mUI.AddLabel Me, "t", title, PAD + 24, PAD - 2, Me.InsideWidth - PAD * 2 - 30, 20, True, 12
    mUI.AddLabel Me, "t2", "Tabelle „" & mCfg.TableName(tid) & "“  ·  " & IIf(multi, "Mehrfachauswahl: Einträge anklicken", "Einfachauswahl: einen Eintrag wählen"), _
                 PAD + 24, PAD + 18, Me.InsideWidth - PAD * 2 - 30, 14, False, 8.5, CMuted()
    y = PAD + 42
    Set txtSearch = mUI.AddText(Me, "q", PAD, y, Me.InsideWidth - PAD * 2 - 180)
    txtSearch.ControlTipText = "Bezeichnung oder ID eingeben, die Liste filtert sofort"
    Set chkArch = mUI.AddCheck(Me, "arch", "Archivierte anzeigen", Me.InsideWidth - PAD - 170, y, 170)
    y = y + ROW_H + 6
    Set lblInfo = mUI.AddLabel(Me, "info", "", PAD, y, Me.InsideWidth - PAD * 2, 14, False, 8.5, CMuted())
    y = y + 18
    h = 300
    If h > mUI.MaxFormHeight - y - 120 Then h = mUI.MaxFormHeight - y - 120
    Set lst = mUI.AddList(Me, "lst", PAD, y, Me.InsideWidth - PAD * 2, h, multi, 3, _
                          CStr(Int((Me.InsideWidth - PAD * 2) * 0.45)) & ";" & CStr(Int((Me.InsideWidth - PAD * 2) * 0.4)) & ";60")
    y = y + h + 6
    Set lblSel = mUI.AddLabel(Me, "sel", "", PAD, y, Me.InsideWidth - PAD * 2, 28, False, 8.5, CText())
    y = y + 32
    Set btnOk = mUI.AddButton(Me, "ok", "Übernehmen", PAD, y, 110, "primary")
    Set btnNew = mUI.AddButton(Me, "new", "Neu anlegen…", PAD + 118, y, 120)
    Set btnClear = mUI.AddButton(Me, "clr", "Auswahl leeren", PAD + 246, y, 120)
    Set btnCancel = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 100, y, 100)
    btnCancel.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    LoadRecords
    Refill
End Sub

Private Sub LoadRecords()
    Dim t As cTableData, r As Long, lf As String, f As cField, c As Long
    Set t = mData.TData(mTid)
    mN = 0
    ReDim mIds(1 To Application.Max(1, t.RowCount))
    ReDim mLabels(1 To Application.Max(1, t.RowCount))
    lf = mCfg.LabelField(mTid)
    If mCfg.FieldExists(lf) Then Set f = mCfg.Field(lf)
    For r = 1 To t.RowCount
        If Not t.IsArchived(r) Or chkArch.Value Or ListHas(mSel, t.IdAt(r)) Then
            mN = mN + 1
            mIds(mN) = t.IdAt(r)
            If f Is Nothing Then
                mLabels(mN) = t.IdAt(r)
            Else
                mLabels(mN) = mFmt.DisplayText(f, t.Val_(r, lf), t.IdAt(r))
                If Len(mLabels(mN)) = 0 Then mLabels(mN) = "(ohne Bezeichnung)"
            End If
            If t.IsArchived(r) Then mLabels(mN) = mLabels(mN) & " (archiviert)"
        End If
    Next
End Sub

Private Sub Refill()
    Dim q As String, i As Long, n As Long, out() As Variant, maxShow As Long, hits As Long
    q = LCase$(Trim$(txtSearch.Text))
    maxShow = 400
    ReDim out(0 To maxShow - 1, 0 To 2)
    ReDim mShown(0 To maxShow - 1)
    For i = 1 To mN
        If Len(q) = 0 Or InStr(1, LCase$(mLabels(i)), q) > 0 Or InStr(1, LCase$(mIds(i)), q) > 0 Then
            hits = hits + 1
            If n < maxShow Then
                out(n, 0) = mLabels(i)
                out(n, 1) = mData.RecordHint(mTid, mIds(i))
                out(n, 2) = mIds(i)
                mShown(n) = mIds(i)
                n = n + 1
            End If
        End If
    Next
    mLoading = True
    lst.Clear
    If n > 0 Then
        ReDim Preserve out(0 To maxShow - 1, 0 To 2)
        Dim o2() As Variant, j As Long
        ReDim o2(0 To n - 1, 0 To 2)
        For i = 0 To n - 1
            For j = 0 To 2: o2(i, j) = out(i, j): Next
        Next
        lst.List = o2
        For i = 0 To n - 1
            If ListHas(mSel, mShown(i)) Then
                If mMulti Then lst.Selected(i) = True Else lst.ListIndex = i
            End If
        Next
    End If
    mLoading = False
    lblInfo.caption = FmtInt(hits) & " Treffer" & IIf(hits > maxShow, ", angezeigt werden die ersten " & maxShow & "; bitte Suche verfeinern", "") & _
                      IIf(chkArch.Value, "", "  ·  archivierte Einträge ausgeblendet")
    UpdateSel
End Sub

Private Sub UpdateSel()
    Dim a As Variant, i As Long, s As String
    a = Split(mSel, ";")
    For i = 0 To UBound(a)
        If i < 6 Then s = s & IIf(Len(s) > 0, ", ", "") & mData.RecordLabel(mTid, CStr(a(i)))
    Next
    If UBound(a) >= 6 Then s = s & " … (+" & (UBound(a) - 5) & ")"
    If Len(s) = 0 Then lblSel.caption = "Ausgewählt: nichts" Else lblSel.caption = "Ausgewählt (" & (UBound(a) + 1) & "): " & s
End Sub

Private Sub lst_Change()
    Dim i As Long
    If mLoading Then Exit Sub
    If mMulti Then
        For i = 0 To lst.ListCount - 1
            If lst.Selected(i) Then mSel = ListAdd(mSel, mShown(i)) Else mSel = ListRemove(mSel, mShown(i))
        Next
    Else
        If lst.ListIndex >= 0 Then mSel = mShown(lst.ListIndex)
    End If
    UpdateSel
End Sub

Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mMulti And lst.ListIndex >= 0 Then
        mSel = mShown(lst.ListIndex)
        btnOk_Click
    End If
End Sub

Private Sub txtSearch_Change()
    Refill
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDown And lst.ListCount > 0 Then
        lst.SetFocus
        If lst.ListIndex < 0 Then lst.ListIndex = 0
        KeyCode = 0
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = 0
        If Not mMulti And lst.ListCount = 1 Then
            mSel = mShown(0)
            btnOk_Click
        End If
    End If
End Sub

Private Sub chkArch_Click()
    LoadRecords
    Refill
End Sub

Private Sub btnOk_Click()
    Result = mSel
    OK = True
    Me.Hide
End Sub

Private Sub btnClear_Click()
    mSel = ""
    Refill
End Sub

Private Sub btnNew_Click()
    Dim rid As String
    If Not mApp.RequireWritable("Die Schnellanlage") Then Exit Sub
    rid = mForms.NewRecordStacked(mTid, Nothing)
    If Len(rid) > 0 Then
        If mMulti Then mSel = ListAdd(mSel, rid) Else mSel = rid
        LoadRecords
        txtSearch.Text = ""
        Refill
    End If
End Sub

Private Sub btnCancel_Click()
    OK = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCancel_Click
    End If
End Sub
