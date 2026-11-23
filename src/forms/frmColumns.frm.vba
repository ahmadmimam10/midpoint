'==============================================================================
' frmColumns - Spaltenauswahl einer Tabellenansicht: Sichtbarkeit, Reihenfolge,
' Breite und optionale Zusammenfassung je Spalte. Wird dauerhaft in der
' Ansicht gespeichert.
'==============================================================================
Option Explicit

Public OK As Boolean
Public ColSpec As String
Public SumSpec As String
Private mVid As String
Private mFid() As String
Private mVis() As Boolean
Private mWid() As Single
Private mSum() As String
Private mN As Long
Private mLoading As Boolean
Private WithEvents lst As MSForms.ListBox
Private WithEvents btnUp As MSForms.CommandButton
Private WithEvents btnDown As MSForms.CommandButton
Private WithEvents txtW As MSForms.TextBox
Private WithEvents cboSum As MSForms.ComboBox
Private WithEvents btnOk As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private WithEvents btnAll As MSForms.CommandButton
Private WithEvents chkVis As MSForms.CheckBox
Private lblInfo As MSForms.Label

Public Sub Setup(ByVal vid As String)
    Dim tid As String, cols As Collection, c As Variant, f As cField, seen As New cMap, i As Long, y As Single, sums As String
    mVid = vid
    tid = mCfg.ViewGet(vid, "Tabelle")
    sums = mCfg.ViewOpt(vid, "summen", "")
    Set cols = mViewTable.ViewColumns(vid)
    ReDim mFid(1 To mCfg.Fields(tid).Count + 1): ReDim mVis(1 To mCfg.Fields(tid).Count + 1)
    ReDim mWid(1 To mCfg.Fields(tid).Count + 1): ReDim mSum(1 To mCfg.Fields(tid).Count + 1)
    For Each c In cols
        mN = mN + 1
        mFid(mN) = c(0): mVis(mN) = True: mWid(mN) = c(1)
        seen.SetVal CStr(c(0)), True
    Next
    For Each f In mCfg.Fields(tid)
        If Not seen.Exists(f.ID) Then
            mN = mN + 1
            mFid(mN) = f.ID: mVis(mN) = False: mWid(mN) = mViewTable.DefaultWidth(f, False)
        End If
    Next
    For i = 1 To mN
        mSum(i) = SumOf(sums, mFid(i))
    Next
    mUI.StyleForm Me
    Me.Width = 540
    Me.caption = "Spalten auswählen"
    mUI.AddLabel Me, "t", "Spalten der Ansicht „" & mCfg.ViewGet(vid, "Name") & "“", PAD, PAD - 2, 480, 20, True, 12
    mUI.AddLabel Me, "t2", "Spalte anklicken = auswählen (blendet nichts ein oder aus). Sichtbarkeit rechts mit „Sichtbar“ oder per Doppelklick. " & _
                 "Das Bezeichnungsfeld bleibt immer als erste, fixierte Spalte sichtbar.", PAD, PAD + 20, Me.InsideWidth - PAD * 2, 28, False, 8.5, CMuted()
    y = PAD + 54
    ' Einfachauswahl: Klick waehlt nur die Zeile; Spalte 1 zeigt den Haken fuer "sichtbar"
    Set lst = mUI.AddList(Me, "lst", PAD, y, 300, 300, False, 2, "22;270")
    Set chkVis = mUI.AddCheck(Me, "vis", "Sichtbar", PAD + 310, y + 2, 150)
    Set btnUp = mUI.AddButton(Me, "up", "Nach oben", PAD + 310, y + 30, 120)
    Set btnDown = mUI.AddButton(Me, "down", "Nach unten", PAD + 310, y + 62, 120)
    Set btnAll = mUI.AddButton(Me, "all", "Alle anzeigen", PAD + 310, y + 94, 120)
    mUI.AddLabel Me, "lw", "Breite (Punkte)", PAD + 310, y + 134, 150, 15, True
    Set txtW = mUI.AddText(Me, "w", PAD + 310, y + 151, 80)
    mUI.AddLabel Me, "ls", "Zusammenfassung", PAD + 310, y + 184, 150, 15, True
    Set cboSum = mUI.AddCombo(Me, "s", PAD + 310, y + 201, 170)
    cboSum.AddItem "(keine)": cboSum.AddItem "Anzahl": cboSum.AddItem "Ausgefüllt": cboSum.AddItem "Leer"
    cboSum.AddItem "Summe": cboSum.AddItem "Durchschnitt": cboSum.AddItem "Minimum": cboSum.AddItem "Maximum"
    Set lblInfo = mUI.AddLabel(Me, "info", "", PAD + 310, y + 229, 180, 60, False, 8, CMuted())
    y = y + 312
    Set btnOk = mUI.AddButton(Me, "ok", "Übernehmen", PAD, y, 120, "primary")
    Set btnCancel = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 104, y, 104)
    btnCancel.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    Refill 1
End Sub

Private Function SumOf(ByVal spec As String, ByVal fid As String) As String
    Dim a As Variant, i As Long, p As Variant
    a = Split(spec, ",")
    For i = 0 To UBound(a)
        p = Split(a(i), ":")
        If UBound(p) >= 1 Then If p(0) = fid Then SumOf = p(1)
    Next
End Function

Private Function SumLabel(ByVal code As String) As String
    Select Case code
        Case "ANZAHL": SumLabel = "Anzahl"
        Case "GEFUELLT": SumLabel = "Ausgefüllt"
        Case "LEER": SumLabel = "Leer"
        Case "SUMME": SumLabel = "Summe"
        Case "MITTEL": SumLabel = "Durchschnitt"
        Case "MIN": SumLabel = "Minimum"
        Case "MAX": SumLabel = "Maximum"
        Case Else: SumLabel = "(keine)"
    End Select
End Function

Private Function SumCode(ByVal lbl As String) As String
    Select Case lbl
        Case "Anzahl": SumCode = "ANZAHL"
        Case "Ausgefüllt": SumCode = "GEFUELLT"
        Case "Leer": SumCode = "LEER"
        Case "Summe": SumCode = "SUMME"
        Case "Durchschnitt": SumCode = "MITTEL"
        Case "Minimum": SumCode = "MIN"
        Case "Maximum": SumCode = "MAX"
    End Select
End Function

Private Sub Refill(ByVal sel As Long)
    Dim i As Long, lf As String
    mLoading = True
    lst.Clear
    lf = mCfg.LabelField(mCfg.ViewGet(mVid, "Tabelle"))
    For i = 1 To mN
        If mFid(i) = lf Then mVis(i) = True
        lst.AddItem IIf(mVis(i), ChrW(&H2713), "")
        lst.List(i - 1, 1) = mCfg.FieldName(mFid(i)) & IIf(mFid(i) = lf, "  (Bezeichnung, fixiert)", "") & _
                             IIf(mVis(i), "", "  (ausgeblendet)") & IIf(Len(mSum(i)) > 0, "  · " & SumLabel(mSum(i)), "")
    Next
    If sel >= 1 And sel <= mN Then lst.ListIndex = sel - 1
    mLoading = False
    ShowCurrent
End Sub

Private Sub ShowCurrent()
    Dim i As Long, f As cField
    i = lst.ListIndex + 1
    If i < 1 Then Exit Sub
    mLoading = True
    txtW.Text = FmtNum(mWid(i), 0)
    cboSum.Value = SumLabel(mSum(i))
    chkVis.Value = mVis(i)
    chkVis.Enabled = (mFid(i) <> mCfg.LabelField(mCfg.ViewGet(mVid, "Tabelle")))
    Set f = mCfg.Field(mFid(i))
    lblInfo.caption = "Typ: " & f.TypeName_ & IIf(f.IsNumeric_, "", vbLf & "Summe/Durchschnitt nur bei Zahlenfeldern sinnvoll.")
    mLoading = False
End Sub

Private Sub lst_Change()
    If mLoading Then Exit Sub
    ShowCurrent
End Sub

Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ToggleCurrent
End Sub

Private Sub lst_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeySpace Then ToggleCurrent: KeyCode = 0
End Sub

Private Sub chkVis_Click()
    Dim i As Long
    If mLoading Then Exit Sub
    i = lst.ListIndex + 1
    If i < 1 Then Exit Sub
    If mVis(i) <> CBool(chkVis.Value) Then ToggleCurrent
End Sub

Private Sub ToggleCurrent()
    Dim i As Long
    i = lst.ListIndex + 1
    If i < 1 Then Exit Sub
    If mFid(i) = mCfg.LabelField(mCfg.ViewGet(mVid, "Tabelle")) Then Exit Sub
    mVis(i) = Not mVis(i)
    Refill i
End Sub

Private Sub Swap(ByVal a As Long, ByVal b As Long)
    Dim s As String, v As Boolean, w As Single, sm As String
    s = mFid(a): mFid(a) = mFid(b): mFid(b) = s
    v = mVis(a): mVis(a) = mVis(b): mVis(b) = v
    w = mWid(a): mWid(a) = mWid(b): mWid(b) = w
    sm = mSum(a): mSum(a) = mSum(b): mSum(b) = sm
End Sub

Private Sub btnUp_Click()
    Dim i As Long
    i = lst.ListIndex + 1
    If i <= 1 Then Exit Sub
    Swap i, i - 1
    Refill i - 1
End Sub

Private Sub btnDown_Click()
    Dim i As Long
    i = lst.ListIndex + 1
    If i < 1 Or i >= mN Then Exit Sub
    Swap i, i + 1
    Refill i + 1
End Sub

Private Sub btnAll_Click()
    Dim i As Long
    For i = 1 To mN: mVis(i) = True: Next
    Refill lst.ListIndex + 1
End Sub

Private Sub txtW_Change()
    Dim i As Long, d As Double
    If mLoading Then Exit Sub
    i = lst.ListIndex + 1
    If i < 1 Then Exit Sub
    If ParseNumberDE(txtW.Text, d) Then
        If d >= 40 And d <= 600 Then mWid(i) = d
    End If
End Sub

Private Sub cboSum_Change()
    Dim i As Long
    If mLoading Then Exit Sub
    i = lst.ListIndex + 1
    If i < 1 Then Exit Sub
    mSum(i) = SumCode(cboSum.Text)
    Refill i
End Sub

Private Sub btnOk_Click()
    Dim i As Long
    ColSpec = "": SumSpec = ""
    For i = 1 To mN
        If mVis(i) Then ColSpec = ColSpec & IIf(Len(ColSpec) > 0, ";", "") & mFid(i) & ":" & CLng(mWid(i))
        If Len(mSum(i)) > 0 And mVis(i) Then SumSpec = SumSpec & IIf(Len(SumSpec) > 0, ",", "") & mFid(i) & ":" & mSum(i)
    Next
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
