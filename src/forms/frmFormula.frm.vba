'==============================================================================
' frmFormula - Berechnung per Auswahl von Funktion, Feldern und Werten (keine
' eigene Formelsyntax). Ergebnisvorschau an echten Datensaetzen, verstaendliche
' Fehlermeldungen, Zirkelbezugspruefung, lesbare Beschreibung.
'==============================================================================
Option Explicit

Public OK As Boolean
Public Definition As String
Public ResultType As String
Private mTid As String
Private mFid As String
Private mEvts As Collection
Private mFns As Variant
Private mArgs As Variant
Private mLoading As Boolean
Private cboFn As MSForms.ComboBox
Private lblHelp As MSForms.Label
Private lblPrev As MSForms.Label
Private lblErr As MSForms.Label
Private Const MAXA As Long = 9
Private mY0 As Single

Public Sub Setup(ByVal tid As String, ByVal fid As String, ByVal def As String)
    Dim i As Long, y As Single, b As MSForms.CommandButton
    mTid = tid: mFid = fid
    Set mEvts = New Collection
    mFns = mCalc.FnCodes()
    mUI.StyleForm Me
    Me.Width = 700
    If Me.Width > mUI.MaxFormWidth Then Me.Width = mUI.MaxFormWidth
    Me.caption = "Berechnung festlegen"
    mUI.AddLabel Me, "t", "Berechnung festlegen", PAD, PAD - 2, 500, 20, True, 13
    mUI.AddLabel Me, "t2", "Funktion wählen, dann Felder oder feste Werte zuordnen. Berechnungen können weitere berechnete Felder verwenden. " & _
                 "Leere Werte ergeben ein leeres Ergebnis (nicht 0).", PAD, PAD + 20, Me.InsideWidth - PAD * 2, 28, False, 8.5, CMuted()
    y = PAD + 54
    mUI.AddLabel Me, "lf", "Funktion", PAD, y + 3, 160, 15, True
    Set cboFn = mUI.AddCombo(Me, "fn", PAD + 170, y, Me.InsideWidth - PAD * 2 - 170)
    For i = 0 To UBound(mFns)
        cboFn.AddItem mCalc.FnLabel(CStr(mFns(i)))
    Next
    Bind cboFn, "fn"
    y = y + ROW_H + 6
    Set lblHelp = mUI.AddLabel(Me, "help", "", PAD + 170, y, Me.InsideWidth - PAD * 2 - 170, 30, False, 8.5, CMuted())
    y = y + 34
    mY0 = y
    For i = 1 To MAXA
        BuildArgRow i, mY0 + (i - 1) * 30
    Next
    y = mY0 + MAXA * 30 + 6
    Set b = mUI.AddButton(Me, "prev", "Vorschau berechnen", PAD, y, 160, , ROW_H): Bind b, "preview"
    Set lblPrev = mUI.AddLabel(Me, "pv", "", PAD + 170, y, Me.InsideWidth - PAD * 2 - 170, 64, False, 8.5, CText())
    y = y + 70
    Set lblErr = mUI.AddLabel(Me, "err", "", PAD, y, Me.InsideWidth - PAD * 2, 28, False, 8.5, CDanger())
    y = y + 32
    Set b = mUI.AddButton(Me, "ok", "Übernehmen", PAD, y, 120, "primary"): Bind b, "ok"
    Set b = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 104, y, 104): Bind b, "cancel"
    b.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    LoadDef def
End Sub

Private Sub Bind(ByVal c As Object, ByVal key As String)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, c, key
    mEvts.Add e
End Sub

Private Sub BuildArgRow(ByVal i As Long, ByVal y As Single)
    Dim c As Object
    mUI.AddLabel Me, "al" & i, "", PAD, y + 3, 165, 15, True
    Set c = mUI.AddCombo(Me, "as" & i, PAD + 170, y, 110)
    c.AddItem "Feld": c.AddItem "Fester Wert": c.AddItem "(leer)"
    Bind c, "src:" & i
    Set c = mUI.AddCombo(Me, "af" & i, PAD + 286, y, Me.InsideWidth - PAD * 2 - 286)
    Bind c, "fld:" & i
    Set c = mUI.AddText(Me, "av" & i, PAD + 286, y, Me.InsideWidth - PAD * 2 - 286)
    Set c = mUI.AddCombo(Me, "ak" & i, PAD + 170, y, Me.InsideWidth - PAD * 2 - 170)
    Bind c, "key:" & i
End Sub

Private Function FnCode() As String
    If cboFn.ListIndex >= 0 Then FnCode = mFns(cboFn.ListIndex)
End Function

Private Function ArgKind(ByVal i As Long) As String
    If Not IsArray(mArgs) Then Exit Function
    If i - 1 <= UBound(mArgs) Then ArgKind = Split(mArgs(i - 1), ":")(0)
End Function

' Felder fuer eine Argumentart
Private Function FieldsFor(ByVal kind As String, ByVal i As Long) As Collection
    Dim res As New Collection, f As cField, tt As String, lf As cField
    Select Case kind
        Case "link"
            For Each f In mCfg.Fields(mTid)
                If f.IsLinkType Then res.Add f
            Next
        Case "tfield"
            Set lf = LinkFieldOfRow()
            If Not lf Is Nothing Then
                If lf.FType = FT_LINK Then tt = lf.TargetTable Else tt = mCfg.Field(lf.PairField).TableID
                For Each f In mCfg.Fields(tt)
                    If Not f.IsLinkType Then res.Add f
                Next
            End If
        Case "status"
            For Each f In mCfg.Fields(mTid)
                If f.FType = FT_STATUS Then res.Add f
            Next
        Case Else
            For Each f In mCfg.Fields(mTid)
                If f.ID <> mFid And Not f.IsLinkType Then
                    Select Case kind
                        Case "num": If f.IsNumeric_ Or f.IsDate_ Then res.Add f
                        Case "date": If f.IsDate_ Then res.Add f
                        Case Else: res.Add f
                    End Select
                End If
            Next
    End Select
    Set FieldsFor = res
End Function

Private Function LinkFieldOfRow() As cField
    Dim i As Long, nm As String, f As cField
    For i = 1 To MAXA
        If ArgKind(i) = "link" Then
            nm = Me.Controls("af" & i).Text
            For Each f In mCfg.Fields(mTid)
                If f.IsLinkType And f.Name = nm Then Set LinkFieldOfRow = f: Exit Function
            Next
        End If
    Next
End Function

Private Sub FillFieldCombo(ByVal i As Long, ByVal kind As String)
    Dim c As MSForms.ComboBox, f As cField
    Set c = Me.Controls("af" & i)
    c.Clear
    If kind = "tfield" Or kind = "status" Then c.AddItem "(keins)"
    For Each f In FieldsFor(kind, i)
        c.AddItem f.Name
    Next
End Sub

Private Sub FillKeyCombo(ByVal i As Long, ByVal kind As String)
    Dim c As MSForms.ComboBox, v As Variant
    Set c = Me.Controls("ak" & i)
    c.Clear
    Select Case kind
        Case "op": c.AddItem "+ plus": c.AddItem "" & ChrW(&H2212) & " minus": c.AddItem "× mal": c.AddItem "÷ geteilt durch"
        Case "cmp"
            For Each v In mCalc.CmpCodes()
                c.AddItem mCalc.CmpText(CStr(v))
            Next
        Case "logic": c.AddItem "(keine zweite Bedingung)": c.AddItem "und": c.AddItem "oder"
        Case "unit": c.AddItem "Tage": c.AddItem "Wochen": c.AddItem "Monate": c.AddItem "Jahre"
        Case "agg"
            c.AddItem "Anzahl": c.AddItem "Anzahl mit Bedingung": c.AddItem "Anzahl ausgefüllter Werte"
            c.AddItem "Summe": c.AddItem "Durchschnitt": c.AddItem "Minimum": c.AddItem "Maximum"
    End Select
    If c.ListCount > 0 Then c.ListIndex = 0
End Sub

Private Function KeyCode_(ByVal kind As String, ByVal idx As Long) As String
    Dim a As Variant
    Select Case kind
        Case "op": a = Array("+", "-", "*", "/")
        Case "cmp": a = mCalc.CmpCodes()
        Case "logic": a = Array("", "UND", "ODER")
        Case "unit": a = Array("Tage", "Wochen", "Monate", "Jahre")
        Case "agg": a = Array("ANZAHL", "ANZAHLWENN", "GEFUELLT", "SUMME", "MITTEL", "MIN", "MAX")
        Case Else: Exit Function
    End Select
    If idx >= 0 And idx <= UBound(a) Then KeyCode_ = CStr(a(idx)) Else KeyCode_ = Chr$(1)
End Function

Private Function KeyIndex(ByVal kind As String, ByVal code As String) As Long
    Dim i As Long
    For i = 0 To 8
        If KeyCode_(kind, i) = code Then KeyIndex = i: Exit Function
    Next
End Function

Private Sub LayoutArgs()
    Dim i As Long, kind As String, lbl As String, src As Long, a As Variant
    mArgs = mCalc.FnArgs(FnCode())
    lblHelp.caption = mCalc.FnHelp(FnCode())
    For i = 1 To MAXA
        kind = ArgKind(i)
        If Len(kind) = 0 Then
            Me.Controls("al" & i).Visible = False: Me.Controls("as" & i).Visible = False
            Me.Controls("af" & i).Visible = False: Me.Controls("av" & i).Visible = False: Me.Controls("ak" & i).Visible = False
        Else
            Me.Controls("al" & i).caption = Split(mArgs(i - 1), ":")(1)
            Me.Controls("al" & i).Visible = True
            Select Case kind
                Case "op", "cmp", "logic", "unit", "agg"
                    Me.Controls("as" & i).Visible = False: Me.Controls("af" & i).Visible = False: Me.Controls("av" & i).Visible = False
                    Me.Controls("ak" & i).Visible = True
                Case "link", "tfield", "status"
                    Me.Controls("as" & i).Visible = False: Me.Controls("av" & i).Visible = False: Me.Controls("ak" & i).Visible = False
                    Me.Controls("af" & i).Visible = True
                    Me.Controls("af" & i).Left = PAD + 170: Me.Controls("af" & i).Width = Me.InsideWidth - PAD * 2 - 170
                Case "digits", "text"
                    Me.Controls("as" & i).Visible = False: Me.Controls("af" & i).Visible = False: Me.Controls("ak" & i).Visible = False
                    Me.Controls("av" & i).Visible = True
                    Me.Controls("av" & i).Left = PAD + 170: Me.Controls("av" & i).Width = 140
                Case Else
                    Me.Controls("ak" & i).Visible = False
                    Me.Controls("as" & i).Visible = True
                    Me.Controls("af" & i).Left = PAD + 286: Me.Controls("af" & i).Width = Me.InsideWidth - PAD * 2 - 286
                    Me.Controls("av" & i).Left = PAD + 286: Me.Controls("av" & i).Width = Me.InsideWidth - PAD * 2 - 286
                    src = Me.Controls("as" & i).ListIndex
                    Me.Controls("af" & i).Visible = (src = 0)
                    Me.Controls("av" & i).Visible = (src = 1)
            End Select
        End If
    Next
    ' bedingte Zeilen (zweite Bedingung) ausblenden, wenn keine Verknuepfung gewaehlt
    If FnCode() = "WENN" Or FnCode() = "PRUEFEN" Then
        If Me.Controls("ak4").ListIndex <= 0 Then
            For i = 5 To 7
                Me.Controls("al" & i).Visible = False: Me.Controls("as" & i).Visible = False
                Me.Controls("af" & i).Visible = False: Me.Controls("av" & i).Visible = False: Me.Controls("ak" & i).Visible = False
            Next
        End If
    End If
End Sub

Private Sub ResetArgs()
    Dim i As Long, kind As String
    mLoading = True
    mArgs = mCalc.FnArgs(FnCode())
    For i = 1 To MAXA
        kind = ArgKind(i)
        Me.Controls("av" & i).Text = ""
        Select Case kind
            Case "op", "cmp", "logic", "unit", "agg": FillKeyCombo i, kind
            Case "link", "tfield", "status": FillFieldCombo i, kind
            Case "digits": Me.Controls("av" & i).Text = "0"
            Case "text": Me.Controls("av" & i).Text = " "
            Case ""
            Case Else
                FillFieldCombo i, kind
                Me.Controls("as" & i).ListIndex = 0
        End Select
    Next
    If FnCode() = "ROLLUP" And Me.Controls("af1").ListCount > 0 Then Me.Controls("af1").ListIndex = 0: FillFieldCombo 3, "tfield": FillFieldCombo 4, "tfield"
    mLoading = False
    LayoutArgs
End Sub

Private Sub LoadDef(ByVal def As String)
    Dim p As Variant, i As Long, a As String, kind As String, fn As String, k As Long
    mLoading = True
    p = Split(def, "|")
    If UBound(p) < 0 Then
        cboFn.ListIndex = 0
        mLoading = False
        ResetArgs
        Exit Sub
    End If
    fn = p(0)
    For k = 0 To UBound(mFns)
        If mFns(k) = fn Then cboFn.ListIndex = k
    Next
    mLoading = False
    ResetArgs
    mLoading = True
    For i = 1 To MAXA
        kind = ArgKind(i)
        a = mCalc.ArgAt(p, i)
        Select Case kind
            Case "op", "cmp", "logic", "unit", "agg"
                Me.Controls("ak" & i).ListIndex = KeyIndex(kind, Mid$(a, 3))
                If kind = "link" Then FillFieldCombo 3, "tfield"
            Case "link", "tfield", "status"
                If Left$(a, 2) = "F:" Then
                    On Error Resume Next
                    Me.Controls("af" & i).Value = mCfg.FieldName(Mid$(a, 3))
                    On Error GoTo 0
                    If kind = "link" Then FillFieldCombo 3, "tfield": FillFieldCombo 4, "tfield"
                End If
            Case "digits"
                Me.Controls("av" & i).Text = Mid$(a, 3)
            Case "text"
                Me.Controls("av" & i).Text = Unesc(Mid$(a, 3))
            Case ""
            Case Else
                If Left$(a, 2) = "F:" Then
                    Me.Controls("as" & i).ListIndex = 0
                    On Error Resume Next
                    Me.Controls("af" & i).Value = mCfg.FieldName(Mid$(a, 3))
                    On Error GoTo 0
                ElseIf Len(a) = 0 Then
                    Me.Controls("as" & i).ListIndex = 2
                Else
                    Me.Controls("as" & i).ListIndex = 1
                    Select Case Left$(a, 2)
                        Case "N:": Me.Controls("av" & i).Text = FmtNum(InvNum(Mid$(a, 3)), 6, True)
                        Case "D:": Me.Controls("av" & i).Text = FmtDate(InvDate(Mid$(a, 3)))
                        Case "B:": Me.Controls("av" & i).Text = IIf(Mid$(a, 3) = "1", "Ja", "Nein")
                        Case "O:": Me.Controls("av" & i).Text = mCfg.OptName(Mid$(a, 3))
                        Case Else: Me.Controls("av" & i).Text = Unesc(Mid$(a, 3))
                    End Select
                End If
        End Select
    Next
    ' ROLLUP: Zielfelder nach Laden der Verknuepfung erneut setzen
    If fn = "ROLLUP" Then
        For i = 3 To 4
            a = mCalc.ArgAt(p, i)
            If Left$(a, 2) = "F:" Then
                On Error Resume Next
                Me.Controls("af" & i).Value = mCfg.FieldName(Mid$(a, 3))
                On Error GoTo 0
            End If
        Next
    End If
    mLoading = False
    LayoutArgs
End Sub

' Feld-ID aus dem Feldauswahlfeld einer Zeile
Private Function RowFieldId(ByVal i As Long) As String
    Dim kind As String, f As cField, nm As String
    kind = ArgKind(i)
    nm = Me.Controls("af" & i).Text
    If Len(nm) = 0 Or nm = "(keins)" Then Exit Function
    For Each f In FieldsFor(kind, i)
        If f.Name = nm Then RowFieldId = f.ID: Exit Function
    Next
End Function

' Bezugsfeld fuer Konstanten (Vergleichspartner)
Private Function PartnerField(ByVal i As Long) As cField
    Dim fn As String, j As Long
    fn = FnCode()
    Select Case fn
        Case "WENN", "PRUEFEN"
            If i = 3 Then j = 1
            If i = 7 Then j = 5
        Case "ROLLUP"
            If i = 6 Then j = 4
    End Select
    If j > 0 Then
        If Len(RowFieldId(j)) > 0 Then Set PartnerField = mCfg.Field(RowFieldId(j))
    End If
End Function

Private Function BuildDef(ByRef msg As String) As String
    Dim fn As String, i As Long, kind As String, s As String, a As String, txt As String, d As Double, dt As Date, ht As Boolean, pf As cField, oid As String
    fn = FnCode()
    If Len(fn) = 0 Then msg = "Bitte eine Funktion wählen.": Exit Function
    s = fn
    For i = 1 To UBound(mArgs) + 1
        kind = ArgKind(i)
        a = ""
        Select Case kind
            Case "op", "cmp", "logic", "unit", "agg"
                If Me.Controls("ak" & i).Visible Or kind = "logic" Then
                    If Me.Controls("ak" & i).ListIndex >= 0 Then a = "K:" & KeyCode_(kind, Me.Controls("ak" & i).ListIndex)
                    If a = "K:" Then a = ""
                End If
            Case "link"
                If Len(RowFieldId(i)) = 0 Then msg = "Bitte ein Verknüpfungsfeld wählen (die Tabelle braucht dafür ein Feld vom Typ Verknüpfung).": Exit Function
                a = "F:" & RowFieldId(i)
            Case "tfield", "status"
                If Len(RowFieldId(i)) > 0 Then a = "F:" & RowFieldId(i)
            Case "digits"
                If Not IsAllDigits(Trim$(Me.Controls("av" & i).Text)) Then msg = "Nachkommastellen: bitte eine ganze Zahl von 0 bis 6 eingeben.": Exit Function
                a = "N:" & Trim$(Me.Controls("av" & i).Text)
            Case "text"
                a = "T:" & Esc(Me.Controls("av" & i).Text)
            Case Else
                If Me.Controls("as" & i).Visible Then
                    Select Case Me.Controls("as" & i).ListIndex
                        Case 0
                            If Len(RowFieldId(i)) = 0 Then msg = "„" & Me.Controls("al" & i).caption & "“: bitte ein Feld wählen.": Exit Function
                            a = "F:" & RowFieldId(i)
                        Case 1
                            txt = Trim$(Me.Controls("av" & i).Text)
                            Set pf = PartnerField(i)
                            oid = ""
                            If Not pf Is Nothing Then
                                If pf.IsOptionType Then oid = mCfg.OptIdByName(pf.ID, txt)
                            End If
                            If Len(oid) > 0 Then
                                a = "O:" & oid
                            ElseIf kind = "num" Then
                                If Not ParseNumberDE(txt, d) Then msg = "„" & Me.Controls("al" & i).caption & "“: bitte eine Zahl eingeben.": Exit Function
                                a = "N:" & NumInv(d)
                            ElseIf kind = "date" Then
                                If Not ParseDateDE(txt, dt, ht) Then msg = "„" & Me.Controls("al" & i).caption & "“: bitte ein Datum TT.MM.JJJJ eingeben.": Exit Function
                                a = "D:" & DateInv(dt)
                            ElseIf LCase$(txt) = "ja" Or LCase$(txt) = "nein" Then
                                a = "B:" & IIf(LCase$(txt) = "ja", "1", "0")
                            ElseIf ParseNumberDE(txt, d) Then
                                a = "N:" & NumInv(d)
                            ElseIf ParseDateDE(txt, dt, ht) Then
                                a = "D:" & DateInv(dt)
                            Else
                                If Not pf Is Nothing Then
                                    If pf.IsOptionType And Len(txt) > 0 Then msg = "„" & txt & "“ ist keine Auswahlmöglichkeit von „" & pf.Name & "“.": Exit Function
                                End If
                                a = "T:" & Esc(txt)
                            End If
                        Case Else
                            a = ""
                    End Select
                End If
        End Select
        s = s & "|" & a
    Next
    BuildDef = s
End Function

Private Sub Preview()
    Dim def As String, msg As String, t As cTableData, r As Long, v As Variant, s As String, tmp As New cField, path As String
    lblErr.caption = ""
    def = BuildDef(msg)
    If Len(msg) > 0 Then lblErr.caption = ChrW(&H26A0) & "  " & msg: Exit Sub
    If Len(mFid) > 0 Then
        If mCalc.HasCycle(mFid, def, path) Then lblErr.caption = ChrW(&H26A0) & "  Zirkelbezug: " & path & ". Bitte ein anderes Feld wählen.": Exit Sub
    End If
    tmp.FType = FT_FORMEL: tmp.ResultType = mCalc.DefaultResultType(def, mTid): tmp.Decimals = 2
    s = "Beschreibung: " & mCalc.Describe(def) & vbLf
    Set t = mData.TData(mTid)
    If t.RowCount = 0 Then
        s = s & "Vorschau: noch keine Datensätze vorhanden."
    Else
        s = s & "Vorschau:"
        For r = 1 To Application.Min(3, t.RowCount)
            v = mCalc.EvalDef(def, t, r)
            s = s & "  " & mData.RecordLabel(mTid, t.IdAt(r)) & " " & ChrW(&H2192) & " " & IIf(IsBlank(v), "(leer)", mFmt.DisplayText(tmp, v)) & IIf(r < Application.Min(3, t.RowCount), " ·", "")
        Next
    End If
    lblPrev.caption = s
End Sub

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    Dim p As Long, kind As String, i As Long, def As String, msg As String, path As String
    On Error GoTo failed
    If mLoading Then Exit Sub
    p = InStr(key, ":")
    If p > 0 Then kind = Left$(key, p - 1): i = CLng(Mid$(key, p + 1)) Else kind = key
    Select Case kind
        Case "fn": If evt = "Change" Then ResetArgs
        Case "src": If evt = "Change" Then LayoutArgs
        Case "key": If evt = "Change" Then LayoutArgs
        Case "fld"
            If evt = "Change" And ArgKind(i) = "link" Then
                mLoading = True
                FillFieldCombo 3, "tfield": FillFieldCombo 4, "tfield"
                mLoading = False
            End If
        Case "preview": If evt = "Click" Then Preview
        Case "ok"
            If evt = "Click" Then
                def = BuildDef(msg)
                If Len(msg) > 0 Then lblErr.caption = ChrW(&H26A0) & "  " & msg: Exit Sub
                If Len(mFid) > 0 Then
                    If mCalc.HasCycle(mFid, def, path) Then lblErr.caption = ChrW(&H26A0) & "  Zirkelbezug: " & path & ". Bitte ein anderes Feld wählen.": Exit Sub
                End If
                Definition = def
                ResultType = mCalc.DefaultResultType(def, mTid)
                OK = True
                Me.Hide
            End If
        Case "cancel": If evt = "Click" Then OK = False: Me.Hide
    End Select
    Exit Sub
failed:
    mApp.Boundary "Berechnung"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True: OK = False: Me.Hide
End Sub

Private Sub UserForm_Terminate()
    Dim e As cEvt
    If Not mEvts Is Nothing Then
        For Each e In mEvts: e.Unbind: Next
    End If
End Sub
