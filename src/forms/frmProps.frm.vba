'==============================================================================
' frmProps - Allgemeiner Einstellungsdialog (Text, Auswahl, Ja/Nein, Datum,
' Hinweise). Wird von Einstellungen, Ansichten, Dashboards u. a. genutzt.
' Fehler werden am Feld angezeigt; der Dialog bleibt bis zur Korrektur offen.
'==============================================================================
Option Explicit

Public OK As Boolean
Private mDefs As Collection        ' Array(key, kind, label, value, items, help)
Private mCtl As cMap
Private mErr As cMap
Private mEvts As Collection
Private mTitle As String
Private mIntro As String
Private mW As Single
Private mBuilt As Boolean
Private WithEvents btnOk As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private WithEvents btnBack As MSForms.CommandButton
Private mOkCaption As String
Public ChangedKey As String
Public PreviewMacro As String
Public AllowBack As Boolean         ' mehrstufiger Assistent: Schaltflaeche "Zurueck"
Public WentBack As Boolean
Private mBlocks As Collection       ' Array(Schlüssel, Steuerelementnamen, Abstände, Höhe, hat Steuerelemente)
Private mDep As cMap                ' Schlüssel -> Array(steuernder Schlüssel, Werte)
Private mFrTop As Single
Private mListSel As cMap            ' Schlüssel -> zuletzt gewählte Indizes einer Mehrfachliste
Private mStyleVal As cMap           ' Schlüssel -> gewählter Farbstil
Private mHelper As cMap             ' Schlüssel -> Array(Beschriftung, Funktion) der Hilfe-Schaltfläche

Public Sub Init(ByVal title As String, ByVal intro As String, Optional ByVal width As Single = 500)
    Set mDefs = New Collection
    Set mCtl = New cMap
    Set mErr = New cMap
    Set mEvts = New Collection
    mTitle = title: mIntro = intro: mW = width
End Sub

Public Sub AddText(ByVal key As String, ByVal label As String, ByVal value As String, Optional ByVal help As String = "", _
                   Optional ByVal multiline As Boolean = False, Optional ByVal password As Boolean = False)
    mDefs.Add Array(key, IIf(multiline, "multi", IIf(password, "pw", "text")), label, value, Empty, help)
End Sub

Public Sub AddCombo(ByVal key As String, ByVal label As String, ByVal items As Variant, ByVal value As String, Optional ByVal help As String = "")
    mDefs.Add Array(key, "combo", label, value, items, help)
End Sub

Public Sub AddCheck(ByVal key As String, ByVal label As String, ByVal value As Boolean, Optional ByVal help As String = "")
    mDefs.Add Array(key, "check", label, value, Empty, help)
End Sub

Public Sub AddDate(ByVal key As String, ByVal label As String, ByVal value As String, Optional ByVal help As String = "")
    mDefs.Add Array(key, "date", label, value, Empty, help)
End Sub

Public Sub AddMulti(ByVal key As String, ByVal label As String, ByVal items As Variant, ByVal selected As String, Optional ByVal help As String = "")
    ' items: Anzeigetexte; selected: Indizes ";"-getrennt
    mDefs.Add Array(key, "list", label, selected, items, help)
End Sub

' Farbstil mit Vorschau der Farben; "Ändern…" öffnet die Stilauswahl (frmStylePick).
' inheritLabel: zusätzlicher erster Eintrag ohne eigene Farben, z. B. "(wie Dashboard)"
Public Sub AddStyle(ByVal key As String, ByVal label As String, ByVal value As String, Optional ByVal help As String = "", _
                    Optional ByVal inheritLabel As String = "")
    If mStyleVal Is Nothing Then Set mStyleVal = New cMap
    If Len(value) = 0 Or (Not mDocStyle.IsStyleName(value) And value <> inheritLabel) Then
        ' frühere Einzelfarbe oder Palette: als nächstliegender Stil anzeigen
        If value <> inheritLabel Then value = "Akzentfarbe"
    End If
    mStyleVal.SetVal key, value
    mDefs.Add Array(key, "style", label, value, inheritLabel, help)
End Sub

' Schaltfläche rechts neben einer Auswahl (z. B. "Welches Diagramm passt?"). Die Funktion
' erhält den aktuellen Wert und liefert den neuen (leer = unverändert).
Public Sub SetComboHelper(ByVal key As String, ByVal caption As String, ByVal funcName As String)
    If mHelper Is Nothing Then Set mHelper = New cMap
    mHelper.SetVal key, Array(caption, funcName)
End Sub

Public Sub AddInfo(ByVal text As String)
    mDefs.Add Array("", "info", text, "", Empty, "")
End Sub

Private Sub Build()
    Dim d As Variant, y As Single, c As Object, lbl As MSForms.Label, i As Long, b As MSForms.CommandButton, e As cEvt, h As Single, maxH As Single
    Dim fr As MSForms.Frame, iw As Single, a As Variant, n0 As Long, h0 As Single, nms() As String, offs() As Single
    mUI.StyleForm Me
    Me.caption = mTitle
    If mW > mUI.MaxFormWidth Then mW = mUI.MaxFormWidth
    Me.Width = mW
    mUI.AddLabel Me, "t", mTitle, PAD, PAD - 2, Me.InsideWidth - PAD * 2, 20, True, 12
    y = PAD + 24
    If Len(mIntro) > 0 Then
        Set lbl = mUI.AddLabel(Me, "intro", mIntro, PAD, y, Me.InsideWidth - PAD * 2, 20, False, 9, CMuted())
        lbl.AutoSize = True
        If lbl.Width > Me.InsideWidth - PAD * 2 Then lbl.Width = Me.InsideWidth - PAD * 2
        y = y + lbl.Height + 10
    End If
    Set fr = mUI.AddFrame(Me, "fr", PAD - 4, y, Me.InsideWidth - PAD * 2 + 8, 100, "", True)
    mFrTop = y
    iw = fr.Width - 24
    h = 6
    Set mBlocks = New Collection
    For Each d In mDefs
        n0 = fr.Controls.Count: h0 = h
        Select Case d(1)
            Case "info"
                Set lbl = mUI.AddLabel(fr, "i" & h, CStr(d(2)), 4, h, iw, 20, False, 9, CMuted())
                lbl.AutoSize = True
                If lbl.Width > iw Then lbl.Width = iw
                h = h + lbl.Height + 8
            Case "check"
                Set c = mUI.AddCheck(fr, "c_" & d(0), CStr(d(2)), 4, h, iw)
                c.Value = CBool(d(3))
                mCtl.SetVal CStr(d(0)), c
                Set e = New cEvt: e.Bind Me, c, "chk:" & d(0): mEvts.Add e
                h = h + ROW_H
                If Len(d(5)) > 0 Then mUI.AddLabel fr, "h_" & d(0), CStr(d(5)), 22, h - 2, iw - 18, 14, False, 8, CMuted(): h = h + 14
                h = h + 6
            Case Else
                mUI.AddLabel fr, "l_" & d(0), CStr(d(2)), 4, h, iw, 15, True, 9.5
                h = h + 17
                Select Case d(1)
                    Case "text", "pw"
                        Set c = mUI.AddText(fr, "c_" & d(0), 4, h, iw)
                        c.Text = CStr(d(3))
                        If d(1) = "pw" Then c.PasswordChar = "*"
                        h = h + ROW_H
                    Case "multi"
                        Set c = mUI.AddText(fr, "c_" & d(0), 4, h, iw, 60, True)
                        c.Text = CStr(d(3))
                        h = h + 60
                    Case "combo"
                        If HasHelper(CStr(d(0))) Then
                            Set c = mUI.AddCombo(fr, "c_" & d(0), 4, h, iw - 186)
                            Set b = mUI.AddButton(fr, "hb_" & d(0), CStr(mHelper.GetVal(CStr(d(0)))(0)), iw - 176, h, 176, , ROW_H)
                            Set e = New cEvt: e.Bind Me, b, "hlp:" & d(0): mEvts.Add e
                        Else
                            Set c = mUI.AddCombo(fr, "c_" & d(0), 4, h, iw)
                        End If
                        If IsArray(d(4)) Then
                            For i = LBound(d(4)) To UBound(d(4))
                                c.AddItem CStr(d(4)(i))
                            Next
                        End If
                        On Error Resume Next
                        c.Value = CStr(d(3))
                        On Error GoTo 0
                        Set e = New cEvt: e.Bind Me, c, CStr(d(0)): mEvts.Add e
                        h = h + ROW_H
                    Case "style"
                        ' Rahmen mit Stilname, Farbleiste und "Ändern…"
                        ' oben der Name, darunter die acht Farben; rechts der Button
                        Set lbl = mUI.AddLabel(fr, "sb_" & d(0), "", 4, h, iw, 44)
                        lbl.BackStyle = fmBackStyleOpaque: lbl.BackColor = CWhite()
                        lbl.BorderStyle = fmBorderStyleSingle: lbl.BorderColor = CBorderStrong()
                        Set c = mUI.AddLabel(fr, "c_" & d(0), CStr(d(3)), 12, h + 5, iw - 120, 15, True, 9.5)
                        c.WordWrap = False
                        For i = 0 To 7
                            Set lbl = mUI.AddLabel(fr, "sw_" & d(0) & "_" & i, "", 12 + i * 18, h + 23, 15, 15)
                            lbl.BackStyle = fmBackStyleOpaque
                            lbl.BorderStyle = fmBorderStyleSingle
                        Next
                        Set b = mUI.AddButton(fr, "sp_" & d(0), "Ändern…", iw - 100, h + 10, 92, , 24)
                        Set e = New cEvt: e.Bind Me, b, "sty:" & d(0): mEvts.Add e
                        Set e = New cEvt: e.Bind Me, c, "sty:" & d(0): mEvts.Add e
                        h = h + 44
                    Case "date"
                        Set c = mUI.AddText(fr, "c_" & d(0), 4, h, 140)
                        c.Text = CStr(d(3))
                        Set b = mUI.AddButton(fr, "d_" & d(0), "Datum…", 150, h, 76, , ROW_H)
                        Set e = New cEvt: e.Bind Me, b, "date:" & d(0): mEvts.Add e
                        h = h + ROW_H
                    Case "list"
                        Set c = mUI.AddList(fr, "c_" & d(0), 4, h, iw, 120, True)
                        If IsArray(d(4)) Then
                            If ArrLen(d(4)) > 0 Then c.List = d(4)
                        End If
                        ' Mehr als neun Einträge: nur ganze Zeilen zeigen (120 pt werden 9 Zeilen, 117 pt). Mit 9,2 Zeilen blieb der letzte
                        ' Eintrag angeschnitten und ließ sich nicht hineinrollen, weil MSForms die angeschnittene Zeile als sichtbar zählt.
                        If c.ListCount > 9 Then
                            c.IntegralHeight = True
                            c.Height = 120
                        End If
                        For Each a In Split(CStr(d(3)), ";")
                            If Len(a) > 0 Then If CLng(a) < c.ListCount Then c.Selected(CLng(a)) = True
                        Next
                        ' Auswahl merken: Mehrfachlisten verlieren ihre Haken beim Ausblenden des Dialogs
                        If mListSel Is Nothing Then Set mListSel = New cMap
                        mListSel.SetVal CStr(d(0)), SelectedIdx(c)
                        Set e = New cEvt: e.Bind Me, c, "lst:" & d(0): mEvts.Add e
                        h = h + c.Height
                End Select
                mCtl.SetVal CStr(d(0)), c
                If Len(d(5)) > 0 Then
                    Set lbl = mUI.AddLabel(fr, "h_" & d(0), CStr(d(5)), 4, h + 2, iw, 14, False, 8, CMuted())
                    lbl.AutoSize = True
                    If lbl.Width > iw Then lbl.Width = iw
                    h = h + lbl.Height + 2
                End If
                Set lbl = mUI.AddLabel(fr, "e_" & d(0), "", 4, h + 2, iw, 14, False, 8.5, CDanger())
                lbl.Visible = False
                mErr.SetVal CStr(d(0)), lbl
                h = h + 22
        End Select
        ' Block merken: Steuerelemente mit Abstand zum Blockanfang (für Ein-/Ausblenden)
        ReDim nms(0 To Application.Max(0, fr.Controls.Count - n0 - 1)): ReDim offs(0 To Application.Max(0, fr.Controls.Count - n0 - 1))
        For i = n0 To fr.Controls.Count - 1
            nms(i - n0) = fr.Controls(i).Name
            offs(i - n0) = fr.Controls(i).Top - h0
        Next
        mBlocks.Add Array(CStr(d(0)), nms, offs, h - h0, fr.Controls.Count > n0)
    Next
    If AllowBack Then
        Set btnBack = mUI.AddButton(Me, "back", ChrW(&H2190) & "  Zurück", PAD, 0, 100)
        Set btnOk = mUI.AddButton(Me, "ok", IIf(Len(mOkCaption) > 0, mOkCaption, "Übernehmen"), PAD + 108, 0, 120, "primary")
    Else
        Set btnOk = mUI.AddButton(Me, "ok", IIf(Len(mOkCaption) > 0, mOkCaption, "Übernehmen"), PAD, 0, 120, "primary")
    End If
    Set btnCancel = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 104, 0, 104)
    btnCancel.Cancel = True
    mBuilt = True
    For Each d In mDefs
        If d(1) = "style" Then ShowStyle CStr(d(0))
    Next
    Layout
End Sub

Private Function HasHelper(ByVal key As String) As Boolean
    If mHelper Is Nothing Then Exit Function
    HasHelper = mHelper.Exists(key)
End Function

' Name und Farbleiste eines Stilfeldes anzeigen
Private Sub ShowStyle(ByVal key As String)
    Dim fr As MSForms.Frame, nm As String, cols As Variant, i As Long, inherit As Boolean
    Set fr = Me.Controls("fr")
    nm = CStr(mStyleVal.GetVal(key, "Akzentfarbe"))
    inherit = Not mDocStyle.IsStyleName(nm)
    fr.Controls("c_" & key).caption = nm
    cols = mDocStyle.StyleColors(nm)
    For i = 0 To 7
        With fr.Controls("sw_" & key & "_" & i)
            .Visible = Not inherit
            .BackColor = cols(i)
            .BorderColor = Blend(cols(i), RGB(0, 0, 0), 0.15)
        End With
    Next
End Sub

Private Sub PickStyle(ByVal key As String)
    Dim f As frmStylePick, d As Variant, inheritLabel As String, ttl As String
    For Each d In mDefs
        If CStr(d(0)) = key Then inheritLabel = CStr(d(4)): ttl = CStr(d(2))
    Next
    Set f = New frmStylePick
    f.Setup ttl, CStr(mStyleVal.GetVal(key, "Akzentfarbe")), inheritLabel
    f.Show
    If f.OK Then
        mStyleVal.SetVal key, f.Result
        ShowStyle key
        ' wie eine geänderte Auswahl behandeln (Vorschau, abhängige Angaben)
        HandleEvent key, "Change"
    End If
    Unload f
End Sub

' Abhängige Angabe: key nur zeigen, wenn das Kästchen ctrlKey angehakt ist bzw. die Auswahl
' ctrlKey einen der Werte in values (";"-getrennt) hat. Vereinfacht die Dialoge.
Public Sub ShowWhen(ByVal key As String, ByVal ctrlKey As String, Optional ByVal values As String = "")
    If mDep Is Nothing Then Set mDep = New cMap
    mDep.SetVal key, Array(ctrlKey, values)
End Sub

Private Function BlockVisible(ByVal key As String) As Boolean
    Dim dp As Variant, c As Object, v As Variant
    BlockVisible = True
    If mDep Is Nothing Or Len(key) = 0 Then Exit Function
    If Not mDep.Exists(key) Then Exit Function
    dp = mDep.GetVal(key)
    If Not mCtl.Exists(CStr(dp(0))) Then Exit Function
    Set c = mCtl.GetVal(CStr(dp(0)))
    If TypeName(c) = "CheckBox" Then
        BlockVisible = CBool(c.Value)
    Else
        BlockVisible = False
        For Each v In Split(CStr(dp(1)), ";")
            If StrComp(CStr(v), c.Text, vbTextCompare) = 0 Then BlockVisible = True
        Next
    End If
End Function

' Blöcke untereinander setzen (ausgeblendete rücken zusammen), Rahmen und Schaltflächen anpassen
Private Sub Layout()
    Dim fr As MSForms.Frame, b As Variant, h As Single, vis As Boolean, i As Long, y As Single, maxH As Single
    Set fr = Me.Controls("fr")
    h = 6
    For Each b In mBlocks
        vis = BlockVisible(CStr(b(0)))
        If b(4) Then
            For i = 0 To UBound(b(1))
                ' nur Geaendertes setzen: Listenfelder verlieren sonst ihre Mehrfachauswahl
                With fr.Controls(CStr(b(1)(i)))
                    If Abs(.Top - (h + b(2)(i))) > 0.1 Then .Top = h + b(2)(i)
                    If Left$(.Name, 2) = "e_" Then
                        If Not vis And .Visible Then .Visible = False
                    ElseIf Left$(.Name, 3) = "sw_" Then
                        ' Farbleiste nur bei eigenem Stil, nicht bei "(wie Dashboard)"
                        If .Visible <> (vis And mDocStyle.IsStyleName(CStr(mStyleVal.GetVal(Mid$(.Name, 4, InStrRev(.Name, "_") - 4), "")))) Then
                            .Visible = (vis And mDocStyle.IsStyleName(CStr(mStyleVal.GetVal(Mid$(.Name, 4, InStrRev(.Name, "_") - 4), ""))))
                        End If
                    ElseIf .Visible <> vis Then
                        .Visible = vis
                    End If
                End With
            Next
        End If
        If vis Then h = h + b(3)
    Next
    maxH = mUI.MaxFormHeight() - mFrTop - 90
    If h + 6 > maxH Then
        fr.Height = maxH
        fr.ScrollHeight = h + 10
        fr.ScrollBars = fmScrollBarsVertical
    Else
        fr.Height = h + 6
        fr.ScrollBars = fmScrollBarsNone
    End If
    y = fr.Top + fr.Height + 12
    btnOk.Top = y
    btnCancel.Top = y
    If Not btnBack Is Nothing Then btnBack.Top = y
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    If Me.Visible Then mUI.KeepOnScreen Me
End Sub

Public Function ShowDialog(Optional ByVal okCaption As String = "") As Boolean
    mOkCaption = okCaption
    If Not mBuilt Then Build
    OK = False
    WentBack = False
    RestoreLists
    If Len(PreviewMacro) > 0 Then
        ' Vorschau-Dialoge rechts unten, damit der Inhalt dahinter sichtbar bleibt
        On Error Resume Next
        Me.StartUpPosition = 0
        Me.Left = Application.Left + Application.Width - Me.Width - 24
        Me.Top = Application.Top + Application.Height - Me.Height - 40
        On Error GoTo 0
    End If
    Me.Show
    ShowDialog = OK
End Function

' Gemerkte Auswahl der Mehrfachlisten wieder anzeigen (nach erneutem Öffnen, z. B. bei Fehlern)
Private Sub RestoreLists()
    Dim k As Variant, lb As MSForms.ListBox, i As Long, sel As String
    If mListSel Is Nothing Then Exit Sub
    For Each k In mListSel.Keys
        Set lb = mCtl.GetVal(CStr(k))
        sel = ";" & CStr(mListSel.GetVal(CStr(k))) & ";"
        For i = 0 To lb.ListCount - 1
            If lb.Selected(i) <> (InStr(sel, ";" & i & ";") > 0) Then lb.Selected(i) = (InStr(sel, ";" & i & ";") > 0)
        Next
    Next
End Sub

Public Function GetText(ByVal key As String) As String
    Dim c As Object
    If Not mStyleVal Is Nothing Then
        If mStyleVal.Exists(key) Then GetText = CStr(mStyleVal.GetVal(key)): Exit Function
    End If
    If Not mCtl.Exists(key) Then Exit Function
    Set c = mCtl.GetVal(key)
    Select Case TypeName(c)
        Case "CheckBox": GetText = IIf(c.Value, "Ja", "Nein")
        Case "ComboBox": GetText = c.Text
        Case "ListBox"
            If Not mListSel Is Nothing Then
                If mListSel.Exists(key) Then GetText = CStr(mListSel.GetVal(key)): Exit Function
            End If
            GetText = SelectedIdx(c)
        Case Else: GetText = c.Text
    End Select
End Function

Public Function GetBool(ByVal key As String) As Boolean
    If mCtl.Exists(key) Then GetBool = CBool(mCtl.GetVal(key).Value)
End Function

Public Function GetIndex(ByVal key As String) As Long
    If mCtl.Exists(key) Then GetIndex = mCtl.GetVal(key).ListIndex Else GetIndex = -1
End Function

Private Function SelectedIdx(ByVal lb As MSForms.ListBox) As String
    Dim i As Long, s As String
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then s = s & IIf(Len(s) > 0, ";", "") & i
    Next
    SelectedIdx = s
End Function

Public Sub SetError(ByVal key As String, ByVal msg As String)
    Dim lbl As MSForms.Label
    If Not mErr.Exists(key) Then mUI.Alert msg, mTitle: Exit Sub
    Set lbl = mErr.GetVal(key)
    lbl.caption = ChrW(&H26A0) & "  " & msg
    lbl.Visible = True
    mUI.MarkError mCtl.GetVal(key), True
    On Error Resume Next
    mCtl.GetVal(key).SetFocus
End Sub

Public Sub ClearErrors()
    Dim k As Variant
    For Each k In mErr.Keys
        mErr.GetVal(CStr(k)).Visible = False
        mUI.MarkError mCtl.GetVal(CStr(k)), False
    Next
End Sub

Public Sub SetComboItems(ByVal key As String, ByVal items As Variant, ByVal value As String)
    Dim c As MSForms.ComboBox, i As Long
    Set c = mCtl.GetVal(key)
    c.Clear
    For i = LBound(items) To UBound(items)
        c.AddItem CStr(items(i))
    Next
    On Error Resume Next
    c.Value = value
End Sub

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    Dim res As String, ok_ As Boolean, k As String
    If Left$(key, 4) = "sty:" Then
        If evt = "Click" Then PickStyle Mid$(key, 5)
        Exit Sub
    End If
    If Left$(key, 4) = "hlp:" Then
        If evt = "Click" Then
            k = Mid$(key, 5)
            res = CStr(Application.Run(mRender.MacroRef(CStr(mHelper.GetVal(k)(1))), mCtl.GetVal(k).Text))
            If Len(res) > 0 Then
                On Error Resume Next
                mCtl.GetVal(k).Value = res
                On Error GoTo 0
            End If
        End If
        Exit Sub
    End If
    If Left$(key, 4) = "lst:" Then
        If evt = "Change" Then
            k = Mid$(key, 5)
            If mCtl.Exists(k) Then mListSel.SetVal k, SelectedIdx(mCtl.GetVal(k))
        End If
        Exit Sub
    End If
    If Left$(key, 4) = "chk:" Then
        If Not mDep Is Nothing And mBuilt Then Layout
        Exit Sub
    End If
    If evt = "Change" And Not mDep Is Nothing And mBuilt Then Layout
    If Left$(key, 5) = "date:" Then
        k = Mid$(key, 6)
        res = mForms.PickDate(mCtl.GetVal(k).Text, False, "", ok_)
        If ok_ Then mCtl.GetVal(k).Text = res
    ElseIf evt = "Change" Then
        ChangedKey = key
        ' Live-Vorschau: der Aufrufer zeigt die Auswahl sofort hinter dem Dialog
        If Len(PreviewMacro) > 0 And mBuilt Then
            On Error Resume Next
            Application.Run mRender.MacroRef(PreviewMacro), Me
            On Error GoTo 0
        End If
    End If
End Sub

Private Sub btnOk_Click()
    OK = True
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    OK = False
    Me.Hide
End Sub

Private Sub btnBack_Click()
    OK = False
    WentBack = True
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
