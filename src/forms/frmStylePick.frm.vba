'==============================================================================
' frmStylePick - Farbstil wählen: je Stil Name, kurze Beschreibung und die acht
' Farben als Musterleiste. Klick wählt, Doppelklick übernimmt sofort.
' Optional ein erster Eintrag "(wie Dashboard)" ohne eigene Farben.
'==============================================================================
Option Explicit

Public Result As String
Public OK As Boolean
Private mNames As Collection
Private mSel As Long
Private mEvts As Collection
Private WithEvents btnOk As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton

' Zeilenabstand ROW_HT + 4 = 48 pt: bei 100 bis 175 % Skalierung ganze Pixel. Mit 46 pt (61,33 Pixel) zeichnete Excel jede dritte
' Beschreibung eine Stufe kleiner.
Private Const ROW_HT As Single = 44
Private Const SW As Single = 17

Public Sub Setup(ByVal title As String, ByVal current As String, Optional ByVal inheritLabel As String = "")
    Dim s As Variant, i As Long, y As Single, w As Single
    Set mNames = New Collection
    Set mEvts = New Collection
    If Len(inheritLabel) > 0 Then mNames.Add inheritLabel
    For Each s In mDocStyle.StyleNames()
        mNames.Add CStr(s)
    Next
    mUI.StyleForm Me
    Me.Width = 560
    Me.caption = title
    w = Me.InsideWidth - PAD * 2
    mUI.AddLabel Me, "t", title, PAD, PAD - 2, w, 20, True, 12
    mUI.AddLabel Me, "t2", "Die Farben rechts zeigen, wie Kopfzeilen und Diagramme aussehen. Helle Stile erhalten automatisch dunkle Schrift.", _
                 PAD, PAD + 20, w, 28, False, 8.5, CMuted()
    y = PAD + 52
    mSel = 1
    For i = 1 To mNames.Count
        AddRow i, y, w, inheritLabel
        If mNames(i) = current Then mSel = i
        y = y + ROW_HT + 4
    Next
    y = y + 8
    Set btnOk = mUI.AddButton(Me, "ok", "Übernehmen", PAD, y, 120, "primary")
    Set btnCancel = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 104, y, 104)
    btnCancel.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    If Me.Height > mUI.MaxFormHeight() Then
        ' kleiner Bildschirm: Fenster begrenzen, Liste rollt
        Me.Height = mUI.MaxFormHeight()
        Me.ScrollBars = fmScrollBarsVertical
        Me.ScrollHeight = y + BTN_H + PAD
    End If
    MarkSel
End Sub

Private Sub AddRow(ByVal i As Long, ByVal y As Single, ByVal w As Single, ByVal inheritLabel As String)
    Dim bg As MSForms.Label, lbl As MSForms.Label, cols As Variant, k As Long, nm As String, x As Single
    nm = mNames(i)
    Set bg = mUI.AddLabel(Me, "bg" & i, "", PAD, y, w, ROW_HT)
    bg.BackStyle = fmBackStyleOpaque
    bg.BackColor = CWhite()
    bg.BorderStyle = fmBorderStyleSingle
    bg.BorderColor = CBorder()
    Bind bg, i
    Set lbl = mUI.AddLabel(Me, "nm" & i, nm, PAD + 12, y + 6, 230, 16, True, 10)
    Bind lbl, i
    If Len(inheritLabel) > 0 And i = 1 Then
        Set lbl = mUI.AddLabel(Me, "hp" & i, "Übernimmt den Farbstil des Dashboards.", PAD + 12, y + 23, 300, 14, False, 8, CMuted())
        Bind lbl, i
        Exit Sub
    End If
    Set lbl = mUI.AddLabel(Me, "hp" & i, mDocStyle.StyleHelp(nm), PAD + 12, y + 23, w - 8 * (SW + 3) - 36, 14, False, 8, CMuted())
    lbl.WordWrap = False
    Bind lbl, i
    cols = mDocStyle.StyleColors(nm)
    x = PAD + w - 8 * (SW + 3) - 10
    For k = 0 To 7
        Set lbl = mUI.AddLabel(Me, "sw" & i & "_" & k, "", x + k * (SW + 3), y + (ROW_HT - SW) / 2, SW, SW)
        lbl.BackStyle = fmBackStyleOpaque
        lbl.BackColor = cols(k)
        lbl.BorderStyle = fmBorderStyleSingle
        lbl.BorderColor = Blend(cols(k), RGB(0, 0, 0), 0.15)
        Bind lbl, i
    Next
End Sub

Private Sub Bind(ByVal c As Object, ByVal i As Long)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, c, CStr(i)
    mEvts.Add e
End Sub

Private Sub MarkSel()
    Dim i As Long
    For i = 1 To mNames.Count
        With Me.Controls("bg" & i)
            .BackColor = IIf(i = mSel, CAccentSoft(), CWhite())
            .BorderColor = IIf(i = mSel, CAccent(), CBorder())
        End With
        Me.Controls("nm" & i).ForeColor = IIf(i = mSel, CText(), CText())
    Next
End Sub

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    Dim i As Long
    i = CLng(Val(key))
    If i < 1 Or i > mNames.Count Then Exit Sub
    If evt = "Click" Then
        mSel = i
        MarkSel
    ElseIf evt = "DblClick" Then
        mSel = i
        btnOk_Click
    End If
End Sub

Private Sub btnOk_Click()
    Result = mNames(mSel)
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
