'==============================================================================
' frmDesignPick - Dokumentstil wählen: je Vorlage Name (in ihrer Schrift),
' Beschreibung, Navigations- und Akzentfarbe sowie die Farben des Farbstils.
' Letzte Zeile "Eigene Auswahl…": Farben und Schrift selbst zusammenstellen.
'==============================================================================
Option Explicit

Public Result As String          ' Vorlagenname oder "Eigene Auswahl"
Public OK As Boolean
Private Const CUSTOM As String = "Eigene Auswahl"
Private mNames As Collection
Private mSel As Long
Private mEvts As Collection
Private WithEvents btnOk As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton

Private Const ROW_HT As Single = 44
Private Const SW As Single = 15

Public Sub Setup(ByVal current As String)
    Dim s As Variant, i As Long, y As Single, w As Single
    Set mNames = New Collection
    Set mEvts = New Collection
    For Each s In mDocStyle.DesignNames()
        mNames.Add CStr(s)
    Next
    mNames.Add CUSTOM
    mUI.StyleForm Me
    Me.Width = 600
    Me.caption = "Dokumentstil"
    w = Me.InsideWidth - PAD * 2
    mUI.AddLabel Me, "t", "Dokumentstil der ganzen Datei", PAD, PAD - 2, w, 20, True, 12
    mUI.AddLabel Me, "t2", "Eine Vorlage legt Navigationsfarbe, Akzentfarbe, Farbstil (Dashboards, Erfassungsdateien, Export) und Schrift gemeinsam fest. " & _
                 "Mit „Eigene Auswahl“ stellen Sie alles selbst zusammen.", PAD, PAD + 20, w, 28, False, 8.5, CMuted()
    y = PAD + 54
    mSel = mNames.Count
    For i = 1 To mNames.Count
        AddRow i, y, w
        If mNames(i) = current Then mSel = i
        y = y + ROW_HT + 4
    Next
    y = y + 8
    Set btnOk = mUI.AddButton(Me, "ok", "Übernehmen", PAD, y, 120, "primary")
    Set btnCancel = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 104, y, 104)
    btnCancel.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    If Me.Height > mUI.MaxFormHeight() Then Me.Height = mUI.MaxFormHeight(): Me.ScrollBars = fmScrollBarsVertical: Me.ScrollHeight = y + BTN_H + PAD
    MarkSel
End Sub

Private Sub AddRow(ByVal i As Long, ByVal y As Single, ByVal w As Single)
    Dim bg As MSForms.Label, lbl As MSForms.Label, d As Variant, cols As Variant, k As Long, nm As String, x As Single, c As Long
    nm = mNames(i)
    Set bg = mUI.AddLabel(Me, "bg" & i, "", PAD, y, w, ROW_HT)
    bg.BackStyle = fmBackStyleOpaque
    bg.BackColor = CWhite()
    bg.BorderStyle = fmBorderStyleSingle
    bg.BorderColor = CBorder()
    Bind bg, i
    d = mDocStyle.DesignSpec(nm)
    Set lbl = mUI.AddLabel(Me, "nm" & i, nm & IIf(IsEmpty(d), "…", ""), PAD + 12, y + 5, 250, 18, True, 10.5)
    If Not IsEmpty(d) Then lbl.Font.Name = CStr(d(3)): lbl.Font.Bold = True
    Bind lbl, i
    If IsEmpty(d) Then
        Set lbl = mUI.AddLabel(Me, "hp" & i, "Navigation, Akzent, Farbstil und Schrift einzeln wählen.", PAD + 12, y + 25, w - 24, 14, False, 8, CMuted())
        Bind lbl, i
        Exit Sub
    End If
    Set lbl = mUI.AddLabel(Me, "hp" & i, CStr(d(3)) & " · " & CStr(d(4)), PAD + 12, y + 25, w - 9 * (SW + 2) - 50, 14, False, 8, CMuted())
    lbl.WordWrap = False
    Bind lbl, i
    ' Navigation und Akzent als größere Felder, danach sechs Farben des Farbstils
    x = PAD + w - 9 * (SW + 2) - 24
    c = mCmd.NavValue(CStr(d(0)))
    AddSwatch "sn" & i, x, y + (ROW_HT - SW - 4) / 2, SW + 4, c, i
    x = x + SW + 7
    c = mCmd.AccentValue(CStr(d(1)))
    AddSwatch "sa" & i, x, y + (ROW_HT - SW - 4) / 2, SW + 4, c, i
    x = x + SW + 12
    cols = mDocStyle.StyleColors(CStr(d(2)))
    For k = 0 To 5
        AddSwatch "sw" & i & "_" & k, x + k * (SW + 2), y + (ROW_HT - SW) / 2, SW, cols(k), i
    Next
End Sub

Private Sub AddSwatch(ByVal nm As String, ByVal x As Single, ByVal y As Single, ByVal sz As Single, ByVal c As Long, ByVal i As Long)
    Dim lbl As MSForms.Label
    Set lbl = mUI.AddLabel(Me, nm, "", x, y, sz, sz)
    lbl.BackStyle = fmBackStyleOpaque
    lbl.BackColor = c
    lbl.BorderStyle = fmBorderStyleSingle
    lbl.BorderColor = Blend(c, RGB(0, 0, 0), 0.15)
    Bind lbl, i
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
