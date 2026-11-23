'==============================================================================
' frmCaptureIntro - "So funktioniert's": Erfassungsdateien in einem Bild.
' Vier Schritte als Kacheln mit Symbol und Pfeil, darunter die beiden Arten
' (neue Einträge sammeln / Bestand prüfen) in je zwei Sätzen.
'==============================================================================
Option Explicit

Public StartNew As Boolean
Private WithEvents btnStart As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

Private Const CW As Single = 150      ' Breite einer Schritt-Kachel
Private Const CH As Single = 122
Private Const GAP As Single = 26

Public Sub Setup()
    Dim i As Long, x As Single, y As Single, w As Single, acc As Long, lbl As MSForms.Label
    Dim icons As Variant, ttl As Variant, txt As Variant
    icons = Array(ChrW(&HE8FD), ChrW(&HE715), ChrW(&HE8B5), ChrW(&HE9D9))
    ttl = Array("Liste festlegen", "Datei verschicken", "Rücklauf einlesen", "Gesamtliste")
    txt = Array("Welche Spalten sollen ausgefüllt werden? Sie wählen sie aus Ihrer Tabelle.", _
                "Eine normale Excel-Datei ohne Makros, per E-Mail an alle, z. B. 60 Kitas.", _
                "Ausgefüllte Dateien zurückbekommen und mehrere auf einmal auswählen.", _
                "Alle Antworten stehen in Ihrer Tabelle, mit dem Namen des Absenders.")
    acc = CAccentFill()
    mUI.StyleForm Me
    Me.Width = PAD * 2 + CW * 4 + GAP * 3 + (Me.Width - Me.InsideWidth)
    Me.caption = "So funktioniert's"
    w = Me.InsideWidth - PAD * 2
    mUI.AddLabel Me, "t", "Erfassungsdateien: Daten von vielen einsammeln", PAD, PAD - 2, w, 22, True, 13
    mUI.AddLabel Me, "t2", "Sie bauen einmal eine Liste, schicken sie an alle Beteiligten und lesen die Antworten wieder ein. " & _
                 "Die Empfänger brauchen nur Excel, kein Programm, keine Makros.", PAD, PAD + 24, w, 30, False, 9.5, CMuted()
    y = PAD + 64
    For i = 0 To 3
        x = PAD + i * (CW + GAP)
        Set lbl = mUI.AddLabel(Me, "c" & i, "", x, y, CW, CH)
        lbl.BackStyle = fmBackStyleOpaque
        lbl.BackColor = Blend(acc, RGB(255, 255, 255), 0.9)
        lbl.BorderStyle = fmBorderStyleSingle
        lbl.BorderColor = Blend(acc, RGB(255, 255, 255), 0.6)
        Set lbl = mUI.AddLabel(Me, "n" & i, CStr(i + 1), x + 10, y + 10, 24, 24, True, 11, COnAccent())
        lbl.BackStyle = fmBackStyleOpaque
        lbl.BackColor = acc
        lbl.TextAlign = fmTextAlignCenter
        mUI.AddIcon Me, "i" & i, CStr(icons(i)), x + CW - 36, y + 8, 20, CAccent()
        mUI.AddLabel Me, "h" & i, CStr(ttl(i)), x + 10, y + 42, CW - 20, 18, True, 10.5
        Set lbl = mUI.AddLabel(Me, "x" & i, CStr(txt(i)), x + 10, y + 62, CW - 20, CH - 66, False, 8.5, CMuted())
        lbl.WordWrap = True
        If i < 3 Then mUI.AddIcon Me, "a" & i, ChrW(&HE72A), x + CW + 4, y + CH / 2 - 10, 16, Blend(acc, RGB(255, 255, 255), 0.35)
    Next
    y = y + CH + 18
    ' die beiden Arten
    ModeCard "m1", PAD, y, (w - GAP) / 2, "Neue Einträge sammeln", "am häufigsten", _
             "Alle bekommen dieselbe leere Liste. Jede Stelle trägt ihre Einträge ein, etwa alle Kinder mit Förderbedarf. " & _
             "Beim Einlesen entstehen neue Einträge in Ihrer Tabelle.", acc
    ModeCard "m2", PAD + (w - GAP) / 2 + GAP, y, (w - GAP) / 2, "Bestand prüfen", "erweitert", _
             "Jede Stelle bekommt ihre vorhandenen Einträge, prüft sie und ergänzt fehlende Angaben. " & _
             "Beim Einlesen werden die Einträge aktualisiert.", RGB(100, 116, 139)
    y = y + 84
    Set lbl = mUI.AddLabel(Me, "tip", ChrW(&H2192) & "  Tipp: Legen Sie fest, wer antworten soll (z. B. alle Kitas). Dann sehen Sie unter „Rückmeldungen“, wer noch fehlt.", _
                           PAD, y, w, 18, False, 9, CText())
    y = y + 28
    Set btnStart = mUI.AddButton(Me, "start", "Neue Erfassung starten", PAD, y, 190, "primary")
    Set btnClose = mUI.AddButton(Me, "close", "Schließen", Me.InsideWidth - PAD - 104, y, 104)
    btnClose.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
End Sub

Private Sub ModeCard(ByVal nm As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal ttl As String, ByVal tag As String, _
                     ByVal txt As String, ByVal col As Long)
    Dim lbl As MSForms.Label
    Set lbl = mUI.AddLabel(Me, nm & "bg", "", x, y, w, 72)
    lbl.BackStyle = fmBackStyleOpaque: lbl.BackColor = CWhite()
    lbl.BorderStyle = fmBorderStyleSingle: lbl.BorderColor = CBorder()
    Set lbl = mUI.AddLabel(Me, nm & "bar", "", x, y, 4, 72)
    lbl.BackStyle = fmBackStyleOpaque: lbl.BackColor = col
    mUI.AddLabel Me, nm & "t", ttl, x + 14, y + 10, w - 120, 18, True, 10.5
    Set lbl = mUI.AddLabel(Me, nm & "g", tag, x + w - 100, y + 11, 88, 16, False, 8, CMuted())
    lbl.TextAlign = fmTextAlignRight
    Set lbl = mUI.AddLabel(Me, nm & "x", txt, x + 14, y + 32, w - 26, 50, False, 8.5, CMuted())
    lbl.WordWrap = True
End Sub

Private Sub btnStart_Click()
    StartNew = True
    Me.Hide
End Sub

Private Sub btnClose_Click()
    StartNew = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True: btnClose_Click
End Sub
