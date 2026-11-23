'==============================================================================
' frmMsg - Meldungen, Rueckfragen, Eingaben, Listenauswahl und Berichte im
' einheitlichen Design. Wird vollstaendig zur Laufzeit aufgebaut.
'==============================================================================
Option Explicit

Private mIsReport As Boolean      ' Bericht: Liste dient nur zur Anzeige, keine Auswahl noetig

Public Result As Long
Public InputValue As String
Public SelectedIndex As Long
Public SelectedList As String

Private WithEvents b1 As MSForms.CommandButton
Private WithEvents b2 As MSForms.CommandButton
Private WithEvents b3 As MSForms.CommandButton
Private WithEvents txt As MSForms.TextBox
Private WithEvents lst As MSForms.ListBox
Private mW As Single
Private mY As Single
Private mButtons As Variant
Private mDanger As Boolean
Private mHasList As Boolean

Private Sub Header(ByVal title As String, ByVal msg As String, ByVal kind As String, ByVal w As Single)
    Dim glyph As String, col As Long, lbl As MSForms.Label
    mW = w
    mUI.StyleForm Me
    Me.caption = title
    Me.Width = w
    Select Case kind
        Case "error": glyph = ChrW(&HEA39): col = CDanger()
        Case "warn": glyph = ChrW(&HE7BA): col = CWarn()
        Case "question": glyph = ChrW(&HE9CE): col = CAccent()
        Case Else: glyph = ChrW(&HE946): col = CAccent()
    End Select
    mUI.AddIcon Me, "icn", glyph, PAD - 4, PAD - 2, 16, col
    mUI.AddLabel Me, "lblTitle", title, PAD + 24, PAD, w - PAD * 2 - 30, 20, True, 12
    mY = PAD + 30
    If Len(msg) > 0 Then
        Set lbl = mUI.AddLabel(Me, "lblMsg", msg, PAD, mY, Me.InsideWidth - PAD * 2, 20)
        lbl.AutoSize = True
        If lbl.Width > Me.InsideWidth - PAD * 2 Then lbl.Width = Me.InsideWidth - PAD * 2
        mY = mY + lbl.Height + 12
    End If
End Sub

Private Sub AddButtons(ByVal buttons As Variant)
    Dim i As Long, n As Long, x As Single, bw As Single, c As MSForms.CommandButton, kind As String
    mButtons = buttons
    n = UBound(buttons) - LBound(buttons) + 1
    ' Platzbedarf aller Schaltflächen: bei Bedarf das Fenster verbreitern (sonst ragte die
    ' linke Schaltfläche über den Rand hinaus)
    Dim need As Single
    For i = 1 To n
        need = need + BtnWidth(CStr(buttons(LBound(buttons) + i - 1))) + IIf(i > 1, GAP, 0)
    Next
    If need + PAD * 2 > Me.InsideWidth Then Me.Width = Me.Width + (need + PAD * 2 - Me.InsideWidth)
    x = Me.InsideWidth - PAD
    For i = n To 1 Step -1
        bw = BtnWidth(CStr(buttons(LBound(buttons) + i - 1)))
        x = x - bw
        If i = 1 Then
            kind = IIf(mDanger, "danger", "primary")
        Else
            kind = "secondary"
        End If
        Set c = mUI.AddButton(Me, "b" & i, CStr(buttons(LBound(buttons) + i - 1)), x, mY + 4, bw, kind)
        Select Case i
            Case 1: Set b1 = c
            Case 2: Set b2 = c
            Case 3: Set b3 = c
        End Select
        x = x - GAP
    Next
    If n >= 2 Then
        Me.Controls("b" & n).Cancel = True
    Else
        b1.Cancel = True
    End If
    If txt Is Nothing And Not mHasList Then b1.Default = True
    mY = mY + 4 + BTN_H + PAD
    Me.Height = mY + (Me.Height - Me.InsideHeight)
End Sub

Private Function BtnWidth(ByVal cap As String) As Single
    BtnWidth = 96
    If Len(cap) > 12 Then BtnWidth = 7.2 * Len(cap) + 28
End Function

Public Sub Setup(ByVal title As String, ByVal msg As String, ByVal buttons As Variant, Optional ByVal kind As String = "info", Optional ByVal danger As Boolean = False)
    mDanger = danger
    Header title, msg, kind, IIf(Len(msg) > 400, 560, 440)
    AddButtons buttons
End Sub

Public Sub SetupInput(ByVal title As String, ByVal msg As String, ByVal dflt As String, ByVal password As Boolean)
    Header title, msg, "question", 440
    Set txt = mUI.AddText(Me, "txtIn", PAD, mY, Me.InsideWidth - PAD * 2)
    txt.Text = dflt
    If password Then txt.PasswordChar = "*"
    mY = mY + ROW_H + 12
    AddButtons Array("OK", "Abbrechen")
End Sub

' withBack: mehrstufiger Assistent: Schaltflaeche "Zurueck" liefert Result = 2 (Abbrechen = 3)
Public Sub SetupList(ByVal title As String, ByVal msg As String, ByVal items As Variant, ByVal multi As Boolean, ByVal okCaption As String, _
                     ByVal cols As Long, ByVal colWidths As String, ByVal preIndex As Long, Optional ByVal withBack As Boolean = False)
    Dim h As Single
    mHasList = True
    Header title, msg, "question", IIf(cols > 1, 620, 460)
    h = 14 * 16
    If h > mUI.MaxFormHeight - mY - 120 Then h = mUI.MaxFormHeight - mY - 120
    Set lst = mUI.AddList(Me, "lst", PAD, mY, Me.InsideWidth - PAD * 2, h, multi, cols, colWidths)
    If IsArray(items) Then
        If ArrLen(items) > 0 Then lst.List = items
    End If
    If preIndex >= 0 And preIndex < lst.ListCount Then lst.ListIndex = preIndex
    mY = mY + h + 12
    If withBack Then
        AddButtons Array(okCaption, ChrW(&H2190) & "  Zurück", "Abbrechen")
    Else
        AddButtons Array(okCaption, "Abbrechen")
    End If
End Sub

Public Sub SetupReport(ByVal title As String, ByVal msg As String, ByVal items As Variant, ByVal cols As Long, ByVal colWidths As String, _
                       ByVal buttons As Variant, ByVal danger As Boolean, Optional ByVal headers As Variant)
    Dim h As Single, sumW As Single, w As Variant, x As Single, i As Long, lbl As MSForms.Label, fw As Single
    mHasList = True
    mIsReport = True
    mDanger = danger
    ' Fensterbreite nach den Spalten, damit nichts seitlich verschwindet
    For Each w In Split(colWidths, ";")
        sumW = sumW + Val(w)
    Next
    fw = Application.Max(640, sumW + PAD * 2 + 36)
    If fw > mUI.MaxFormWidth() Then fw = mUI.MaxFormWidth()
    Header title, msg, IIf(danger, "warn", "info"), fw
    ' Spaltenköpfe genau über den Spalten der Liste
    If Not IsMissing(headers) Then
        If IsArray(headers) Then
            x = PAD + 7
            i = 0
            For Each w In Split(colWidths, ";")
                If i <= UBound(headers) Then
                    Set lbl = mUI.AddLabel(Me, "hd" & i, CStr(headers(i)), x, mY, Val(w) - 4, 15, True, 8.5, CMuted())
                    lbl.WordWrap = False
                End If
                x = x + Val(w)
                i = i + 1
            Next
            mY = mY + 17
        End If
    End If
    h = 260
    If h > mUI.MaxFormHeight - mY - 110 Then h = mUI.MaxFormHeight - mY - 110
    If h < 80 Then h = 80
    Set lst = mUI.AddList(Me, "lst", PAD, mY, Me.InsideWidth - PAD * 2, h, False, cols, colWidths)
    If Not IsEmpty(items) Then
        On Error Resume Next
        lst.List = items
        On Error GoTo 0
    End If
    mY = mY + h + 12
    AddButtons buttons
End Sub

Private Sub Finish(ByVal r As Long)
    Dim i As Long
    Result = r
    If Not txt Is Nothing Then InputValue = txt.Text
    If Not lst Is Nothing Then
        SelectedIndex = lst.ListIndex
        SelectedList = ""
        If lst.MultiSelect <> fmMultiSelectSingle Then
            For i = 0 To lst.ListCount - 1
                If lst.Selected(i) Then SelectedList = SelectedList & IIf(Len(SelectedList) > 0, ";", "") & i
            Next
        End If
    End If
    Me.Hide
End Sub

Public Sub PreselectMany(ByVal idxList As String)
    Dim v As Variant
    For Each v In Split(idxList, ";")
        If Len(v) > 0 Then
            If CLng(v) < lst.ListCount Then lst.Selected(CLng(v)) = True
        End If
    Next
End Sub

Private Sub b1_Click()
    If mHasList And Not lst Is Nothing And Not mIsReport Then
        If lst.ListCount > 0 And lst.ListIndex < 0 And UBound(mButtons) = 1 And lst.MultiSelect = fmMultiSelectSingle Then
            If InStr(1, b1.caption, "Übernehmen") > 0 Or InStr(1, b1.caption, "Auswählen") > 0 Or InStr(1, b1.caption, "Öffnen") > 0 Then
                mUI.Alert "Bitte zuerst einen Eintrag in der Liste auswählen.", "Auswahl"
                Exit Sub
            End If
        End If
    End If
    Finish 1
End Sub

Private Sub b2_Click()
    Finish 2
End Sub

Private Sub b3_Click()
    Finish 3
End Sub

Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lst.ListIndex >= 0 Then Finish 1
End Sub

Private Sub txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = 0: Finish 1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Finish 0
    End If
End Sub
