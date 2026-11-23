'==============================================================================
' frmDate - Lokaler Kalender mit optionaler Uhrzeit (hh:mm). Uebernehmen
' aendert nur das Eingabefeld des aufrufenden Formulars; gespeichert wird dort.
'==============================================================================
Option Explicit

Public Result As String
Public OK As Boolean
Private mSel As Date
Private mMonth As Date
Private mWithTime As Boolean
Private mEvts As Collection
Private mLoading As Boolean
Private WithEvents cboMonth As MSForms.ComboBox
Private WithEvents cboYear As MSForms.ComboBox
Private WithEvents chkTime As MSForms.CheckBox
Private WithEvents btnOk As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private txtH As MSForms.TextBox
Private txtM As MSForms.TextBox
Private lblSel As MSForms.Label
Private lblErr As MSForms.Label
Private Const CELL_W As Single = 46
Private Const CELL_H As Single = 28

Public Sub Setup(ByVal cur As String, ByVal allowTime As Boolean, ByVal fieldName As String)
    Dim d As Date, ht As Boolean, y As Single, i As Long, b As MSForms.CommandButton, x As Single, lbl As MSForms.Label, e As cEvt
    mLoading = True
    Set mEvts = New Collection
    mUI.StyleForm Me
    Me.Width = PAD * 2 + CELL_W * 7 + 6 * 4 + 12
    Me.caption = "Datum" & IIf(allowTime, " und Uhrzeit", "")
    If mUtil.ParseDateDE(cur, d, ht) Then
        mSel = d
    Else
        mSel = Date
        ht = False
    End If
    mMonth = DateSerial(Year(mSel), Month(mSel), 1)
    mUI.AddLabel Me, "t", "Datum" & IIf(allowTime, " und Uhrzeit", "") & IIf(Len(fieldName) > 0, ": " & fieldName, ""), PAD, PAD - 2, Me.InsideWidth - PAD * 2, 20, True, 12
    y = PAD + 28
    Set b = mUI.AddButton(Me, "prev", ChrW(&H2039), PAD, y, 30, , ROW_H): BindE b, "prev"
    Set cboMonth = mUI.AddCombo(Me, "cm", PAD + 36, y, 140)
    For i = 1 To 12: cboMonth.AddItem MonthName_DE(i): Next
    Set cboYear = mUI.AddCombo(Me, "cy", PAD + 182, y, 80)
    For i = Year(Date) - 30 To Year(Date) + 20: cboYear.AddItem CStr(i): Next
    Set b = mUI.AddButton(Me, "next", ChrW(&H203A), Me.InsideWidth - PAD - 30, y, 30, , ROW_H): BindE b, "next"
    y = y + ROW_H + 10
    For i = 1 To 7
        mUI.AddLabel Me, "wd" & i, WeekdayShort_DE(i), PAD + (i - 1) * (CELL_W + 4), y, CELL_W, 14, True, 8.5, CMuted()
    Next
    y = y + 16
    For i = 0 To 41
        x = PAD + (i Mod 7) * (CELL_W + 4)
        Set b = mUI.AddButton(Me, "d" & i, "", x, y + (i \ 7) * (CELL_H + 4), CELL_W, , CELL_H)
        b.Font.Name = UI_FONT_SEMI
        BindE b, "d" & i
    Next
    y = y + 6 * (CELL_H + 4) + 4
    Set lblSel = mUI.AddLabel(Me, "sel", "", PAD, y, Me.InsideWidth - PAD * 2, 15, False, 9.5)
    y = y + 22
    Set b = mUI.AddButton(Me, "today", "Heute", PAD, y, 100): BindE b, "today"
    Set b = mUI.AddButton(Me, "tomorrow", "Morgen", PAD + 108, y, 100): BindE b, "tomorrow"
    Set b = mUI.AddButton(Me, "plus7", "+7 Tage", PAD + 216, y, 100): BindE b, "plus7"
    y = y + BTN_H + 12
    mWithTime = allowTime
    If allowTime Then
        Set chkTime = mUI.AddCheck(Me, "ct", "Uhrzeit angeben", PAD, y, 130)
        Set txtH = mUI.AddText(Me, "th", PAD + 140, y, 40)
        mUI.AddLabel Me, "c1", ":", PAD + 183, y + 3, 8, 15
        Set txtM = mUI.AddText(Me, "tm", PAD + 192, y, 40)
        txtH.MaxLength = 2: txtM.MaxLength = 2
        txtH.TextAlign = fmTextAlignCenter: txtM.TextAlign = fmTextAlignCenter
        mUI.AddLabel Me, "tl", "Stunden : Minuten", PAD + 140, y + ROW_H + 2, 200, 14, False, 8, CMuted()
        chkTime.Value = ht
        txtH.Text = Format$(Hour(mSel), "00"): txtM.Text = Format$(Minute(mSel), "00")
        Call TimeState
        y = y + ROW_H + 22
    End If
    Set lblErr = mUI.AddLabel(Me, "err", "", PAD, y, Me.InsideWidth - PAD * 2, 15, False, 8.5, CDanger())
    y = y + 18
    mUI.AddLabel Me, "hint", "Übernehmen ändert das Eingabefeld. Gespeichert wird erst mit „Speichern“ im Eintrag.", PAD, y, Me.InsideWidth - PAD * 2, 26, False, 8.5, CMuted()
    y = y + 30
    Set btnOk = mUI.AddButton(Me, "ok", "Übernehmen", PAD, y, 120, "primary")
    Set btnCancel = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 110, y, 110)
    btnCancel.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    mLoading = False
    Call DrawMonth
End Sub

Private Sub BindE(ByVal c As Object, ByVal key As String)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, c, key
    mEvts.Add e
End Sub

Private Sub TimeState()
    Dim en As Boolean
    If chkTime Is Nothing Then Exit Sub
    en = chkTime.Value
    txtH.Enabled = en: txtM.Enabled = en
    txtH.Locked = False: txtM.Locked = False
    txtH.BackColor = IIf(en, CWhite(), CSurface())
    txtM.BackColor = IIf(en, CWhite(), CSurface())
End Sub

Private Sub DrawMonth()
    Dim first As Date, start As Date, i As Long, d As Date, b As MSForms.CommandButton
    mLoading = True
    cboMonth.ListIndex = Month(mMonth) - 1
    On Error Resume Next
    cboYear.Value = CStr(Year(mMonth))
    On Error GoTo 0
    mLoading = False
    first = mMonth
    start = first - Weekday(first, vbMonday) + 1
    For i = 0 To 41
        d = start + i
        Set b = Me.Controls("d" & i)
        b.caption = CStr(Day(d))
        b.Tag = CStr(CLng(d))
        If Int(CDbl(d)) = Int(CDbl(mSel)) Then
            b.BackColor = CAccentFill(): b.ForeColor = COnAccent()
        ElseIf Month(d) <> Month(mMonth) Then
            b.BackColor = CWhite(): b.ForeColor = RGB(170, 170, 170)
        ElseIf d = Date Then
            b.BackColor = CAccentSoft(): b.ForeColor = CText()
        Else
            b.BackColor = CWhite(): b.ForeColor = CText()
        End If
    Next
    lblSel.caption = "Gewählt: " & DateLong_DE(Int(mSel))
End Sub

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    Dim t As Date
    If evt <> "Click" Then Exit Sub
    Select Case key
        Case "prev": mMonth = DateSerial(Year(mMonth), Month(mMonth) - 1, 1)
        Case "next": mMonth = DateSerial(Year(mMonth), Month(mMonth) + 1, 1)
        Case "today": SetDay Date
        Case "tomorrow": SetDay Date + 1
        Case "plus7": SetDay Date + 7
        Case Else
            If Left$(key, 1) = "d" Then SetDay CDate(CLng(Me.Controls(key).Tag))
    End Select
    Call DrawMonth
End Sub

Private Sub SetDay(ByVal d As Date)
    mSel = Int(d) + (CDbl(mSel) - Int(CDbl(mSel)))
    mMonth = DateSerial(Year(d), Month(d), 1)
End Sub

Private Sub cboMonth_Change()
    If mLoading Or cboMonth.ListIndex < 0 Then Exit Sub
    mMonth = DateSerial(Year(mMonth), cboMonth.ListIndex + 1, 1)
    Call DrawMonth
End Sub

Private Sub cboYear_Change()
    If mLoading Or cboYear.ListIndex < 0 Then Exit Sub
    mMonth = DateSerial(CLng(cboYear.Value), Month(mMonth), 1)
    Call DrawMonth
End Sub

Private Sub chkTime_Click()
    Call TimeState
    If chkTime.Value Then
        On Error Resume Next
        txtH.SetFocus
    End If
End Sub

Private Sub btnOk_Click()
    Dim h As Long, m As Long, s As Long
    If mWithTime Then
        If chkTime.Value Then
            If Not (IsAllDigits(txtH.Text) And IsAllDigits(txtM.Text)) Then
                lblErr.caption = "Bitte die Uhrzeit nur mit Ziffern angeben (00 bis 23 : 00 bis 59).": Exit Sub
            End If
            h = CLng(txtH.Text): m = CLng(txtM.Text)
            If h > 23 Or m > 59 Then lblErr.caption = "Ungültige Uhrzeit: Stunden 00 bis 23, Minuten 00 bis 59.": Exit Sub
            Result = FmtDate(mSel) & " " & Format$(h, "00") & ":" & Format$(m, "00")
        Else
            Result = FmtDate(mSel)
        End If
    Else
        Result = FmtDate(mSel)
    End If
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
