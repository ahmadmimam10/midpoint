Attribute VB_Name = "mUI"
'==============================================================================
' mUI - Einheitlicher Formularbaukasten (Raster, Schrift, Farben) und
' Standarddialoge im Anwendungsdesign (Meldung, Rueckfrage, Eingabe, Liste).
'==============================================================================
Option Explicit

Public Const PAD As Single = 18
Public Const ROW_H As Single = 22
Public Const BTN_H As Single = 26
Public Const LBL_H As Single = 15
Public Const GAP As Single = 8
Public Const FONT_SIZE As Single = 9.75

' Nur fuer automatisierte Pruefungen (mSelfTest): Meldungen werden protokolliert statt angezeigt,
' Rueckfragen gelten als bestaetigt. Im Normalbetrieb immer False.
Public gQuiet As Boolean
Public gQuietLog As String

'--- Grundelemente ----------------------------------------------------------------
Public Sub StyleForm(ByVal frm As Object)
    frm.BackColor = CWhite()
    frm.Font.Name = UI_FONT
    frm.Font.Size = FONT_SIZE
End Sub

Public Function AddLabel(ByVal container As Object, ByVal nm As String, ByVal caption As String, ByVal l As Single, ByVal t As Single, _
                         ByVal w As Single, Optional ByVal h As Single = LBL_H, Optional ByVal bold As Boolean = False, _
                         Optional ByVal size As Single = FONT_SIZE, Optional ByVal color As Long = -1) As MSForms.Label
    Dim c As MSForms.Label
    Set c = container.Controls.Add("Forms.Label.1", nm)
    With c
        .caption = caption
        .Left = l: .Top = t: .Width = w: .Height = h
        .BackStyle = fmBackStyleTransparent
        .Font.Name = IIf(bold, UI_FONT_SEMI, UI_FONT)
        .Font.Size = size
        .ForeColor = IIf(color = -1, CText(), color)
        .WordWrap = True
    End With
    Set AddLabel = c
End Function

Public Function AddIcon(ByVal container As Object, ByVal nm As String, ByVal glyph As String, ByVal l As Single, ByVal t As Single, _
                        Optional ByVal size As Single = 14, Optional ByVal color As Long = -1) As MSForms.Label
    Dim c As MSForms.Label
    Set c = container.Controls.Add("Forms.Label.1", nm)
    With c
        .caption = glyph
        .Left = l: .Top = t: .Width = size + 8: .Height = size + 6
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ICON_FONT
        .Font.Size = size
        .ForeColor = IIf(color = -1, CAccent(), color)
        .TextAlign = fmTextAlignCenter
    End With
    Set AddIcon = c
End Function

Public Function AddText(ByVal container As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, _
                        Optional ByVal h As Single = ROW_H, Optional ByVal multiline As Boolean = False) As MSForms.TextBox
    Dim c As MSForms.TextBox
    Set c = container.Controls.Add("Forms.TextBox.1", nm)
    With c
        .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Name = UI_FONT: .Font.Size = FONT_SIZE
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = CBorderStrong()
        .ForeColor = CText()
        .BackColor = CWhite()
        If multiline Then
            .MultiLine = True
            .WordWrap = True
            .EnterKeyBehavior = True
            .ScrollBars = fmScrollBarsVertical
        End If
    End With
    Set AddText = c
End Function

Public Function AddCombo(ByVal container As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, _
                         Optional ByVal dropList As Boolean = True) As MSForms.ComboBox
    Dim c As MSForms.ComboBox
    Set c = container.Controls.Add("Forms.ComboBox.1", nm)
    With c
        .Left = l: .Top = t: .Width = w: .Height = ROW_H
        .Font.Name = UI_FONT: .Font.Size = FONT_SIZE
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = CBorderStrong()
        .ForeColor = CText()
        .ShowDropButtonWhen = fmShowDropButtonWhenAlways
        .ListRows = 14
        If dropList Then
            .Style = fmStyleDropDownList
        Else
            .Style = fmStyleDropDownCombo
            .MatchEntry = fmMatchEntryComplete
        End If
    End With
    Set AddCombo = c
End Function

Public Function AddList(ByVal container As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                        Optional ByVal multi As Boolean = False, Optional ByVal cols As Long = 1, Optional ByVal colWidths As String = "") As MSForms.ListBox
    Dim c As MSForms.ListBox
    Set c = container.Controls.Add("Forms.ListBox.1", nm)
    With c
        ' vor der Hoehe abschalten, sonst wird sie auf ganze Zeilen abgerundet
        .IntegralHeight = False
        .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Name = UI_FONT: .Font.Size = FONT_SIZE
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = CBorderStrong()
        .ForeColor = CText()
        .ColumnCount = cols
        If Len(colWidths) > 0 Then .ColumnWidths = colWidths
        If multi Then
            .MultiSelect = fmMultiSelectMulti
            .ListStyle = fmListStyleOption
        End If
        .IntegralHeight = False
    End With
    Set AddList = c
End Function

Public Function AddCheck(ByVal container As Object, ByVal nm As String, ByVal caption As String, ByVal l As Single, ByVal t As Single, ByVal w As Single) As MSForms.CheckBox
    Dim c As MSForms.CheckBox
    Set c = container.Controls.Add("Forms.CheckBox.1", nm)
    With c
        .caption = caption
        .Left = l: .Top = t: .Width = w: .Height = ROW_H
        .Font.Name = UI_FONT: .Font.Size = FONT_SIZE
        .ForeColor = CText()
        .BackStyle = fmBackStyleTransparent
        .SpecialEffect = fmButtonEffectFlat
    End With
    Set AddCheck = c
End Function

Public Function AddOption(ByVal container As Object, ByVal nm As String, ByVal caption As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal grp As String) As MSForms.OptionButton
    Dim c As MSForms.OptionButton
    Set c = container.Controls.Add("Forms.OptionButton.1", nm)
    With c
        .caption = caption
        .Left = l: .Top = t: .Width = w: .Height = ROW_H
        .Font.Name = UI_FONT: .Font.Size = FONT_SIZE
        .ForeColor = CText()
        .BackStyle = fmBackStyleTransparent
        .GroupName = grp
        .SpecialEffect = fmButtonEffectFlat
    End With
    Set AddOption = c
End Function

' Art: "primary", "secondary", "danger", "ghost"
Public Function AddButton(ByVal container As Object, ByVal nm As String, ByVal caption As String, ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, Optional ByVal kind As String = "secondary", Optional ByVal h As Single = BTN_H) As MSForms.CommandButton
    Dim c As MSForms.CommandButton
    Set c = container.Controls.Add("Forms.CommandButton.1", nm)
    With c
        .caption = caption
        .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Name = UI_FONT_SEMI: .Font.Size = FONT_SIZE
        .WordWrap = False
        .TakeFocusOnClick = True
    End With
    StyleButton c, kind
    Set AddButton = c
End Function

Public Sub StyleButton(ByVal c As MSForms.CommandButton, ByVal kind As String)
    Select Case kind
        Case "primary": c.BackColor = CAccentFill(): c.ForeColor = COnAccent()
        Case "danger": c.BackColor = CDanger(): c.ForeColor = CWhite()
        Case "ghost": c.BackColor = CWhite(): c.ForeColor = CAccent()
        Case Else: c.BackColor = CSurface2(): c.ForeColor = CText()
    End Select
End Sub

Public Function AddFrame(ByVal container As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                         Optional ByVal caption As String = "", Optional ByVal scroll As Boolean = False) As MSForms.Frame
    Dim c As MSForms.Frame
    Set c = container.Controls.Add("Forms.Frame.1", nm)
    With c
        .Left = l: .Top = t: .Width = w: .Height = h
        .caption = caption
        .BackColor = CWhite()
        .Font.Name = UI_FONT: .Font.Size = FONT_SIZE
        .ForeColor = CMuted()
        If Len(caption) = 0 Then
            .SpecialEffect = fmSpecialEffectFlat
            .BorderStyle = fmBorderStyleNone
        Else
            .SpecialEffect = fmSpecialEffectFlat
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = CBorder()
        End If
        If scroll Then
            .ScrollBars = fmScrollBarsVertical
            .KeepScrollBarsVisible = fmScrollBarsNone
        End If
    End With
    Set AddFrame = c
End Function

' Duenne Trennlinie
Public Function AddLine(ByVal container As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, ByVal w As Single) As MSForms.Label
    Dim c As MSForms.Label
    Set c = container.Controls.Add("Forms.Label.1", nm)
    With c
        .caption = ""
        .Left = l: .Top = t: .Width = w: .Height = 1
        .BackStyle = fmBackStyleOpaque
        .BackColor = CBorder()
    End With
    Set AddLine = c
End Function

' Farbmarke (Optionsfarbe)
Public Function AddChip(ByVal container As Object, ByVal nm As String, ByVal caption As String, ByVal paletteKey As String, _
                        ByVal l As Single, ByVal t As Single, ByVal w As Single) As MSForms.Label
    Dim c As MSForms.Label
    Set c = container.Controls.Add("Forms.Label.1", nm)
    With c
        .caption = " " & caption
        .Left = l: .Top = t: .Width = w: .Height = 18
        .BackStyle = fmBackStyleOpaque
        .BackColor = PaletteBg(paletteKey)
        .ForeColor = PaletteFg(paletteKey)
        .Font.Name = UI_FONT: .Font.Size = 9
        .WordWrap = False
    End With
    Set AddChip = c
End Function

' Formulare richten sich nach dem Arbeitsbereich des Bildschirms (nicht nach dem Excel-Fenster)
Public Function MaxFormHeight() As Single
    MaxFormHeight = mWin.WorkAreaHeightPt() * 0.92
    If MaxFormHeight < 300 Then MaxFormHeight = 300
End Function

' Dialog, der waehrend der Anzeige gewachsen ist, vollstaendig sichtbar halten
' (sonst liegen die Schaltflaechen unten unter der Taskleiste)
Public Sub KeepOnScreen(ByVal frm As Object)
    Dim maxB As Single
    On Error Resume Next
    maxB = mWin.WorkAreaHeightPt() - 6
    If frm.Top + frm.Height > maxB Then frm.Top = Application.Max(4, maxB - frm.Height)
End Sub

Public Function MaxFormWidth() As Single
    MaxFormWidth = mWin.WorkAreaWidthPt() * 0.94
    If MaxFormWidth < 420 Then MaxFormWidth = 420
End Function

Public Sub MarkError(ByVal c As Object, ByVal isErr As Boolean)
    On Error Resume Next
    If isErr Then
        c.BorderColor = CDanger()
        c.BackColor = RGB(255, 247, 247)
    Else
        c.BorderColor = CBorderStrong()
        c.BackColor = CWhite()
    End If
End Sub

'--- Standarddialoge ------------------------------------------------------------------
Public Sub Alert(ByVal msg As String, Optional ByVal title As String = "Hinweis", Optional ByVal isError As Boolean = False)
    Dim f As frmMsg
    If gQuiet Then gQuietLog = gQuietLog & title & ": " & msg & vbLf: Exit Sub
    Set f = New frmMsg
    f.Setup title, msg, Array("OK"), IIf(isError, "error", "info")
    f.Show: mNav.FocusGrid
    Unload f
End Sub

Public Function Confirm(ByVal msg As String, Optional ByVal title As String = "Bitte bestätigen", Optional ByVal yesCaption As String = "OK", _
                        Optional ByVal noCaption As String = "Abbrechen", Optional ByVal danger As Boolean = False) As Boolean
    Dim f As frmMsg
    If gQuiet Then gQuietLog = gQuietLog & title & " (bestätigt): " & msg & vbLf: Confirm = True: Exit Function
    Set f = New frmMsg
    f.Setup title, msg, Array(yesCaption, noCaption), IIf(danger, "warn", "question"), danger
    f.Show: mNav.FocusGrid
    Confirm = (f.Result = 1)
    Unload f
End Function

' Drei Antworten; liefert 1..3, 0 bei Schliessen
Public Function Ask3(ByVal msg As String, ByVal title As String, ByVal b1 As String, ByVal b2 As String, ByVal b3 As String) As Long
    Dim f As frmMsg
    Set f = New frmMsg
    f.Setup title, msg, Array(b1, b2, b3), "question"
    f.Show: mNav.FocusGrid
    Ask3 = f.Result
    Unload f
End Function

' Wie Ask3, aber rechts unten, damit eine Vorschau dahinter sichtbar bleibt
Public Function Ask3Corner(ByVal msg As String, ByVal title As String, ByVal b1 As String, ByVal b2 As String, ByVal b3 As String) As Long
    Dim f As frmMsg
    If gQuiet Then gQuietLog = gQuietLog & title & " (Vorschau): " & msg & vbLf: Ask3Corner = 1: Exit Function
    Set f = New frmMsg
    f.Setup title, msg, Array(b1, b2, b3), "question"
    On Error Resume Next
    f.StartUpPosition = 0
    f.Left = Application.Left + Application.Width - f.Width - 24
    f.Top = Application.Top + Application.Height - f.Height - 40
    On Error GoTo 0
    f.Show: mNav.FocusGrid
    Ask3Corner = f.Result
    Unload f
End Function

Public Function AskText(ByVal msg As String, ByVal title As String, Optional ByVal dflt As String = "", Optional ByVal password As Boolean = False, _
                        Optional ByRef ok As Boolean) As String
    Dim f As frmMsg
    Set f = New frmMsg
    f.SetupInput title, msg, dflt, password
    f.Show: mNav.FocusGrid
    ok = (f.Result = 1)
    If ok Then AskText = f.InputValue
    Unload f
End Function

' Auswahl aus Liste; liefert 0-basierten Index oder -1
Public Function ChooseOne(ByVal msg As String, ByVal title As String, ByVal items As Variant, Optional ByVal okCaption As String = "Übernehmen", _
                          Optional ByVal preIndex As Long = -1) As Long
    Dim f As frmMsg
    Set f = New frmMsg
    f.SetupList title, msg, items, False, okCaption, 1, "", preIndex
    f.Show: mNav.FocusGrid
    If f.Result = 1 Then ChooseOne = f.SelectedIndex Else ChooseOne = -1
    Unload f
End Function

' Mehrfachauswahl; liefert ";"-getrennte Indizes, ok=False bei Abbruch
' Mit allowBack (Assistent): Schaltflaeche "Zurueck"; wentBack meldet sie (ok ist dann False)
Public Function ChooseMany(ByVal msg As String, ByVal title As String, ByVal items As Variant, ByVal preIdx As String, ByRef ok As Boolean, _
                           Optional ByVal allowBack As Boolean = False, Optional ByRef wentBack As Boolean, Optional ByVal okCaption As String = "Übernehmen") As String
    Dim f As frmMsg
    Set f = New frmMsg
    f.SetupList title, msg, items, True, okCaption, 1, "", -1, allowBack
    f.PreselectMany preIdx
    f.Show: mNav.FocusGrid
    ok = (f.Result = 1)
    wentBack = (allowBack And f.Result = 2)
    If ok Then ChooseMany = f.SelectedList
    Unload f
End Function

' Bericht/Vorschau mit Tabelle; liefert gedrueckten Button (1..n)
Public Function ShowReport(ByVal title As String, ByVal msg As String, ByVal items As Variant, ByVal cols As Long, ByVal colWidths As String, _
                           ByVal buttons As Variant, Optional ByVal danger As Boolean = False, Optional ByVal headers As Variant) As Long
    Dim f As frmMsg
    If gQuiet Then gQuietLog = gQuietLog & title & " (Bericht): " & msg & vbLf: ShowReport = 1: Exit Function
    Set f = New frmMsg
    If IsMissing(headers) Then
        f.SetupReport title, msg, items, cols, colWidths, buttons, danger
    Else
        f.SetupReport title, msg, items, cols, colWidths, buttons, danger, headers
    End If
    f.Show: mNav.FocusGrid
    ShowReport = f.Result
    Unload f
End Function
