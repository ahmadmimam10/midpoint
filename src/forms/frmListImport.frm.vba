'==============================================================================
' frmListImport - Assistent "Liste übernehmen": 1. Art und Datei, 2. Tabelle bzw.
' Befragung, 3. Rückfragen, 4. Ziel und Felder, 5. Prüfung und Übernahme.
' Die Quelldatei wird nur gelesen; übernommen wird erst nach Bestätigung.
'==============================================================================
Option Explicit

Public ResultTid As String
Public ResultMsg As String           ' Erfolgsmeldung, angezeigt nach dem Schließen (mCmd.Cmd_ListImport)
Private mEvts As Collection
Private mQEvts As Collection
Private mQSlots As Long
Private mStep As Long
Private mRead As Boolean
Private mTabIdx As Long
Private mFieldKeys() As String
Private mValRaw() As String
Private mLabelKey As String
Private mNameAuto As Boolean
Private mPlan As cMap
Private mValidRows As Collection
Private mErrs As Collection
Private mLoading As Boolean
Private mFH As Single
Private mValY As Single, mValW As Single   ' Lage des Wertebereichs in Schritt 4
Private mMultiVals As Boolean              ' Wertebereich zeigt die Häkchenliste (Mehrfachauswahl)
Private fr1 As MSForms.Frame, fr2 As MSForms.Frame, fr3 As MSForms.Frame, fr4 As MSForms.Frame, fr5 As MSForms.Frame
Private frQ As MSForms.Frame
Private lblStep As MSForms.Label

Public Sub Setup()
    Dim w As Single, b As MSForms.CommandButton, y As Single, top0 As Single
    Set mEvts = New Collection
    Set mQEvts = New Collection
    mUI.StyleForm Me
    Me.Width = 840
    If Me.Width > mUI.MaxFormWidth Then Me.Width = mUI.MaxFormWidth
    Me.caption = "Liste übernehmen"
    mFH = 440
    If mUI.MaxFormHeight - 150 < mFH Then mFH = Application.Max(340, mUI.MaxFormHeight - 150)
    mUI.AddLabel Me, "t", "Liste übernehmen", PAD, PAD - 2, 500, 20, True, 13
    Set lblStep = mUI.AddLabel(Me, "step", "", PAD, PAD + 20, Me.InsideWidth - PAD * 2, 15, False, 9, CAccent())
    w = Me.InsideWidth - PAD * 2
    top0 = PAD + 42
    BuildStep1 w, top0
    BuildStep2 w, top0
    BuildStep3 w, top0
    BuildStep4 w, top0
    BuildStep5 w, top0
    y = top0 + mFH + 10
    Set b = mUI.AddButton(Me, "back", "Zurück", PAD, y, 100): Bind b, "back"
    Set b = mUI.AddButton(Me, "next", "Weiter", PAD + 108, y, 160, "primary"): Bind b, "next"
    Set b = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 104, y, 104): Bind b, "cancel"
    b.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    mLoading = True
    fr1.Controls("optList").Value = True
    mLoading = False
    mListImport.LiReset
    ShowStep 1
End Sub

Private Sub Bind(ByVal c As Object, ByVal key As String)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, c, key
    mEvts.Add e
End Sub

'--- Aufbau der Schritte ----------------------------------------------------------
Private Sub BuildStep1(ByVal w As Single, ByVal top0 As Single)
    Dim c As Object, b As MSForms.CommandButton
    Set fr1 = mUI.AddFrame(Me, "fr1", PAD - 4, top0, w + 8, mFH)
    mUI.AddLabel fr1, "lmode", "Was möchten Sie übernehmen?", 4, 4, 400, 15, True
    Set c = mUI.AddOption(fr1, "optAsk", "Befragung aus 2ask (kodierter Excel-Export mit Codeplan)", 4, 24, 520, "modus"): Bind c, "modus"
    mUI.AddLabel fr1, "hAsk", "In 2ask: Ergebnisse, Ergebnisse herunterladen, „Kodierte Ergebnisse/Antworten“, Format Excel (XLSX) und Haken bei " & _
                 "„Codeplan hinzufügen“. Fragen, Antwortoptionen und Codes werden genau übernommen.", 24, 46, w - 30, 28, False, 8.5, CMuted()
    Set c = mUI.AddOption(fr1, "optList", "Sonstige Liste (Excel oder CSV)", 4, 80, 520, "modus"): Bind c, "modus"
    mUI.AddLabel fr1, "hList", "MidPoint sucht die Tabellen in der Datei, erkennt Kopfzeilen, Summen- und Gruppenzeilen und die Art jeder Spalte " & _
                 "und fragt nach, wo es nicht sicher ist.", 24, 102, w - 30, 28, False, 8.5, CMuted()
    mUI.AddLabel fr1, "lp", "Datei", 4, 142, 300, 15, True
    Set c = mUI.AddText(fr1, "path", 4, 159, w - 130): Bind c, "path"
    Set b = mUI.AddButton(fr1, "browse", "Durchsuchen…", w - 118, 158, 118, , ROW_H): Bind b, "browse"
    Set b = mUI.AddButton(fr1, "read", "Datei lesen", 4, 192, 120, "primary", ROW_H): Bind b, "read"
    mUI.AddLabel fr1, "info", "Die Datei wird nur gelesen (schreibgeschützt, Makros deaktiviert) und nicht verändert.", 134, 196, w - 140, 15, False, 8.5, CMuted()
    mUI.AddLabel fr1, "res", "", 4, 232, w, mFH - 238, False, 10, CText()
End Sub

Private Sub BuildStep2(ByVal w As Single, ByVal top0 As Single)
    Dim c As Object, b As MSForms.CommandButton, h As Single, y As Single
    Set fr2 = mUI.AddFrame(Me, "fr2", PAD - 4, top0, w + 8, mFH)
    ' Sonstige Liste: Tabellen und Vorschau
    mUI.AddLabel fr2, "lt", "Gefundene Tabellen: bitte die Tabelle wählen, die übernommen werden soll", 4, 4, w, 15, True
    Set c = mUI.AddList(fr2, "tabs", 4, 21, w, 100, False, 5, "120;80;60;55;" & (w - 330)): Bind c, "tabs"
    mUI.AddLabel fr2, "lv", "Vorschau: links steht, wie MidPoint die Zeile versteht (Kopf, Zeile, Gruppe, Summe, Titel, Notiz)", 4, 129, w, 15, True
    h = mFH - 146 - 36
    Set c = mUI.AddList(fr2, "prev", 4, 146, w, h, False, 9)
    y = 146 + h + 9
    mUI.AddLabel fr2, "lh", "Kopfzeile stimmt nicht? Kopf in Zeile", 4, y + 4, 200, 15
    Set c = mUI.AddText(fr2, "hdr", 206, y, 50)
    Set c = mUI.AddCombo(fr2, "hdrn", 264, y, 150)
    c.AddItem "über 1 Zeile": c.AddItem "über 2 Zeilen": c.ListIndex = 0
    Set b = mUI.AddButton(fr2, "rehdr", "Neu erkennen", 422, y - 1, 120, , ROW_H): Bind b, "rehdr"
    ' 2ask: erkannte Fragen
    mUI.AddLabel fr2, "asum", "", 4, 4, w, 48, False, 10, CText()
    mUI.AddLabel fr2, "laq", "Fragen der Befragung (Nummer, Frage, Frageart, entstehende Felder)", 4, 56, w, 15, True
    Set c = mUI.AddList(fr2, "askq", 4, 73, w, mFH - 79, False, 4, "50;" & (w - 280) & ";160;60")
End Sub

Private Sub BuildStep3(ByVal w As Single, ByVal top0 As Single)
    Set fr3 = mUI.AddFrame(Me, "fr3", PAD - 4, top0, w + 8, mFH)
    mUI.AddLabel fr3, "lq", "Vorgeschlagen ist jeweils die empfohlene Antwort. Jede Änderung wirkt sofort auf die Felder im nächsten Schritt.", _
                 4, 4, w, 30, False, 9.5, CText()
    Set frQ = mUI.AddFrame(fr3, "qs", 0, 34, w + 8, mFH - 36, "", True)
    mUI.AddLabel frQ, "none", "Keine Rückfragen: MidPoint hat die Liste eindeutig erkannt. Weiter mit „Weiter“.", 4, 4, w - 20, 20, False, 10, CMuted()
End Sub

Private Sub BuildStep4(ByVal w As Single, ByVal top0 As Single)
    Dim c As Object, b As MSForms.CommandButton, ids As Variant, i As Long, y As Single, hL As Single, code As Variant
    Set fr4 = mUI.AddFrame(Me, "fr4", PAD - 4, top0, w + 8, mFH)
    Set c = mUI.AddOption(fr4, "optNew", "Als neue Tabelle:", 4, 4, 130, "ziel"): Bind c, "ziel"
    Set c = mUI.AddText(fr4, "newname", 138, 4, 230): Bind c, "newname"
    Set c = mUI.AddOption(fr4, "optOld", "An bestehende Tabelle anhängen:", 386, 4, 200, "ziel"): Bind c, "ziel"
    Set c = mUI.AddCombo(fr4, "tbl", 590, 4, w - 590)
    ids = mCfg.TableIds()
    For i = 0 To ArrLen(ids) - 1
        c.AddItem mCfg.TableName(CStr(ids(i)))
    Next
    Bind c, "tbl"
    mUI.AddLabel fr4, "lk", "Schlüssel zum Aktualisieren:", 386, 35, 200, 15
    Set c = mUI.AddCombo(fr4, "key", 590, 31, w - 590)
    mUI.AddLabel fr4, "lf", "Felder: Zeile anklicken und unten anpassen (" & ChrW(&H2713) & " = wird übernommen, " & ChrW(&H2605) & " = Bezeichnungsfeld)", 4, 58, w, 15, True
    hL = mFH - 75 - 190
    Set c = mUI.AddList(fr4, "flds", 4, 75, w, hL, False, 5, "20;220;135;210;" & (w - 600)): Bind c, "flds"
    y = 75 + hL + 8
    Set c = mUI.AddCheck(fr4, "incl", "Übernehmen", 4, y + 17, 110): Bind c, "incl"
    mUI.AddLabel fr4, "lfn", "Feldname", 120, y, 200, 15, True
    Set c = mUI.AddText(fr4, "fname", 120, y + 17, 260): Bind c, "fname"
    mUI.AddLabel fr4, "lft", "Typ", 390, y, 150, 15, True
    Set c = mUI.AddCombo(fr4, "ftype", 390, y + 17, 190)
    For Each code In mListImport.LiTypeCodes()
        c.AddItem FieldTypeName(CStr(code))
    Next
    Bind c, "ftype"
    Set c = mUI.AddCheck(fr4, "islabel", "Bezeichnungsfeld", 592, y + 17, 160): Bind c, "islabel"
    y = y + 50
    mUI.AddLabel fr4, "lvals", "Werte dieses Felds (Wert in der Liste, Anzahl, wird übernommen als)", 4, y, 470, 15, True
    Set c = mUI.AddList(fr4, "vals", 4, y + 17, 470, mFH - y - 21, False, 3, "200;55;200"): Bind c, "vals"
    mUI.AddLabel fr4, "lvt", "Markierten Wert übernehmen als", 486, y, w - 486, 15, True
    Set c = mUI.AddText(fr4, "vtarget", 486, y + 17, w - 486 - 92)
    Set b = mUI.AddButton(fr4, "vapply", "Ändern", w - 84, y + 16, 84, , ROW_H): Bind b, "vapply"
    mUI.AddLabel fr4, "lvh", "Gleiche Bedeutung zusammenlegen: denselben Wert eintragen, z. B. „in Bearb.“ als „in Arbeit“. Leer: Wert wird nicht " & _
                 "übernommen.", 486, y + 46, w - 486, 64, False, 8.5, CMuted()
    ' Mehrfachauswahl: dem markierten Wert mehrere Werte per Häkchen zuordnen (sichtbar nur bei diesem Typ)
    Set c = mUI.AddList(fr4, "vopts", 486, y + 17, w - 486, mFH - y - 21 - ROW_H - 6, True)
    c.Visible = False
    mUI.AddLabel(fr4, "lvn", "Neuer Wert:", 486, mFH - 4 - ROW_H + 4, 64, 15).Visible = False
    mValY = y
    mValW = w
End Sub

' Bereich rechts neben der Werteliste: Textfeld (Auswählen) oder Häkchenliste (Mehrfachauswahl)
Private Sub LayoutValueArea(ByVal multi As Boolean)
    Dim y As Single, w As Single, ty As Single
    y = mValY: w = mValW
    mMultiVals = multi
    fr4.Controls("vopts").Visible = multi
    fr4.Controls("lvn").Visible = multi
    fr4.Controls("lvh").Visible = Not multi
    If multi Then
        ty = mFH - 4 - ROW_H
        fr4.Controls("lvt").caption = "Markiertem Wert zuordnen (mehrere Häkchen möglich)"
        fr4.Controls("vtarget").Left = 486 + 66: fr4.Controls("vtarget").Top = ty: fr4.Controls("vtarget").Width = w - 486 - 66 - 92
        fr4.Controls("vtarget").ControlTipText = "Weiteren Wert ergänzen, mehrere mit Semikolon trennen"
        fr4.Controls("vapply").Top = ty - 1
    Else
        fr4.Controls("lvt").caption = "Markierten Wert übernehmen als"
        fr4.Controls("vtarget").Left = 486: fr4.Controls("vtarget").Top = y + 17: fr4.Controls("vtarget").Width = w - 486 - 92
        fr4.Controls("vtarget").ControlTipText = ""
        fr4.Controls("vapply").Top = y + 16
    End If
End Sub

Private Sub BuildStep5(ByVal w As Single, ByVal top0 As Single)
    Dim c As Object, b As MSForms.CommandButton
    Set fr5 = mUI.AddFrame(Me, "fr5", PAD - 4, top0, w + 8, mFH)
    mUI.AddLabel fr5, "sum", "", 4, 4, w, 60, False, 10, CText()
    mUI.AddLabel fr5, "lr", "Bericht: fehlerhafte Einträge (werden nicht übernommen) und Hinweise zu einzelnen Werten", 4, 68, w, 15, True
    Set c = mUI.AddList(fr5, "errs", 4, 85, w, mFH - 85 - 38, False, 4, "170;170;160;" & (w - 500))
    Set b = mUI.AddButton(fr5, "saveerr", "Bericht speichern (CSV)…", 4, mFH - 30, 220, , ROW_H): Bind b, "saveerr"
End Sub

'--- Ablauf -------------------------------------------------------------------------
Private Sub ShowStep(ByVal s As Long)
    Dim isAsk As Boolean
    mStep = s
    isAsk = (mListImport.liMode = "2ask")
    fr1.Visible = (s = 1): fr2.Visible = (s = 2): fr3.Visible = (s = 3): fr4.Visible = (s = 4): fr5.Visible = (s = 5)
    Me.Controls("back").Enabled = (s > 1)
    Me.Controls("next").Enabled = True
    Me.Controls("next").caption = "Weiter"
    Select Case s
        Case 1
            lblStep.caption = "Schritt 1 von 5: Art der Liste und Datei wählen"
        Case 2
            ShowStep2Controls isAsk
            If isAsk Then
                lblStep.caption = "Schritt 2 von 5: Erkannte Fragen der Befragung prüfen"
                FillAsk
            Else
                lblStep.caption = "Schritt 2 von 5: Tabelle wählen und Vorschau prüfen"
                FillTables
            End If
        Case 3
            lblStep.caption = "Schritt 3 von 5: Rückfragen beantworten (" & Plural(mListImport.liQs.Count, "Rückfrage", "Rückfragen") & ")"
            RefreshQuestions
        Case 4
            lblStep.caption = "Schritt 4 von 5: Ziel und Felder festlegen"
            InitTarget
            FillFields 0
            Me.Controls("next").caption = "Prüfen"
        Case 5
            lblStep.caption = "Schritt 5 von 5: Prüfergebnis. Übernommen wird erst mit „Übernehmen“."
            Me.Controls("next").caption = "Übernehmen"
            Me.Controls("next").Enabled = (mValidRows.Count > 0)
    End Select
End Sub

Private Sub ShowStep2Controls(ByVal isAsk As Boolean)
    Dim nm As Variant
    For Each nm In Array("lt", "tabs", "lv", "prev", "lh", "hdr", "hdrn", "rehdr")
        fr2.Controls(CStr(nm)).Visible = Not isAsk
    Next
    For Each nm In Array("asum", "laq", "askq")
        fr2.Controls(CStr(nm)).Visible = isAsk
    Next
End Sub

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    On Error GoTo failed
    If mLoading Then Exit Sub
    Select Case key
        Case "modus": If evt = "Click" Then ModeChanged
        Case "path": If evt = "Change" Then mRead = False
        Case "browse": If evt = "Click" Then Browse
        Case "read": If evt = "Click" Then ReadFile
        Case "tabs"
            If evt = "Click" Or evt = "Change" Then FillPreview
            If evt = "DblClick" Then NextStep
        Case "rehdr": If evt = "Click" Then ReHeader
        Case "ziel", "tbl": If evt = "Click" Or evt = "Change" Then TargetChanged
        Case "newname": If evt = "Change" Then mNameAuto = False
        Case "flds": If evt = "Click" Or evt = "Change" Then ShowField
        Case "incl": If evt = "Click" Then FieldEdited "incl"
        Case "fname": If evt = "Change" Then FieldEdited "name"
        Case "ftype": If evt = "Change" Then FieldEdited "type"
        Case "islabel": If evt = "Click" Then FieldEdited "label"
        Case "vals": If evt = "Click" Or evt = "Change" Then ShowValue
        Case "vapply": If evt = "Click" Then ApplyValue
        Case "saveerr": If evt = "Click" Then mIO.ExportErrorReport
        Case "back": If evt = "Click" And mStep > 1 Then ShowStep mStep - 1
        Case "next": If evt = "Click" Then NextStep
        Case "cancel": If evt = "Click" Then Me.Hide
        Case Else
            If Left$(key, 2) = "q:" And evt = "Change" Then QuestionChanged CLng(Mid$(key, 3))
    End Select
    Exit Sub
failed:
    mLoading = False
    mApp.Boundary "Liste übernehmen"
End Sub

Private Sub NextStep()
    Select Case mStep
        Case 1
            If Not mRead Then
                If Not ReadFile() Then Exit Sub
            End If
            ShowStep 2
        Case 2
            If mListImport.liMode = "liste" Then
                If Not ChooseTable() Then Exit Sub
            End If
            ShowStep 3
        Case 3
            ShowStep 4
        Case 4
            If Not RunCheck() Then Exit Sub
            ShowStep 5
        Case 5
            DoImport
    End Select
End Sub

'--- Schritt 1: Datei -----------------------------------------------------------------
Private Sub ModeChanged()
    mRead = False
    mTabIdx = 0
    fr1.Controls("res").caption = ""
End Sub

Private Sub Browse()
    Dim res As Variant
    res = Application.GetOpenFilename("Excel- und CSV-Dateien (*.xlsx;*.xlsm;*.xls;*.csv;*.txt),*.xlsx;*.xlsm;*.xls;*.csv;*.txt", , "Liste wählen")
    If VarType(res) <> vbString Then Exit Sub
    mLoading = True
    fr1.Controls("path").Text = res
    mLoading = False
    mRead = False
    ReadFile
End Sub

Private Function ReadFile() As Boolean
    Dim p As String, mode As String, res As String
    p = Trim$(fr1.Controls("path").Text)
    If Len(p) = 0 Then mUI.Alert "Bitte zuerst eine Datei wählen (Durchsuchen).", "Liste übernehmen": Exit Function
    If Not FileExists(p) Then mUI.Alert "Die Datei wurde nicht gefunden:" & vbLf & p, "Liste übernehmen": Exit Function
    mode = IIf(fr1.Controls("optAsk").Value, "2ask", "liste")
    mRead = False
    mTabIdx = 0
    fr1.Controls("res").caption = ""
    mApp.BeginWork "Datei wird gelesen …"
    mListImport.LiReadFile p, mode
    mApp.EndWork
    If mode = "liste" And mListImport.liLooksAsk Then
        If mUI.Confirm("Die Datei sieht aus wie ein kodierter 2ask-Export mit Codeplan. Als Befragung übernehmen? Dann werden Fragen, " & _
                       "Antwortoptionen und Codes genau übernommen.", "Liste übernehmen", "Als Befragung übernehmen", "Als Liste lesen") Then
            mLoading = True
            fr1.Controls("optAsk").Value = True
            mLoading = False
            mApp.BeginWork "Datei wird gelesen …"
            mListImport.LiReadFile p, "2ask"
            mApp.EndWork
        End If
    End If
    If mListImport.liMode = "2ask" Then
        mApp.BeginWork
        mListImport.LiPlan
        mApp.EndWork
        mLabelKey = mListImport.LiSuggestLabel()
        mNameAuto = True
        res = mListImport.LiAskSummary()
        mRead = True
    ElseIf mListImport.liTables.Count = 0 Then
        res = "In der Datei wurde keine Tabelle mit Kopfzeile und Datenzeilen gefunden."
    Else
        res = Plural(mListImport.liTables.Count, "Tabelle", "Tabellen") & " gefunden. Im nächsten Schritt wählen Sie die Tabelle, die übernommen werden soll."
        mRead = True
    End If
    fr1.Controls("res").caption = res
    ReadFile = mRead
End Function

'--- Schritt 2: Tabelle bzw. Befragung ------------------------------------------------
Private Sub FillTables()
    Dim lb As MSForms.ListBox, out() As Variant, i As Long, a As Variant, j As Long, n As Long, sel As Long
    Set lb = fr2.Controls("tabs")
    n = mListImport.liTables.Count
    If n = 0 Then lb.Clear: Exit Sub
    ReDim out(0 To n - 1, 0 To 4)
    For i = 1 To n
        a = mListImport.LiTableRow(i)
        For j = 0 To 4: out(i - 1, j) = a(j): Next
    Next
    sel = mTabIdx
    If sel < 1 Or sel > n Then sel = mListImport.LiBestTable()
    mLoading = True
    lb.List = out
    lb.ListIndex = sel - 1
    mLoading = False
    FillPreview
End Sub

Private Sub FillPreview()
    Dim i As Long, arr As Variant, lb As MSForms.ListBox, nc As Long, cw As String, k As Long, t As cLiTab, wd As Single
    i = fr2.Controls("tabs").ListIndex + 1
    Set lb = fr2.Controls("prev")
    If i < 1 Then lb.Clear: Exit Sub
    nc = mListImport.LiPreviewColCount(i, 8) + 1
    wd = (lb.Width - 70) / Application.Max(1, nc - 1)
    If wd < 60 Then wd = 60
    ' Texte passend zur Spaltenbreite mit „…“ kuerzen (etwa 5 pt je Zeichen), statt sie hart abzuschneiden
    arr = mListImport.LiPreview(i, 40, 8, Application.Max(12, Int(wd / 5)))
    nc = UBound(arr, 2) + 1
    cw = "62"
    For k = 2 To nc: cw = cw & ";" & Format$(wd, "0"): Next
    lb.Clear
    lb.ColumnCount = nc
    lb.ColumnWidths = cw
    lb.List = arr
    Set t = mListImport.liTables(i)
    mLoading = True
    If t.Kind = "liste" Or t.Kind = "raster" Then
        fr2.Controls("hdr").Text = CStr(t.HdrRow)
        fr2.Controls("hdrn").ListIndex = IIf(t.HdrRows = 2, 1, 0)
    Else
        fr2.Controls("hdr").Text = ""
    End If
    fr2.Controls("hdr").Enabled = (t.Kind = "liste")
    fr2.Controls("hdrn").Enabled = (t.Kind = "liste")
    fr2.Controls("rehdr").Enabled = (t.Kind = "liste")
    mLoading = False
End Sub

Private Sub ReHeader()
    Dim i As Long, r As Long
    i = fr2.Controls("tabs").ListIndex + 1
    If i < 1 Then Exit Sub
    r = CLng(Val(fr2.Controls("hdr").Text))
    If r < 1 Then mUI.Alert "Bitte die Zeilennummer der Kopfzeile eintragen.", "Liste übernehmen": Exit Sub
    mListImport.LiSetHeader i, r, fr2.Controls("hdrn").ListIndex + 1
    mTabIdx = 0
    FillTables
    fr2.Controls("tabs").ListIndex = i - 1
End Sub

Private Function ChooseTable() As Boolean
    Dim i As Long
    i = fr2.Controls("tabs").ListIndex + 1
    If i < 1 Then mUI.Alert "Bitte eine Tabelle wählen.", "Liste übernehmen": Exit Function
    If i <> mTabIdx Then
        mApp.BeginWork "Liste wird ausgewertet …"
        mListImport.LiChooseTable i
        mApp.EndWork
        mTabIdx = i
        mLabelKey = mListImport.LiSuggestLabel()
        mNameAuto = True
    End If
    ChooseTable = True
End Function

Private Sub FillAsk()
    Dim lb As MSForms.ListBox, out() As Variant, a As Variant, i As Long, n As Long
    fr2.Controls("asum").caption = mListImport.LiAskSummary()
    Set lb = fr2.Controls("askq")
    n = mListImport.liAskQuestions.Count
    If n = 0 Then lb.Clear: Exit Sub
    ReDim out(0 To n - 1, 0 To 3)
    i = 0
    For Each a In mListImport.liAskQuestions
        out(i, 0) = a(0): out(i, 1) = Truncate(CStr(a(1)), 90): out(i, 2) = a(2): out(i, 3) = CStr(a(3))
        i = i + 1
    Next
    lb.List = out
End Sub

'--- Schritt 3: Rückfragen --------------------------------------------------------------
' Die Steuerelemente werden wiederverwendet (nicht im laufenden Ereignis zerstört).
Private Sub AddQSlot(ByVal i As Long, ByVal w As Single)
    Dim cb As MSForms.ComboBox, e As cEvt, lb As MSForms.Label
    mUI.AddLabel frQ, "qt" & i, "", 4, 0, w, 15, True, 9.5, CAccent()
    Set lb = mUI.AddLabel(frQ, "qx" & i, "", 4, 0, w, 15, False, 9.5, CText())
    Set cb = mUI.AddCombo(frQ, "qc" & i, 4, 0, w)
    cb.ListRows = 8
    Set e = New cEvt
    e.Bind Me, cb, "q:" & i
    mQEvts.Add e
    mQSlots = i
End Sub

Private Sub RefreshQuestions()
    Dim q As cLiQ, y As Single, i As Long, w As Single, lt As MSForms.Label, lx As MSForms.Label, cb As MSForms.ComboBox
    Dim k As Long, same As Boolean
    mLoading = True
    w = frQ.Width - 30
    y = 4
    frQ.Controls("none").Visible = (mListImport.liQs.Count = 0)
    For Each q In mListImport.liQs
        i = i + 1
        If i > mQSlots Then AddQSlot i, w
        Set lt = frQ.Controls("qt" & i)
        Set lx = frQ.Controls("qx" & i)
        Set cb = frQ.Controls("qc" & i)
        lt.caption = q.Topic
        lt.Top = y: lt.Visible = True
        y = y + 17
        lx.AutoSize = False
        lx.Width = w
        lx.caption = q.Text
        lx.AutoSize = True
        lx.Top = y: lx.Visible = True
        y = y + lx.Height + 4
        same = (CStr(cb.Tag) = q.ID And cb.ListCount = UBound(q.Choices) + 1)
        If same Then
            For k = 0 To UBound(q.Choices)
                If CStr(cb.List(k)) <> CStr(q.Choices(k)) Then same = False: Exit For
            Next
        End If
        If Not same Then
            cb.Clear
            For k = 0 To UBound(q.Choices)
                cb.AddItem CStr(q.Choices(k))
            Next
            cb.Tag = q.ID
        End If
        If cb.ListIndex <> q.Answer Then cb.ListIndex = q.Answer
        cb.Top = y: cb.Visible = True
        y = y + ROW_H + 18
    Next
    For k = i + 1 To mQSlots
        frQ.Controls("qt" & k).Visible = False
        frQ.Controls("qx" & k).Visible = False
        frQ.Controls("qc" & k).Visible = False
    Next
    frQ.ScrollHeight = y + 4
    If frQ.ScrollTop > Application.Max(0, y + 4 - frQ.InsideHeight) Then frQ.ScrollTop = Application.Max(0, y + 4 - frQ.InsideHeight)
    lblStep.caption = "Schritt 3 von 5: Rückfragen beantworten (" & Plural(mListImport.liQs.Count, "Rückfrage", "Rückfragen") & ")"
    mLoading = False
End Sub

Private Sub QuestionChanged(ByVal idx As Long)
    Dim cb As MSForms.ComboBox
    Set cb = frQ.Controls("qc" & idx)
    If cb.ListIndex < 0 Or Len(CStr(cb.Tag)) = 0 Then Exit Sub
    mListImport.LiSetAnswer CStr(cb.Tag), cb.ListIndex
    mApp.BeginWork
    mListImport.LiPlan
    mApp.EndWork
    RefreshQuestions
End Sub

'--- Schritt 4: Ziel und Felder -----------------------------------------------------------
Private Function TargetTid() As String
    If fr4.Controls("optOld").Value Then TargetTid = mCfg.TableIdByName(fr4.Controls("tbl").Text)
End Function

Private Sub InitTarget()
    mLoading = True
    If mNameAuto Or Len(Trim$(fr4.Controls("newname").Text)) = 0 Then
        fr4.Controls("newname").Text = mListImport.LiSuggestTableName()
        mNameAuto = True
    End If
    If Not fr4.Controls("optOld").Value Then fr4.Controls("optNew").Value = True
    mLoading = False
    FillKeyCombo
End Sub

Private Sub TargetChanged()
    Dim sel As Long
    mLoading = True
    If fr4.Controls("optNew").Value = False And fr4.Controls("optOld").Value = False Then fr4.Controls("optNew").Value = True
    mLoading = False
    FillKeyCombo
    sel = fr4.Controls("flds").ListIndex
    FillFields IIf(sel < 0, 0, sel)
End Sub

Private Sub FillKeyCombo()
    Dim kc As MSForms.ComboBox, tid As String, col As cLiCol
    Set kc = fr4.Controls("key")
    tid = TargetTid()
    mLoading = True
    kc.Clear
    kc.AddItem "(keiner: nur neue Einträge anlegen)"
    If Len(tid) > 0 Then
        For Each col In mListImport.liCols
            If col.Include Then
                If Len(mCfg.FieldIdByName(tid, col.Name)) > 0 Then kc.AddItem col.Name
            End If
        Next
    End If
    kc.ListIndex = 0
    kc.Enabled = (Len(tid) > 0)
    fr4.Controls("lk").ForeColor = IIf(Len(tid) > 0, CText(), CMuted())
    fr4.Controls("newname").Enabled = (Len(tid) = 0)
    mLoading = False
End Sub

Private Sub FillFields(ByVal sel As Long)
    Dim lb As MSForms.ListBox, col As cLiCol, out() As Variant, i As Long, n As Long, tid As String, a As Variant, j As Long
    Set lb = fr4.Controls("flds")
    n = mListImport.liCols.Count
    mLoading = True
    If n = 0 Then
        lb.Clear
        mLoading = False
        Exit Sub
    End If
    ReDim out(0 To n - 1, 0 To 4)
    ReDim mFieldKeys(0 To n - 1)
    tid = TargetTid()
    i = 0
    For Each col In mListImport.liCols
        mFieldKeys(i) = col.Key
        a = FieldRow(col, tid)
        For j = 0 To 4: out(i, j) = a(j): Next
        i = i + 1
    Next
    lb.List = out
    If sel < 0 Then sel = 0
    If sel > n - 1 Then sel = n - 1
    lb.ListIndex = sel
    mLoading = False
    ShowField
End Sub

Private Function FieldRow(ByVal col As cLiCol, ByVal tid As String) As Variant
    Dim h As String, nm As String
    h = col.Hint
    If Len(tid) > 0 And col.Include Then
        If Len(mCfg.FieldIdByName(tid, col.Name)) > 0 Then h = "in vorhandenes Feld" & IIf(Len(h) > 0, "; " & h, "") Else h = "neues Feld" & IIf(Len(h) > 0, "; " & h, "")
    End If
    If col.Unsure Then h = "bitte prüfen" & IIf(Len(h) > 0, "; " & h, "")
    nm = col.Name
    If col.Key = mLabelKey And Len(tid) = 0 And col.Include Then nm = nm & "  " & ChrW(&H2605)
    FieldRow = Array(IIf(col.Include, ChrW(&H2713), ""), nm, FieldTypeName(col.FType), col.Sample, h)
End Function

Private Sub UpdateFieldRows()
    Dim lb As MSForms.ListBox, i As Long, col As cLiCol, a As Variant, j As Long, tid As String
    Set lb = fr4.Controls("flds")
    tid = TargetTid()
    For i = 0 To lb.ListCount - 1
        Set col = mListImport.LiFieldByKey(mFieldKeys(i))
        If Not col Is Nothing Then
            a = FieldRow(col, tid)
            For j = 0 To 4
                If CStr(lb.List(i, j)) <> CStr(a(j)) Then lb.List(i, j) = a(j)
            Next
        End If
    Next
End Sub

Private Sub ShowField()
    Dim i As Long, col As cLiCol, vals As Collection, out() As Variant, j As Long, a As Variant, editable As Boolean, cb As MSForms.ComboBox
    Dim lv As MSForms.ListBox
    i = fr4.Controls("flds").ListIndex
    Set lv = fr4.Controls("vals")
    mLoading = True
    If i < 0 Or i > UBound(mFieldKeys) Then mLoading = False: Exit Sub
    Set col = mListImport.LiFieldByKey(mFieldKeys(i))
    If col Is Nothing Then mLoading = False: Exit Sub
    fr4.Controls("incl").Value = col.Include
    fr4.Controls("fname").Text = col.Name
    Set cb = fr4.Controls("ftype")
    cb.ListIndex = -1
    For j = 0 To cb.ListCount - 1
        If cb.List(j) = FieldTypeName(col.FType) Then cb.ListIndex = j: Exit For
    Next
    fr4.Controls("islabel").Value = (col.Key = mLabelKey)
    fr4.Controls("islabel").Enabled = (Len(TargetTid()) = 0 And col.CanBeLabel)
    ' Werte eines Auswahlfelds
    Set vals = mListImport.LiFieldValues(col.Key)
    editable = mListImport.LiValuesEditable(col.Key)
    lv.Clear
    fr4.Controls("vtarget").Text = ""
    If vals.Count > 0 Then
        ReDim out(0 To vals.Count - 1, 0 To 2)
        ReDim mValRaw(0 To vals.Count - 1)
        j = 0
        For Each a In vals
            out(j, 0) = a(0): out(j, 1) = a(1)
            out(j, 2) = IIf(Len(a(2)) = 0, "(nicht übernehmen)", Replace(CStr(a(2)), ";", "; "))
            mValRaw(j) = a(0)
            j = j + 1
        Next
        lv.List = out
        fr4.Controls("lvals").caption = "Werte dieses Felds (Wert in der Liste, Anzahl, wird übernommen als)"
    Else
        ReDim mValRaw(0 To 0)
        fr4.Controls("lvals").caption = "Werte gibt es nur bei „Auswählen“ und „Mehrfachauswahl“"
    End If
    ' Mehrfachauswahl: alle Werte des Felds als Häkchenliste
    LayoutValueArea (col.FType = FT_MEHRFACH And editable And vals.Count > 0)
    If mMultiVals Then
        With fr4.Controls("vopts")
            .Clear
            For Each a In col.Opts
                .AddItem CStr(a)
            Next
        End With
    End If
    fr4.Controls("vtarget").Enabled = editable And vals.Count > 0
    fr4.Controls("vapply").Enabled = editable And vals.Count > 0
    fr4.Controls("lvt").ForeColor = IIf(editable And vals.Count > 0, CText(), CMuted())
    mLoading = False
End Sub

Private Sub FieldEdited(ByVal part As String)
    Dim i As Long, key As String, col As cLiCol, code As String
    i = fr4.Controls("flds").ListIndex
    If i < 0 Then Exit Sub
    key = mFieldKeys(i)
    Set col = mListImport.LiFieldByKey(key)
    If col Is Nothing Then Exit Sub
    Select Case part
        Case "incl"
            mListImport.LiSetFieldPart key, "incl", CBool(fr4.Controls("incl").Value)
            col.Include = CBool(fr4.Controls("incl").Value)
            UpdateFieldRows
            FillKeyCombo
        Case "name"
            mListImport.LiSetFieldPart key, "name", fr4.Controls("fname").Text
            col.Name = mListImport.LiCleanName(fr4.Controls("fname").Text)
            UpdateFieldRows
        Case "type"
            code = FieldTypeFromName(fr4.Controls("ftype").Text)
            If Len(code) = 0 Or code = col.FType Then Exit Sub
            mListImport.LiSetFieldPart key, "type", code
            mApp.BeginWork
            mListImport.LiPlan
            mApp.EndWork
            FillFields i
        Case "label"
            If fr4.Controls("islabel").Value Then
                mLabelKey = key
            ElseIf mLabelKey = key Then
                mLabelKey = ""
            End If
            UpdateFieldRows
    End Select
End Sub

Private Sub ShowValue()
    Dim j As Long, lv As MSForms.ListBox, cur As String, k As Long, vo As MSForms.ListBox
    Set lv = fr4.Controls("vals")
    j = lv.ListIndex
    If j < 0 Then Exit Sub
    mLoading = True
    If mMultiVals Then
        ' Häkchen bei allen Werten, denen der markierte Wert zugeordnet ist
        Set vo = fr4.Controls("vopts")
        cur = ";" & Replace(CStr(lv.List(j, 2)), "; ", ";") & ";"
        For k = 0 To vo.ListCount - 1
            vo.Selected(k) = (InStr(1, cur, ";" & vo.List(k) & ";", vbTextCompare) > 0)
        Next
        fr4.Controls("vtarget").Text = ""
    Else
        fr4.Controls("vtarget").Text = IIf(CStr(lv.List(j, 2)) = "(nicht übernehmen)", "", CStr(lv.List(j, 2)))
    End If
    mLoading = False
End Sub

Private Sub ApplyValue()
    Dim i As Long, j As Long, key As String, t As String, k As Long, extra As Variant, vo As MSForms.ListBox
    i = fr4.Controls("flds").ListIndex
    j = fr4.Controls("vals").ListIndex
    If i < 0 Or j < 0 Then mUI.Alert "Bitte zuerst einen Wert in der Liste markieren.", "Liste übernehmen": Exit Sub
    key = mFieldKeys(i)
    If Not mListImport.LiValuesEditable(key) Then Exit Sub
    If mMultiVals Then
        ' angehakte Werte und neu eingetragene (Semikolon) zusammen; keines: Wert wird nicht übernommen
        Set vo = fr4.Controls("vopts")
        For k = 0 To vo.ListCount - 1
            If vo.Selected(k) Then t = t & IIf(Len(t) > 0, ";", "") & vo.List(k)
        Next
        For Each extra In Split(Replace(fr4.Controls("vtarget").Text, ",", ";"), ";")
            If Len(Trim$(extra)) > 0 Then t = t & IIf(Len(t) > 0, ";", "") & Trim$(extra)
        Next
    Else
        t = Trim$(fr4.Controls("vtarget").Text)
    End If
    mListImport.LiSetValueMap key, mValRaw(j), t
    mApp.BeginWork
    mListImport.LiPlan
    mApp.EndWork
    FillFields i
    If j < fr4.Controls("vals").ListCount Then fr4.Controls("vals").ListIndex = j
End Sub

'--- Schritt 5: Prüfung und Übernahme -------------------------------------------------------
Private Function RunCheck() As Boolean
    Dim tid As String, nm As String, updates As Long, e As Variant, out() As Variant, i As Long, rep As New Collection, msg As String
    Dim keyKey As String, col As cLiCol, seen As New cMap, nIncl As Long, total As Long
    tid = TargetTid()
    If fr4.Controls("optOld").Value And Len(tid) = 0 Then mUI.Alert "Bitte eine bestehende Tabelle wählen.", "Liste übernehmen": Exit Function
    nm = Trim$(fr4.Controls("newname").Text)
    If Len(tid) = 0 Then
        If Len(nm) = 0 Then mUI.Alert "Bitte einen Namen für die neue Tabelle angeben.", "Liste übernehmen": Exit Function
        If Len(mCfg.TableIdByName(nm)) > 0 Then
            mUI.Alert "Eine Tabelle „" & nm & "“ gibt es bereits. Bitte einen anderen Namen wählen oder „An bestehende Tabelle anhängen“.", "Liste übernehmen"
            Exit Function
        End If
    End If
    For Each col In mListImport.liCols
        If col.Include Then
            nIncl = nIncl + 1
            If Len(Trim$(col.Name)) = 0 Then mUI.Alert "Bitte für jedes übernommene Feld einen Namen angeben.", "Liste übernehmen": Exit Function
            If seen.Exists(col.Name) Then
                mUI.Alert "Der Feldname „" & col.Name & "“ kommt mehrfach vor. Bitte eines der Felder umbenennen.", "Liste übernehmen"
                Exit Function
            End If
            seen.SetVal col.Name, True
            If fr4.Controls("key").ListIndex > 0 And col.Name = fr4.Controls("key").Text Then keyKey = col.Key
        End If
    Next
    If nIncl = 0 Then mUI.Alert "Bitte mindestens ein Feld zum Übernehmen auswählen.", "Liste übernehmen": Exit Function
    mApp.BeginWork "Einträge werden geprüft …"
    Set mPlan = mListImport.LiBuildPlan(nm, tid, keyKey, IIf(Len(tid) = 0, mLabelKey, ""))
    Set mErrs = mIO.CheckImport(mPlan, mValidRows, updates)
    mApp.EndWork
    total = UBound(mPlan.GetVal("data"), 1) - 1
    For Each e In mErrs
        rep.Add Array(mListImport.LiSourceRow(CLng(e(0))), e(1), e(2), "nicht übernommen: " & e(3))
    Next
    For Each e In mListImport.liNotes
        rep.Add Array(e(0), e(1), e(2), e(3))
    Next
    mIO.StoreErrorReport rep, mListImport.liPath
    msg = FmtInt(mValidRows.Count) & " von " & FmtInt(total) & " Einträgen werden übernommen" & IIf(updates > 0, " (davon " & updates & " Aktualisierungen)", "") & _
          IIf(mErrs.Count > 0, ", " & FmtInt(mErrs.Count) & " sind fehlerhaft und werden ausgelassen", "") & "." & vbLf
    If mListImport.liNotes.Count > 0 Then msg = msg & Plural(mListImport.liNotes.Count, "Hinweis", "Hinweise") & " zu einzelnen Werten; siehe Bericht." & vbLf
    msg = msg & IIf(Len(tid) = 0, "Die neue Tabelle „" & nm & "“ wird erst beim Übernehmen angelegt. ", "") & "Vorher wird automatisch eine Sicherung erstellt."
    fr5.Controls("sum").caption = msg
    fr5.Controls("errs").Clear
    If rep.Count > 0 Then
        ReDim out(0 To Application.Min(rep.Count, 500) - 1, 0 To 3)
        i = 0
        For Each e In rep
            If i >= 500 Then Exit For
            out(i, 0) = e(0): out(i, 1) = e(1): out(i, 2) = Truncate(CStr(e(2)), 40): out(i, 3) = e(3)
            i = i + 1
        Next
        fr5.Controls("errs").List = out
    End If
    fr5.Controls("saveerr").Enabled = (rep.Count > 0)
    RunCheck = True
End Function

Private Sub DoImport()
    Dim tname As String
    If mValidRows Is Nothing Then Exit Sub
    If mValidRows.Count = 0 Then Exit Sub
    If Not mApp.RequireWritable("Die Übernahme") Then Exit Sub
    If Len(TargetTid()) = 0 Then tname = Trim$(fr4.Controls("newname").Text) Else tname = fr4.Controls("tbl").Text
    If Not mUI.Confirm(FmtInt(mValidRows.Count) & " Einträge in " & IIf(Len(TargetTid()) = 0, "die neue Tabelle", "die Tabelle") & " „" & tname & "“ übernehmen?" & vbLf & _
                       "Vorher wird automatisch eine Sicherung erstellt.", "Liste übernehmen", "Übernehmen", "Zurück") Then Exit Sub
    ResultTid = mIO.ExecuteImport(mPlan, mValidRows)
    If Len(ResultTid) > 0 Then
        ' die Meldung zeigt mCmd.Cmd_ListImport nach dem Schließen und dem Öffnen der neuen Tabelle
        ResultMsg = "Übernommen: " & FmtInt(mValidRows.Count) & " Einträge in „" & mCfg.TableName(ResultTid) & "“." & _
                    IIf(fr5.Controls("errs").ListCount > 0, " Der Bericht mit allen Hinweisen bleibt unter Einrichten > Letzter Importbericht verfügbar.", "")
        Me.Hide
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True: Me.Hide
End Sub

Private Sub UserForm_Terminate()
    Dim e As cEvt
    If Not mEvts Is Nothing Then
        For Each e In mEvts: e.Unbind: Next
    End If
    If Not mQEvts Is Nothing Then
        For Each e In mQEvts: e.Unbind: Next
    End If
End Sub
