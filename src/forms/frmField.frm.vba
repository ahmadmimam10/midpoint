'==============================================================================
' frmField - Feld anlegen oder bearbeiten: Name, Typ, Hilfetext, Pflicht,
' Sichtbarkeit, Eindeutigkeit, Dublettenpruefung, Register, Standardwert und
' typabhaengige Einstellungen (Laenge, Nachkommastellen, Grenzen, Optionen mit
' Farben, Verknuepfungsziel, Berechnung, Datumsfolge).
'==============================================================================
Option Explicit

Public ResultFid As String
Private mTid As String
Private mFid As String
Private mType As String
Private mKind As String
Private mEvts As Collection
Private mFormula As String
Private mFormulaRT As String
Private mFormulaDec As Long
Private fr As MSForms.Frame
Private tp As MSForms.Frame
Private txtName As MSForms.TextBox
Private lblSelType As MSForms.Label
Private lblTypeHelp As MSForms.Label
Private txtHelp As MSForms.TextBox
Private chkReq As MSForms.CheckBox
Private chkVis As MSForms.CheckBox
Private chkUniq As MSForms.CheckBox
Private chkDup As MSForms.CheckBox
Private cboGroup As MSForms.ComboBox
Private txtDefault As MSForms.TextBox
Private lblDefHint As MSForms.Label
Private lblErr As MSForms.Label
Private lstOpt As MSForms.ListBox
Private mOptIds() As String
Private mLoading As Boolean
Private mTpH As Single
Private advF As MSForms.Frame
Private mAdvTop As Single
Private mAdvH As Single
Private mXl As cXlEditor            ' Excel-Formel (Formelspalten)
Private mXlMode As Boolean
Private mRtAuto As Boolean          ' Ergebnistyp noch automatisch (neue Formelspalte)
Private mRollInline As Boolean      ' Rollup direkt im Dialog (ohne Bedingung)
Private mAdvBase As Single          ' Höhe "Weitere Einstellungen" ohne Zusatzzeile
Private txtMaxPick As MSForms.TextBox
Private mRuLinks() As String        ' Verknuepfungsfelder in der Reihenfolge der Auswahl
Private mRuFlds() As String         ' Felder der verknuepften Tabelle (Index 0 = keins)

Private Const RU_AGGS As String = "ANZAHL;SUMME;MITTEL;MIN;MAX;GEFUELLT"
Private Const RU_NAMES As String = "Anzahl der verknüpften Einträge;Summe;Durchschnitt;Kleinster Wert;Größter Wert;Anzahl mit ausgefülltem Feld"

Public Sub Setup(ByVal tid As String, ByVal fid As String)
    Dim f As cField, y As Single, codes As Variant, i As Long, b As MSForms.CommandButton, w As Single, g As cField, groups As New cMap
    mTid = tid: mFid = fid
    Set mEvts = New Collection
    mUI.StyleForm Me
    Me.Width = 660
    If Me.Width > mUI.MaxFormWidth Then Me.Width = mUI.MaxFormWidth
    If Len(fid) > 0 Then Set f = mCfg.Field(fid)
    Me.caption = IIf(f Is Nothing, "Neue Eigenschaft", "Eigenschaft bearbeiten") & " (" & mCfg.TableName(tid) & ")"
    If f Is Nothing Then
        mUI.AddLabel Me, "t", "Neue Eigenschaft in „" & mCfg.TableName(tid) & "“", PAD, PAD - 2, 560, 20, True, 13
    Else
        mUI.AddLabel Me, "t", "Eigenschaft „" & f.Name & "“", PAD, PAD - 2, 560, 20, True, 13
    End If
    Set fr = mUI.AddFrame(Me, "fr", PAD - 4, PAD + 26, Me.InsideWidth - PAD * 2 + 8, 300, "", True)
    w = fr.Width - 28
    mLoading = True
    y = 6
    mUI.AddLabel fr, "ln", "Name der Eigenschaft *", 4, y, 280, 15, True
    mUI.AddLabel fr, "lt", "Art der Eigenschaft", 300, y, 280, 15, True
    Set txtName = mUI.AddText(fr, "name", 4, y + 17, 280)
    If f Is Nothing Then
        mKind = FT_TEXT: mType = FT_TEXT
        Set lblSelType = mUI.AddLabel(fr, "typ", "", 300, y + 21, w - 296, 15, True, 9.5, CAccent())
        y = y + 46
        mUI.AddLabel fr, "ltg", "Art wählen wie in Notion: ein Klick genügt", 4, y, w, 15, False, 8.5, CMuted()
        y = y + 16
        y = BuildTypeGrid(y, w)
        Set lblTypeHelp = mUI.AddLabel(fr, "tgh", "", 4, y + 2, w, 15, False, 8.5, CMuted())
        MarkType
        y = y + 22
    Else
        txtName.Text = f.Name
        mType = f.FType: mKind = f.Kind
        mUI.AddIcon fr, "typi", FieldTypeIcon(mKind), 300, y + 19, 12, CAccent()
        mUI.AddLabel fr, "typ", f.TypeName_ & IIf(f.FType = FT_LINK, " " & ChrW(&H2192) & " " & mCfg.TableName(f.TargetTable), ""), 318, y + 20, 200, 15, True
        If mKind = mType And (mStruct.CanConvert(f.FType, FT_TEXT) Or mStruct.CanConvert(f.FType, FT_ZAHL)) Then
            Set b = mUI.AddButton(fr, "chtype", "Typ ändern…", w - 110, y + 16, 114, , ROW_H)
            Bind b, "chtype"
        End If
        mFormula = f.Formula: mFormulaRT = f.ResultType: mFormulaDec = f.Decimals
        y = y + 46
    End If
    ' Haeufiges sichtbar, Seltenes unter "Weitere Einstellungen" (erst auf Klick)
    Set chkReq = mUI.AddCheck(fr, "req", "Pflichtangabe", 4, y, 130)
    Set b = mUI.AddButton(fr, "adv", "Weitere Einstellungen  " & ChrW(&H25B8), w - 186, y - 2, 190, , ROW_H)
    Bind b, "adv"
    y = y + ROW_H + 6
    Set advF = mUI.AddFrame(fr, "advF", 0, y, fr.Width - 20, 150)
    mAdvTop = y
    mUI.AddLabel advF, "lh", "Hilfetext (erscheint in der Eingabemaske)", 4, 0, 400, 15, True
    Set txtHelp = mUI.AddText(advF, "help", 4, 17, w)
    Set chkVis = mUI.AddCheck(advF, "vis", "In der Maske sichtbar", 4, 46, 170)
    Set chkUniq = mUI.AddCheck(advF, "uniq", "Eindeutiger Wert", 180, 46, 140)
    Set chkDup = mUI.AddCheck(advF, "dup", "Dublettenprüfung", 330, 46, 150)
    chkVis.Value = True
    mUI.AddLabel advF, "lg", "Register in der Maske", 4, 46 + ROW_H + 8, 280, 15, True
    mUI.AddLabel advF, "ldf", "Standardwert für neue Einträge", 300, 46 + ROW_H + 8, 300, 15, True
    Set cboGroup = mUI.AddCombo(advF, "group", 4, 46 + ROW_H + 25, 280, False)
    For Each g In mCfg.Fields(tid, False)
        If Len(g.Group) > 0 And Not groups.Exists(g.Group) Then groups.SetVal g.Group, True: cboGroup.AddItem g.Group
    Next
    Set txtDefault = mUI.AddText(advF, "def", 300, 46 + ROW_H + 25, w - 296)
    Set lblDefHint = mUI.AddLabel(advF, "defh", "", 300, 46 + ROW_H + 49, w - 296, 14, False, 8, CMuted())
    mAdvBase = 46 + ROW_H + 66
    ' nur bei Mehrfachauswahl: Höchstzahl wählbarer Optionen
    mUI.AddLabel advF, "lms", "Höchstens auswählbar", 4, mAdvBase, 280, 15, True
    Set txtMaxPick = mUI.AddText(advF, "maxsel", 4, mAdvBase + 17, 80)
    mUI.AddLabel advF, "lmsh", "Leer = beliebig viele. Beispiel: 6 Möglichkeiten, höchstens 3 ankreuzen. Gilt in der Maske, beim Einfügen, Import und in Erfassungsdateien.", _
                 96, mAdvBase + 17, w - 92, 28, False, 8, CMuted()
    mAdvH = mAdvBase
    advF.Height = mAdvH
    advF.Visible = False
    If Not f Is Nothing Then
        txtHelp.Text = f.Help
        chkReq.Value = f.Required
        chkVis.Value = f.Visible
        chkUniq.Value = f.Unique
        chkDup.Value = ListHas(mCfg.GetS(SH_TAB, tid, "Dublettenfelder"), fid)
        cboGroup.Text = f.Group
        txtDefault.Text = DefaultForDisplay(f)
        If f.FType = FT_MEHRFACH And mValid.MaxPicks(f) > 0 Then txtMaxPick.Text = CStr(mValid.MaxPicks(f))
    End If
    Set tp = mUI.AddFrame(fr, "tp", 0, y, fr.Width - 20, 200)
    mLoading = False
    BuildTypePanel
    Set lblErr = mUI.AddLabel(Me, "err", "", PAD, 0, Me.InsideWidth - PAD * 2, 28, False, 8.5, CDanger())
    Set b = mUI.AddButton(Me, "save", "Speichern", PAD, 0, 120, "primary"): Bind b, "save"
    Set b = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 104, 0, 104): Bind b, "cancel"
    b.Cancel = True
    Relayout
End Sub

' Kachelraster aller Eigenschaftsarten (Symbol + Name), Auswahl per Klick
Private Function BuildTypeGrid(ByVal y As Single, ByVal w As Single) As Single
    Dim codes As Variant, i As Long, cw As Single, x As Single, yy As Single, c As String, hit As MSForms.Label, lbl As MSForms.Label
    codes = FieldTypeCodes()
    cw = Int(w / 4)
    For i = 0 To UBound(codes)
        c = CStr(codes(i))
        x = 4 + (i Mod 4) * cw
        yy = y + (i \ 4) * 24
        Set hit = mUI.AddLabel(fr, "tk_" & c, "", x, yy, cw - 4, 22)
        hit.BackStyle = fmBackStyleOpaque
        hit.BackColor = CWhite()
        hit.SpecialEffect = fmSpecialEffectFlat
        hit.BorderStyle = fmBorderStyleSingle
        hit.BorderColor = CBorder()
        Bind hit, "tk:" & c
        Set lbl = mUI.AddIcon(fr, "tki_" & c, FieldTypeIcon(c), x + 6, yy + 4, 11, CMuted())
        Bind lbl, "tk:" & c
        Set lbl = mUI.AddLabel(fr, "tkn_" & c, FieldTypeName(c), x + 24, yy + 4, cw - 32, 15, False, 9)
        lbl.WordWrap = False
        Bind lbl, "tk:" & c
    Next
    BuildTypeGrid = y + ((UBound(codes) \ 4) + 1) * 24
End Function

Private Sub MarkType()
    Dim codes As Variant, i As Long, c As String, sel As Boolean
    codes = FieldTypeCodes()
    For i = 0 To UBound(codes)
        c = CStr(codes(i))
        sel = (c = mKind)
        With fr.Controls("tk_" & c)
            .BackColor = IIf(sel, CAccentSoft(), CWhite())
            .BorderColor = IIf(sel, CAccent(), CBorder())
        End With
        fr.Controls("tki_" & c).ForeColor = IIf(sel, CAccent(), CMuted())
        fr.Controls("tkn_" & c).Font.Bold = sel
    Next
    lblSelType.caption = FieldTypeName(mKind)
    lblTypeHelp.caption = FieldTypeHelp(mKind)
End Sub

Private Sub Relayout()
    Dim h As Single, maxH As Single
    tp.Top = mAdvTop + IIf(advF.Visible, mAdvH + 4, 0)
    h = tp.Top + mTpH + 10
    maxH = mUI.MaxFormHeight() - fr.Top - 100
    If h > maxH Then
        fr.Height = maxH
        fr.ScrollBars = fmScrollBarsVertical
        fr.ScrollHeight = h
    Else
        fr.Height = h
        fr.ScrollBars = fmScrollBarsNone
    End If
    lblErr.Top = fr.Top + fr.Height + 6
    Me.Controls("save").Top = lblErr.Top + 30
    Me.Controls("cancel").Top = lblErr.Top + 30
    Me.Height = Me.Controls("save").Top + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    If Me.Visible Then mUI.KeepOnScreen Me
End Sub

Private Sub Bind(ByVal c As Object, ByVal key As String)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, c, key
    mEvts.Add e
End Sub

Private Function DefaultForDisplay(ByVal f As cField) As String
    Select Case f.FType
        Case FT_AUSWAHL, FT_STATUS: DefaultForDisplay = mCfg.OptName(f.DefaultVal)
        Case FT_ZAHL, FT_GELD: If Len(f.DefaultVal) > 0 Then DefaultForDisplay = FmtNum(InvNum(f.DefaultVal), 6, True)
        Case FT_PROZENT: If Len(f.DefaultVal) > 0 Then DefaultForDisplay = FmtNum(InvNum(f.DefaultVal) * 100, 6, True)
        Case FT_DATUM, FT_DATUMZEIT
            If LCase$(Left$(f.DefaultVal, 5)) = "heute" Then DefaultForDisplay = f.DefaultVal Else If Len(f.DefaultVal) > 0 Then DefaultForDisplay = FmtDate(InvDate(f.DefaultVal))
        Case Else: DefaultForDisplay = f.DefaultVal
    End Select
End Function

'--- Typabhaengiger Bereich -----------------------------------------------------------
Private Sub ClearPanel()
    Dim i As Long
    Set lstOpt = Nothing
    If Not mXl Is Nothing Then mXl.Teardown: Set mXl = Nothing
    mXlMode = False
    For i = tp.Controls.Count - 1 To 0 Step -1
        tp.Controls.Remove tp.Controls(i).Name
    Next
End Sub

Private Sub BuildTypePanel()
    Dim f As cField, y As Single, w As Single, c As Object, b As MSForms.CommandButton, g As cField, ids As Variant, i As Long
    ClearPanel
    w = tp.Width - 8
    If Len(mFid) > 0 Then Set f = mCfg.Field(mFid)
    y = 4
    chkUniq.Enabled = (mType = FT_TEXT Or mType = FT_ZAHL Or mType = FT_DATUM Or mType = FT_PFAD)
    chkDup.Enabled = Not (mType = FT_LINK Or mType = FT_BACKLINK Or mType = FT_LTEXT)
    txtDefault.Enabled = Not (mType = FT_LINK Or mType = FT_BACKLINK Or mType = FT_FORMEL Or mType = FT_MEHRFACH)
    lblDefHint.caption = DefaultHint(mType)
    chkReq.Enabled = Not IsAutoKind(mKind)
    If IsAutoKind(mKind) Then
        chkReq.Value = False: chkUniq.Enabled = False: chkDup.Enabled = False
        txtDefault.Enabled = False
        lblDefHint.caption = "wird automatisch gesetzt"
    End If
    If mKind = FT_KONTROLL Then lblDefHint.caption = "Ja oder Nein (leer = nicht angehakt)"
    If mKind = FT_PERSON Then lblDefHint.caption = "für diesen Typ nicht verfügbar"
    If mKind = FT_DATEIEN Then txtDefault.Enabled = False: lblDefHint.caption = "für diesen Typ nicht verfügbar"
    ' Höchstzahl nur bei Mehrfachauswahl (unter "Weitere Einstellungen")
    advF.Controls("lms").Visible = (mType = FT_MEHRFACH)
    txtMaxPick.Visible = (mType = FT_MEHRFACH)
    advF.Controls("lmsh").Visible = (mType = FT_MEHRFACH)
    mAdvH = mAdvBase + IIf(mType = FT_MEHRFACH, 50, 0)
    advF.Height = mAdvH
    If BuildKindPanel(f, y, w) Then GoTo done
    Select Case mType
        Case FT_TEXT, FT_PFAD
            mUI.AddLabel tp, "l1", "Maximale Länge (Zeichen, leer = unbegrenzt)", 4, y, 300, 15, True
            Set c = mUI.AddText(tp, "maxlen", 4, y + 17, 120)
            If Not f Is Nothing Then If f.MaxLen > 0 Then c.Text = CStr(f.MaxLen)
            y = y + 46
        Case FT_ZAHL, FT_GELD, FT_PROZENT
            If mType <> FT_GELD Then
                mUI.AddLabel tp, "l1", "Nachkommastellen (0 = Ganzzahl)", 4, y, 200, 15, True
                Set c = mUI.AddText(tp, "dec", 4, y + 17, 80)
                If Not f Is Nothing Then c.Text = CStr(f.Decimals) Else c.Text = IIf(mType = FT_ZAHL, "0", "0")
            End If
            mUI.AddLabel tp, "l2", "Minimum" & IIf(mType = FT_PROZENT, " (in %)", ""), 220, y, 150, 15, True
            mUI.AddLabel tp, "l3", "Maximum" & IIf(mType = FT_PROZENT, " (in %)", ""), 380, y, 150, 15, True
            Set c = mUI.AddText(tp, "min", 220, y + 17, 140)
            If Not f Is Nothing Then If Not IsEmpty(f.MinVal) Then c.Text = FmtNum(CDbl(f.MinVal), 6, True)
            Set c = mUI.AddText(tp, "max", 380, y + 17, 140)
            If Not f Is Nothing Then If Not IsEmpty(f.MaxVal) Then c.Text = FmtNum(CDbl(f.MaxVal), 6, True)
            mUI.AddLabel tp, "l4", IIf(mType = FT_GELD, "Geldbeträge werden centgenau (2 Nachkommastellen) in Euro gespeichert.", _
                                   IIf(mType = FT_PROZENT, "Prozent: Eingabe 25 bedeutet 25 %. Gespeichert wird der Anteil (0,25), ohne mehrdeutige Faktoren.", "")), _
                         4, y + 44, w, 28, False, 8, CMuted()
            y = y + 74
        Case FT_DATUM, FT_DATUMZEIT
            mUI.AddLabel tp, "l1", "Darf nicht vor diesem Datumsfeld liegen (optional, z. B. Ende nicht vor Beginn)", 4, y, w, 15, True
            Set c = mUI.AddCombo(tp, "notbefore", 4, y + 17, 280)
            c.AddItem "(keine Regel)"
            c.ListIndex = 0
            For Each g In mCfg.Fields(mTid)
                If (g.FType = FT_DATUM Or g.FType = FT_DATUMZEIT) And g.ID <> mFid Then
                    c.AddItem g.Name
                    If Not f Is Nothing Then If f.NotBefore = g.ID Then c.ListIndex = c.ListCount - 1
                End If
            Next
            y = y + 46
        Case FT_JANEIN
            Set c = mUI.AddCheck(tp, "allowempty", "Leere Angabe erlaubt (sonst ist „Nein“ vorbelegt)", 4, y, w)
            If Not f Is Nothing Then c.Value = f.AllowEmpty
            y = y + ROW_H + 6
        Case FT_AUSWAHL, FT_STATUS, FT_MEHRFACH
            If f Is Nothing Then
                mUI.AddLabel tp, "l1", "Auswahlmöglichkeiten (eine je Zeile)" & IIf(mType = FT_STATUS, " (ein * am Zeilenende kennzeichnet einen abgeschlossenen Zustand)", ""), 4, y, w, 15, True
                Set c = mUI.AddText(tp, "opts", 4, y + 17, 320, 110, True)
                If mType = FT_STATUS Then c.Text = "Offen" & vbCrLf & "In Arbeit" & vbCrLf & "Erledigt *"
                If mKind = FT_PERSON Then c.Text = UserId(): tp.Controls("l1").caption = "Personen, eine je Zeile (z. B. Windows-Kennung oder Name); weitere jederzeit über „Neue Person…“ in der Maske"
                mUI.AddLabel tp, "l2", "Farben werden automatisch vergeben und lassen sich nach dem Speichern ändern. Optionen können später umbenannt, sortiert und für neue Eingaben deaktiviert werden, ohne bestehende Zuordnungen zu verlieren.", _
                             334, y + 17, w - 334, 90, False, 8, CMuted()
                y = y + 134
            Else
                mUI.AddLabel tp, "l1", "Auswahlmöglichkeiten (Änderungen werden sofort gespeichert)", 4, y, w, 15, True
                Set lstOpt = mUI.AddList(tp, "optlist", 4, y + 17, 330, 150, False, 4, "140;70;60;55")
                Bind lstOpt, "optlist"
                Set b = mUI.AddButton(tp, "oadd", "Neu…", 344, y + 17, 110, , ROW_H): Bind b, "oadd"
                Set b = mUI.AddButton(tp, "oren", "Umbenennen…", 344, y + 43, 110, , ROW_H): Bind b, "oren"
                Set b = mUI.AddButton(tp, "ocol", "Farbe…", 344, y + 69, 110, , ROW_H): Bind b, "ocol"
                Set b = mUI.AddButton(tp, "oup", "Nach oben", 460, y + 17, 110, , ROW_H): Bind b, "oup"
                Set b = mUI.AddButton(tp, "odown", "Nach unten", 460, y + 43, 110, , ROW_H): Bind b, "odown"
                Set b = mUI.AddButton(tp, "oact", "Aktiv / inaktiv", 460, y + 69, 110, , ROW_H): Bind b, "oact"
                If mType = FT_STATUS Then
                    Set b = mUI.AddButton(tp, "odone", "Abgeschlossen ein/aus", 344, y + 95, 226, , ROW_H): Bind b, "odone"
                End If
                Set b = mUI.AddButton(tp, "odel", "Löschen…", 344, y + 121, 110, "danger", ROW_H): Bind b, "odel"
                y = y + 176
                RefreshOptions ""
            End If
        Case FT_LINK
            If f Is Nothing Then
                mUI.AddLabel tp, "l0", "Eine Verknüpfung verbindet Einträge dieser Tabelle mit Einträgen einer anderen Tabelle: Statt Namen abzutippen, wird ausgewählt. " & _
                             "Beispiel: In „Aufgaben“ die Verknüpfung „Projekt“ zur Tabelle „Projekte“: Jede Aufgabe wählt ihr Projekt, und beim Projekt erscheinen automatisch alle seine Aufgaben.", _
                             4, y, w, 40, False, 8.5, CMuted()
                y = y + 44
                mUI.AddLabel tp, "l1", "1. Mit welcher Tabelle verbinden? *", 4, y, 250, 15, True
                Set c = mUI.AddCombo(tp, "target", 4, y + 17, 250)
                ids = mCfg.TableIds()
                For i = 0 To ArrLen(ids) - 1
                    c.AddItem mCfg.TableName(CStr(ids(i)))
                Next
                Bind c, "target"
                mUI.AddLabel tp, "l2", "2. Wie viele Einträge je Eintrag?", 270, y, 250, 15, True
                Set c = mUI.AddCheck(tp, "multi", "Mehrere auswählen erlaubt", 270, y + 17, w - 270)
                Bind c, "multi"
                y = y + 46
                mUI.AddLabel tp, "l3", "3. Name des Feldes auf der Gegenseite", 4, y, 300, 15, True
                Set c = mUI.AddText(tp, "backname", 4, y + 17, 250)
                c.Text = mCfg.TableName(mTid)
                Bind c, "backname"
                mUI.AddLabel tp, "l4", "Dort entsteht automatisch ein Feld, das die verknüpften Einträge von hier zeigt, ohne doppelte Erfassung.", _
                             270, y + 17, w - 270, 28, False, 8, CMuted()
                y = y + 48
                ' Klartext, was entsteht (aktualisiert sich bei jeder Auswahl)
                Set c = mUI.AddLabel(tp, "linksent", "", 4, y, w, 42, False, 9, CAccent())
                y = y + 46
                LinkSentence
            Else
                mUI.AddLabel tp, "l1", "Ziel: " & mCfg.TableName(f.TargetTable) & "  ·  " & IIf(f.Multi, "Mehrfachzuordnung (n:m)", "Einfachzuordnung (n:1)") & _
                             "  ·  Gegenseite: „" & mCfg.FieldName(f.PairField) & "“", 4, y, w, 15
                y = y + 24
            End If
        Case FT_BACKLINK
            mUI.AddLabel tp, "l1", "Rückverknüpfung zu „" & mCfg.FieldName(f.PairField) & "“ in „" & mCfg.TableName(mCfg.Field(f.PairField).TableID) & "“. " & _
                         "Zuordnungen können auf beiden Seiten gepflegt werden.", 4, y, w, 30, False, 9
            y = y + 34
        Case FT_FORMEL
            ' Excel-Formel direkt eingeben (Standard); Rollups und bestehende Baukasten-
            ' Berechnungen weiter ueber den Assistenten
            mRollInline = RollupInline()
            mXlMode = Not mRollInline And (mKind <> FT_ROLLUP) And Left$(mFormula, 6) <> "ROLLUP" And (Len(mFormula) = 0 Or mCalc.IsXl(mFormula))
            If mRollInline Then
                BuildRollupPanel y, w
            ElseIf mXlMode Then
                Set mXl = New cXlEditor
                y = y + mXl.Build(tp, mTid, mFid, mFormula, 4, y, w, Me)
            Else
                mUI.AddLabel tp, "l1", "Berechnung", 4, y, 200, 15, True
                Set b = mUI.AddButton(tp, "fx", "Berechnung festlegen…", 4, y + 17, 180, "primary", ROW_H): Bind b, "fx"
                mUI.AddLabel tp, "fxd", IIf(Len(mFormula) > 0, mCalc.Describe(mFormula), "Noch keine Berechnung festgelegt."), 194, y + 17, w - 194, 40, False, 8.5, CMuted()
                If mKind <> FT_ROLLUP Then
                    Set b = mUI.AddButton(tp, "toxl", "Stattdessen Excel-Formel schreiben", 4, y + 44, 240, , ROW_H): Bind b, "toxl"
                    y = y + 28
                End If
                y = y + 62
            End If
            mUI.AddLabel tp, "l2", "Ergebnis anzeigen als", 4, y, 200, 15, True
            mUI.AddLabel tp, "l3", "Nachkommastellen", 220, y, 200, 15, True
            Set c = mUI.AddCombo(tp, "rt", 4, y + 17, 200)
            For Each ids In Array(FT_ZAHL, FT_GELD, FT_PROZENT, FT_DATUM, FT_TEXT, FT_JANEIN)
                c.AddItem FieldTypeName(CStr(ids))
            Next
            c.Value = FieldTypeName(IIf(Len(mFormulaRT) > 0, mFormulaRT, FT_ZAHL))
            Bind c, "rt"
            Set c = mUI.AddText(tp, "fdec", 220, y + 17, 80)
            c.Text = CStr(mFormulaDec)
            Bind c, "fdec"
            mRtAuto = (Len(mFid) = 0)
            If mXlMode Then
                Set b = mUI.AddButton(tp, "fx", "Assistent statt Formel…", w - 186, y + 15, 190, , ROW_H): Bind b, "fx"
                mXl.ResultType = FieldTypeFromName(tp.Controls("rt").Text)
                mXl.Decimals = mFormulaDec
                mXl.Refresh
            End If
            y = y + 46
            If mRollInline Then RollupChanged ""
    End Select
done:
    tp.Height = y + 4
    mTpH = y + 4
End Sub

' Einstellungen der Notion-Arten, die ueber den Grundtyp hinausgehen; True = Bereich vollstaendig
Private Function BuildKindPanel(ByVal f As cField, ByRef y As Single, ByVal w As Single) As Boolean
    Dim c As Object, i As Long, p As Variant, tf As cField, n As Long, info As String
    Select Case mKind
        Case FT_KONTROLL
            info = "Ein Häkchen pro Eintrag. In der Tabelle schaltet ein Doppelklick auf die Zelle das Häkchen um; in Filtern und Dashboards zählt „angehakt“ / „nicht angehakt“."
        Case FT_URL
            info = "Webadresse, z. B. https://intranet/seite. Die Adresse wird geprüft und erst nach Rückfrage im Browser geöffnet. Die Anwendung selbst ruft nie das Internet auf."
        Case FT_EMAIL
            info = "E-Mail-Adresse mit Formatprüfung. „E-Mail schreiben“ öffnet nach Rückfrage das lokale E-Mail-Programm; es wird nichts automatisch versendet."
        Case FT_TELEFON
            info = "Telefonnummer; erlaubt sind Ziffern, Leerzeichen, +, -, / und Klammern."
        Case FT_ORT
            info = "Adresse oder Ort als Text. Es wird bewusst kein Kartendienst verwendet (offline)."
        Case FT_DATEIEN
            info = "Mehrere lokale Dateien je Eintrag. Gespeichert werden nur Verweise (Pfade), keine Kopien. Öffnen erfolgt erst nach Rückfrage."
        Case FT_ERST_UM, FT_AEND_UM, FT_ERST_VON, FT_AEND_VON
            info = FieldTypeHelp(mKind) & ". Die Werte werden von der Anwendung gepflegt und können nicht von Hand geändert werden; vorhandene Einträge werden beim Anlegen nachgetragen."
        Case FT_AUTOID
            mUI.AddLabel tp, "l1", "Präfix (optional, z. B. AUF ergibt AUF-1, AUF-2 …)", 4, y, 320, 15, True
            Set c = mUI.AddText(tp, "prefix", 4, y + 17, 120)
            If Not f Is Nothing Then c.Text = f.DefaultVal: c.Enabled = False
            mUI.AddLabel tp, "l2", "Die Nummer wird beim Anlegen vergeben, ist eindeutig und wird nie wiederverwendet. Vorhandene Einträge erhalten beim Anlegen des Feldes fortlaufende Nummern.", _
                         140, y + 17, w - 140, 40, False, 8, CMuted()
            y = y + 62
            BuildKindPanel = True
        Case FT_BUTTON
            mUI.AddLabel tp, "l1", "Aktionen beim Klick (bis zu drei Felder setzen)", 4, y, w, 15, True
            y = y + 18
            p = Split(mFormula, "|")
            For i = 1 To 3
                Set c = mUI.AddCombo(tp, "bt" & i, 4, y, 250)
                c.AddItem "(keine Aktion)"
                For Each tf In mCfg.Fields(mTid)
                    If tf.IsEditable And Not tf.IsLinkType And tf.ID <> mFid Then c.AddItem tf.Name
                Next
                c.ListIndex = 0
                Bind c, "bt" & i
                mUI.AddLabel tp, "lq" & i, "=", 260, y + 4, 14, 15, True
                Set c = mUI.AddCombo(tp, "bv" & i, 278, y, w - 278, False)
                n = 2 * i - 1
                If UBound(p) >= n + 1 Then
                    If Left$(p(n), 2) = "F:" And mCfg.FieldExists(Mid$(p(n), 3)) Then
                        mLoading = True
                        tp.Controls("bt" & i).Text = mCfg.FieldName(Mid$(p(n), 3))
                        mLoading = False
                        FillButtonValues i
                        c.Text = Unesc(CStr(p(n + 1)))
                    End If
                End If
                y = y + ROW_H + 6
            Next
            mUI.AddLabel tp, "l2", "Werte: Name einer Option, ein Text oder Zahl sowie heute, jetzt, ich (eigene Kennung), leer oder umschalten (Kontrollkästchen). " & _
                         "Der Button erscheint in der Eingabemaske; in der Tabelle löst ein Doppelklick auf die Zelle ihn aus.", 4, y, w, 40, False, 8, CMuted()
            y = y + 44
            BuildKindPanel = True
    End Select
    If Len(info) > 0 Then
        mUI.AddLabel tp, "l0", info, 4, y, w, 40, False, 8.5, CMuted()
        y = y + 44
        If mKind <> FT_URL And mKind <> FT_EMAIL And mKind <> FT_TELEFON And mKind <> FT_ORT Then BuildKindPanel = True
    End If
End Function

' Vorschlaege fuer den Wert einer Button-Aktion je nach Zielfeld
Private Sub FillButtonValues(ByVal i As Long)
    Dim cb As MSForms.ComboBox, tf As cField, v As Variant, fid As String
    Set cb = tp.Controls("bv" & i)
    cb.Clear
    fid = mCfg.FieldIdByName(mTid, tp.Controls("bt" & i).Text)
    If Len(fid) = 0 Then Exit Sub
    Set tf = mCfg.Field(fid)
    Select Case tf.FType
        Case FT_DATUM, FT_DATUMZEIT: cb.AddItem "heute": cb.AddItem "jetzt"
        Case FT_JANEIN: cb.AddItem "umschalten": cb.AddItem "Ja": cb.AddItem "Nein"
        Case FT_AUSWAHL, FT_STATUS, FT_MEHRFACH
            For Each v In mCfg.OptionIds(fid)
                cb.AddItem mCfg.OptName(CStr(v))
            Next
            If tf.Kind = FT_PERSON Then cb.AddItem "ich"
        Case FT_TEXT: cb.AddItem "ich"
    End Select
    cb.AddItem "leer"
    If cb.ListCount > 0 Then cb.ListIndex = 0
End Sub

Private Function DefaultHint(ByVal t As String) As String
    Select Case t
        Case FT_DATUM, FT_DATUMZEIT: DefaultHint = "„heute“, „heute+7“ oder ein Datum"
        Case FT_JANEIN: DefaultHint = "Ja oder Nein"
        Case FT_AUSWAHL, FT_STATUS: DefaultHint = "Name einer Auswahlmöglichkeit"
        Case FT_ZAHL, FT_GELD: DefaultHint = "Zahl, z. B. 0"
        Case FT_PROZENT: DefaultHint = "Prozentwert, z. B. 0"
        Case FT_LINK, FT_BACKLINK, FT_FORMEL, FT_MEHRFACH: DefaultHint = "für diesen Typ nicht verfügbar"
    End Select
End Function

Private Sub RefreshOptions(ByVal selOid As String)
    Dim v As Variant, n As Long, sel As Long
    If lstOpt Is Nothing Then Exit Sub
    lstOpt.Clear
    ReDim mOptIds(0 To mCfg.OptionIds(mFid).Count)
    sel = -1
    For Each v In mCfg.OptionIds(mFid)
        lstOpt.AddItem mCfg.OptName(CStr(v))
        lstOpt.List(n, 1) = PaletteName(mCfg.OptColor(CStr(v)))
        lstOpt.List(n, 2) = IIf(mCfg.OptActive(CStr(v)), "aktiv", "inaktiv")
        lstOpt.List(n, 3) = IIf(mCfg.OptDone(CStr(v)), "erledigt", "")
        mOptIds(n) = CStr(v)
        If CStr(v) = selOid Then sel = n
        n = n + 1
    Next
    If sel >= 0 Then lstOpt.ListIndex = sel
End Sub

Private Function SelOid() As String
    If lstOpt Is Nothing Then Exit Function
    If lstOpt.ListIndex >= 0 Then SelOid = mOptIds(lstOpt.ListIndex)
End Function

'--- Ereignisse -------------------------------------------------------------------------
Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    Dim oid As String, nm As String, ok As Boolean, items() As Variant, keys As Variant, i As Long, sel As Long
    On Error GoTo failed
    If Left$(key, 3) = "tk:" Then
        If evt = "Click" And Mid$(key, 4) <> mKind Then
            mKind = Mid$(key, 4)
            mType = BaseTypeOf(mKind)
            If mKind = FT_ROLLUP Then mFormula = "": mFormulaRT = FT_ZAHL
            MarkType
            BuildTypePanel
            Relayout
        End If
        Exit Sub
    End If
    Select Case key
        Case "bt1", "bt2", "bt3"
            If evt = "Change" And Not mLoading Then FillButtonValues CLng(Right$(key, 1))
        Case "chtype": If evt = "Click" Then ChangeType
        Case "adv"
            If evt = "Click" Then
                advF.Visible = Not advF.Visible
                fr.Controls("adv").caption = "Weitere Einstellungen  " & IIf(advF.Visible, ChrW(&H25BE), ChrW(&H25B8))
                Relayout
            End If
        Case "fx": If evt = "Click" Then EditFormula
        Case "target", "multi", "backname"
            If (evt = "Change" Or evt = "Click") And Not mLoading Then LinkSentence
        Case "ru_link", "ru_agg", "ru_fld"
            If evt = "Change" And Not mLoading Then RollupChanged key
        Case "toxl"
            If evt = "Click" Then
                mFormula = ""
                BuildTypePanel
                Relayout
            End If
        Case "rt", "fdec"
            If evt = "Change" And Not mLoading And key = "rt" Then mRtAuto = False
            If evt = "Change" And Not mLoading And Not mXl Is Nothing Then
                mXl.ResultType = FieldTypeFromName(tp.Controls("rt").Text)
                mXl.Decimals = CLng(Val(tp.Controls("fdec").Text))
                mXl.Refresh
            End If
        Case "save": If evt = "Click" Then DoSave
        Case "cancel": If evt = "Click" Then Me.Hide
        Case "oadd"
            If evt = "Click" Then
                nm = mUI.AskText("Name der neuen Auswahlmöglichkeit:", "Neue Option", "", False, ok)
                If ok And Len(Trim$(nm)) > 0 Then
                    oid = mStruct.AddOption(mFid, nm, "", False)
                    mStruct.RebuildLists
                    RefreshOptions oid
                End If
            End If
        Case "oren"
            oid = SelOid
            If evt = "Click" And Len(oid) > 0 Then
                nm = mUI.AskText("Neuer Name (bestehende Zuordnungen bleiben erhalten, da intern die ID gespeichert ist):", "Option umbenennen", mCfg.OptName(oid), False, ok)
                If ok And Len(Trim$(nm)) > 0 Then
                    mStruct.UpdateOption oid, nm, mCfg.OptColor(oid), mCfg.OptActive(oid), mCfg.OptDone(oid)
                    mStruct.RebuildLists
                    RefreshOptions oid
                End If
            End If
        Case "ocol"
            oid = SelOid
            If evt = "Click" And Len(oid) > 0 Then
                keys = PaletteKeys()
                ReDim items(0 To UBound(keys))
                For i = 0 To UBound(keys): items(i) = PaletteName(CStr(keys(i))): Next
                sel = mUI.ChooseOne("Farbe für „" & mCfg.OptName(oid) & "“ (zusätzlich zur Beschriftung):", "Farbe", items, "Übernehmen")
                If sel >= 0 Then
                    mStruct.UpdateOption oid, mCfg.OptName(oid), CStr(keys(sel)), mCfg.OptActive(oid), mCfg.OptDone(oid)
                    RefreshOptions oid
                End If
            End If
        Case "oup", "odown"
            oid = SelOid
            If evt = "Click" And Len(oid) > 0 Then
                mStruct.MoveInOrder SH_OPT, "Feld", mFid, oid, IIf(key = "oup", -1, 1)
                mStruct.RebuildLists
                RefreshOptions oid
            End If
        Case "oact"
            oid = SelOid
            If evt = "Click" And Len(oid) > 0 Then
                mStruct.UpdateOption oid, mCfg.OptName(oid), mCfg.OptColor(oid), Not mCfg.OptActive(oid), mCfg.OptDone(oid)
                mStruct.RebuildLists
                RefreshOptions oid
            End If
        Case "odone"
            oid = SelOid
            If evt = "Click" And Len(oid) > 0 Then
                mStruct.UpdateOption oid, mCfg.OptName(oid), mCfg.OptColor(oid), mCfg.OptActive(oid), Not mCfg.OptDone(oid)
                RefreshOptions oid
            End If
        Case "odel"
            oid = SelOid
            If evt = "Click" And Len(oid) > 0 Then
                If mUI.Confirm("Option „" & mCfg.OptName(oid) & "“ löschen?", "Option löschen", "Löschen", "Abbrechen", True) Then
                    mStruct.DeleteOption oid
                    mStruct.RebuildLists
                    RefreshOptions ""
                End If
            End If
    End Select
    Exit Sub
failed:
    mApp.Boundary "Feld bearbeiten"
End Sub

Private Sub ChangeType()
    Dim res As String
    res = mCmd.ChangeTypeFlow(mFid)
    If Len(res) > 0 Then
        mType = res: mKind = res
        Me.Controls("typ").caption = FieldTypeName(res)
        BuildTypePanel
        Relayout
    End If
End Sub

Private Sub EditFormula()
    Dim f As frmFormula
    Set f = New frmFormula
    f.Setup mTid, mFid, IIf(mCalc.IsXl(mFormula), "", mFormula)
    f.Show
    If f.OK Then
        mFormula = f.Definition
        mFormulaRT = f.ResultType
        If mXlMode Or mRollInline Then
            ' vom Formelfeld/Rollup-Bereich zum Assistenten gewechselt
            BuildTypePanel
            Relayout
        Else
            tp.Controls("fxd").caption = mCalc.Describe(mFormula)
            On Error Resume Next
            tp.Controls("rt").Value = FieldTypeName(mFormulaRT)
        End If
    End If
    Unload f
End Sub

'--- Verknuepfung: Klartext ---------------------------------------------------------------
Private Sub LinkSentence()
    Dim tn As String, me_ As String, back As String, multi As Boolean
    On Error Resume Next
    tn = tp.Controls("target").Text
    multi = tp.Controls("multi").Value
    back = Trim$(tp.Controls("backname").Text)
    me_ = mCfg.TableName(mTid)
    If Len(tn) = 0 Then
        tp.Controls("linksent").caption = "Bitte zuerst die Tabelle wählen, mit der verbunden werden soll."
    Else
        tp.Controls("linksent").caption = ChrW(&H2192) & "  Jeder Eintrag in „" & me_ & "“ wählt " & IIf(multi, "beliebig viele Einträge", "genau einen Eintrag") & _
            " aus „" & tn & "“." & vbLf & ChrW(&H2190) & "  In „" & tn & "“ zeigt das Feld „" & IIf(Len(back) > 0, back, me_) & _
            "“ automatisch alle Einträge aus „" & me_ & "“, die ihn ausgewählt haben."
    End If
End Sub

'--- Rollup direkt im Dialog -----------------------------------------------------------------
' Ohne Bedingung (Bedingungen weiter ueber den Assistenten)
Private Function RollupInline() As Boolean
    Dim p As Variant
    If mKind <> FT_ROLLUP And Left$(mFormula, 6) <> "ROLLUP" Then Exit Function
    p = Split(mFormula, "|")
    If UBound(p) >= 4 Then
        If Len(p(4)) > 0 Then Exit Function
    End If
    RollupInline = True
End Function

Private Sub BuildRollupPanel(ByRef y As Single, ByVal w As Single)
    Dim c As MSForms.ComboBox, g As cField, n As Long, p As Variant, i As Long, b As MSForms.CommandButton
    mUI.AddLabel tp, "ru0", "Ein Rollup fasst die verknüpften Einträge zu einem Wert zusammen. Beispiel: In „Projekte“ über die Verknüpfung „Aufgaben“ " & _
                 "die Anzahl der Aufgaben oder die Summe ihrer Kosten. Der Wert aktualisiert sich automatisch.", 4, y, w, 40, False, 8.5, CMuted()
    y = y + 44
    ReDim mRuLinks(0 To mCfg.Fields(mTid).Count)
    For Each g In mCfg.Fields(mTid)
        If g.IsLinkType Then mRuLinks(n) = g.ID: n = n + 1
    Next
    If n = 0 Then
        mUI.AddLabel tp, "ru_none", ChrW(&H26A0) & "  Diese Tabelle hat noch keine Verknüpfung. Bitte zuerst eine Eigenschaft der Art „Verknüpfung“ anlegen " & _
                     "(z. B. „Aufgaben“ " & ChrW(&H2192) & " Tabelle Aufgaben); danach kann der Rollup sie auswerten.", 4, y, w, 30, False, 9, CWarn()
        y = y + 36
        Exit Sub
    End If
    ReDim Preserve mRuLinks(0 To n - 1)
    mUI.AddLabel tp, "rul1", "1. Über welche Verknüpfung?", 4, y, 200, 15, True
    mUI.AddLabel tp, "rul2", "2. Was berechnen?", 214, y, 200, 15, True
    mUI.AddLabel tp, "rul3", "3. Von welchem Feld?", 424, y, w - 420, 15, True
    mLoading = True
    Set c = mUI.AddCombo(tp, "ru_link", 4, y + 17, 200)
    For i = 0 To n - 1
        c.AddItem mCfg.FieldName(mRuLinks(i)) & "  (" & mCfg.TableName(RuTarget(mRuLinks(i))) & ")"
    Next
    Bind c, "ru_link"
    Set c = mUI.AddCombo(tp, "ru_agg", 214, y + 17, 200)
    For Each p In Split(RU_NAMES, ";")
        c.AddItem CStr(p)
    Next
    Bind c, "ru_agg"
    Set c = mUI.AddCombo(tp, "ru_fld", 424, y + 17, w - 420)
    Bind c, "ru_fld"
    ' bestehende Einstellung uebernehmen
    p = Split(mFormula, "|")
    tp.Controls("ru_link").ListIndex = 0
    tp.Controls("ru_agg").ListIndex = 0
    If UBound(p) >= 2 Then
        For i = 0 To n - 1
            If "F:" & mRuLinks(i) = p(1) Then tp.Controls("ru_link").ListIndex = i
        Next
        i = IndexIn(RU_AGGS, Mid$(p(2), 3))
        If i < 0 And Mid$(p(2), 3) = "ANZAHLWENN" Then i = 0
        If i >= 0 Then tp.Controls("ru_agg").ListIndex = i
    End If
    FillRollupFields IIf(UBound(p) >= 3, Mid$(NzS(ArrItem(p, 3)), 3), "")
    mLoading = False
    y = y + 46
    Set c = Nothing
    mUI.AddLabel tp, "ru_sent", "", 4, y, w, 30, False, 9, CAccent()
    y = y + 34
    Set b = mUI.AddButton(tp, "fx", "Mit Bedingung (erweitert)…", 4, y, 200, , ROW_H): Bind b, "fx"
    mUI.AddLabel tp, "ru_fxh", "z. B. nur Aufgaben mit Status „Erledigt“ zählen", 212, y + 4, w - 212, 15, False, 8.5, CMuted()
    y = y + ROW_H + 12
End Sub

Private Function ArrItem(ByVal a As Variant, ByVal i As Long) As Variant
    If i <= UBound(a) Then ArrItem = a(i)
End Function

Private Function IndexIn(ByVal lst As String, ByVal code As String) As Long
    Dim a As Variant, i As Long
    a = Split(lst, ";")
    IndexIn = -1
    For i = 0 To UBound(a)
        If a(i) = code Then IndexIn = i: Exit Function
    Next
End Function

Private Function RuTarget(ByVal linkFid As String) As String
    Dim lf As cField
    Set lf = mCfg.Field(linkFid)
    If lf.FType = FT_LINK Then RuTarget = lf.TargetTable Else RuTarget = mCfg.Field(lf.PairField).TableID
End Function

Private Function RuAgg() As String
    RuAgg = Split(RU_AGGS, ";")(Application.Max(0, tp.Controls("ru_agg").ListIndex))
End Function

' Felder der verknuepften Tabelle passend zur Auswertung (Summe/Durchschnitt nur Zahlen)
Private Sub FillRollupFields(ByVal keepFid As String)
    Dim c As MSForms.ComboBox, tf As cField, agg As String, n As Long, ok As Boolean, sel As Long, tt As String
    Set c = tp.Controls("ru_fld")
    c.Clear
    agg = RuAgg()
    tt = RuTarget(mRuLinks(tp.Controls("ru_link").ListIndex))
    ReDim mRuFlds(0 To mCfg.Fields(tt).Count)
    If agg = "ANZAHL" Then
        c.AddItem "(nicht nötig, es wird gezählt)"
        mRuFlds(0) = ""
        c.ListIndex = 0
        c.Enabled = False
        Exit Sub
    End If
    c.Enabled = True
    sel = 0
    For Each tf In mCfg.Fields(tt)
        Select Case agg
            Case "SUMME", "MITTEL": ok = (tf.ValueType = FT_ZAHL Or tf.ValueType = FT_GELD Or tf.ValueType = FT_PROZENT)
            Case "MIN", "MAX": ok = (tf.ValueType = FT_ZAHL Or tf.ValueType = FT_GELD Or tf.ValueType = FT_PROZENT Or tf.ValueType = FT_DATUM Or tf.ValueType = FT_DATUMZEIT)
            Case Else: ok = Not tf.IsLinkType
        End Select
        If ok Then
            c.AddItem tf.Name
            mRuFlds(n) = tf.ID
            If tf.ID = keepFid Then sel = n
            n = n + 1
        End If
    Next
    If n = 0 Then
        c.AddItem "(keine passenden Felder in „" & mCfg.TableName(tt) & "“)"
        mRuFlds(0) = ""
    End If
    c.ListIndex = sel
End Sub

' Auswahl geaendert: Felder nachfuehren, Formel bilden, Klartext und Ergebnistyp
Private Sub RollupChanged(ByVal key As String)
    Dim lfid As String, agg As String, fld As String, tt As String, s As String, tf As cField, rt As String
    On Error GoTo done
    If tp.Controls("ru_link") Is Nothing Then Exit Sub
    mLoading = True
    If key = "ru_link" Or key = "ru_agg" Then
        If tp.Controls("ru_fld").ListIndex >= 0 Then fld = mRuFlds(tp.Controls("ru_fld").ListIndex)
        FillRollupFields fld
    End If
    lfid = mRuLinks(tp.Controls("ru_link").ListIndex)
    agg = RuAgg()
    fld = ""
    If tp.Controls("ru_fld").ListIndex >= 0 Then fld = mRuFlds(tp.Controls("ru_fld").ListIndex)
    tt = RuTarget(lfid)
    mFormula = "ROLLUP|F:" & lfid & "|K:" & agg & IIf(Len(fld) > 0, "|F:" & fld, "")
    If agg = "ANZAHL" Then
        s = "Anzahl der verknüpften Einträge aus „" & mCfg.TableName(tt) & "“"
    ElseIf Len(fld) = 0 Then
        s = "Bitte ein Feld wählen."
        mFormula = ""
    Else
        Set tf = mCfg.Field(fld)
        s = Split(RU_NAMES, ";")(IndexIn(RU_AGGS, agg)) & " von „" & tf.Name & "“ über alle verknüpften Einträge aus „" & mCfg.TableName(tt) & "“"
    End If
    tp.Controls("ru_sent").caption = ChrW(&H2192) & "  Ergebnis je Eintrag in „" & mCfg.TableName(mTid) & "“: " & s & "."
    ' Ergebnistyp vorschlagen (nur solange nicht selbst geaendert)
    If mRtAuto Then
        rt = FT_ZAHL
        If Not tf Is Nothing And agg <> "GEFUELLT" Then
            If tf.ValueType = FT_GELD Or tf.ValueType = FT_DATUM Or tf.ValueType = FT_PROZENT Then rt = tf.ValueType
        End If
        tp.Controls("rt").Value = FieldTypeName(rt)
        If agg = "ANZAHL" Or agg = "GEFUELLT" Then tp.Controls("fdec").Text = "0"
        If agg = "MITTEL" And rt = FT_ZAHL Then tp.Controls("fdec").Text = "2"
    End If
done:
    mLoading = False
End Sub

' Vorschau der Formel: Ergebnistyp bei neuen Formelspalten passend vorschlagen
Public Sub OnFormulaPreview(ByVal v As Variant)
    Dim nm As String
    If Not mRtAuto Then Exit Sub
    Select Case VarType(v)
        Case vbBoolean: nm = FieldTypeName(FT_JANEIN)
        Case vbString: If Not IsErrVal(v) Then nm = FieldTypeName(FT_TEXT)
        Case vbDouble: nm = FieldTypeName(FT_ZAHL)
    End Select
    If Len(nm) = 0 Or tp.Controls("rt").Text = nm Then Exit Sub
    mLoading = True
    tp.Controls("rt").Value = nm
    mLoading = False
    mXl.ResultType = FieldTypeFromName(nm)
End Sub

'--- Speichern -----------------------------------------------------------------------------
Private Function UsesBasePanel() As Boolean
    UsesBasePanel = Not (IsAutoKind(mKind) Or mKind = FT_KONTROLL Or mKind = FT_DATEIEN)
End Function

' Zusatzeinstellungen der Notion-Arten in props uebernehmen; False bei Eingabefehler
Private Function KindProps(ByVal f As cField, ByVal props As cMap) As Boolean
    Dim i As Long, fid As String, s As String, spec As String, tf As cField, v As Variant, msg As String
    Select Case mKind
        Case FT_KONTROLL
            props.SetVal "LeerErlaubt", "Nein"
        Case FT_BUTTON
            spec = "AKTION"
            For i = 1 To 3
                fid = mCfg.FieldIdByName(mTid, tp.Controls("bt" & i).Text)
                If Len(fid) > 0 Then
                    s = Trim$(tp.Controls("bv" & i).Text)
                    Set tf = mCfg.Field(fid)
                    If Len(s) = 0 Then lblErr.caption = ChrW(&H26A0) & "  Aktion " & i & ": bitte einen Wert für „" & tf.Name & "“ angeben (oder „leer“).": Exit Function
                    If LCase$(s) <> "ich" Then
                        If Not mForms.ButtonValue(tf, s, Empty, v, msg) Then lblErr.caption = ChrW(&H26A0) & "  Aktion " & i & " (" & tf.Name & "): " & msg: Exit Function
                    End If
                    spec = spec & "|F:" & fid & "|" & Esc(s)
                End If
            Next
            If spec = "AKTION" Then lblErr.caption = ChrW(&H26A0) & "  Bitte mindestens eine Aktion festlegen.": Exit Function
            props.SetVal "Formel", spec
            props.SetVal "Beschreibung", ""
    End Select
    KindProps = True
End Function

' Fehler in "Weitere Einstellungen": Bereich aufklappen, Feld markieren
Private Sub ShowAdvError(ByVal msg As String)
    lblErr.caption = ChrW(&H26A0) & "  " & msg
    If Not advF.Visible Then
        advF.Visible = True
        fr.Controls("adv").caption = "Weitere Einstellungen  " & ChrW(&H25BE)
        Relayout
    End If
    mUI.MarkError txtMaxPick, True
    On Error Resume Next
    txtMaxPick.SetFocus
End Sub

Private Function ReadNum(ByVal ctlName As String, ByRef v As Variant, ByVal label As String) As Boolean
    Dim d As Double, s As String
    v = Empty
    On Error Resume Next
    s = Trim$(tp.Controls(ctlName).Text)
    On Error GoTo 0
    If Len(s) = 0 Then ReadNum = True: Exit Function
    If Not ParseNumberDE(s, d) Then lblErr.caption = ChrW(&H26A0) & "  " & label & ": bitte eine Zahl eingeben.": Exit Function
    v = d
    ReadNum = True
End Function

Private Sub DoSave()
    Dim props As New cMap, nm As String, v As Variant, mn As Variant, mx As Variant, dec As Variant, msg As String, f As cField
    Dim fid As String, t As cTableData, empties As Long, r As Long, c As Long, defV As Variant, ans As Long, lines As Variant, i As Long
    Dim ln As String, done As Boolean, other As String, tgt As String, tmpF As New cField, parsed As Variant, dmsg As String
    lblErr.caption = ""
    mUI.MarkError txtName, False
    nm = Trim$(txtName.Text)
    If Len(nm) = 0 Then lblErr.caption = ChrW(&H26A0) & "  Bitte einen Feldnamen eingeben.": mUI.MarkError txtName, True: Exit Sub
    other = mCfg.FieldIdByName(mTid, nm)
    If Len(other) > 0 And other <> mFid Then lblErr.caption = ChrW(&H26A0) & "  Diesen Feldnamen gibt es in der Tabelle bereits.": mUI.MarkError txtName, True: Exit Sub
    If Len(mFid) > 0 Then Set f = mCfg.Field(mFid)
    props.SetVal "Name", nm
    props.SetVal "Hilfetext", Trim$(txtHelp.Text)
    props.SetVal "Pflicht", IIf(chkReq.Value, "Ja", "Nein")
    props.SetVal "Sichtbar", IIf(chkVis.Value, "Ja", "Nein")
    props.SetVal "Eindeutig", IIf(chkUniq.Value And chkUniq.Enabled, "Ja", "Nein")
    props.SetVal "Gruppe", Trim$(cboGroup.Text)
    ' typabhaengig
    If Not KindProps(f, props) Then Exit Sub
    If UsesBasePanel() Then
    Select Case mType
        Case FT_TEXT, FT_PFAD
            If Not ReadNum("maxlen", v, "Maximale Länge") Then Exit Sub
            props.SetVal "MaxLaenge", IIf(IsEmpty(v), Empty, CLng(Val(v)))
        Case FT_ZAHL, FT_GELD, FT_PROZENT
            If mType <> FT_GELD Then
                If Not ReadNum("dec", dec, "Nachkommastellen") Then Exit Sub
                If IsEmpty(dec) Then dec = 0
                If dec < 0 Or dec > 6 Then lblErr.caption = ChrW(&H26A0) & "  Nachkommastellen: 0 bis 6.": Exit Sub
                props.SetVal "Dezimalstellen", CLng(dec)
            Else
                props.SetVal "Dezimalstellen", 2
            End If
            If Not ReadNum("min", mn, "Minimum") Then Exit Sub
            If Not ReadNum("max", mx, "Maximum") Then Exit Sub
            If Not IsEmpty(mn) And Not IsEmpty(mx) Then
                If mn > mx Then lblErr.caption = ChrW(&H26A0) & "  Das Minimum darf nicht größer als das Maximum sein.": Exit Sub
            End If
            props.SetVal "Minimum", IIf(IsEmpty(mn), Empty, NumInv(CDbl(IIf(IsEmpty(mn), 0, mn))))
            props.SetVal "Maximum", IIf(IsEmpty(mx), Empty, NumInv(CDbl(IIf(IsEmpty(mx), 0, mx))))
        Case FT_DATUM, FT_DATUMZEIT
            If tp.Controls("notbefore").ListIndex > 0 Then
                props.SetVal "NichtVor", mCfg.FieldIdByName(mTid, tp.Controls("notbefore").Text)
            Else
                props.SetVal "NichtVor", ""
            End If
        Case FT_JANEIN
            props.SetVal "LeerErlaubt", IIf(tp.Controls("allowempty").Value, "Ja", "Nein")
        Case FT_LINK
            If f Is Nothing Then
                If tp.Controls("target").ListIndex < 0 Then lblErr.caption = ChrW(&H26A0) & "  Bitte eine Zieltabelle wählen.": Exit Sub
                tgt = mCfg.TableIdByName(tp.Controls("target").Text)
                props.SetVal "Zieltabelle", tgt
                props.SetVal "Mehrfach", IIf(tp.Controls("multi").Value, "Ja", "Nein")
                props.SetVal "Gegenname", Trim$(tp.Controls("backname").Text)
            End If
        Case FT_FORMEL
            If mXlMode Then
                mFormula = mXl.Definition(dmsg)
                If Len(dmsg) > 0 Then lblErr.caption = ChrW(&H26A0) & "  Formel: " & dmsg: Exit Sub
            End If
            If Len(mFormula) = 0 Then lblErr.caption = ChrW(&H26A0) & "  Bitte zuerst die Berechnung festlegen.": Exit Sub
            props.SetVal "Formel", mFormula
            props.SetVal "Beschreibung", mCalc.Describe(mFormula)
            props.SetVal "Ergebnistyp", FieldTypeFromName(tp.Controls("rt").Text)
            props.SetVal "Dezimalstellen", CLng(Val(tp.Controls("fdec").Text))
    End Select
    End If
    ' Höchstzahl der Auswahl (Mehrfachauswahl; im Feld "Maximum" gespeichert)
    If mType = FT_MEHRFACH Then
        mUI.MarkError txtMaxPick, False
        If Len(Trim$(txtMaxPick.Text)) = 0 Then
            props.SetVal "Maximum", Empty
        ElseIf Not IsNumeric(Trim$(txtMaxPick.Text)) Then
            ShowAdvError "Höchstens auswählbar: bitte eine ganze Zahl ab 1 eingeben (oder leer lassen).": Exit Sub
        ElseIf CDbl(Trim$(txtMaxPick.Text)) < 1 Or CDbl(Trim$(txtMaxPick.Text)) <> Int(CDbl(Trim$(txtMaxPick.Text))) Then
            ShowAdvError "Höchstens auswählbar: bitte eine ganze Zahl ab 1 eingeben (oder leer lassen).": Exit Sub
        Else
            props.SetVal "Maximum", NumInv(CDbl(Trim$(txtMaxPick.Text)))
        End If
    End If
    ' Standardwert pruefen
    If txtDefault.Enabled And Len(Trim$(txtDefault.Text)) > 0 Then
        tmpF.ID = mFid: tmpF.FType = mType: tmpF.Decimals = 6: tmpF.TableID = mTid
        Select Case mType
            Case FT_DATUM, FT_DATUMZEIT
                If LCase$(Left$(Trim$(txtDefault.Text), 5)) = "heute" Then
                    props.SetVal "Standard", LCase$(Replace(Trim$(txtDefault.Text), " ", ""))
                ElseIf mValid.ParseInput(tmpF, txtDefault.Text, parsed, dmsg) Then
                    props.SetVal "Standard", DateInv(CDate(parsed))
                Else
                    lblErr.caption = ChrW(&H26A0) & "  Standardwert: " & dmsg: Exit Sub
                End If
            Case FT_ZAHL, FT_GELD, FT_PROZENT
                If mValid.ParseInput(tmpF, txtDefault.Text, parsed, dmsg) Then props.SetVal "Standard", NumInv(CDbl(parsed)) Else lblErr.caption = ChrW(&H26A0) & "  Standardwert: " & dmsg: Exit Sub
            Case FT_JANEIN
                If mValid.ParseInput(tmpF, txtDefault.Text, parsed, dmsg) Then props.SetVal "Standard", IIf(parsed, "Ja", "Nein") Else lblErr.caption = ChrW(&H26A0) & "  Standardwert: " & dmsg: Exit Sub
            Case FT_AUSWAHL, FT_STATUS
                If f Is Nothing Then
                    props.SetVal "StandardName", Trim$(txtDefault.Text)
                ElseIf Len(mCfg.OptIdByName(mFid, txtDefault.Text)) = 0 Then
                    lblErr.caption = ChrW(&H26A0) & "  Standardwert: keine Auswahlmöglichkeit dieses Namens.": Exit Sub
                Else
                    props.SetVal "Standard", mCfg.OptIdByName(mFid, txtDefault.Text)
                End If
            Case Else
                props.SetVal "Standard", Trim$(txtDefault.Text)
        End Select
    Else
        props.SetVal "Standard", ""
    End If
    If mKind = FT_AUTOID Then
        If f Is Nothing Then props.SetVal "Standard", Trim$(tp.Controls("prefix").Text) Else props.SetVal "Standard", f.DefaultVal
    End If

    mApp.BeginWork
    On Error GoTo failed
    If f Is Nothing Then
        ' Optionen
        If mType = FT_AUSWAHL Or mType = FT_STATUS Or mType = FT_MEHRFACH Then
            lines = Split(Replace(tp.Controls("opts").Text, vbCr, ""), vbLf)
            For i = 0 To UBound(lines)
                ln = Trim$(lines(i))
                done = (Right$(ln, 1) = "*")
                If done Then ln = Trim$(Left$(ln, Len(ln) - 1))
                If Len(ln) > 0 Then
                    If InStr(ln, ";") > 0 Or InStr(ln, ",") > 0 Then mApp.EndWork: lblErr.caption = ChrW(&H26A0) & "  Optionsnamen dürfen kein ; oder , enthalten: " & ln: Exit Sub
                    props.SetVal "Optionen", NzS(props.GetVal("Optionen", "")) & IIf(Len(NzS(props.GetVal("Optionen", ""))) > 0, ",", "") & ln & ":" & IIf(done, ":x", "")
                End If
            Next
        End If
        fid = mStruct.AddField(mTid, nm, mKind, props)
        If props.Exists("StandardName") Then
            mCfg.SetV SH_FLD, fid, "Standard", mCfg.OptIdByName(fid, CStr(props.GetVal("StandardName")))
        End If
        If mType = FT_FORMEL Then mCalc.RecalcAll
    Else
        fid = mFid
        ' Eindeutigkeit neu verlangt: vorhandene Dubletten pruefen
        If props.GetVal("Eindeutig") = "Ja" And Not f.Unique Then
            msg = DuplicateReport(fid)
            If Len(msg) > 0 Then
                mApp.EndWork
                lblErr.caption = ChrW(&H26A0) & "  Eindeutigkeit nicht möglich: Es gibt bereits doppelte Werte: " & msg
                Exit Sub
            End If
        End If
        mStruct.UpdateField fid, props
        If mType = FT_FORMEL Then
            If f.Formula <> mFormula Or f.ResultType <> NzS(props.GetVal("Ergebnistyp", "")) Then mCalc.RecalcAll
        End If
    End If
    ' Dublettenpruefung
    If chkDup.Enabled Then
        If chkDup.Value Then
            mCfg.SetV SH_TAB, mTid, "Dublettenfelder", ListAdd(mCfg.GetS(SH_TAB, mTid, "Dublettenfelder"), fid)
        Else
            mCfg.SetV SH_TAB, mTid, "Dublettenfelder", ListRemove(mCfg.GetS(SH_TAB, mTid, "Dublettenfelder"), fid)
        End If
    End If
    mApp.EndWork
    ' Pflichtfeld in gefuellter Tabelle: Behandlung bestehender leerer Werte
    Dim wasReq As Boolean
    If Not f Is Nothing Then wasReq = f.Required
    If chkReq.Value And Not wasReq And mType <> FT_FORMEL Then HandleRequiredEmpties fid
    ResultFid = fid
    Me.Hide
    Exit Sub
failed:
    msg = Err.Description
    mApp.ResetWork
    lblErr.caption = ChrW(&H26A0) & "  " & msg
End Sub

Private Function DuplicateReport(ByVal fid As String) As String
    Dim t As cTableData, c As Long, r As Long, seen As New cMap, k As String, n As Long
    Set t = mData.TData(mTid)
    c = t.ColOf(fid)
    For r = 1 To t.RowCount
        k = NormKey(t.Cell(r, c))
        If Len(k) > 0 Then
            If seen.Exists(k) Then
                n = n + 1
                If n <= 3 Then DuplicateReport = DuplicateReport & IIf(n > 1, ", ", "") & "„" & mFmt.DisplayText(mCfg.Field(fid), t.Cell(r, c)) & "“"
            Else
                seen.SetVal k, True
            End If
        End If
    Next
    If n > 3 Then DuplicateReport = DuplicateReport & " und " & (n - 3) & " weitere"
End Function

' Bei neuem Pflichtfeld in gefuellter Tabelle: Standardwert eintragen oder Nachpflege
Private Sub HandleRequiredEmpties(ByVal fid As String)
    Dim t As cTableData, c As Long, r As Long, n As Long, ans As Long, f As cField, defV As Variant, vals() As Variant
    Set f = mCfg.Field(fid)
    Set t = mData.TData(mTid)
    c = t.ColOf(fid)
    If c = 0 Then Exit Sub
    For r = 1 To t.RowCount
        If f.IsLinkType Then
            If Len(mData.LinkTargets(fid, t.IdAt(r))) = 0 Then n = n + 1
        ElseIf IsBlank(t.Cell(r, c)) Then
            n = n + 1
        End If
    Next
    If n = 0 Then Exit Sub
    defV = mValid.DefaultValue(f)
    If IsBlank(defV) Or f.IsLinkType Then
        mUI.Alert Plural(n, "bestehender Eintrag hat", "bestehende Einträge haben") & " in „" & f.Name & "“ noch keinen Wert. " & _
                  "Die Pflicht gilt für diese Einträge ab der nächsten Bearbeitung (Nachpflege). Tipp: In der Tabellenansicht nach „" & f.Name & " ist leer“ filtern.", "Pflichtfeld ergänzt"
        Exit Sub
    End If
    ans = mUI.Ask3(Plural(n, "bestehender Eintrag hat", "bestehende Einträge haben") & " in „" & f.Name & "“ noch keinen Wert. Wie sollen diese behandelt werden?", _
                   "Pflichtfeld ergänzt", "Standardwert eintragen (" & mFmt.DisplayText(f, defV) & ")", "Leer lassen, später nachpflegen", "Abbrechen")
    If ans <> 1 Then
        If ans = 2 Then mUI.Alert "Die Pflicht gilt für diese Einträge ab der nächsten Bearbeitung.", "Pflichtfeld ergänzt"
        Exit Sub
    End If
    ReDim vals(1 To t.RowCount)
    For r = 1 To t.RowCount
        If IsBlank(t.Cell(r, c)) Then vals(r) = defV Else vals(r) = t.Cell(r, c)
    Next
    mData.WriteColumnRaw mTid, fid, vals
    mData.ClearUndo
    mLog.Add "Standardwert eingetragen", mTid, fid, fid, n & " Einträge"
    mCalc.RecalcAll
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True: Me.Hide
End Sub

Private Sub UserForm_Terminate()
    Dim e As cEvt
    If Not mEvts Is Nothing Then
        For Each e In mEvts: e.Unbind: Next
    End If
    If Not mXl Is Nothing Then mXl.Teardown: Set mXl = Nothing
End Sub
