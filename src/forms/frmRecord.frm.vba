'==============================================================================
' frmRecord - Dynamische Detailmaske je Tabelle. Felder, Register (Gruppen),
' Pflichtkennzeichen, Hilfetexte, Fehler am Feld, Verknuepfungen mit Auswahl,
' Schnellanlage und Wechsel zu verknuepften Eintraegen (Entwurf bleibt erhalten).
' Speichern prueft alle Register, speichert als Transaktion und schliesst.
'==============================================================================
Option Explicit

Public SavedRid As String
Public PreviewOnly As Boolean
Private mTid As String
Private mRid As String
Private mIsNew As Boolean
Private mPresets As cMap
Private mFields As Collection
Private mCtl As cMap            ' FeldID -> Eingabe-Steuerelement
Private mErrLbl As cMap         ' FeldID -> Fehlerbeschriftung
Private mPageOf As cMap         ' FeldID -> Registerindex
Private mLinks As cMap          ' FeldID -> "R1;R2"
Private mMulti As cMap          ' FeldID -> "O1;O2"
Private mMultiIds As cMap       ' FeldID -> Array der sichtbaren Options-IDs in der Liste
Private mChip As cMap           ' FeldID -> Farbmarke
Private mEvts As Collection
Private mLoading As Boolean
Private mDirty As Boolean
Private mp As MSForms.MultiPage
Private lblFoot As MSForms.Label
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnSaveNew As MSForms.CommandButton
Private WithEvents btnDup As MSForms.CommandButton
Private WithEvents btnArch As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private WithEvents btnMore As MSForms.CommandButton
Private mFW As Single
Private mSaveAndNew As Boolean

' Register fuer Rueckverknuepfungen (Gegenrichtung: andere Tabellen verweisen hierher)
Private Const BACK_GROUP As String = "Verknüpft aus anderen Tabellen"

'==============================================================================
' Aufbau
'==============================================================================
Public Function OpenFor(ByVal tid As String, ByVal rid As String, Optional ByVal presets As cMap = Nothing) As String
    mTid = tid
    mRid = rid
    mIsNew = (Len(rid) = 0)
    If presets Is Nothing Then Set mPresets = New cMap Else Set mPresets = presets
    BuildAll
    Me.Show
    OpenFor = SavedRid
End Function

' Nur fuer Laufzeitmessung (mSelfTest): Maske aufbauen, ohne sie anzuzeigen
Public Sub PrepareOnly(ByVal tid As String, ByVal rid As String)
    mTid = tid
    mRid = rid
    mIsNew = (Len(rid) = 0)
    Set mPresets = New cMap
    BuildAll
End Sub

Private Sub BuildAll()
    Dim f As cField, groups As New Collection, gname As String, i As Long, pg As MSForms.Page, y() As Single, gi As Long
    Dim title As String, meta As String, w As Single, h As Single, maxH As Single, contentMax As Single, e As cEvt
    mLoading = True
    Set mCtl = New cMap: Set mErrLbl = New cMap: Set mPageOf = New cMap
    Set mLinks = New cMap: Set mMulti = New cMap: Set mMultiIds = New cMap: Set mChip = New cMap
    If Not mEvts Is Nothing Then
        For Each e In mEvts: e.Unbind: Next
    End If
    Set mEvts = New Collection
    Set btnSaveNew = Nothing: Set btnMore = Nothing: Set btnDup = Nothing: Set btnArch = Nothing
    Me.Controls.Clear
    mUI.StyleForm Me
    w = 700
    If w > mUI.MaxFormWidth Then w = mUI.MaxFormWidth
    Me.Width = w
    Set mFields = mCfg.Fields(mTid)

    ' Kopf (knapp; Details stehen im Register "Verlauf")
    If mIsNew Then
        title = "Neuer Eintrag"
        meta = mCfg.TableName(mTid)
    Else
        title = mData.RecordLabel(mTid, mRid)
        meta = mCfg.TableName(mTid) & "  ·  zuletzt geändert " & FmtStamp(mData.GetValue(mTid, mRid, C_CHANGED))
        For Each f In mFields
            If f.Kind = FT_AUTOID Then meta = mFmt.DisplayText(f, mData.GetValue(mTid, mRid, f.ID), mRid) & "  ·  " & meta: Exit For
        Next
    End If
    Me.caption = mCfg.TableName(mTid) & IIf(mIsNew, ": neuer Eintrag", ": Eintrag bearbeiten")
    mUI.AddIcon Me, "hIcon", mCfg.TableIcon(mTid), PAD - 4, PAD - 3, 16
    mUI.AddLabel(Me, "hTitle", Truncate(title, 70), PAD + 24, PAD - 2, Me.InsideWidth - PAD * 2 - 30, 22, True, 13).WordWrap = False
    mUI.AddLabel Me, "hMeta", meta, PAD + 24, PAD + 21, Me.InsideWidth - PAD * 2 - 30, 14, False, 8.5, CMuted()
    If Not mIsNew Then
        If mData.IsArchivedRec(mTid, mRid) Then
            mUI.AddChip Me, "hArch", "Archiviert am " & FmtDate(mData.GetValue(mTid, mRid, C_ARCH)) & "; wiederherstellen über die Schaltfläche unten", "orange", PAD + 24, PAD + 38, 420
        End If
    End If

    ' Gruppen bestimmen (nur aus sichtbaren Eigenschaften; Berechnetes bleibt ausgeblendet)
    For Each f In mFields
        If Not ShowInForm(f) Then GoTo nextG
        gname = f.Group
        If Len(gname) = 0 Then gname = IIf(f.FType = FT_BACKLINK, BACK_GROUP, "Angaben")
        If Not CollHas(groups, gname) Then groups.Add gname, gname
nextG:
    Next
    If groups.Count = 0 Then groups.Add "Angaben", "Angaben"
    If Not mIsNew Then groups.Add "Verlauf", "Verlauf"

    Set mp = Me.Controls.Add("Forms.MultiPage.1", "mp")
    mp.Left = PAD: mp.Top = PAD + 58
    mp.Width = Me.InsideWidth - PAD * 2
    mp.Font.Name = UI_FONT: mp.Font.Size = FONT_SIZE
    mp.BackColor = CWhite()
    mp.Style = fmTabStyleTabs
    Do While mp.Pages.Count > 0: mp.Pages.Remove 0: Loop
    For i = 1 To groups.Count
        Set pg = mp.Pages.Add("pg" & i, groups(i))
    Next
    mFW = mp.Width - 26
    ReDim y(0 To groups.Count - 1)
    For i = 0 To groups.Count - 1: y(i) = 12: Next

    For Each f In mFields
        If ShowInForm(f) Then
            gname = f.Group
            If Len(gname) = 0 Then gname = IIf(f.FType = FT_BACKLINK, BACK_GROUP, "Angaben")
            gi = CollIndex(groups, gname) - 1
            mPageOf.SetVal f.ID, gi
            If gname = BACK_GROUP And y(gi) = 12 Then
                mUI.AddLabel mp.Pages(gi), "backInfo", "Hier sehen Sie Einträge aus anderen Tabellen, die diesen Eintrag ausgewählt haben, also die Gegenrichtung einer Verknüpfung. " & _
                             "Beispiel: Ist eine Aufgabe mit diesem Projekt verknüpft, erscheint sie hier beim Projekt. Hinzufügen oder Entfernen ändert die Verknüpfung auf beiden Seiten.", _
                             12, 10, mFW - 24, 40, False, 8.5, CMuted()
                y(gi) = 56
            End If
            y(gi) = AddFieldRow(mp.Pages(gi), f, y(gi))
        End If
    Next
    ' nur ein Register: ohne Registerleiste wirkt die Maske ruhiger
    If groups.Count = 1 Then mp.Style = fmTabStyleNone
    If Not mIsNew Then
        gi = groups.Count - 1
        y(gi) = AddHistory(mp.Pages(gi), y(gi))
    End If

    ' Hoehen
    For i = 0 To groups.Count - 1
        If y(i) > contentMax Then contentMax = y(i)
    Next
    maxH = mUI.MaxFormHeight() - (PAD + 58) - 110
    h = contentMax + 34
    If h > maxH Then h = maxH
    If h < 180 Then h = 180
    mp.Height = h
    For i = 0 To groups.Count - 1
        mp.Pages(i).ScrollBars = fmScrollBarsNone
        If y(i) + 10 > h - 26 Then
            mp.Pages(i).ScrollBars = fmScrollBarsVertical
            mp.Pages(i).ScrollHeight = y(i) + 16
            mp.Pages(i).KeepScrollBarsVisible = fmScrollBarsNone
        End If
    Next

    ' Fuss
    Dim fy As Single, bx As Single, b As MSForms.CommandButton
    fy = mp.Top + mp.Height + 8
    Set lblFoot = mUI.AddLabel(Me, "foot", IIf(HasRequired(), "* Pflichtangabe", ""), PAD, fy, Me.InsideWidth - PAD * 2, 15, False, 8.5, CMuted())
    fy = fy + 24
    Set btnSave = mUI.AddButton(Me, "bSave", "Speichern", PAD, fy, 110, "primary")
    bx = PAD + 118
    If mIsNew Then
        Set btnSaveNew = mUI.AddButton(Me, "bSaveNew", "Speichern & neu", bx, fy, 130)
    Else
        ' Seltenes erst auf Klick ("Mehr")
        Set btnMore = mUI.AddButton(Me, "bMore", "Mehr  " & ChrW(&H25BE), bx, fy, 84)
        bx = bx + 92
        Set btnDup = mUI.AddButton(Me, "bDup", "Duplizieren", bx, fy, 100)
        btnDup.Visible = False
        If mData.IsArchivedRec(mTid, mRid) Then
            Set btnArch = mUI.AddButton(Me, "bArch", "Wiederherstellen", bx + 108, fy, 120)
        Else
            Set btnArch = mUI.AddButton(Me, "bArch", "Archivieren", bx + 108, fy, 104)
            btnArch.Visible = False
        End If
    End If
    Set btnCancel = mUI.AddButton(Me, "bCancel", "Abbrechen", Me.InsideWidth - PAD - 104, fy, 104)
    btnCancel.Cancel = True
    If PreviewOnly Then
        btnSave.Enabled = False
        If Not btnSaveNew Is Nothing Then btnSaveNew.Enabled = False
        lblFoot.caption = "Vorschau der Eingabemaske: So sehen Nutzerinnen und Nutzer das Formular. Speichern ist in der Vorschau deaktiviert."
        lblFoot.ForeColor = CAccent()
        Me.caption = "Vorschau: " & Me.caption
    End If
    If mApp.gReadOnly Then
        btnSave.Enabled = False
        If Not btnSaveNew Is Nothing Then btnSaveNew.Enabled = False
        If Not btnDup Is Nothing Then btnDup.Enabled = False
        If Not btnArch Is Nothing Then btnArch.Enabled = False
        lblFoot.caption = "Lesemodus: Die Datei ist schreibgeschützt geöffnet. Änderungen können nicht gespeichert werden."
        lblFoot.ForeColor = CWarn()
    End If
    Me.Height = fy + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    mp.Value = 0
    mLoading = False
    mDirty = False
End Sub

' Automatisch Berechnetes (Formel, Rollup, Erstellt/Geaendert, ID) zeigt die Maske nicht;
' es steht in der Tabelle bzw. im Kopf und im Verlauf.
Private Function ShowInForm(ByVal f As cField) As Boolean
    If f.FType = FT_FORMEL Then Exit Function
    If IsAutoKind(f.Kind) And f.Kind <> FT_BUTTON Then Exit Function
    If f.Kind = FT_BUTTON And mIsNew Then Exit Function
    ShowInForm = True
End Function

Private Function CollHas(ByVal c As Collection, ByVal k As String) As Boolean
    Dim v As Variant
    On Error GoTo nope
    v = c(k)
    CollHas = True
nope:
End Function

Private Function CollIndex(ByVal c As Collection, ByVal k As String) As Long
    Dim i As Long
    For i = 1 To c.Count
        If c(i) = k Then CollIndex = i: Exit Function
    Next
End Function

Private Sub Bind(ByVal ctl As Object, ByVal key As String)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, ctl, key
    mEvts.Add e
End Sub

' Erzeugt Beschriftung, Eingabe und Fehlerzeile eines Feldes; liefert neue y-Position
Private Function AddFieldRow(ByVal pg As Object, ByVal f As cField, ByVal y As Single) As Single
    Dim lbl As MSForms.Label, cap As String, v As Variant, tb As MSForms.TextBox, cb As MSForms.ComboBox, lb As MSForms.ListBox
    Dim ck As MSForms.CheckBox, b As MSForms.CommandButton, x As Single, oid As Variant, i As Long, h As Single, ids() As String
    Dim lst As String, a As Variant, n As Long, hint As String, lblE As MSForms.Label, cur As String
    cap = f.Name & IIf(f.Required, " *", "")
    Set lbl = mUI.AddLabel(pg, "l_" & f.ID, cap, 12, y, mFW - 24, 15, True, 9.5)
    y = y + 17
    If Len(f.Help) > 0 Or f.FType = FT_BACKLINK Then
        If Len(f.Help) > 0 Then
            hint = f.Help
        Else
            hint = "Einträge aus „" & mCfg.TableName(mCfg.Field(f.PairField).TableID) & "“, die im Feld „" & mCfg.FieldName(f.PairField) & "“ diesen Eintrag ausgewählt haben"
        End If
        Set lbl = mUI.AddLabel(pg, "h_" & f.ID, hint, 12, y, mFW - 24, 13, False, 8.5, CMuted())
        hint = ""
        lbl.AutoSize = True
        If lbl.Width > mFW - 24 Then lbl.Width = mFW - 24
        y = y + lbl.Height + 3
    End If
    v = CurrentValue(f)
    ' Sonderarten
    If f.Kind = FT_BUTTON Then
        Set b = mUI.AddButton(pg, "bb_" & f.ID, ChrW(&H25B6) & "  " & f.Name, 12, y, 200, IIf(mIsNew Or mApp.gReadOnly Or PreviewOnly, "secondary", "primary"), ROW_H + 2)
        Bind b, "bb:" & f.ID
        b.Enabled = Not mIsNew And Not mApp.gReadOnly And Not PreviewOnly
        mUI.AddLabel pg, "u_" & f.ID, IIf(mIsNew, "Verfügbar nach dem ersten Speichern. ", "") & "Setzt: " & mForms.ButtonDescription(f), 222, y + 4, mFW - 240, 28, False, 8.5, CMuted()
        AddFieldRow = y + ROW_H + 30
        Exit Function
    ElseIf IsAutoKind(f.Kind) Then
        Set tb = mUI.AddText(pg, "in_" & f.ID, 12, y, 270)
        If mIsNew Then tb.Text = "(wird automatisch gesetzt)" Else tb.Text = mFmt.DisplayText(f, v, mRid)
        tb.Locked = True: tb.BackColor = CSurface(): tb.ForeColor = CMuted(): tb.TabStop = False
        mCtl.SetVal f.ID, tb
        AddFieldRow = y + ROW_H + 18
        Exit Function
    ElseIf f.Kind = FT_DATEIEN Then
        Set tb = mUI.AddText(pg, "in_" & f.ID, 12, y, 10)
        tb.Visible = False
        tb.Text = NzS(v)
        mCtl.SetVal f.ID, tb
        Set lb = mUI.AddList(pg, "fl_" & f.ID, 12, y, mFW - 24 - 128, 15 * 4 + 8)
        Bind lb, "fl:" & f.ID
        FillFileList f.ID
        x = mFW - 12 - 120
        Set b = mUI.AddButton(pg, "fa_" & f.ID, "Hinzufügen…", x, y, 120, , ROW_H): Bind b, "fa:" & f.ID
        Set b = mUI.AddButton(pg, "fo_" & f.ID, "Öffnen", x, y + ROW_H + 4, 120, , ROW_H): Bind b, "fo:" & f.ID
        Set b = mUI.AddButton(pg, "fr_" & f.ID, "Entfernen", x, y + (ROW_H + 4) * 2, 120, , ROW_H): Bind b, "fr:" & f.ID
        Set lblE = mUI.AddLabel(pg, "e_" & f.ID, "", 12, y + 70, mFW - 24, 14, False, 8.5, CDanger())
        lblE.Visible = False
        mErrLbl.SetVal f.ID, lblE
        AddFieldRow = y + 96
        Exit Function
    End If
    Select Case f.FType
        Case FT_TEXT, FT_PFAD
            Set tb = mUI.AddText(pg, "in_" & f.ID, 12, y, IIf(f.FType = FT_PFAD, mFW - 24 - 230, mFW - 24))
            tb.Text = mFmt.InputText(f, v)
            If f.MaxLen > 0 Then tb.MaxLength = f.MaxLen
            tb.ControlTipText = f.Help
            mCtl.SetVal f.ID, tb
            Bind tb, "in:" & f.ID
            If f.Kind = FT_URL Or f.Kind = FT_EMAIL Then
                tb.Width = mFW - 24 - 128
                Set b = mUI.AddButton(pg, "bu_" & f.ID, IIf(f.Kind = FT_URL, "Im Browser öffnen", "E-Mail schreiben"), mFW - 12 - 120, y, 120, , ROW_H): Bind b, "bu:" & f.ID
            End If
            If f.FType = FT_PFAD Then
                x = 12 + tb.Width + 6
                Set b = mUI.AddButton(pg, "bf_" & f.ID, "Datei…", x, y, 70, , ROW_H): Bind b, "bf:" & f.ID
                Set b = mUI.AddButton(pg, "bo_" & f.ID, "Ordner…", x + 76, y, 74, , ROW_H): Bind b, "bd:" & f.ID
                Set b = mUI.AddButton(pg, "bp_" & f.ID, "Öffnen", x + 156, y, 68, , ROW_H): Bind b, "po:" & f.ID
            End If
            y = y + ROW_H
        Case FT_LTEXT
            Set tb = mUI.AddText(pg, "in_" & f.ID, 12, y, mFW - 24, 70, True)
            tb.Text = NzS(v)
            mCtl.SetVal f.ID, tb
            Bind tb, "in:" & f.ID
            y = y + 70
        Case FT_ZAHL, FT_GELD, FT_PROZENT
            Set tb = mUI.AddText(pg, "in_" & f.ID, 12, y, 170)
            tb.Text = mFmt.InputText(f, v)
            tb.ControlTipText = mFmt.InputHint(f)
            mCtl.SetVal f.ID, tb
            Bind tb, "in:" & f.ID
            ' nur die Einheit neben dem Feld; Details stehen im Tooltip
            If f.FType = FT_GELD Then hint = "€"
            If f.FType = FT_PROZENT Then hint = "%"
            If Len(hint) > 0 Then mUI.AddLabel pg, "u_" & f.ID, hint, 188, y + 4, 30, 15, False, 9.5, CMuted()
            y = y + ROW_H
        Case FT_DATUM, FT_DATUMZEIT
            Set tb = mUI.AddText(pg, "in_" & f.ID, 12, y, IIf(f.FType = FT_DATUMZEIT, 170, 130))
            tb.Text = mFmt.InputText(f, v)
            tb.ControlTipText = mFmt.InputHint(f)
            mCtl.SetVal f.ID, tb
            Bind tb, "in:" & f.ID
            x = 12 + tb.Width + 6
            Set b = mUI.AddButton(pg, "bc_" & f.ID, "Datum…", x, y, 76, , ROW_H): Bind b, "dt:" & f.ID
            Set b = mUI.AddButton(pg, "bt_" & f.ID, "Heute", x + 82, y, 76, , ROW_H): Bind b, "td:" & f.ID
            y = y + ROW_H
        Case FT_JANEIN
            If f.AllowEmpty Then
                Set cb = mUI.AddCombo(pg, "in_" & f.ID, 12, y, 140)
                cb.AddItem "(keine Angabe)": cb.AddItem "Ja": cb.AddItem "Nein"
                If IsBlank(v) Then cb.ListIndex = 0 Else cb.ListIndex = IIf(CBool(v), 1, 2)
                mCtl.SetVal f.ID, cb
                Bind cb, "in:" & f.ID
            Else
                Set ck = mUI.AddCheck(pg, "in_" & f.ID, "Ja", 12, y, 200)
                If Not IsBlank(v) Then ck.Value = CBool(v)
                mCtl.SetVal f.ID, ck
                Bind ck, "in:" & f.ID
            End If
            y = y + ROW_H
        Case FT_AUSWAHL, FT_STATUS
            Set cb = mUI.AddCombo(pg, "in_" & f.ID, 12, y, 270, False)
            cur = NzS(v)
            FillOptionCombo cb, f, cur
            mCtl.SetVal f.ID, cb
            Bind cb, "in:" & f.ID
            mChip.SetVal f.ID, mUI.AddChip(pg, "chip_" & f.ID, "", "grau", 290, y + 2, 120)
            UpdateChip f.ID
            Set b = mUI.AddButton(pg, "bn_" & f.ID, "+", 418, y, ROW_H + 2, , ROW_H): Bind b, "nw:" & f.ID
            b.ControlTipText = IIf(f.Kind = FT_PERSON, "Neue Person anlegen", "Neue Auswahlmöglichkeit anlegen")
            y = y + ROW_H
        Case FT_MEHRFACH
            n = mCfg.OptionIds(f.ID).Count
            If n > 8 Then
                Set tb = mUI.AddText(pg, "ms_" & f.ID, 12, y, 270)
                tb.ControlTipText = "Optionen durchsuchen"
                Bind tb, "ms:" & f.ID
                mUI.AddLabel pg, "msl_" & f.ID, "durchsuchen", 290, y + 4, 200, 15, False, 8.5, CMuted()
                y = y + ROW_H + 4
            End If
            h = Application.Min(Application.Max(n, 2), 8) * 15 + 8
            Set lb = mUI.AddList(pg, "in_" & f.ID, 12, y, 270, h, True)
            mCtl.SetVal f.ID, lb
            mMulti.SetVal f.ID, NzS(v)
            FillMultiList f.ID, ""
            Bind lb, "ml:" & f.ID
            Set b = mUI.AddButton(pg, "bn_" & f.ID, "+", 290, y, ROW_H + 2, , ROW_H): Bind b, "nw:" & f.ID
            b.ControlTipText = IIf(f.Kind = FT_PERSON, "Neue Person anlegen", "Neue Auswahlmöglichkeit anlegen")
            mUI.AddLabel pg, "sel_" & f.ID, "", 290, y + ROW_H + 6, mFW - 310, h - ROW_H - 6, False, 8.5, CMuted()
            UpdateMultiSummary f.ID
            y = y + h
        Case FT_LINK, FT_BACKLINK
            If Not mIsNew Then lst = mData.LinkTargets(f.ID, mRid) Else lst = NzS(mPresets.GetVal(f.ID, ""))
            mLinks.SetVal f.ID, lst
            If f.FType = FT_LINK And Not f.Multi Then h = ROW_H Else h = 15 * 5 + 8
            Set lb = mUI.AddList(pg, "in_" & f.ID, 12, y, mFW - 24 - 128, h, False, 2, (mFW - 24 - 132) * 0.55 & ";" & (mFW - 24 - 132) * 0.45)
            mCtl.SetVal f.ID, lb
            Bind lb, "ll:" & f.ID
            FillLinkList f.ID
            x = mFW - 12 - 120
            Set b = mUI.AddButton(pg, "bk_" & f.ID, "Auswählen…", x, y, 120, , ROW_H): Bind b, "pk:" & f.ID
            If h > ROW_H Then
                Set b = mUI.AddButton(pg, "bl_" & f.ID, "Öffnen", x, y + ROW_H + 4, 120, , ROW_H): Bind b, "ol:" & f.ID
                Set b = mUI.AddButton(pg, "br_" & f.ID, "Entfernen", x, y + (ROW_H + 4) * 2, 120, , ROW_H): Bind b, "rl:" & f.ID
            Else
                ' einzeilig: Oeffnen/Entfernen in eigener Zeile
                y = y + ROW_H + 4
                Set b = mUI.AddButton(pg, "bl_" & f.ID, "Öffnen", 12, y, 90, , ROW_H): Bind b, "ol:" & f.ID
                Set b = mUI.AddButton(pg, "br_" & f.ID, "Entfernen", 108, y, 90, , ROW_H): Bind b, "rl:" & f.ID
                h = ROW_H
            End If
            y = y + h
        Case FT_FORMEL
            Set tb = mUI.AddText(pg, "in_" & f.ID, 12, y, 270)
            tb.Text = mFmt.DisplayText(f, v, mRid)
            tb.Locked = True
            tb.BackColor = CSurface()
            tb.ForeColor = IIf(IsErrVal(v), CDanger(), CMuted())
            tb.TabStop = False
            tb.ControlTipText = "Berechnet: " & f.FormulaDesc
            mCtl.SetVal f.ID, tb
            Set lbl = mUI.AddLabel(pg, "u_" & f.ID, "Berechnet: " & f.FormulaDesc, 290, y, mFW - 310, 30, False, 8, CMuted())
            If mIsNew Then tb.Text = "(wird beim Speichern berechnet)"
            y = y + Application.Max(ROW_H, 30)
    End Select
    Set lblE = mUI.AddLabel(pg, "e_" & f.ID, "", 12, y + 2, mFW - 24, 14, False, 8.5, CDanger())
    lblE.Visible = False
    mErrLbl.SetVal f.ID, lblE
    AddFieldRow = y + 24
End Function

Private Function CurrentValue(ByVal f As cField) As Variant
    If mPresets.Exists(f.ID) And Not f.IsLinkType Then
        CurrentValue = mPresets.GetVal(f.ID)
    ElseIf mIsNew Then
        CurrentValue = mValid.DefaultValue(f)
    Else
        CurrentValue = mData.GetValue(mTid, mRid, f.ID)
    End If
End Function

Private Function AddHistory(ByVal pg As Object, ByVal y As Single) As Single
    Dim t As cTableData, r As Long, items() As Variant, n As Long, lb As MSForms.ListBox, a As Variant, i As Long, flds As String
    Set t = mCfg.TD(SH_LOG)
    ReDim items(0 To 199, 0 To 3)
    For r = t.RowCount To 1 Step -1
        If CStr(t.Cell(r, 5)) = mRid Then
            items(n, 0) = FmtStamp(t.Cell(r, 1))
            items(n, 1) = NzS(t.Cell(r, 2))
            items(n, 2) = NzS(t.Cell(r, 3))
            flds = ""
            a = Split(NzS(t.Cell(r, 6)), ";")
            For i = 0 To UBound(a)
                If Len(a(i)) > 0 Then flds = flds & IIf(Len(flds) > 0, ", ", "") & mCfg.FieldName(CStr(a(i)))
            Next
            items(n, 3) = flds & IIf(Len(NzS(t.Cell(r, 7))) > 0, IIf(Len(flds) > 0, " · ", "") & NzS(t.Cell(r, 7)), "")
            n = n + 1
            If n >= 200 Then Exit For
        End If
    Next
    mUI.AddLabel pg, "histInfo", "Änderungsverlauf dieses Eintrags (neueste zuerst). Protokolliert werden Zeitpunkt, Windows-Kennung und betroffene Felder, keine Feldinhalte.", _
                 12, y, mFW - 24, 28, False, 8.5, CMuted()
    y = y + 32
    Set lb = mUI.AddList(pg, "histList", 12, y, mFW - 24, 220, False, 4, "95;70;80;" & (mFW - 24 - 260))
    If n > 0 Then
        Dim out() As Variant, j As Long
        ReDim out(0 To n - 1, 0 To 3)
        For i = 0 To n - 1
            For j = 0 To 3: out(i, j) = items(i, j): Next
        Next
        lb.List = out
    End If
    AddHistory = y + 230
End Function

'--- Auswahl-Hilfen ------------------------------------------------------------------------
Private Sub FillOptionCombo(ByVal cb As MSForms.ComboBox, ByVal f As cField, ByVal cur As String)
    Dim oid As Variant, i As Long
    cb.Clear
    cb.AddItem "(keine Angabe)"
    i = 0
    For Each oid In mCfg.OptionIds(f.ID)
        If mCfg.OptActive(CStr(oid)) Or CStr(oid) = cur Then
            cb.AddItem mCfg.OptName(CStr(oid)) & IIf(mCfg.OptActive(CStr(oid)), "", " (deaktiviert)")
            i = i + 1
            If CStr(oid) = cur Then cb.ListIndex = i
        End If
    Next
    If Len(cur) = 0 Then cb.ListIndex = 0
End Sub

Private Function ComboOptionId(ByVal f As cField, ByVal txt As String) As String
    txt = Trim$(Replace(txt, " (deaktiviert)", ""))
    If txt = "(keine Angabe)" Or Len(txt) = 0 Then Exit Function
    ComboOptionId = mCfg.OptIdByName(f.ID, txt)
End Function

Private Sub UpdateChip(ByVal fid As String)
    Dim cb As MSForms.ComboBox, oid As String, chip As MSForms.Label
    Set cb = mCtl.GetVal(fid)
    Set chip = mChip.GetVal(fid)
    oid = ComboOptionId(mCfg.Field(fid), cb.Text)
    If Len(oid) = 0 Then
        chip.Visible = False
    Else
        chip.Visible = True
        chip.caption = " " & mCfg.OptName(oid) & IIf(mCfg.OptDone(oid), "  " & ChrW(&H2713), "")
        chip.BackColor = PaletteBg(mCfg.OptColor(oid))
        chip.ForeColor = PaletteFg(mCfg.OptColor(oid))
    End If
End Sub

Private Sub FillMultiList(ByVal fid As String, ByVal filterText As String)
    Dim lb As MSForms.ListBox, oid As Variant, sel As String, ids() As String, n As Long, i As Long
    Set lb = mCtl.GetVal(fid)
    sel = mMulti.GetVal(fid, "")
    mLoading = True
    lb.Clear
    ReDim ids(0 To mCfg.OptionIds(fid).Count)
    For Each oid In mCfg.OptionIds(fid)
        If mCfg.OptActive(CStr(oid)) Or ListHas(sel, CStr(oid)) Then
            If Len(filterText) = 0 Or InStr(1, mCfg.OptName(CStr(oid)), filterText, vbTextCompare) > 0 Then
                lb.AddItem mCfg.OptName(CStr(oid)) & IIf(mCfg.OptActive(CStr(oid)), "", " (deaktiviert)")
                ids(n) = CStr(oid)
                If ListHas(sel, CStr(oid)) Then lb.Selected(n) = True
                n = n + 1
            End If
        End If
    Next
    mMultiIds.SetVal fid, ids
    mLoading = False
End Sub

Private Sub UpdateMultiSummary(ByVal fid As String)
    Dim sel As String, a As Variant, i As Long, s As String
    sel = mMulti.GetVal(fid, "")
    a = Split(sel, ";")
    For i = 0 To UBound(a)
        s = s & IIf(Len(s) > 0, ", ", "") & mCfg.OptName(CStr(a(i)))
    Next
    If Len(s) = 0 Then s = "Nichts ausgewählt" Else s = "Ausgewählt (" & (UBound(a) + 1) & "): " & s
    If mValid.MaxPicks(mCfg.Field(fid)) > 0 Then
        s = s & vbLf & "Höchstens " & mValid.MaxPicks(mCfg.Field(fid)) & " auswählbar"
        Me.Controls("sel_" & fid).ForeColor = IIf(mValid.MultiCount(sel) > mValid.MaxPicks(mCfg.Field(fid)), CDanger(), CMuted())
    End If
    Me.Controls("sel_" & fid).caption = s
End Sub

Private Sub FillLinkList(ByVal fid As String)
    Dim lb As MSForms.ListBox, lst As String, a As Variant, i As Long, f As cField, tt As String, out() As Variant, n As Long
    Set f = mCfg.Field(fid)
    If f.FType = FT_LINK Then tt = f.TargetTable Else tt = mCfg.Field(f.PairField).TableID
    Set lb = mCtl.GetVal(fid)
    lst = mLinks.GetVal(fid, "")
    lb.Clear
    If Len(lst) = 0 Then
        lb.AddItem "(keine Verknüpfung)"
        lb.List(0, 1) = ""
        Exit Sub
    End If
    a = Split(lst, ";")
    ReDim out(0 To UBound(a), 0 To 1)
    For i = 0 To UBound(a)
        out(i, 0) = mData.RecordLabel(tt, CStr(a(i))) & IIf(mData.IsArchivedRec(tt, CStr(a(i))), " (archiviert)", "")
        out(i, 1) = mData.RecordHint(tt, CStr(a(i)))
    Next
    lb.List = out
End Sub

'==============================================================================
' Ereignisse
'==============================================================================
Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    Dim p As Long, kind As String, fid As String
    On Error GoTo failed
    p = InStr(key, ":")
    kind = Left$(key, p - 1)
    fid = Mid$(key, p + 1)
    Select Case kind
        Case "in"
            If evt = "Change" Or evt = "Click" Then
                If Not mLoading Then mDirty = True: ClearError fid
                If mCfg.Field(fid).IsSingleOption Then UpdateChip fid
            ElseIf evt = "KeyUp" Then
                If arg = vbKeyTab Or arg = vbKeyReturn Then EnsureVisible fid
            End If
        Case "ml"
            If evt = "Change" And Not mLoading Then SyncMulti fid: mDirty = True: ClearError fid
        Case "ms"
            If evt = "Change" Then FillMultiList fid, Me.Controls("ms_" & fid).Text
        Case "dt"
            If evt = "Click" Then PickDate fid
        Case "td"
            If evt = "Click" Then
                If mCfg.Field(fid).FType = FT_DATUMZEIT Then
                    mCtl.GetVal(fid).Text = FmtDateTime(Now)
                Else
                    mCtl.GetVal(fid).Text = FmtDate(Date)
                End If
            End If
        Case "nw"
            If evt = "Click" Then QuickAddOption fid
        Case "pk"
            If evt = "Click" Then PickLinks fid
        Case "ol"
            If evt = "Click" Then OpenLinked fid
        Case "ll"
            If evt = "DblClick" Then OpenLinked fid
        Case "rl"
            If evt = "Click" Then RemoveLinked fid
        Case "bb"
            If evt = "Click" Then RunButtonHere fid
        Case "bu"
            If evt = "Click" Then OpenLink fid
        Case "fa"
            If evt = "Click" Then AddFiles fid
        Case "fo"
            If evt = "Click" Then OpenFile fid
        Case "fl"
            If evt = "DblClick" Then OpenFile fid
        Case "fr"
            If evt = "Click" Then RemoveFile fid
        Case "bf", "bd"
            If evt = "Click" Then BrowsePath fid, (kind = "bd")
        Case "po"
            If evt = "Click" Then OpenPath fid
    End Select
    Exit Sub
failed:
    mApp.Boundary "Formular"
End Sub

Private Sub SyncMulti(ByVal fid As String)
    Dim lb As MSForms.ListBox, ids As Variant, i As Long, sel As String
    Set lb = mCtl.GetVal(fid)
    ids = mMultiIds.GetVal(fid)
    sel = mMulti.GetVal(fid, "")
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then sel = ListAdd(sel, CStr(ids(i))) Else sel = ListRemove(sel, CStr(ids(i)))
    Next
    mMulti.SetVal fid, sel
    UpdateMultiSummary fid
End Sub

Private Sub EnsureVisible(ByVal fid As String)
    Dim c As Object, pg As Object, gi As Long
    On Error Resume Next
    gi = mPageOf.GetVal(fid, 0)
    Set pg = mp.Pages(gi)
    Set c = mCtl.GetVal(fid)
    If c.Top + c.Height > pg.ScrollTop + mp.Height - 40 Then pg.ScrollTop = c.Top - 40
    If c.Top < pg.ScrollTop Then pg.ScrollTop = c.Top - 40
End Sub

Private Sub ClearError(ByVal fid As String)
    Dim lbl As MSForms.Label
    If Not mErrLbl.Exists(fid) Then Exit Sub
    Set lbl = mErrLbl.GetVal(fid)
    If lbl.Visible Then
        lbl.Visible = False
        MarkCtl fid, False
    End If
End Sub

' Mehrfachlisten verlieren beim Umfärben ihre Haken und zeichnen den Rahmen in ihrem eigenen
' Ereignis nicht neu: dort zeigt ein roter Rahmen dahinter den Fehler an
Private Sub MarkCtl(ByVal fid As String, ByVal isErr As Boolean)
    Dim lb As MSForms.ListBox, fr As MSForms.Label
    If Not mMultiIds.Exists(fid) Then
        mUI.MarkError mCtl.GetVal(fid), isErr
        Exit Sub
    End If
    Set lb = mCtl.GetVal(fid)
    On Error Resume Next
    Set fr = lb.Parent.Controls("ef_" & fid)
    On Error GoTo 0
    If fr Is Nothing Then
        Set fr = mUI.AddLabel(lb.Parent, "ef_" & fid, "", lb.Left - 2, lb.Top - 2, lb.Width + 4, lb.Height + 4)
        fr.BackStyle = fmBackStyleOpaque
        fr.BackColor = CDanger()
        fr.ZOrder fmZOrderBack
    End If
    fr.Visible = isErr
    RestoreMulti fid
End Sub

Private Sub ShowError(ByVal fid As String, ByVal msg As String)
    Dim lbl As MSForms.Label
    If Not mErrLbl.Exists(fid) Then
        ' Eigenschaft ist in der Maske ausgeblendet (z. B. berechnet): Meldung im Fuss
        lblFoot.caption = ChrW(&H26A0) & "  " & mCfg.FieldName(fid) & ": " & msg
        lblFoot.ForeColor = CDanger()
        Exit Sub
    End If
    Set lbl = mErrLbl.GetVal(fid)
    lbl.caption = ChrW(&H26A0) & "  " & msg
    lbl.Visible = True
    MarkCtl fid, True
End Sub

' Mehrfachliste verliert beim Umfärben ihre Haken: aus dem gemerkten Stand wieder setzen
Private Sub RestoreMulti(ByVal fid As String)
    Dim lb As MSForms.ListBox, ids As Variant, i As Long, sel As String, wasLoading As Boolean
    If Not mMultiIds.Exists(fid) Then Exit Sub
    Set lb = mCtl.GetVal(fid)
    ids = mMultiIds.GetVal(fid)
    sel = mMulti.GetVal(fid, "")
    wasLoading = mLoading
    mLoading = True
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) <> ListHas(sel, CStr(ids(i))) Then lb.Selected(i) = ListHas(sel, CStr(ids(i)))
    Next
    mLoading = wasLoading
End Sub

Private Sub PickDate(ByVal fid As String)
    Dim tb As MSForms.TextBox, res As String, ok As Boolean
    Set tb = mCtl.GetVal(fid)
    res = mForms.PickDate(tb.Text, mCfg.Field(fid).FType = FT_DATUMZEIT, mCfg.Field(fid).Name, ok)
    If ok Then tb.Text = res: mDirty = True
End Sub

Private Sub QuickAddOption(ByVal fid As String)
    Dim nm As String, ok As Boolean, oid As String, f As cField
    Set f = mCfg.Field(fid)
    If Not mApp.RequireWritable("Das Anlegen einer Option") Then Exit Sub
    nm = mUI.AskText("Neue Auswahlmöglichkeit für „" & f.Name & "“. Sie steht danach in allen Einträgen dieser Tabelle zur Verfügung." & vbLf & _
                     "Farben und Reihenfolge lassen sich unter Einrichten ändern.", "Neue Option", "", False, ok)
    If Not ok Or Len(Trim$(nm)) = 0 Then Exit Sub
    If Len(mCfg.OptIdByName(fid, nm)) > 0 Then
        mUI.Alert "Diese Option gibt es bereits.", "Neue Option"
    Else
        oid = mStruct.AddOption(fid, Trim$(nm), "", False)
        mLog.Add "Option angelegt", f.TableID, oid, fid, Trim$(nm)
        mStruct.RebuildLists
    End If
    oid = mCfg.OptIdByName(fid, Trim$(nm))
    If f.FType = FT_MEHRFACH Then
        mMulti.SetVal fid, ListAdd(mMulti.GetVal(fid, ""), oid)
        FillMultiList fid, ""
        UpdateMultiSummary fid
    Else
        FillOptionCombo mCtl.GetVal(fid), f, oid
        UpdateChip fid
    End If
    mDirty = True
End Sub

Private Sub PickLinks(ByVal fid As String)
    Dim f As cField, tt As String, multi As Boolean, res As String, ok As Boolean
    Set f = mCfg.Field(fid)
    If f.FType = FT_LINK Then
        tt = f.TargetTable: multi = f.Multi
    Else
        tt = mCfg.Field(f.PairField).TableID: multi = True
    End If
    res = mForms.PickRecords(tt, mLinks.GetVal(fid, ""), multi, "Einträge für „" & f.Name & "“ auswählen", ok)
    If ok Then
        mLinks.SetVal fid, res
        FillLinkList fid
        mDirty = True
        ClearError fid
    End If
End Sub

Private Function SelectedLinkId(ByVal fid As String) As String
    Dim lb As MSForms.ListBox, a As Variant
    Set lb = mCtl.GetVal(fid)
    a = Split(mLinks.GetVal(fid, ""), ";")
    If UBound(a) < 0 Then Exit Function
    If lb.ListIndex < 0 Then
        If UBound(a) = 0 Then SelectedLinkId = a(0)
        Exit Function
    End If
    If lb.ListIndex <= UBound(a) Then SelectedLinkId = a(lb.ListIndex)
End Function

Private Sub OpenLinked(ByVal fid As String)
    Dim rid As String, f As cField, tt As String
    rid = SelectedLinkId(fid)
    If Len(rid) = 0 Then mUI.Alert "Bitte zuerst einen verknüpften Eintrag in der Liste auswählen.", "Öffnen": Exit Sub
    Set f = mCfg.Field(fid)
    If f.FType = FT_LINK Then tt = f.TargetTable Else tt = mCfg.Field(f.PairField).TableID
    ' Dieser Entwurf bleibt geoeffnet; der verknuepfte Eintrag oeffnet sich darueber.
    mForms.OpenRecordStacked tt, rid
    FillLinkList fid
End Sub

Private Sub RemoveLinked(ByVal fid As String)
    Dim rid As String
    rid = SelectedLinkId(fid)
    If Len(rid) = 0 Then mUI.Alert "Bitte zuerst einen verknüpften Eintrag in der Liste auswählen.", "Entfernen": Exit Sub
    mLinks.SetVal fid, ListRemove(mLinks.GetVal(fid, ""), rid)
    FillLinkList fid
    mDirty = True
End Sub

Private Sub BrowsePath(ByVal fid As String, ByVal folder As Boolean)
    Dim res As Variant, fd As FileDialog
    If folder Then
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        fd.title = "Ordner auswählen"
        If fd.Show = -1 Then mCtl.GetVal(fid).Text = fd.SelectedItems(1): mDirty = True
    Else
        res = Application.GetOpenFilename("Alle Dateien (*.*),*.*", , "Datei auswählen")
        If VarType(res) = vbString Then mCtl.GetVal(fid).Text = res: mDirty = True
    End If
End Sub

Private Sub RunButtonHere(ByVal fid As String)
    If mDirty Then
        If Not mUI.Confirm("Vor der Button-Aktion müssen die Änderungen dieses Eintrags gespeichert werden. Jetzt speichern?", "Button", "Speichern und ausführen", "Abbrechen") Then Exit Sub
        If Not DoSave() Then Exit Sub
        mDirty = False
    End If
    mForms.RunButton mTid, mRid, fid
    SavedRid = mRid
    BuildAll
End Sub

Private Sub OpenLink(ByVal fid As String)
    Dim s As String, f As cField
    Set f = mCfg.Field(fid)
    s = Trim$(mCtl.GetVal(fid).Text)
    If Len(s) = 0 Then Exit Sub
    If f.Kind = FT_EMAIL Then
        If Not mUI.Confirm("Das E-Mail-Programm dieses Rechners öffnen, um an " & s & " zu schreiben? Es wird nichts automatisch versendet.", "E-Mail", "Öffnen", "Abbrechen") Then Exit Sub
        s = "mailto:" & s
    Else
        If LCase$(Left$(s, 4)) = "www." Then s = "https://" & s
        If Not mUI.Confirm("Diese Adresse im Browser öffnen? Dafür wird eine Internetverbindung benötigt; die Anwendung selbst bleibt offline." & vbLf & vbLf & s, "Link öffnen", "Öffnen", "Abbrechen") Then Exit Sub
    End If
    On Error GoTo failed
    ThisWorkbook.FollowHyperlink Address:=s, NewWindow:=True
    Exit Sub
failed:
    mUI.Alert "Die Adresse konnte nicht geöffnet werden: " & Err.Description, "Link öffnen", True
End Sub

Private Sub FillFileList(ByVal fid As String)
    Dim lb As MSForms.ListBox, a As Variant, i As Long
    Set lb = Me.Controls("fl_" & fid)
    lb.Clear
    a = Split(mCtl.GetVal(fid).Text, "|")
    For i = 0 To UBound(a)
        If Len(Trim$(a(i))) > 0 Then lb.AddItem a(i)
    Next
    If lb.ListCount = 0 Then lb.AddItem "(keine Dateien)"
End Sub

Private Sub AddFiles(ByVal fid As String)
    Dim res As Variant, i As Long, s As String
    res = Application.GetOpenFilename("Alle Dateien (*.*),*.*", , "Dateien auswählen", , True)
    If Not IsArray(res) Then Exit Sub
    s = mCtl.GetVal(fid).Text
    For i = LBound(res) To UBound(res)
        s = ListAdd(s, CStr(res(i)), "|")
    Next
    mCtl.GetVal(fid).Text = s
    FillFileList fid
    mDirty = True
End Sub

Private Function SelectedFile(ByVal fid As String) As String
    Dim lb As MSForms.ListBox
    Set lb = Me.Controls("fl_" & fid)
    If lb.ListIndex >= 0 Then SelectedFile = lb.List(lb.ListIndex)
    If SelectedFile = "(keine Dateien)" Then SelectedFile = ""
End Function

Private Sub OpenFile(ByVal fid As String)
    Dim p As String
    p = SelectedFile(fid)
    If Len(p) = 0 Then mUI.Alert "Bitte zuerst eine Datei in der Liste auswählen.", "Datei öffnen": Exit Sub
    If Not FileExists(p) Then mUI.Alert "Die Datei wurde nicht gefunden:" & vbLf & p, "Datei öffnen", True: Exit Sub
    If Not mUI.Confirm("Diese lokale Datei mit dem zugeordneten Programm öffnen?" & vbLf & vbLf & p, "Datei öffnen", "Öffnen", "Abbrechen") Then Exit Sub
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=p, NewWindow:=True
End Sub

Private Sub RemoveFile(ByVal fid As String)
    Dim p As String
    p = SelectedFile(fid)
    If Len(p) = 0 Then Exit Sub
    mCtl.GetVal(fid).Text = ListRemove(mCtl.GetVal(fid).Text, p, "|")
    FillFileList fid
    mDirty = True
End Sub

Private Sub OpenPath(ByVal fid As String)
    Dim p As String
    p = Trim$(mCtl.GetVal(fid).Text)
    If Len(p) = 0 Then Exit Sub
    If Not (FileExists(p) Or DirExists(p)) Then
        mUI.Alert "Der Pfad wurde nicht gefunden:" & vbLf & p, "Lokaler Pfad", True
        Exit Sub
    End If
    If mUI.Confirm("Soll dieser lokale Pfad mit dem zugeordneten Windows-Programm geöffnet werden?" & vbLf & vbLf & p, "Pfad öffnen", "Öffnen", "Abbrechen") Then
        On Error GoTo failed
        ThisWorkbook.FollowHyperlink Address:=p, NewWindow:=True
    End If
    Exit Sub
failed:
    mUI.Alert "Der Pfad konnte nicht geöffnet werden: " & Err.Description, "Lokaler Pfad", True
End Sub

'==============================================================================
' Speichern
'==============================================================================
Private Function CollectValues(ByRef vals As cMap, ByRef firstErr As String) As Boolean
    Dim f As cField, c As Object, outV As Variant, msg As String, txt As String, cur As String, ok As Boolean
    Set vals = New cMap
    ok = True
    For Each f In mFields
        If f.IsEditable And Not f.IsLinkType Then
            Set c = mCtl.GetVal(f.ID)
            If mIsNew Then cur = "" Else cur = NzS(mData.GetValue(mTid, mRid, f.ID))
            msg = ""
            Select Case f.FType
                Case FT_JANEIN
                    If TypeName(c) = "CheckBox" Then
                        vals.SetVal f.ID, CBool(c.Value)
                    Else
                        Select Case c.ListIndex
                            Case 1: vals.SetVal f.ID, True
                            Case 2: vals.SetVal f.ID, False
                            Case Else: vals.SetVal f.ID, Empty
                        End Select
                    End If
                Case FT_AUSWAHL, FT_STATUS
                    txt = Trim$(Replace(c.Text, " (deaktiviert)", ""))
                    If txt = "(keine Angabe)" Or Len(txt) = 0 Then
                        vals.SetVal f.ID, Empty
                    ElseIf mValid.ParseInput(f, txt, outV, msg, cur) Then
                        vals.SetVal f.ID, outV
                    End If
                Case FT_MEHRFACH
                    If Len(mMulti.GetVal(f.ID, "")) > 0 Then vals.SetVal f.ID, mMulti.GetVal(f.ID, "") Else vals.SetVal f.ID, Empty
                Case Else
                    If mValid.ParseInput(f, c.Text, outV, msg, cur) Then vals.SetVal f.ID, outV
            End Select
            If Len(msg) > 0 Then
                ShowError f.ID, msg
                If Len(firstErr) = 0 Then firstErr = f.ID
                ok = False
            End If
        End If
    Next
    CollectValues = ok
End Function

Private Sub FocusField(ByVal fid As String)
    On Error Resume Next
    mp.Value = mPageOf.GetVal(fid, 0)
    EnsureVisible fid
    mCtl.GetVal(fid).SetFocus
End Sub

Private Function DoSave() As Boolean
    Dim vals As cMap, firstErr As String, errs As cMap, k As Variant, dup As String, ans As Long, rid As String
    Dim t As cTableData, a As Variant, c As Long, f As cField, note As String, links As cMap, changedF As String
    If Not mApp.RequireWritable("Das Speichern") Then Exit Function
    ' 1) Eingaben umwandeln
    If Not CollectValues(vals, firstErr) Then
        lblFoot.caption = ChrW(&H26A0) & "  Bitte die markierten Angaben korrigieren. Es wurde noch nichts gespeichert."
        lblFoot.ForeColor = CDanger()
        FocusField firstErr
        Exit Function
    End If
    ' 2) Datensatzregeln
    Set links = New cMap
    For Each k In mLinks.Keys
        links.SetVal CStr(k), mLinks.GetVal(CStr(k))
    Next
    If Not mValid.ValidateRecord(mTid, IIf(mIsNew, "", mRid), vals, links, errs) Then
        For Each k In errs.Keys
            ShowError CStr(k), CStr(errs.GetVal(CStr(k)))
            If Len(firstErr) = 0 Then firstErr = CStr(k)
        Next
        lblFoot.caption = ChrW(&H26A0) & "  Bitte die markierten Angaben korrigieren. Es wurde noch nichts gespeichert."
        lblFoot.ForeColor = CDanger()
        FocusField firstErr
        Exit Function
    End If
    ' 3) Dublettenhinweis
    dup = mValid.FindDuplicateRecord(mTid, IIf(mIsNew, "", mRid), vals)
    If Len(dup) > 0 Then
        ans = mUI.Ask3("Es gibt bereits einen Eintrag mit denselben Angaben in den Dublettenfeldern:" & vbLf & vbLf & _
                       "„" & mData.RecordLabel(mTid, dup) & "“  (" & mData.RecordHint(mTid, dup) & ")" & vbLf & vbLf & _
                       "Möchten Sie den bestehenden Eintrag ansehen oder trotzdem speichern?", "Mögliche Dublette", _
                       "Trotzdem speichern", "Bestehenden öffnen", "Zurück")
        If ans = 2 Then mForms.OpenRecordStacked mTid, dup: Exit Function
        If ans <> 1 Then Exit Function
    End If
    ' 4) Transaktion
    On Error GoTo failed
    mApp.BeginWork
    If mIsNew Then
        mData.TxBegin "Neuer Eintrag in " & mCfg.TableName(mTid)
        rid = mData.NewRecordId()
        Set t = mData.TData(mTid)
        a = mData.NewRowArray(mTid, rid)
        For Each k In vals.Keys
            c = t.ColOf(CStr(k))
            If c > 0 Then a(c) = vals.GetVal(CStr(k))
        Next
        mData.InsertRecord mTid, a
        changedF = Join(vals.Keys, ";")
    Else
        mData.TxBegin "Änderung an „" & mData.RecordLabel(mTid, mRid) & "“"
        rid = mRid
        mData.UpdateRecord mTid, rid, vals
        changedF = Join(vals.Keys, ";")
    End If
    For Each k In mLinks.Keys
        Set f = mCfg.Field(CStr(k))
        If mData.LinkTargets(CStr(k), rid) <> mLinks.GetVal(CStr(k)) Then
            note = note & mData.SetLinks(CStr(k), rid, mLinks.GetVal(CStr(k)))
            changedF = changedF & IIf(Len(changedF) > 0, ";", "") & CStr(k)
        End If
    Next
    mLog.Add IIf(mIsNew, "Angelegt", "Geändert"), mTid, rid, changedF
    mData.TxCommit
    mApp.EndWork
    SavedRid = rid
    If Len(note) > 0 Then mUI.Alert "Gespeichert. Hinweis zu geänderten Zuordnungen:" & vbLf & note, "Zuordnungen"
    DoSave = True
    Exit Function
failed:
    Dim emsg As String
    emsg = Err.Description
    mData.TxAbortIfOpen
    mApp.ResetWork
    lblFoot.caption = ChrW(&H26A0) & "  Speichern fehlgeschlagen; es wurde nichts übernommen (" & emsg & "). Ihre Eingaben sind erhalten."
    lblFoot.ForeColor = CDanger()
End Function

Private Sub btnSave_Click()
    If DoSave() Then
        mDirty = False
        Me.Hide
    End If
End Sub

Private Sub btnSaveNew_Click()
    If DoSave() Then
        mForms.NotifySaved mTid
        mRid = ""
        mIsNew = True
        Set mPresets = New cMap
        BuildAll
        On Error Resume Next
        mp.Value = 0
    End If
End Sub

Private Sub btnDup_Click()
    Dim pre As New cMap, f As cField, v As Variant, vals As cMap, firstErr As String
    If mDirty Then
        If Not mUI.Confirm("Ungespeicherte Änderungen dieses Eintrags werden nicht in die Kopie übernommen, solange Sie nicht vorher speichern. Trotzdem eine Kopie des gespeicherten Standes öffnen?", _
                           "Duplizieren", "Kopie öffnen", "Zurück") Then Exit Sub
    End If
    For Each f In mFields
        If f.IsEditable Then
            If f.FType = FT_LINK Then
                pre.SetVal f.ID, mData.LinkTargets(f.ID, mRid)
            ElseIf Not f.IsLinkType Then
                v = mData.GetValue(mTid, mRid, f.ID)
                If f.Unique Then v = Empty
                pre.SetVal f.ID, v
            End If
        End If
    Next
    mForms.NewRecordStacked mTid, pre
End Sub

Private Sub btnArch_Click()
    Dim arch As Boolean
    If Not mApp.RequireWritable("Das Archivieren") Then Exit Sub
    arch = Not mData.IsArchivedRec(mTid, mRid)
    If mDirty Then
        If Not mUI.Confirm("Es gibt ungespeicherte Änderungen. Diese gehen beim " & IIf(arch, "Archivieren", "Wiederherstellen") & " verloren. Fortfahren?", _
                           "Ungespeicherte Änderungen", "Fortfahren", "Zurück") Then Exit Sub
    End If
    If arch Then
        If Not mUI.Confirm("„" & mData.RecordLabel(mTid, mRid) & "“ archivieren?" & vbLf & vbLf & _
                           "Archivierte Einträge erscheinen nicht mehr in den Ansichten und Auswertungen, bleiben aber mit allen Verknüpfungen erhalten " & _
                           "und können über „Ansicht > Archiv anzeigen“ wiederhergestellt werden.", "Archivieren", "Archivieren", "Abbrechen") Then Exit Sub
    End If
    mForms.SetArchived mTid, mRid, arch
    SavedRid = mRid
    mDirty = False
    Me.Hide
End Sub

Private Sub btnMore_Click()
    btnDup.Visible = Not btnDup.Visible
    btnArch.Visible = btnDup.Visible
End Sub

Private Function HasRequired() As Boolean
    Dim f As cField
    For Each f In mFields
        If f.Required And ShowInForm(f) Then HasRequired = True: Exit Function
    Next
End Function

Private Sub btnCancel_Click()
    If mDirty Then
        If Not mUI.Confirm("Es gibt ungespeicherte Änderungen. Wirklich schließen und die Änderungen verwerfen?", "Änderungen verwerfen", _
                           "Verwerfen", "Weiter bearbeiten") Then Exit Sub
    End If
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCancel_Click
    End If
End Sub

Private Sub UserForm_Terminate()
    Dim e As cEvt
    If Not mEvts Is Nothing Then
        For Each e In mEvts: e.Unbind: Next
    End If
End Sub
