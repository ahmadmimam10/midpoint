'==============================================================================
' frmTableEditor - Assistent fuer neue Tabellen und Bearbeitung bestehender
' Tabellen: Name, Beschreibung, Symbol, Felder (anlegen, bearbeiten, ordnen,
' ausblenden, entfernen), Bezeichnungsfeld, Dublettenpruefung, Maskenvorschau.
' Jede Aenderung wird sofort in der Konfiguration gespeichert.
'==============================================================================
Option Explicit

Public TableId As String
Private mIcon As Long
Private mEvts As Collection
Private mStep As Long
Private txtName As MSForms.TextBox
Private txtDesc As MSForms.TextBox
Private lst As MSForms.ListBox
Private cboLabel As MSForms.ComboBox
Private lblDup As MSForms.Label
Private lblErr As MSForms.Label
Private mFieldIds() As String
Private mLoading As Boolean
Private fr1 As MSForms.Frame
Private fr2 As MSForms.Frame

Public Sub Setup(ByVal tid As String)
    Dim y As Single, i As Long, codes As Variant, b As MSForms.CommandButton, x As Single
    TableId = tid
    Set mEvts = New Collection
    mUI.StyleForm Me
    Me.Width = 720
    If Me.Width > mUI.MaxFormWidth Then Me.Width = mUI.MaxFormWidth
    Me.caption = IIf(Len(tid) = 0, "Neue Tabelle", "Tabelle bearbeiten")
    mUI.AddLabel Me, "t", IIf(Len(tid) = 0, "Neue Tabelle anlegen", "Tabelle „" & mCfg.TableName(tid) & "“ bearbeiten"), PAD, PAD - 2, 600, 20, True, 13
    ' Bereich 1: Eigenschaften
    y = PAD + 30
    mUI.AddLabel Me, "ln", "Name der Tabelle *", PAD, y, 300, 15, True
    mUI.AddLabel Me, "ld", "Beschreibung (optional)", PAD + 320, y, 300, 15, True
    Set txtName = mUI.AddText(Me, "name", PAD, y + 17, 300)
    Set txtDesc = mUI.AddText(Me, "desc", PAD + 320, y + 17, Me.InsideWidth - PAD * 2 - 320)
    y = y + 46
    mUI.AddLabel Me, "li", "Symbol", PAD, y, 300, 15, True
    y = y + 17
    codes = TableIconCodes()
    For i = 0 To UBound(codes)
        x = PAD + i * 34
        Set b = mUI.AddButton(Me, "ic" & i, ChrW(CLng(codes(i))), x, y, 30, , 28)
        b.Font.Name = ICON_FONT
        b.Font.Size = 12
        b.Tag = CStr(codes(i))
        Bind b, "icon:" & i
    Next
    y = y + 36
    If Len(tid) > 0 Then
        txtName.Text = mCfg.TableName(tid)
        txtDesc.Text = mCfg.GetS(SH_TAB, tid, "Beschreibung")
        mIcon = CLng(Val(mCfg.GetS(SH_TAB, tid, "Symbol")))
    Else
        mIcon = CLng(codes(0))
    End If
    MarkIcon
    Set lblErr = mUI.AddLabel(Me, "err", "", PAD, y, Me.InsideWidth - PAD * 2, 15, False, 8.5, CDanger())
    y = y + 18
    If Len(tid) = 0 Then
        Set b = mUI.AddButton(Me, "create", "Tabelle anlegen und Eigenschaften festlegen", PAD, y, 260, "primary")
        Bind b, "create"
        Set b = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 104, y, 104)
        b.Cancel = True
        Bind b, "close"
        mUI.AddLabel Me, "hint", "Im nächsten Schritt fügen Sie Eigenschaften (Spalten) hinzu. Eine Eigenschaft „Name“ wird als Bezeichnungsfeld vorbereitet und kann umbenannt werden.", _
                     PAD, y + 34, Me.InsideWidth - PAD * 2, 28, False, 8.5, CMuted()
        Me.Height = y + 70 + PAD + (Me.Height - Me.InsideHeight)
    Else
        Set b = mUI.AddButton(Me, "saveprops", "Eigenschaften übernehmen", PAD, y, 190)
        Bind b, "saveprops"
        y = y + BTN_H + 14
        BuildFieldsArea y
    End If
End Sub

Private Sub Bind(ByVal c As Object, ByVal key As String)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, c, key
    mEvts.Add e
End Sub

Private Sub MarkIcon()
    Dim i As Long, codes As Variant
    codes = TableIconCodes()
    For i = 0 To UBound(codes)
        If CLng(codes(i)) = mIcon Then
            mUI.StyleButton Me.Controls("ic" & i), "primary"
        Else
            mUI.StyleButton Me.Controls("ic" & i), "secondary"
        End If
    Next
End Sub

Private Sub BuildFieldsArea(ByVal y As Single)
    Dim b As MSForms.CommandButton, x As Single, h As Single
    mUI.AddLine Me, "sep", PAD, y - 6, Me.InsideWidth - PAD * 2
    mUI.AddLabel Me, "lf", "Eigenschaften (Spalten)", PAD, y, 300, 18, True, 11
    mUI.AddLabel Me, "lf2", "Reihenfolge = Reihenfolge in der Eingabemaske. Ausgeblendete Eigenschaften behalten ihre Daten.", PAD + 180, y + 3, Me.InsideWidth - PAD * 2 - 180, 15, False, 8.5, CMuted()
    y = y + 22
    h = mUI.MaxFormHeight() - y - 170
    If h > 260 Then h = 260
    If h < 120 Then h = 120
    Set lst = mUI.AddList(Me, "lst", PAD, y, Me.InsideWidth - PAD * 2 - 150, h, False, 4, "180;140;70;110")
    Bind lst, "lst"
    x = Me.InsideWidth - PAD - 140
    Set b = mUI.AddButton(Me, "fadd", "Neue Eigenschaft…", x, y, 140, "primary"): Bind b, "fadd"
    Set b = mUI.AddButton(Me, "fedit", "Bearbeiten…", x, y + 32, 140): Bind b, "fedit"
    Set b = mUI.AddButton(Me, "fup", "Nach oben", x, y + 64, 140): Bind b, "fup"
    Set b = mUI.AddButton(Me, "fdown", "Nach unten", x, y + 96, 140): Bind b, "fdown"
    Set b = mUI.AddButton(Me, "fhide", "Aus-/Einblenden", x, y + 128, 140): Bind b, "fhide"
    Set b = mUI.AddButton(Me, "fdel", "Spalte entfernen…", x, y + 160, 140, "danger"): Bind b, "fdel"
    y = y + h + 12
    mUI.AddLabel Me, "ll", "Bezeichnungsfeld (für Suche, Auswahl und Verknüpfungen)", PAD, y, 330, 15, True
    mUI.AddLabel Me, "ldp", "Dublettenprüfung", PAD + 350, y, 300, 15, True
    Set cboLabel = mUI.AddCombo(Me, "label", PAD, y + 17, 330)
    Bind cboLabel, "label"
    Set lblDup = mUI.AddLabel(Me, "dup", "", PAD + 350, y + 20, 200, 30, False, 8.5, CMuted())
    Set b = mUI.AddButton(Me, "dupb", "Festlegen…", Me.InsideWidth - PAD - 140, y + 15, 140): Bind b, "dup"
    y = y + 52
    Set b = mUI.AddButton(Me, "preview", "Vorschau der Eingabemaske", PAD, y, 200): Bind b, "preview"
    Set b = mUI.AddButton(Me, "tdel", "Tabelle löschen…", PAD + 208, y, 140, "danger"): Bind b, "tdel"
    Set b = mUI.AddButton(Me, "done", "Fertig", Me.InsideWidth - PAD - 120, y, 120, "primary"): Bind b, "close"
    b.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    mUI.KeepOnScreen Me
    RefreshFields
End Sub

Private Sub RefreshFields(Optional ByVal selFid As String = "")
    Dim f As cField, n As Long, out() As Variant, i As Long, sel As Long, lf As String, dups As String, s As String, v As Variant
    mLoading = True
    lst.Clear
    lf = mCfg.LabelField(TableId)
    ReDim mFieldIds(0 To mCfg.Fields(TableId, False).Count)
    sel = -1
    cboLabel.Clear
    For Each f In mCfg.Fields(TableId, False)
        lst.AddItem f.Name & IIf(f.ID = lf, "  " & ChrW(&H2605) & "", "")
        If f.FType = FT_LINK Then
            lst.List(n, 1) = f.TypeName_ & " " & ChrW(&H2192) & " " & mCfg.TableName(f.TargetTable)
        ElseIf f.FType = FT_BACKLINK Then
            lst.List(n, 1) = f.TypeName_ & " " & ChrW(&H2190) & " " & mCfg.TableName(mCfg.Field(f.PairField).TableID)
        Else
            lst.List(n, 1) = f.TypeName_
        End If
        lst.List(n, 2) = IIf(f.Required, "Pflicht", "")
        lst.List(n, 3) = IIf(f.IsActive, IIf(f.Visible, "sichtbar", "nur Daten"), "ausgeblendet")
        mFieldIds(n) = f.ID
        If f.ID = selFid Then sel = n
        n = n + 1
        If (f.FType = FT_TEXT Or f.FType = FT_FORMEL) And f.IsActive Then
            cboLabel.AddItem f.Name
            If f.ID = lf Then cboLabel.ListIndex = cboLabel.ListCount - 1
        End If
    Next
    If sel >= 0 Then lst.ListIndex = sel
    dups = mCfg.GetS(SH_TAB, TableId, "Dublettenfelder")
    For Each v In Split(dups, ";")
        s = s & IIf(Len(s) > 0, ", ", "") & mCfg.FieldName(CStr(v))
    Next
    lblDup.caption = IIf(Len(s) = 0, "keine (keine Warnung bei gleichen Einträgen)", "Warnung bei gleichem Wert in: " & s)
    mLoading = False
End Sub

Private Function SelFid() As String
    If lst Is Nothing Then Exit Function
    If lst.ListIndex >= 0 Then SelFid = mFieldIds(lst.ListIndex)
End Function

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    Dim fid As String
    On Error GoTo failed
    If Left$(key, 5) = "icon:" Then
        mIcon = CLng(Me.Controls("ic" & Mid$(key, 6)).Tag)
        MarkIcon
        Exit Sub
    End If
    Select Case key
        Case "create": If evt = "Click" Then DoCreate
        Case "saveprops": If evt = "Click" Then SaveProps
        Case "close": If evt = "Click" Then Me.Hide
        Case "fadd": If evt = "Click" Then EditField ""
        Case "fedit": If evt = "Click" Then If Len(SelFid) > 0 Then EditField SelFid Else mUI.Alert "Bitte zuerst eine Eigenschaft auswählen.", "Eigenschaft"
        Case "lst": If evt = "DblClick" And Len(SelFid) > 0 Then EditField SelFid
        Case "fup", "fdown"
            If evt = "Click" And Len(SelFid) > 0 Then
                fid = SelFid
                mStruct.MoveField fid, IIf(key = "fup", -1, 1)
                RefreshFields fid
            End If
        Case "fhide"
            If evt = "Click" And Len(SelFid) > 0 Then
                fid = SelFid
                mStruct.SetFieldHidden fid, mCfg.Field(fid).IsActive
                RefreshFields fid
            End If
        Case "fdel": If evt = "Click" And Len(SelFid) > 0 Then mCmd.DeleteFieldFlow SelFid: RefreshFields
        Case "label"
            If evt = "Change" And Not mLoading And cboLabel.ListIndex >= 0 Then
                mStruct.SetLabelField TableId, mCfg.FieldIdByName(TableId, cboLabel.Text)
                RefreshFields SelFid
            End If
        Case "dup": If evt = "Click" Then ChooseDup
        Case "preview": If evt = "Click" Then Preview
        Case "tdel"
            If evt = "Click" Then
                ' Rueckfragen erscheinen ueber diesem Dialog; nach dem Loeschen schliesst er
                If mCmd.DeleteTableFlow(TableId) Then TableId = "": Me.Hide
            End If
    End Select
    Exit Sub
failed:
    mApp.Boundary "Tabelle bearbeiten"
End Sub

Private Sub DoCreate()
    Dim props As New cMap, fid As String
    lblErr.caption = ""
    If Len(Trim$(txtName.Text)) = 0 Then lblErr.caption = ChrW(&H26A0) & "  Bitte einen Namen eingeben.": mUI.MarkError txtName, True: Exit Sub
    If Len(mCfg.TableIdByName(txtName.Text)) > 0 Then lblErr.caption = ChrW(&H26A0) & "  Diesen Tabellennamen gibt es bereits.": mUI.MarkError txtName, True: Exit Sub
    mApp.BeginWork
    TableId = mStruct.CreateTable(txtName.Text, txtDesc.Text, mIcon)
    props.SetVal "Pflicht", "Ja"
    props.SetVal "Hilfetext", "Bezeichnung, unter der der Eintrag in Listen, Suche und Verknüpfungen erscheint."
    fid = mStruct.AddField(TableId, "Name", FT_TEXT, props)
    mStruct.SetLabelField TableId, fid
    mStruct.CreateDefaultView TableId
    mApp.EndWork
    ' Bereich fuer Felder aufbauen
    Me.Controls("create").Visible = False
    Me.Controls("cancel").Visible = False
    Me.Controls("hint").Visible = False
    Me.Controls("t").caption = "Tabelle „" & mCfg.TableName(TableId) & "“: Eigenschaften festlegen"
    Dim b As MSForms.CommandButton
    Set b = mUI.AddButton(Me, "saveprops", "Eigenschaften übernehmen", PAD, Me.Controls("create").Top, 190)
    Bind b, "saveprops"
    BuildFieldsArea Me.Controls("create").Top + BTN_H + 14
    mNav.RefreshNav
End Sub

Private Sub SaveProps()
    lblErr.caption = ""
    mUI.MarkError txtName, False
    mStruct.UpdateTable TableId, txtName.Text, txtDesc.Text, CStr(mIcon)
    Me.Controls("t").caption = "Tabelle „" & mCfg.TableName(TableId) & "“ bearbeiten"
    mNav.RefreshNav
    lblErr.ForeColor = CSuccess()
    lblErr.caption = ChrW(&H2713) & "  Eigenschaften gespeichert."
End Sub

Private Sub EditField(ByVal fid As String)
    Dim f As frmField, res As String
    Set f = New frmField
    f.Setup TableId, fid
    f.Show
    res = f.ResultFid
    Unload f
    If Len(res) > 0 Then RefreshFields res Else RefreshFields SelFid
End Sub

Private Sub ChooseDup()
    Dim items() As Variant, ids() As String, f As cField, n As Long, pre As String, res As String, ok As Boolean, v As Variant, s As String
    ReDim items(0 To mCfg.Fields(TableId).Count): ReDim ids(0 To mCfg.Fields(TableId).Count)
    For Each f In mCfg.Fields(TableId)
        If Not f.IsLinkType And f.FType <> FT_LTEXT Then
            items(n) = f.Name: ids(n) = f.ID
            If ListHas(mCfg.GetS(SH_TAB, TableId, "Dublettenfelder"), f.ID) Then pre = pre & IIf(Len(pre) > 0, ";", "") & n
            n = n + 1
        End If
    Next
    If n = 0 Then Exit Sub
    ReDim Preserve items(0 To n - 1)
    res = mUI.ChooseMany("Bei gleichem Wert in ALLEN gewählten Feldern erscheint beim Speichern eine Warnung mit der Möglichkeit, den bestehenden Eintrag zu öffnen. " & _
                         "(Echte Eindeutigkeit wird je Feld über „Eindeutig“ erzwungen.)", "Dublettenprüfung", items, pre, ok)
    If Not ok Then Exit Sub
    For Each v In Split(res, ";")
        If Len(v) > 0 Then s = ListAdd(s, ids(CLng(v)))
    Next
    mCfg.SetV SH_TAB, TableId, "Dublettenfelder", s
    RefreshFields SelFid
End Sub

Private Sub Preview()
    Dim f As frmRecord
    Set f = New frmRecord
    f.PreviewOnly = True
    f.OpenFor TableId, ""
    Unload f
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True: Me.Hide
End Sub

Private Sub UserForm_Terminate()
    Dim e As cEvt
    If Not mEvts Is Nothing Then
        For Each e In mEvts: e.Unbind: Next
    End If
End Sub
