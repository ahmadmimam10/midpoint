'==============================================================================
' frmXlBar - Schwebende Formelleiste fuer Formelspalten (nicht modal): Formel
' wie in Excel schreiben; ein Klick auf eine Spalte der Tabelle fuegt sie als
' [Spalte] an der Schreibmarke ein. Neue Formelspalte oder bestehende bearbeiten.
'==============================================================================
Option Explicit

Private mTid As String
Private mFid As String
Private mAfter As String
Private mEvts As Collection
Private mXl As cXlEditor
Private txtName As MSForms.TextBox
Private cboRt As MSForms.ComboBox
Private txtDec As MSForms.TextBox
Private lblErr As MSForms.Label
Private mRtAuto As Boolean
Private mLoading As Boolean

Public Sub Setup(ByVal tid As String, ByVal fid As String, ByVal afterFid As String)
    Dim f As cField, y As Single, w As Single, b As MSForms.CommandButton, v As Variant, h As Single
    mTid = tid: mFid = fid: mAfter = afterFid
    Set mEvts = New Collection
    mUI.StyleForm Me
    Me.Width = 470
    w = Me.InsideWidth - PAD * 2
    If Len(fid) > 0 Then Set f = mCfg.Field(fid)
    If f Is Nothing Then
        Me.caption = "Neue Formelspalte"
        mUI.AddLabel Me, "t", "Neue Formelspalte in „" & mCfg.TableName(tid) & "“", PAD, PAD - 4, w, 20, True, 12
    Else
        Me.caption = "Formel: " & f.Name
        mUI.AddLabel Me, "t", "Formel von „" & f.Name & "“", PAD, PAD - 4, w, 20, True, 12
    End If
    mUI.AddLabel Me, "t2", "Tipp: Klicken Sie in der Tabelle auf eine Spalte; sie wird an der Schreibmarke eingefügt. Dieses Fenster bleibt dabei offen.", _
                 PAD, PAD + 16, w, 28, False, 8, CMuted()
    y = PAD + 46
    mLoading = True
    Dim def As String, dec As Long, x2 As Single
    dec = 2
    If Not f Is Nothing Then def = f.Formula: dec = f.Decimals
    ' Zuerst die Formel (Schreibmarke steht beim Oeffnen dort), Name und Format darunter
    Set mXl = New cXlEditor
    y = y + mXl.Build(Me, tid, fid, def, PAD, y, w, Me)
    x2 = PAD
    If f Is Nothing Then
        mUI.AddLabel Me, "ln", "Name der Spalte", x2, y, 170, 15, True
        Set txtName = mUI.AddText(Me, "colname", x2, y + 17, 170)
        txtName.Text = FreeName(tid)
        x2 = x2 + 180
    End If
    mUI.AddLabel Me, "lrt", "Ergebnis als", x2, y, 150, 15, True
    mUI.AddLabel Me, "ldec", "Kommastellen", x2 + 160, y, 90, 15, True
    Set cboRt = mUI.AddCombo(Me, "rt", x2, y + 17, 150)
    For Each v In Array(FT_ZAHL, FT_GELD, FT_PROZENT, FT_DATUM, FT_TEXT, FT_JANEIN)
        cboRt.AddItem FieldTypeName(CStr(v))
    Next
    If f Is Nothing Then
        cboRt.Value = FieldTypeName(FT_ZAHL)
    Else
        cboRt.Value = FieldTypeName(IIf(Len(f.ResultType) > 0, f.ResultType, FT_ZAHL))
    End If
    Bind cboRt, "rt"
    Set txtDec = mUI.AddText(Me, "dec", x2 + 160, y + 17, 60)
    txtDec.Text = CStr(dec)
    Bind txtDec, "dec"
    y = y + 46
    Set lblErr = mUI.AddLabel(Me, "err", "", PAD, y, w, 28, False, 8.5, CDanger())
    y = y + 30
    Set b = mUI.AddButton(Me, "ok", "Übernehmen", PAD, y, 130, "primary"): Bind b, "ok"
    Set b = mUI.AddButton(Me, "cancel", "Schließen", Me.InsideWidth - PAD - 104, y, 104): Bind b, "cancel"
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    mRtAuto = (f Is Nothing)
    mLoading = False
    mXl.ResultType = FieldTypeFromName(cboRt.Text)
    mXl.Decimals = CLng(Val(txtDec.Text))
    mXl.Refresh
    ' rechts neben dem Inhalt, damit die Spalten sichtbar bleiben
    Me.StartUpPosition = 0
    Me.Left = Application.Left + Application.Width - Me.Width - 40
    Me.Top = Application.Top + 140
End Sub

' Vorschlag "Formel", "Formel 2", ... (jederzeit aenderbar)
Private Function FreeName(ByVal tid As String) As String
    Dim k As Long
    FreeName = "Formel"
    Do While Len(mCfg.FieldIdByName(tid, FreeName)) > 0
        k = k + 1
        FreeName = "Formel " & (k + 1)
    Loop
End Function

Private Sub Bind(ByVal c As Object, ByVal key As String)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, c, key
    mEvts.Add e
End Sub

Public Function TableId() As String
    TableId = mTid
End Function

' Klick auf eine Spalte der Tabelle
Public Sub InsertFromSheet(ByVal fid As String)
    If Len(fid) = 0 Then Exit Sub
    mXl.InsertField fid
    ' Tastatur zurueck in die Formel, damit direkt weitergeschrieben werden kann
    mWin.ActivateWindow mWin.FormHwnd(Me.caption)
    mXl.FocusText
End Sub

Public Sub OnFormulaPreview(ByVal v As Variant)
    Dim nm As String
    If Not mRtAuto Then Exit Sub
    Select Case VarType(v)
        Case vbBoolean: nm = FieldTypeName(FT_JANEIN)
        Case vbString: If Not IsErrVal(v) Then nm = FieldTypeName(FT_TEXT)
        Case vbDouble: nm = FieldTypeName(FT_ZAHL)
    End Select
    If Len(nm) = 0 Or cboRt Is Nothing Then Exit Sub
    If cboRt.Text = nm Then Exit Sub
    mLoading = True
    cboRt.Value = nm
    mLoading = False
    mXl.ResultType = FieldTypeFromName(nm)
End Sub

Private Sub Commit()
    Dim def As String, msg As String, props As New cMap, nm As String, fid As String, f As cField, recalc As Boolean
    Dim isNew As Boolean, after As String
    lblErr.caption = ""
    If Not mApp.RequireSetup("Das Ändern einer Formelspalte") Then Exit Sub
    If Len(mFid) = 0 Then
        nm = Trim$(txtName.Text)
        If Len(nm) = 0 Then lblErr.caption = ChrW(&H26A0) & "  Bitte einen Namen für die Spalte eingeben.": txtName.SetFocus: Exit Sub
        If Len(mCfg.FieldIdByName(mTid, nm)) > 0 Then lblErr.caption = ChrW(&H26A0) & "  Diese Tabelle hat bereits eine Spalte „" & nm & "“.": Exit Sub
    End If
    def = mXl.Definition(msg)
    If Len(msg) > 0 Then lblErr.caption = ChrW(&H26A0) & "  " & msg: Exit Sub
    props.SetVal "Formel", def
    props.SetVal "Beschreibung", mCalc.Describe(def)
    props.SetVal "Ergebnistyp", FieldTypeFromName(cboRt.Text)
    props.SetVal "Dezimalstellen", CLng(Val(txtDec.Text))
    On Error GoTo failed
    mApp.BeginWork "Formelspalte wird berechnet …"
    If Len(mFid) = 0 Then
        fid = mStruct.AddField(mTid, nm, FT_FORMEL, props)
        recalc = True
    Else
        fid = mFid
        Set f = mCfg.Field(fid)
        recalc = (f.Formula <> def Or f.ResultType <> props.GetVal("Ergebnistyp") Or f.Decimals <> props.GetVal("Dezimalstellen"))
        mStruct.UpdateField fid, props
    End If
    If recalc Then mCalc.RecalcAll
    mApp.EndWork
    mApp.MarkCfgDirty
    ' Werte sichern: Schliessen entlaedt dieses Formular samt seinen Variablen
    isNew = (Len(mFid) = 0): after = mAfter
    Me.Hide
    If isNew Then
        mCmd.AfterNewProperty fid, after
    Else
        mStruct.RebuildLists
        mNav.RenderCurrent
    End If
    mForms.CloseXlBar
    Exit Sub
failed:
    msg = Err.Description
    mApp.ResetWork
    lblErr.caption = ChrW(&H26A0) & "  " & msg
End Sub

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    On Error GoTo failed
    Select Case key
        Case "rt", "dec"
            If evt = "Change" And Not mLoading Then
                If key = "rt" Then mRtAuto = False
                mXl.ResultType = FieldTypeFromName(cboRt.Text)
                mXl.Decimals = CLng(Val(txtDec.Text))
                mXl.Refresh
            End If
        Case "ok": If evt = "Click" Then Commit
        Case "cancel": If evt = "Click" Then mForms.CloseXlBar
    End Select
    Exit Sub
failed:
    mApp.Boundary "Formelspalte"
End Sub

Public Sub Teardown()
    Dim e As cEvt
    If Not mEvts Is Nothing Then
        For Each e In mEvts: e.Unbind: Next
    End If
    Set mEvts = Nothing
    If Not mXl Is Nothing Then mXl.Teardown: Set mXl = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        mForms.CloseXlBar
    End If
End Sub
