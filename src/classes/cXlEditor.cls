VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cXlEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cXlEditor - Eingabe einer Excel-Formel fuer eine Formelspalte: Formelfeld,
' anklickbare Spalten (fuegen [Spalte] an der Schreibmarke ein), haeufige
' Funktionen, sofortige Pruefung und Vorschau an echten Eintraegen.
' Wird in frmField (eingebettet) und frmXlBar (schwebende Formelleiste) genutzt.
'==============================================================================
Option Explicit

Private mTid As String
Private mFid As String
Private mEvts As Collection
Private txt As MSForms.TextBox
Private lblPrev As MSForms.Label
Private lblErr As MSForms.Label
Private mHost As Object             ' optional: erhaelt OnFormulaChanged
Private mHeight As Single
Private mLastSel As Long
Private mInserting As Boolean

Public ResultType As String         ' fuer die Vorschau (Ergebnis anzeigen als)
Public Decimals As Long

Private Const FN_LIST As String = "WENN;UND;ODER;RUNDEN;SUMME;MIN;MAX;ABS;HEUTE;TAGE;DATEDIF;JAHR;MONAT;TEXT;VERKETTEN;LÄNGE;ISTLEER;WENNFEHLER"

' Baut die Steuerelemente in container ab (x, y) mit Breite w; liefert die Hoehe
Public Function Build(ByVal container As Object, ByVal tid As String, ByVal fid As String, ByVal def As String, _
                      ByVal x As Single, ByVal y As Single, ByVal w As Single, Optional ByVal host As Object = Nothing) As Single
    Dim f As cField, cx As Single, cy As Single, lbl As MSForms.Label, nm As String, cw As Single, fn As Variant, y0 As Single
    mTid = tid: mFid = fid
    Set mHost = host
    Set mEvts = New Collection
    ResultType = FT_ZAHL: Decimals = 2
    y0 = y
    mUI.AddLabel container, "xl_l1", "Formel", x, y, 120, 15, True
    mUI.AddLabel container, "xl_l1b", "wie in Excel, Spalten in [ ], z. B. =[Preis]*[Menge] oder =WENN([Frist]<HEUTE();""überfällig"";"""")", _
                 x + 54, y + 1, w - 54, 15, False, 8, CMuted()
    y = y + 17
    ' einzeilig: mehrzeilige Felder nehmen in nicht modalen Fenstern den Fokus per Klick
    ' nicht zuverlaessig an; Formeln sind ohnehin einzeilig
    Set txt = mUI.AddText(container, "xl_txt", x, y, w, 26, False)
    txt.EnterFieldBehavior = fmEnterFieldBehaviorRecallSelection
    txt.Font.Name = "Consolas": txt.Font.Size = 10.5
    txt.Text = IIf(mCalc.IsXl(def), mCalc.XlDisplay(def), "")
    mLastSel = Len(txt.Text)
    Bind txt, "txt"
    y = y + 32
    ' Spalten zum Anklicken
    mUI.AddLabel container, "xl_l2", "Spalte anklicken = einfügen", x, y, w, 14, False, 8, CMuted()
    y = y + 15
    cx = x: cy = y
    For Each f In mCfg.Fields(tid)
        If f.ID <> fid And f.Kind <> FT_BUTTON Then
            nm = f.Name
            cw = Len(nm) * 5.6 + 26
            If cw > w Then cw = w
            If cx + cw > x + w And cx > x Then cx = x: cy = cy + 23
            Chip container, "xl_c_" & f.ID, nm, cx, cy, cw, CAccentSoft(), CText(), "col:" & f.ID
            cx = cx + cw + 4
        End If
    Next
    y = cy + 24
    ' haeufige Funktionen
    mUI.AddLabel container, "xl_l3", "Funktionen", x, y, w, 14, False, 8, CMuted()
    y = y + 15
    cx = x: cy = y
    For Each fn In Split(FN_LIST, ";")
        cw = Len(fn) * 6 + 14
        If cx + cw > x + w And cx > x Then cx = x: cy = cy + 23
        Chip container, "xl_f_" & fn, CStr(fn), cx, cy, cw, CSurface2(), CText(), "fn:" & fn, "Consolas"
        cx = cx + cw + 4
    Next
    y = cy + 26
    Set lblErr = mUI.AddLabel(container, "xl_err", "", x, y, w, 40, False, 8.5, CDanger())
    Set lblPrev = mUI.AddLabel(container, "xl_prev", "", x, y, w, 40, False, 8.5, CText())
    y = y + 44
    mHeight = y - y0
    Build = mHeight
    Refresh
End Function

' Kaestchen = Hintergrundflaeche + darueber senkrecht mittig gesetzte Beschriftung
' (MSForms-Labels richten Text immer oben aus). Beide reagieren auf Klick.
Private Sub Chip(ByVal container As Object, ByVal nm As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, _
                 ByVal w As Single, ByVal bg As Long, ByVal fg As Long, ByVal key As String, Optional ByVal fontName As String = "")
    Const CH As Single = 20
    Dim lbl As MSForms.Label, tx As MSForms.Label, tip As String
    tip = "Klicken fügt „" & caption & "“ an der Schreibmarke ein"
    Set lbl = mUI.AddLabel(container, nm, "", x, y, w, CH)
    lbl.BackStyle = fmBackStyleOpaque
    lbl.BackColor = bg
    lbl.MousePointer = fmMousePointerArrow
    lbl.ControlTipText = tip
    Bind lbl, key
    Set tx = mUI.AddLabel(container, nm & "_t", caption, x, y, w, CH, False, 8.5, fg)
    If Len(fontName) > 0 Then tx.Font.Name = fontName
    tx.WordWrap = False
    tx.AutoSize = True
    tx.AutoSize = False
    tx.Left = x: tx.Width = w
    tx.Top = y + (CH - tx.Height) / 2
    tx.TextAlign = fmTextAlignCenter
    tx.MousePointer = fmMousePointerArrow
    tx.ControlTipText = tip
    Bind tx, key
End Sub

Private Sub Bind(ByVal c As Object, ByVal key As String)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, c, key
    mEvts.Add e
End Sub

Public Sub Teardown()
    Dim e As cEvt
    If mEvts Is Nothing Then Exit Sub
    For Each e In mEvts: e.Unbind: Next
    Set mEvts = Nothing
    Set mHost = Nothing
End Sub

Public Property Get FormulaText() As String
    FormulaText = txt.Text
End Property

Public Sub FocusText()
    On Error Resume Next
    If TypeName(txt.Parent) = "Frame" Then txt.Parent.SetFocus
    txt.SetFocus
    txt.SelStart = mLastSel
End Sub

' Spalte an der Schreibmarke einfuegen (auch aus der Tabelle: Klick auf eine Spalte)
Public Sub InsertField(ByVal fid As String)
    If Not mCfg.FieldExists(fid) Or fid = mFid Then Exit Sub
    If mCfg.Field(fid).TableID <> mTid Then Exit Sub
    InsertText "[" & mCfg.FieldName(fid) & "]"
End Sub

Public Sub InsertText(ByVal s As String)
    Dim p As Long, cur As String
    cur = txt.Text
    p = mLastSel
    If p < 0 Or p > Len(cur) Then p = Len(cur)
    If Len(cur) = 0 And Left$(s, 1) <> "=" Then s = "=" & s
    mInserting = True
    txt.Text = Left$(cur, p) & s & Mid$(cur, p + 1)
    mInserting = False
    mLastSel = p + Len(s)
    On Error Resume Next
    txt.SelStart = mLastSel
End Sub

' Prueft die Eingabe; liefert die gespeicherte Form oder "" mit msg
Public Function Definition(ByRef msg As String) As String
    Dim def As String, path As String
    def = mCalc.XlFromDisplay(mTid, txt.Text, msg)
    If Len(msg) > 0 Then Exit Function
    msg = mCalc.XlCheck(def)
    If Len(msg) > 0 Then Exit Function
    If Len(mFid) > 0 Then
        If mCalc.HasCycle(mFid, def, path) Then msg = "Zirkelbezug: " & path & ".": Exit Function
    End If
    Definition = def
End Function

' Pruefung und Vorschau an bis zu drei Eintraegen
Public Sub Refresh()
    Dim def As String, msg As String, t As cTableData, rows() As Long, n As Long, i As Long, res() As Variant, s As String, tmp As New cField
    On Error GoTo failed
    If lblErr Is Nothing Or txt Is Nothing Then Exit Sub
    lblErr.caption = "": lblPrev.caption = ""
    If Len(Trim$(txt.Text)) = 0 Or Trim$(txt.Text) = "=" Then
        lblPrev.ForeColor = CMuted()
        lblPrev.caption = "Formel eingeben: Das Ergebnis erscheint hier sofort für die ersten Einträge."
        Exit Sub
    End If
    def = Definition(msg)
    If Len(msg) > 0 Then lblErr.caption = ChrW(&H26A0) & "  " & msg: Exit Sub
    Set t = mData.TData(mTid)
    If t.RowCount = 0 Then lblPrev.caption = "Formel ist gültig. Noch keine Einträge für eine Vorschau.": Exit Sub
    For i = 1 To t.RowCount
        If Not t.IsArchived(i) Then
            n = n + 1
            ReDim Preserve rows(1 To n)
            rows(n) = i
            If n >= 3 Then Exit For
        End If
    Next
    If n = 0 Then lblPrev.caption = "Formel ist gültig.": Exit Sub
    res = mCalc.XlEvalRows(def, t, rows, n, ResultType)
    tmp.FType = FT_FORMEL: tmp.ResultType = ResultType: tmp.Decimals = Decimals
    For i = 1 To n
        s = s & IIf(i > 1, "   ·   ", "") & Truncate(mData.RecordLabel(mTid, t.IdAt(rows(i))), 24) & " " & ChrW(&H2192) & " " & _
            IIf(IsBlank(res(i)), "(leer)", mFmt.DisplayText(tmp, res(i)))
    Next
    lblPrev.ForeColor = CText()
    lblPrev.caption = ChrW(&H2713) & "  Vorschau:  " & s
    If Not mHost Is Nothing Then
        ' erster nicht leerer Wert bestimmt den vorgeschlagenen Ergebnistyp
        For i = 1 To n
            If Not IsBlank(res(i)) Then Exit For
        Next
        On Error Resume Next
        If i <= n Then mHost.OnFormulaPreview res(i)
    End If
    Exit Sub
failed:
    lblErr.caption = ChrW(&H26A0) & "  " & Err.Description
End Sub

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    On Error GoTo failed
    Select Case True
        Case key = "txt"
            ' Schreibmarke merken: nach einem Klick in die Tabelle ist sie im Feld nicht mehr gesetzt
            If evt = "MouseUp" Then
                ' In nicht modalen Fenstern nimmt das mehrzeilige Feld den Fokus per Klick
                ' nicht immer an: dann gezielt setzen (Schreibmarke ans Ende)
                If Not txt.Parent.ActiveControl Is txt Then
                    txt.SetFocus
                    txt.SelStart = Len(txt.Text)
                End If
            End If
            If evt = "Change" Or evt = "KeyUp" Or evt = "MouseUp" Then
                If Not mInserting Then mLastSel = txt.SelStart
                If evt = "Change" Then Refresh
            End If
        Case Left$(key, 4) = "col:"
            If evt = "Click" Then InsertField Mid$(key, 5): FocusText
        Case Left$(key, 3) = "fn:"
            If evt = "Click" Then
                If Mid$(key, 4) = "HEUTE" Then InsertText "HEUTE()" Else InsertText Mid$(key, 4) & "("
                FocusText
            End If
    End Select
    Exit Sub
failed:
    mApp.Boundary "Formel"
End Sub
