'==============================================================================
' frmImport - Importassistent: 1. Datei und Vorschau, 2. Ziel und Spalten-
' zuordnung mit Typvorschlag, 3. Pruefung mit Fehlerbericht. Uebernommen wird
' erst nach ausdruecklicher Entscheidung; Abbruch veraendert nichts.
'==============================================================================
Option Explicit

Public ResultTid As String
Private mEvts As Collection
Private mSrcData As Variant
Private mPath As String
Private mStep As Long
Private mSrc() As String
Private mTarget() As String        ' "" = nicht importieren | "NEW:n" | FeldID
Private mType() As String
Private mNew() As String
Private mNC As Long
Private mValidRows As Collection
Private mErrs As Collection
Private mPlan As cMap
Private mLoading As Boolean
Private fr1 As MSForms.Frame, fr2 As MSForms.Frame, fr3 As MSForms.Frame
Private lblStep As MSForms.Label

Public Sub Setup(Optional ByVal targetTid As String = "")
    Dim y As Single, w As Single, c As Object, b As MSForms.CommandButton, ids As Variant, i As Long
    Set mEvts = New Collection
    mUI.StyleForm Me
    Me.Width = 780
    If Me.Width > mUI.MaxFormWidth Then Me.Width = mUI.MaxFormWidth
    Me.caption = "Daten importieren"
    mUI.AddLabel Me, "t", "Daten importieren (XLSX oder CSV)", PAD, PAD - 2, 500, 20, True, 13
    Set lblStep = mUI.AddLabel(Me, "step", "", PAD, PAD + 20, Me.InsideWidth - PAD * 2, 15, False, 9, CAccent())
    w = Me.InsideWidth - PAD * 2
    ' Schritt 1
    Set fr1 = mUI.AddFrame(Me, "fr1", PAD - 4, PAD + 42, w + 8, 400)
    mUI.AddLabel fr1, "lp", "Datei", 4, 4, 300, 15, True
    Set c = mUI.AddText(fr1, "path", 4, 21, w - 130)
    Set b = mUI.AddButton(fr1, "browse", "Durchsuchen…", w - 118, 20, 118, , ROW_H): Bind b, "browse"
    mUI.AddLabel fr1, "ls", "Tabellenblatt (XLSX)", 4, 52, 200, 15, True
    Set c = mUI.AddCombo(fr1, "sheet", 4, 69, 200)
    mUI.AddLabel fr1, "lh", "Kopfzeile", 214, 52, 66, 15, True
    Set c = mUI.AddText(fr1, "hdr", 214, 69, 60): c.Text = "1"
    mUI.AddLabel fr1, "ld", "Trennzeichen (CSV)", 284, 52, 140, 15, True
    Set c = mUI.AddCombo(fr1, "delim", 284, 69, 140)
    c.AddItem "Automatisch": c.AddItem "Semikolon ;": c.AddItem "Komma ,": c.AddItem "Tabulator": c.ListIndex = 0
    mUI.AddLabel fr1, "le", "Zeichensatz (CSV)", 434, 52, 140, 15, True
    Set c = mUI.AddCombo(fr1, "enc", 434, 69, 140)
    c.AddItem "Automatisch": c.AddItem "UTF-8": c.AddItem "Windows-1252": c.ListIndex = 0
    mUI.AddLabel fr1, "ldc", "Dezimalzeichen", 584, 52, 140, 15, True
    Set c = mUI.AddCombo(fr1, "dec", 584, 69, 140)
    c.AddItem "Komma (1.234,5)": c.AddItem "Punkt (1,234.5)": c.ListIndex = 0
    Set b = mUI.AddButton(fr1, "read", "Datei lesen", 4, 102, 120, "primary", ROW_H): Bind b, "read"
    mUI.AddLabel fr1, "info", "Die Quelldatei wird nur gelesen (schreibgeschützt, Makros deaktiviert).", 134, 106, w - 140, 15, False, 8.5, CMuted()
    Set c = mUI.AddList(fr1, "prev", 4, 132, w, 250, False, 6)
    ' Schritt 2
    Set fr2 = mUI.AddFrame(Me, "fr2", PAD - 4, PAD + 42, w + 8, 400)
    Set c = mUI.AddOption(fr2, "optNew", "In neue Tabelle importieren:", 4, 4, 210, "ziel"): Bind c, "ziel"
    Set c = mUI.AddText(fr2, "newname", 220, 4, 220)
    Set c = mUI.AddOption(fr2, "optOld", "An bestehende Tabelle anhängen:", 4, 32, 210, "ziel"): Bind c, "ziel"
    Set c = mUI.AddCombo(fr2, "tbl", 220, 32, 220)
    ids = mCfg.TableIds()
    For i = 0 To ArrLen(ids) - 1
        c.AddItem mCfg.TableName(CStr(ids(i)))
    Next
    Bind c, "tbl"
    mUI.AddLabel fr2, "lk", "Schlüssel für Aktualisierung", 460, 4, 250, 15, True
    Set c = mUI.AddCombo(fr2, "key", 460, 21, w - 460)
    mUI.AddLabel fr2, "lk2", "Nur bei bestehender Tabelle: Zeilen mit vorhandenem Schlüsselwert aktualisieren den Eintrag, andere werden neu angelegt.", _
                 460, 46, w - 460, 40, False, 8, CMuted()
    mUI.AddLabel fr2, "lm", "Spaltenzuordnung (Zeile anklicken und unten Ziel/Typ wählen)", 4, 90, 500, 15, True
    Set c = mUI.AddList(fr2, "map", 4, 107, w, 180, False, 4, "180;200;140;" & (w - 540)): Bind c, "map"
    mUI.AddLabel fr2, "lt", "Ziel der markierten Spalte", 4, 294, 250, 15, True
    Set c = mUI.AddCombo(fr2, "target", 4, 311, 300): Bind c, "target"
    mUI.AddLabel fr2, "lty", "Feldtyp (bei neuem Feld)", 314, 294, 200, 15, True
    Set c = mUI.AddCombo(fr2, "type", 314, 311, 200): Bind c, "type"
    Dim code As Variant
    For Each code In FieldTypeCodes()
        If code <> FT_LINK And code <> FT_FORMEL Then c.AddItem FieldTypeName(CStr(code))
    Next
    Set c = mUI.AddCheck(fr2, "label", "Bezeichnungsfeld", 524, 311, 160): Bind c, "label"
    Set c = mUI.AddCheck(fr2, "newopts", "Unbekannte Auswahlwerte als neue Optionen anlegen", 4, 342, 400)
    c.Value = True
    ' Schritt 3
    Set fr3 = mUI.AddFrame(Me, "fr3", PAD - 4, PAD + 42, w + 8, 400)
    mUI.AddLabel fr3, "sum", "", 4, 4, w, 44, False, 10, CText()
    mUI.AddLabel fr3, "le3", "Fehlerhafte Zeilen (werden nicht übernommen)", 4, 52, 400, 15, True
    Set c = mUI.AddList(fr3, "errs", 4, 69, w, 260, False, 4, "70;150;160;" & (w - 390))
    Set b = mUI.AddButton(fr3, "saveerr", "Fehlerbericht speichern (CSV)…", 4, 336, 230, , ROW_H): Bind b, "saveerr"
    ' Navigation
    y = PAD + 42 + 410
    Set b = mUI.AddButton(Me, "back", "Zurück", PAD, y, 100): Bind b, "back"
    Set b = mUI.AddButton(Me, "next", "Weiter", PAD + 108, y, 140, "primary"): Bind b, "next"
    Set b = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 104, y, 104): Bind b, "cancel"
    b.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    If Len(targetTid) > 0 Then
        fr2.Controls("optOld").Value = True
        fr2.Controls("tbl").Value = mCfg.TableName(targetTid)
    Else
        fr2.Controls("optNew").Value = True
    End If
    ShowStep 1
End Sub

Private Sub Bind(ByVal c As Object, ByVal key As String)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, c, key
    mEvts.Add e
End Sub

Private Sub ShowStep(ByVal s As Long)
    mStep = s
    fr1.Visible = (s = 1): fr2.Visible = (s = 2): fr3.Visible = (s = 3)
    Me.Controls("back").Enabled = (s > 1)
    Select Case s
        Case 1: lblStep.caption = "Schritt 1 von 3: Datei wählen und Vorschau prüfen"
            Me.Controls("next").caption = "Weiter"
        Case 2: lblStep.caption = "Schritt 2 von 3: Ziel und Spaltenzuordnung festlegen (Typvorschläge prüfen)"
            Me.Controls("next").caption = "Prüfen"
        Case 3: lblStep.caption = "Schritt 3 von 3: Prüfergebnis. Erst jetzt wird über die Übernahme entschieden"
            Me.Controls("next").caption = "Gültige Zeilen übernehmen"
    End Select
End Sub

Private Function TargetTid() As String
    If fr2.Controls("optOld").Value Then TargetTid = mCfg.TableIdByName(fr2.Controls("tbl").Text)
End Function

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    On Error GoTo failed
    Select Case key
        Case "browse": If evt = "Click" Then Browse
        Case "read": If evt = "Click" Then ReadFile
        Case "ziel", "tbl": If (evt = "Click" Or evt = "Change") And Not mLoading Then InitMapping
        Case "map": If evt = "Click" Or evt = "Change" Then ShowMapRow
        Case "target", "type", "label": If (evt = "Change" Or evt = "Click") And Not mLoading Then UpdateMapRow
        Case "saveerr": If evt = "Click" Then mIO.ExportErrorReport
        Case "back": If evt = "Click" And mStep > 1 Then ShowStep mStep - 1
        Case "next": If evt = "Click" Then NextStep
        Case "cancel": If evt = "Click" Then Me.Hide
    End Select
    Exit Sub
failed:
    mApp.Boundary "Import"
End Sub

Private Sub Browse()
    Dim res As Variant, sheets As Variant, i As Long, c As MSForms.ComboBox
    res = Application.GetOpenFilename("Excel- und CSV-Dateien (*.xlsx;*.xlsm;*.xls;*.csv;*.txt),*.xlsx;*.xlsm;*.xls;*.csv;*.txt", , "Datei für den Import wählen")
    If VarType(res) <> vbString Then Exit Sub
    fr1.Controls("path").Text = res
    Set c = fr1.Controls("sheet")
    c.Clear
    If LCase$(Right$(res, 4)) <> ".csv" And LCase$(Right$(res, 4)) <> ".txt" Then
        sheets = mIO.XlsxSheets(CStr(res))
        For i = 0 To UBound(sheets)
            c.AddItem sheets(i)
        Next
        c.ListIndex = 0
    End If
    ReadFile
End Sub

Private Sub ReadFile()
    Dim p As String, delim As String, enc As String, lb As MSForms.ListBox, r As Long, c As Long, out() As Variant, nr As Long, nc As Long
    p = Trim$(fr1.Controls("path").Text)
    If Len(p) = 0 Or Not FileExists(p) Then mUI.Alert "Bitte eine vorhandene lokale Datei wählen.", "Import": Exit Sub
    Select Case fr1.Controls("delim").ListIndex
        Case 1: delim = ";"
        Case 2: delim = ","
        Case 3: delim = vbTab
        Case Else: delim = "auto"
    End Select
    enc = fr1.Controls("enc").Text
    mApp.BeginWork "Datei wird gelesen …"
    mSrcData = mIO.ReadSource(p, fr1.Controls("sheet").Text, CLng(Val(fr1.Controls("hdr").Text)), delim, enc)
    mApp.EndWork
    mPath = p
    Set lb = fr1.Controls("prev")
    lb.Clear
    If IsEmpty(mSrcData) Then mUI.Alert "Die Datei enthält keine Daten.", "Import": Exit Sub
    nr = Application.Min(UBound(mSrcData, 1), 15)
    nc = Application.Min(UBound(mSrcData, 2), 6)
    ReDim out(0 To nr - 1, 0 To nc - 1)
    For r = 1 To nr
        For c = 1 To nc
            out(r - 1, c - 1) = Truncate(NzS(mSrcData(r, c)), 40)
        Next
    Next
    lb.ColumnCount = nc
    lb.List = out
    fr1.Controls("info").caption = FmtInt(UBound(mSrcData, 1) - 1) & " Datenzeilen, " & UBound(mSrcData, 2) & " Spalten erkannt (Vorschau: erste Zeilen/Spalten; Zeile 1 = Überschriften)."
    InitMapping
End Sub

Private Sub InitMapping()
    Dim c As Long, tid As String, f As cField, lb As MSForms.ListBox, kc As MSForms.ComboBox, nm As String
    If IsEmpty(mSrcData) Then Exit Sub
    mNC = UBound(mSrcData, 2)
    ReDim mSrc(1 To mNC): ReDim mTarget(1 To mNC): ReDim mType(1 To mNC): ReDim mNew(1 To mNC)
    tid = TargetTid()
    For c = 1 To mNC
        nm = Trim$(NzS(mSrcData(1, c)))
        If Len(nm) = 0 Then nm = "Spalte " & c
        mSrc(c) = nm
        mType(c) = mIO.SuggestType(mSrcData, c)
        mNew(c) = Left$(Replace(Replace(nm, ";", " "), ",", " "), 60)
        If Len(tid) > 0 Then
            mTarget(c) = mCfg.FieldIdByName(tid, nm)
            If Len(mTarget(c)) > 0 Then
                If mCfg.Field(mTarget(c)).FType = FT_FORMEL Or mCfg.Field(mTarget(c)).FType = FT_BACKLINK Then mTarget(c) = ""
            End If
        Else
            mTarget(c) = "NEW:" & c
        End If
    Next
    If Len(fr2.Controls("newname").Text) = 0 Then fr2.Controls("newname").Text = Left$(Replace(Mid$(mPath, InStrRev(mPath, "\") + 1), ".", " "), 40)
    Set kc = fr2.Controls("key")
    kc.Clear
    kc.AddItem "(keiner: nur neue Einträge anlegen)"
    If Len(tid) > 0 Then
        For Each f In mCfg.Fields(tid)
            If Not f.IsLinkType And f.FType <> FT_FORMEL And f.FType <> FT_LTEXT Then kc.AddItem f.Name
        Next
    End If
    kc.ListIndex = 0
    kc.Enabled = (Len(tid) > 0)
    RefreshMap 0
End Sub

Private Sub RefreshMap(ByVal sel As Long)
    Dim lb As MSForms.ListBox, c As Long, out() As Variant, ex As String
    Set lb = fr2.Controls("map")
    mLoading = True
    ReDim out(0 To mNC - 1, 0 To 3)
    For c = 1 To mNC
        out(c - 1, 0) = mSrc(c)
        If Len(mTarget(c)) = 0 Then
            out(c - 1, 1) = "(nicht importieren)"
            out(c - 1, 2) = ""
        ElseIf Left$(mTarget(c), 4) = "NEW:" Then
            out(c - 1, 1) = "Neues Feld „" & mNew(c) & "“" & IIf(LabelCol() = c, " " & ChrW(&H2605) & "", "")
            out(c - 1, 2) = FieldTypeName(mType(c))
        Else
            out(c - 1, 1) = mCfg.FieldName(mTarget(c))
            out(c - 1, 2) = FieldTypeName(mCfg.Field(mTarget(c)).FType)
        End If
        If UBound(mSrcData, 1) >= 2 Then ex = NzS(mSrcData(2, c)) Else ex = ""
        out(c - 1, 3) = "z. B. " & Truncate(ex, 40)
    Next
    lb.List = out
    If sel >= 0 And sel < mNC Then lb.ListIndex = sel
    mLoading = False
    ShowMapRow
End Sub

Private Function LabelCol() As Long
    LabelCol = CLng(Val(fr2.Controls("label").Tag))
End Function

Private Sub ShowMapRow()
    Dim c As Long, cb As MSForms.ComboBox, tid As String, f As cField
    c = fr2.Controls("map").ListIndex + 1
    If c < 1 Or mLoading Then Exit Sub
    mLoading = True
    tid = TargetTid()
    Set cb = fr2.Controls("target")
    cb.Clear
    cb.AddItem "(nicht importieren)"
    cb.AddItem "Neues Feld „" & mNew(c) & "“"
    If Len(tid) > 0 Then
        For Each f In mCfg.Fields(tid)
            If f.FType <> FT_FORMEL And f.FType <> FT_BACKLINK Then cb.AddItem f.Name
        Next
    End If
    If Len(mTarget(c)) = 0 Then
        cb.ListIndex = 0
    ElseIf Left$(mTarget(c), 4) = "NEW:" Then
        cb.ListIndex = 1
    Else
        cb.Value = mCfg.FieldName(mTarget(c))
    End If
    fr2.Controls("type").Value = FieldTypeName(mType(c))
    fr2.Controls("type").Enabled = (Left$(mTarget(c), 4) = "NEW:")
    fr2.Controls("label").Value = (LabelCol() = c)
    fr2.Controls("label").Enabled = (Len(tid) = 0)
    mLoading = False
End Sub

Private Sub UpdateMapRow()
    Dim c As Long, cb As MSForms.ComboBox, tid As String
    c = fr2.Controls("map").ListIndex + 1
    If c < 1 Then Exit Sub
    tid = TargetTid()
    Set cb = fr2.Controls("target")
    Select Case cb.ListIndex
        Case 0: mTarget(c) = ""
        Case 1: mTarget(c) = "NEW:" & c
        Case Is > 1: mTarget(c) = mCfg.FieldIdByName(tid, cb.Text)
    End Select
    If Len(fr2.Controls("type").Text) > 0 Then mType(c) = FieldTypeFromName(fr2.Controls("type").Text)
    If fr2.Controls("label").Value Then
        fr2.Controls("label").Tag = CStr(c)
    ElseIf LabelCol() = c Then
        fr2.Controls("label").Tag = ""
    End If
    RefreshMap c - 1
End Sub

Private Sub NextStep()
    Dim tid As String, m As New Collection, c As Long, mapped As Long, ok As Long, updates As Long, e As Variant, out() As Variant, i As Long, msg As String
    Select Case mStep
        Case 1
            If IsEmpty(mSrcData) Then mUI.Alert "Bitte zuerst eine Datei lesen.", "Import": Exit Sub
            If UBound(mSrcData, 1) < 2 Then mUI.Alert "Die Datei enthält nur eine Überschriftenzeile und keine Daten.", "Import": Exit Sub
            ShowStep 2
        Case 2
            tid = TargetTid()
            If fr2.Controls("optOld").Value And Len(tid) = 0 Then mUI.Alert "Bitte eine bestehende Tabelle wählen.", "Import": Exit Sub
            If fr2.Controls("optNew").Value Then
                If Len(Trim$(fr2.Controls("newname").Text)) = 0 Then mUI.Alert "Bitte einen Namen für die neue Tabelle angeben.", "Import": Exit Sub
                If Len(mCfg.TableIdByName(fr2.Controls("newname").Text)) > 0 Then mUI.Alert "Eine Tabelle mit diesem Namen gibt es bereits.", "Import": Exit Sub
            End If
            For c = 1 To mNC
                If Len(mTarget(c)) > 0 Then
                    mapped = mapped + 1
                    m.Add Array(c, mTarget(c), mType(c), mNew(c))
                End If
            Next
            If mapped = 0 Then mUI.Alert "Bitte mindestens eine Spalte einem Feld zuordnen.", "Import": Exit Sub
            Set mPlan = New cMap
            mPlan.SetVal "data", mSrcData
            mPlan.SetVal "tid", tid
            mPlan.SetVal "newTable", Trim$(fr2.Controls("newname").Text)
            mPlan.SetVal "map", m
            mPlan.SetVal "decPoint", (fr1.Controls("dec").ListIndex = 1)
            mPlan.SetVal "createOptions", CBool(fr2.Controls("newopts").Value)
            mPlan.SetVal "labelCol", LabelCol()
            If fr2.Controls("key").ListIndex > 0 Then mPlan.SetVal "key", mCfg.FieldIdByName(tid, fr2.Controls("key").Text)
            mApp.BeginWork "Zeilen werden geprüft …"
            Set mErrs = mIO.CheckImport(mPlan, mValidRows, updates)
            mApp.EndWork
            mIO.StoreErrorReport mErrs, mPath
            msg = FmtInt(UBound(mSrcData, 1) - 1) & " Datenzeilen geprüft:  " & FmtInt(mValidRows.Count) & " gültig" & _
                  IIf(updates > 0, " (davon " & updates & " Aktualisierungen)", "") & ",  " & FmtInt(mErrs.Count) & " Fehler." & vbLf & _
                  IIf(mErrs.Count > 0, "Fehlerhafte Zeilen werden nicht übernommen. Sie können abbrechen, die Quelle korrigieren und erneut importieren. Der Bericht bleibt unter Einrichten > Importbericht verfügbar.", _
                      "Alle Zeilen sind gültig.") & IIf(Len(tid) = 0, "  Die neue Tabelle wird erst beim Übernehmen angelegt.", "")
            fr3.Controls("sum").caption = msg
            fr3.Controls("errs").Clear
            If mErrs.Count > 0 Then
                ReDim out(0 To Application.Min(mErrs.Count, 500) - 1, 0 To 3)
                For Each e In mErrs
                    If i >= 500 Then Exit For
                    out(i, 0) = e(0): out(i, 1) = e(1): out(i, 2) = Truncate(CStr(e(2)), 30): out(i, 3) = e(3)
                    i = i + 1
                Next
                fr3.Controls("errs").List = out
            End If
            Me.Controls("next").Enabled = (mValidRows.Count > 0)
            ShowStep 3
        Case 3
            If mValidRows.Count = 0 Then Exit Sub
            If Not mApp.RequireWritable("Der Import") Then Exit Sub
            If Not mUI.Confirm(FmtInt(mValidRows.Count) & " gültige Zeilen übernehmen" & IIf(mErrs.Count > 0, " und " & mErrs.Count & " fehlerhafte auslassen", "") & "?" & vbLf & _
                               "Vorher wird automatisch eine Sicherung erstellt. Die Übernahme lässt sich in dieser Sitzung mit „Rückgängig“ zurücknehmen.", _
                               "Import übernehmen", "Übernehmen", "Zurück") Then Exit Sub
            ResultTid = mIO.ExecuteImport(mPlan, mValidRows)
            If Len(ResultTid) > 0 Then
                mUI.Alert "Import abgeschlossen: " & FmtInt(mValidRows.Count) & " Zeilen übernommen.", "Import"
                Me.Hide
            End If
    End Select
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
