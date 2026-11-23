Attribute VB_Name = "mIO"
'==============================================================================
' mIO - Import (lokale XLSX/CSV mit Vorschau, Typvorschlag, Zuordnung,
' Fehlerbericht) und Export (gefilterte Ansicht als XLSX/CSV, Dashboard als
' XLSX/PDF). Quelldateien werden nur gelesen; Makros aus Quellen werden nicht
' ausgefuehrt. Exporte enthalten keine Makros und keine Verbindung zur Arbeitsdatei.
'==============================================================================
Option Explicit

'==============================================================================
' Lesen
'==============================================================================
' Oeffnet eine XLSX-Datei schreibgeschuetzt ohne Makros und Verknuepfungen
' Zahlen der Spalte col (Datenzeilen 2..n+1) alle ganzzahlig?
Private Function ColWhole(ByRef a() As Variant, ByVal n As Long, ByVal col As Long) As Boolean
    Dim k As Long
    For k = 2 To n + 1
        If VarType(a(k, col)) = vbDouble Then
            If a(k, col) <> Fix(a(k, col)) Then Exit Function
        End If
    Next
    ColWhole = True
End Function

Public Function OpenSourceWorkbook(ByVal path As String) As Workbook
    Dim prevSec As MsoAutomationSecurity, prevEv As Boolean, prevAlerts As Boolean, wb As Workbook
    prevSec = Application.AutomationSecurity
    prevEv = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo failed
    Set wb = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True, Notify:=False)
    wb.Windows(1).Visible = False
    Application.AutomationSecurity = prevSec
    Application.EnableEvents = prevEv
    Application.DisplayAlerts = prevAlerts
    Set OpenSourceWorkbook = wb
    Exit Function
failed:
    Application.AutomationSecurity = prevSec
    Application.EnableEvents = prevEv
    Application.DisplayAlerts = prevAlerts
    Err.Raise Err.Number, , "Die Datei konnte nicht gelesen werden: " & Err.Description
End Function

Public Sub CloseSource(ByVal wb As Workbook)
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    ThisWorkbook.Activate
End Sub

Public Function XlsxSheets(ByVal path As String) As Variant
    Dim wb As Workbook, a() As Variant, i As Long
    Set wb = OpenSourceWorkbook(path)
    ReDim a(0 To wb.Worksheets.Count - 1)
    For i = 1 To wb.Worksheets.Count
        a(i - 1) = wb.Worksheets(i).Name
    Next
    CloseSource wb
    XlsxSheets = a
End Function

' Liest eine Quelle als 2D-Array (1..Zeilen, 1..Spalten); Zeile 1 = Ueberschriften
Public Function ReadSource(ByVal path As String, ByVal sheetName As String, ByVal headerRow As Long, ByVal delim As String, _
                           ByVal encoding As String) As Variant
    Dim ext As String, wb As Workbook, ws As Worksheet, rng As Range, v As Variant, lastR As Long, lastC As Long, res() As Variant
    Dim r As Long, c As Long
    ext = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    If ext = "csv" Or ext = "txt" Then
        ReadSource = ReadCsv(path, delim, encoding)
        Exit Function
    End If
    Set wb = OpenSourceWorkbook(path)
    On Error GoTo failed
    If Len(sheetName) > 0 Then Set ws = wb.Worksheets(sheetName) Else Set ws = wb.Worksheets(1)
    Set rng = ws.UsedRange
    lastR = rng.Row + rng.Rows.Count - 1
    lastC = rng.Column + rng.Columns.Count - 1
    If headerRow < 1 Then headerRow = 1
    If lastR < headerRow Then
        CloseSource wb
        ReadSource = Empty
        Exit Function
    End If
    v = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastR, lastC)).Value
    CloseSource wb
    If Not IsArray(v) Then
        ReDim res(1 To 1, 1 To 1)
        res(1, 1) = v
        ReadSource = res
    Else
        ReadSource = TrimEmpty(v)
    End If
    Exit Function
failed:
    Dim msg As String
    msg = Err.Description
    CloseSource wb
    RaiseUser "Das Blatt konnte nicht gelesen werden: " & msg
End Function

' Entfernt vollstaendig leere Zeilen/Spalten am Ende
Private Function TrimEmpty(ByVal v As Variant) As Variant
    Dim r As Long, c As Long, lastR As Long, lastC As Long, res() As Variant
    For r = 1 To UBound(v, 1)
        For c = 1 To UBound(v, 2)
            If Not IsBlank(v(r, c)) Then
                If r > lastR Then lastR = r
                If c > lastC Then lastC = c
            End If
        Next
    Next
    If lastR = 0 Then TrimEmpty = Empty: Exit Function
    ReDim res(1 To lastR, 1 To lastC)
    For r = 1 To lastR
        For c = 1 To lastC
            If IsError(v(r, c)) Then res(r, c) = "#FEHLER" Else res(r, c) = v(r, c)
        Next
    Next
    TrimEmpty = res
End Function

' CSV lesen: Zeichensatz (UTF-8/Windows-1252) und Trennzeichen automatisch oder gewaehlt
Public Function ReadCsv(ByVal path As String, ByVal delim As String, ByVal encoding As String) As Variant
    Dim b() As Byte, text As String, ok As Boolean, startPos As Long
    b = ReadFileBytes(path)
    If ArrLen(b) = 0 Then ReadCsv = Empty: Exit Function
    startPos = 0
    If UBound(b) >= 2 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then startPos = 3
    End If
    If encoding = "Windows-1252" Then
        text = Ansi1252Decode(b, startPos)
    Else
        text = Utf8Decode(b, startPos, ok)
        If Not ok Then
            If encoding = "UTF-8" Then RaiseUser "Die Datei ist kein gültiges UTF-8. Bitte Zeichensatz „Windows-1252“ wählen."
            text = Ansi1252Decode(b, startPos)
        End If
    End If
    If delim = "" Or delim = "auto" Then delim = DetectDelim(text)
    ReadCsv = ParseCsv(text, delim)
End Function

Private Function DetectDelim(ByVal text As String) As String
    Dim firstLine As String, p As Long, nS As Long, nC As Long, nT As Long
    p = InStr(text, vbLf)
    If p > 0 Then firstLine = Left$(text, p) Else firstLine = text
    nS = Len(firstLine) - Len(Replace(firstLine, ";", ""))
    nC = Len(firstLine) - Len(Replace(firstLine, ",", ""))
    nT = Len(firstLine) - Len(Replace(firstLine, vbTab, ""))
    If nT > nS And nT > nC Then
        DetectDelim = vbTab
    ElseIf nC > nS Then
        DetectDelim = ","
    Else
        DetectDelim = ";"
    End If
End Function

' RFC-4180-aehnlich: Anfuehrungszeichen, doppelte Anfuehrungszeichen, Zeilenumbrueche in Feldern
Public Function ParseCsv(ByVal text As String, ByVal delim As String) As Variant
    Dim rows As New Collection, cur As Collection, fld As String, i As Long, n As Long, ch As String, inQ As Boolean
    Dim maxC As Long, res() As Variant, r As Long, c As Long, row As Collection, sb() As String, sbn As Long
    n = Len(text)
    Set cur = New Collection
    i = 1
    Do While i <= n
        ch = Mid$(text, i, 1)
        If inQ Then
            If ch = """" Then
                If i < n And Mid$(text, i + 1, 1) = """" Then
                    fld = fld & """": i = i + 1
                Else
                    inQ = False
                End If
            Else
                fld = fld & ch
            End If
        Else
            If ch = """" And Len(fld) = 0 Then
                inQ = True
            ElseIf ch = delim Then
                cur.Add fld: fld = ""
            ElseIf ch = vbCr Then
                ' ignorieren (CRLF)
            ElseIf ch = vbLf Then
                cur.Add fld: fld = ""
                rows.Add cur
                Set cur = New Collection
            Else
                fld = fld & ch
            End If
        End If
        i = i + 1
    Loop
    If Len(fld) > 0 Or cur.Count > 0 Then
        cur.Add fld
        rows.Add cur
    End If
    For Each row In rows
        If row.Count > maxC Then maxC = row.Count
    Next
    If rows.Count = 0 Or maxC = 0 Then ParseCsv = Empty: Exit Function
    ReDim res(1 To rows.Count, 1 To maxC)
    r = 0
    For Each row In rows
        r = r + 1
        For c = 1 To row.Count
            res(r, c) = row(c)
        Next
    Next
    ParseCsv = TrimEmpty(res)
End Function

' Typvorschlag fuer eine Quellspalte
Public Function SuggestType(ByVal data As Variant, ByVal col As Long) As String
    Dim r As Long, v As Variant, n As Long, nNum As Long, nDate As Long, nBool As Long, nLong As Long, distinct As New cMap, d As Double, dt As Date, ht As Boolean
    Dim nTime As Long, s As String, nEuro As Long, nPct As Long
    For r = 2 To UBound(data, 1)
        v = data(r, col)
        If Not IsBlank(v) Then
            n = n + 1
            If VarType(v) = vbDate Then
                nDate = nDate + 1
                If CDbl(v) <> Int(CDbl(v)) Then nTime = nTime + 1
            ElseIf VarType(v) = vbBoolean Then
                nBool = nBool + 1
            ElseIf VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
                nNum = nNum + 1
            Else
                s = Trim$(CStr(v))
                Select Case LCase$(s)
                    Case "ja", "nein", "wahr", "falsch", "x"
                        nBool = nBool + 1
                    Case Else
                        If ParseDateDE(s, dt, ht) Then
                            nDate = nDate + 1
                            If ht Then nTime = nTime + 1
                        ElseIf ParseNumberDE(s, d) Then
                            nNum = nNum + 1
                            If InStr(s, "€") > 0 Then nEuro = nEuro + 1
                            If InStr(s, "%") > 0 Then nPct = nPct + 1
                        End If
                End Select
                If Len(s) > 120 Or InStr(s, vbLf) > 0 Then nLong = nLong + 1
            End If
            If distinct.Count < 40 Then distinct.SetVal LCase$(Trim$(CStr(v))), True
        End If
        If r > 2000 Then Exit For
    Next
    If n = 0 Then SuggestType = FT_TEXT: Exit Function
    If nDate = n Then SuggestType = IIf(nTime > 0, FT_DATUMZEIT, FT_DATUM): Exit Function
    If nBool = n Then SuggestType = FT_JANEIN: Exit Function
    If nNum = n Then
        If nEuro > 0 Then
            SuggestType = FT_GELD
        ElseIf nPct > 0 Then
            SuggestType = FT_PROZENT
        Else
            SuggestType = FT_ZAHL
        End If
        Exit Function
    End If
    If nLong > 0 Then SuggestType = FT_LTEXT: Exit Function
    If distinct.Count <= 12 And n >= distinct.Count * 2 And n >= 6 Then SuggestType = FT_AUSWAHL: Exit Function
    SuggestType = FT_TEXT
End Function

' Zahlentext mit Dezimalpunkt in deutsche Schreibweise umwandeln
Public Function NormalizeNumber(ByVal s As String, ByVal decPoint As Boolean) As String
    If Not decPoint Then NormalizeNumber = s: Exit Function
    s = Replace(s, ",", "")
    s = Replace(s, ".", ",")
    NormalizeNumber = s
End Function

'==============================================================================
' Import pruefen und ausfuehren
' Plan (cMap): data, newTable (Name oder ""), tid, map (Collection Array(col, target, type, name)),
'              key (FeldID), decPoint (Boolean), createOptions (Boolean), labelCol (Long)
'==============================================================================
' Liefert Collection von Fehlern Array(Quellzeile, Spalte, Wert, Grund) und zaehlt gueltige Zeilen
Public Function CheckImport(ByVal plan As cMap, ByRef validRows As Collection, ByRef updates As Long) As Collection
    Dim errs As New Collection, data As Variant, r As Long, m As Variant, f As cField, v As Variant, outV As Variant, msg As String
    Dim ok As Boolean, rowVals As cMap, linkVals As cMap, tid As String, keyF As String, keyCol As Long, keySeen As New cMap, kv As String
    Dim existing As String, recErrs As cMap, ek As Variant, newOpts As cMap, tt As String, names As Variant, i As Long, lst As String, rid As String
    Dim s As String, tf As cField, uniqSeen As New cMap, anyVal As Boolean, decPoint As Boolean, createOpts As Boolean
    Dim maps As Collection, fl() As cField, mi As Long
    data = plan.GetVal("data")
    tid = plan.GetVal("tid", "")
    keyF = plan.GetVal("key", "")
    decPoint = plan.GetVal("decPoint", False)
    createOpts = plan.GetVal("createOptions", True)
    Set validRows = New Collection
    Set newOpts = New cMap
    ' Feldbeschreibungen einmal je Spalte (nicht je Zelle) bilden
    Set maps = plan.GetVal("map")
    ReDim fl(0 To maps.Count)
    mi = 0
    For Each m In maps
        mi = mi + 1
        If Len(m(1)) > 0 Then Set fl(mi) = ImportField(m, tid)
    Next
    For r = 2 To UBound(data, 1)
        ok = True
        anyVal = False
        Set rowVals = New cMap
        Set linkVals = New cMap
        mi = 0
        For Each m In maps
            mi = mi + 1
            If Len(m(1)) > 0 Then
                Set f = fl(mi)
                v = data(r, m(0))
                If Not IsBlank(v) Then anyVal = True
                If VarType(v) = vbString And (f.IsNumeric_) Then v = NormalizeNumber(CStr(v), decPoint)
                If f.FType = FT_LINK Then
                    ' Verknuepfung ueber Bezeichnung der Zieleintraege
                    lst = ""
                    If Not IsBlank(v) Then
                        names = Split(Replace(CStr(v), ",", ";"), ";")
                        For i = 0 To UBound(names)
                            If Len(Trim$(names(i))) > 0 Then
                                rid = FindByLabel(f.TargetTable, Trim$(names(i)), msg)
                                If Len(rid) = 0 Then
                                    errs.Add Array(r, CStr(data(1, m(0))), CStr(v), msg): ok = False
                                Else
                                    lst = ListAdd(lst, rid)
                                End If
                            End If
                        Next
                        If ListCount(lst) > 1 And Not f.Multi Then errs.Add Array(r, CStr(data(1, m(0))), CStr(v), "Nur ein verknüpfter Eintrag erlaubt."): ok = False
                    End If
                    linkVals.SetVal f.ID, lst
                ElseIf f.IsOptionType And Not IsBlank(v) And (createOpts Or Left$(m(1), 4) = "NEW:") Then
                    ' unbekannte Optionen werden (nach Bestaetigung) angelegt
                    s = ""
                    names = IIf(f.FType = FT_MEHRFACH, Split(Replace(CStr(v), ",", ";"), ";"), Array(CStr(v)))
                    For i = 0 To UBound(names)
                        If Len(Trim$(names(i))) > 0 Then
                            If InStr(Trim$(names(i)), ";") > 0 Then errs.Add Array(r, CStr(data(1, m(0))), CStr(v), "Ungültiger Optionsname."): ok = False
                            If Left$(m(1), 4) = "NEW:" Or Len(mCfg.OptIdByName(f.ID, Trim$(names(i)))) = 0 Then newOpts.SetVal f.ID & "|" & LCase$(Trim$(names(i))), Array(m(1), Trim$(names(i)))
                            s = s & IIf(Len(s) > 0, ";", "") & Trim$(names(i))
                        End If
                    Next
                    rowVals.SetNew "@" & m(1), s
                ElseIf mValid.ParseCell(f, v, outV, msg) Then
                    rowVals.SetNew "@" & m(1), outV
                Else
                    errs.Add Array(r, CStr(data(1, m(0))), NzS(v), msg)
                    ok = False
                End If
            End If
        Next
        If Not anyVal Then ok = False   ' leere Zeile still ueberspringen
        ' Schluessel / Aktualisierung
        If ok And Len(keyF) > 0 And Len(tid) > 0 Then
            kv = NormKey(rowVals.GetVal("@" & keyF, Empty))
            If Len(kv) = 0 Then
                errs.Add Array(r, mCfg.FieldName(keyF), "", "Schlüsselwert fehlt."): ok = False
            ElseIf keySeen.Exists(kv) Then
                errs.Add Array(r, mCfg.FieldName(keyF), "", "Schlüssel kommt in der Datei mehrfach vor (auch Zeile " & keySeen.GetVal(kv) & ")."): ok = False
            Else
                keySeen.SetVal kv, r
                existing = mValid.FindDuplicateValue(tid, "", keyF, rowVals.GetVal("@" & keyF))
                rowVals.SetVal "#rid", existing
                If Len(existing) > 0 Then updates = updates + 1
            End If
        End If
        ' Pflicht/Eindeutigkeit in bestehender Tabelle
        If ok And Len(tid) > 0 Then
            Dim vals As New cMap, k As Variant
            Set vals = New cMap
            For Each k In rowVals.Keys
                If Left$(CStr(k), 1) = "@" Then
                    If Left$(CStr(k), 5) <> "@NEW:" Then
                        If mCfg.Field(Mid$(CStr(k), 2)).IsOptionType Then
                            vals.SetVal Mid$(CStr(k), 2), "x"   ' Wert vorhanden (Option wird angelegt)
                        Else
                            vals.SetVal Mid$(CStr(k), 2), rowVals.GetVal(CStr(k))
                        End If
                    End If
                End If
            Next
            If Not mValid.ValidateRecord(tid, NzS(rowVals.GetVal("#rid", "")), vals, linkVals, recErrs) Then
                For Each ek In recErrs.Keys
                    errs.Add Array(r, mCfg.FieldName(CStr(ek)), "", CStr(recErrs.GetVal(CStr(ek)))): ok = False
                Next
            End If
            ' Eindeutigkeit innerhalb der Importdatei
            If ok Then
                For Each k In vals.Keys
                    If mCfg.Field(CStr(k)).Unique Then
                        kv = CStr(k) & "|" & NormKey(vals.GetVal(CStr(k)))
                        If Len(NormKey(vals.GetVal(CStr(k)))) > 0 Then
                            If uniqSeen.Exists(kv) Then
                                errs.Add Array(r, mCfg.FieldName(CStr(k)), "", "Eindeutiger Wert kommt in der Datei mehrfach vor (auch Zeile " & uniqSeen.GetVal(kv) & ")."): ok = False
                            Else
                                uniqSeen.SetVal kv, r
                            End If
                        End If
                    End If
                Next
            End If
        End If
        If ok Then
            rowVals.SetVal "#row", r
            rowVals.SetVal "#links", linkVals
            validRows.Add rowVals
        End If
    Next
    plan.SetVal "newOptions", newOpts
    Set CheckImport = errs
End Function

' Feldbeschreibung einer Zuordnung (bestehend oder neu anzulegen)
Private Function ImportField(ByVal m As Variant, ByVal tid As String) As cField
    Dim f As cField
    If Left$(m(1), 4) = "NEW:" Then
        Set f = New cField
        f.ID = m(1)
        ' Art (z. B. E-Mail) und Basistyp (Text) wie bei einem angelegten Feld, damit Pruefung und Wert stimmen
        f.Kind = m(2)
        f.FType = BaseTypeOf(CStr(m(2)))
        f.Name = m(3)
        f.Decimals = IIf(f.FType = FT_ZAHL, 6, 2)
        f.AllowEmpty = True
        f.Status = "aktiv"
        Set ImportField = f
    Else
        Set ImportField = mCfg.Field(CStr(m(1)))
    End If
End Function

Private Function FindByLabel(ByVal tid As String, ByVal lbl As String, ByRef msg As String) As String
    Dim t As cTableData, r As Long, lf As String, c As Long, n As Long, found As String
    Set t = mData.TData(tid)
    lf = mCfg.LabelField(tid)
    c = t.ColOf(lf)
    If c = 0 Then msg = "Zieltabelle hat kein Bezeichnungsfeld.": Exit Function
    For r = 1 To t.RowCount
        If StrComp(Trim$(NzS(t.Cell(r, c))), lbl, vbTextCompare) = 0 Or StrComp(t.IdAt(r), lbl, vbTextCompare) = 0 Then
            n = n + 1
            found = t.IdAt(r)
        End If
    Next
    If n = 0 Then msg = "Kein Eintrag „" & lbl & "“ in „" & mCfg.TableName(tid) & "“ gefunden."
    If n > 1 Then msg = "„" & lbl & "“ ist in „" & mCfg.TableName(tid) & "“ nicht eindeutig (" & n & " Treffer); bitte die ID angeben.": found = ""
    FindByLabel = found
End Function

' Fuehrt einen geprueften Import aus (nur gueltige Zeilen). Liefert Tabellen-ID.
Public Function ExecuteImport(ByVal plan As cMap, ByVal validRows As Collection) As String
    Dim tid As String, m As Variant, fmap As New cMap, fid As String, props As cMap, k As Variant, newOpts As cMap, a As Variant
    Dim t As cTableData, rows() As Variant, n As Long, rv As cMap, c As Long, v As Variant, i As Long, rid As String, s As String
    Dim names As Variant, j As Long, upd As cMap, baseNo As Long, rels() As Variant, nr As Long, links As cMap, f As cField, inserts As Long
    Dim ord As cMap, nmv As Variant, pk As Variant
    tid = plan.GetVal("tid", "")
    If Not mBackup.AutoBackup("vor dem Import") Then Exit Function
    mApp.BeginWork "Import läuft …"
    On Error GoTo failed
    ' Struktur (neue Tabelle/Felder), vor der Datentransaktion
    If Len(tid) = 0 Then
        tid = mStruct.CreateTable(CStr(plan.GetVal("newTable")), CStr(plan.GetVal("tableDesc", "Importiert am " & FmtDate(Date))), &HE8B7&)
        mStruct.CreateDefaultView tid
    End If
    For Each m In plan.GetVal("map")
        If Left$(m(1), 4) = "NEW:" Then
            Set props = New cMap
            ' optionale Feldeigenschaften als 5. Element der Zuordnung (z. B. Dezimalstellen)
            If UBound(m) >= 4 Then
                If IsObject(m(4)) Then
                    For Each pk In m(4).Keys
                        props.SetVal CStr(pk), m(4).GetVal(CStr(pk))
                    Next
                End If
            End If
            If m(2) = FT_ZAHL And Not props.Exists("Dezimalstellen") Then props.SetVal "Dezimalstellen", 2
            fid = mStruct.AddField(tid, CStr(m(3)), CStr(m(2)), props)
            fmap.SetVal CStr(m(1)), fid
            If CLng(plan.GetVal("labelCol", 0)) = CLng(m(0)) Then mStruct.SetLabelField tid, fid
        ElseIf Len(m(1)) > 0 Then
            fmap.SetVal CStr(m(1)), CStr(m(1))
        End If
    Next
    ' Optionen in vorgegebener Reihenfolge (z. B. Codeplan einer Befragung), danach weitere aus den Daten;
    ' je Feld in einem Schritt (bei Hunderten Optionen sonst Minuten)
    Dim optAdd As New cMap, lst As Collection
    If plan.Exists("optionOrder") Then
        Set ord = plan.GetVal("optionOrder")
        For Each k In ord.Keys
            fid = CStr(fmap.GetVal(CStr(k), ""))
            If Len(fid) > 0 Then
                If Not optAdd.Exists(fid) Then optAdd.SetVal fid, New Collection
                Set lst = optAdd.GetVal(fid)
                For Each nmv In ord.GetVal(CStr(k))
                    If Len(Trim$(CStr(nmv))) > 0 Then lst.Add Trim$(CStr(nmv))
                Next
            End If
        Next
    End If
    ' neue Optionen
    Set newOpts = plan.GetVal("newOptions")
    For Each k In newOpts.Keys
        a = newOpts.GetVal(CStr(k))
        fid = CStr(fmap.GetVal(CStr(a(0))))
        If Len(fid) > 0 Then
            If Not optAdd.Exists(fid) Then optAdd.SetVal fid, New Collection
            Set lst = optAdd.GetVal(fid)
            lst.Add CStr(a(1))
        End If
    Next
    For Each k In optAdd.Keys
        mStruct.AddOptionsBulk CStr(k), optAdd.GetVal(CStr(k))
    Next
    mStruct.RebuildLists
    ' Options-IDs je Auswahlfeld einmal nachschlagen (Name -> ID, erste Option gleichen Namens wie OptIdByName)
    Dim optMaps As New cMap, om As cMap, oid As Variant, onm As String, uid As String
    For Each k In fmap.Keys
        fid = CStr(fmap.GetVal(CStr(k)))
        Set f = mCfg.Field(fid)
        If Not f Is Nothing Then
            If f.IsOptionType And Not optMaps.Exists(fid) Then
                Set om = New cMap
                For Each oid In mCfg.OptionIds(fid)
                    onm = Trim$(mCfg.OptName(CStr(oid)))
                    If Len(onm) > 0 Then
                        If Not om.Exists(onm) Then om.SetVal onm, CStr(oid)
                    End If
                Next
                optMaps.SetVal fid, om
            End If
        End If
    Next
    uid = UserId()
    ' je Spalte der Pruefung einmal: Zielfeld und ob es ein Auswahlfeld ist (statt je Zelle nachzuschlagen)
    Dim kinfo As New cMap, ki As Variant
    For Each k In fmap.Keys
        fid = CStr(fmap.GetVal(CStr(k)))
        Set f = mCfg.Field(fid)
        If Not f Is Nothing Then kinfo.SetVal "@" & CStr(k), Array(fid, f.IsOptionType)
    Next
    ' Daten
    Set t = mData.TData(tid)
    ReDim rows(1 To Application.Max(1, validRows.Count), 1 To t.ColCount)
    ReDim rels(1 To Application.Max(1, validRows.Count * 4), 1 To 4)
    baseNo = mCfg.ReserveIds("R", validRows.Count)
    mData.TxBegin "Import nach „" & mCfg.TableName(tid) & "“"
    For Each rv In validRows
        rid = NzS(rv.GetVal("#rid", ""))
        Set upd = New cMap
        For Each k In rv.Keys
            If Left$(CStr(k), 1) = "@" Then
                ki = kinfo.GetVal(CStr(k))
                fid = CStr(ki(0))
                v = rv.GetVal(CStr(k))
                If CBool(ki(1)) And Not IsBlank(v) Then
                    Set om = optMaps.GetVal(fid)
                    s = ""
                    names = Split(CStr(v), ";")
                    For j = 0 To UBound(names)
                        onm = Trim$(CStr(names(j)))
                        If Len(onm) > 0 Then s = ListAdd(s, CStr(om.GetVal(onm, "")))
                    Next
                    v = s
                End If
                upd.SetNew fid, v
            End If
        Next
        Set links = rv.GetVal("#links")
        If Len(rid) > 0 Then
            mData.UpdateRecord tid, rid, upd
            For Each k In links.Keys
                mData.SetLinks fmap.GetVal(CStr(k), CStr(k)), rid, links.GetVal(CStr(k))
            Next
        Else
            n = n + 1
            rows(n, 1) = "R" & Format$(baseNo + n - 1, "0000000")
            rows(n, 2) = Now: rows(n, 3) = Now: rows(n, 4) = uid
            For Each k In upd.Keys
                c = t.ColOf(CStr(k))
                If c > 0 Then rows(n, c) = upd.GetVal(CStr(k))
            Next
            For Each k In links.Keys
                For Each a In Split(links.GetVal(CStr(k)), ";")
                    If Len(a) > 0 Then
                        nr = nr + 1
                        If nr > UBound(rels, 1) Then rels = GrowRels(rels)
                        rels(nr, 1) = k: rels(nr, 2) = rows(n, 1): rels(nr, 3) = a: rels(nr, 4) = Now
                    End If
                Next
            Next
        End If
    Next
    If n > 0 Then mData.InsertRecordsBulk tid, mData.SliceRows(rows, n, t.ColCount)
    If nr > 0 Then mData.RelAddBulk mData.SliceRows(rels, nr, 4)
    mLog.Add "Import", tid, "", "", n & " neu, " & (validRows.Count - n) & " aktualisiert"
    mData.TxCommit
    mCalc.RecalcAll
    mApp.EndWork
    ExecuteImport = tid
    Exit Function
failed:
    Dim msg As String
    msg = Err.Description
    mData.TxAbortIfOpen
    mApp.ResetWork
    RaiseUser "Der Import wurde abgebrochen und alle Datenänderungen zurückgenommen: " & msg & _
              " (Neu angelegte Tabellen/Felder bleiben bestehen und können entfernt werden; vorher wurde eine Sicherung erstellt.)"
End Function

Private Function GrowRels(ByVal a As Variant) As Variant
    Dim b() As Variant, r As Long, c As Long
    ReDim b(1 To UBound(a, 1) * 2, 1 To 4)
    For r = 1 To UBound(a, 1)
        For c = 1 To 4: b(r, c) = a(r, c): Next
    Next
    GrowRels = b
End Function

' Fehlerbericht auf Blatt _Importbericht ablegen (bleibt verfuegbar)
Public Sub StoreErrorReport(ByVal errs As Collection, ByVal source As String)
    Dim ws As Worksheet, i As Long, e As Variant, a() As Variant
    Set ws = ThisWorkbook.Worksheets(SH_IMP)
    mProtect.Unprot ws
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "Importbericht " & FmtDateTime(Now) & ", Quelle: " & Mid$(source, InStrRev(source, "\") + 1)
    ws.Range("A2:D2").Value = Array("Quellzeile", "Spalte", "Wert", "Grund")
    If errs.Count > 0 Then
        ReDim a(1 To errs.Count, 1 To 4)
        For Each e In errs
            i = i + 1
            a(i, 1) = e(0): a(i, 2) = e(1): a(i, 3) = "'" & e(2): a(i, 4) = e(3)
        Next
        ws.Range("A3").Resize(errs.Count, 4).Value = a
    End If
End Sub

Public Function ErrorReportRows() As Variant
    Dim ws As Worksheet, lastR As Long
    Set ws = ThisWorkbook.Worksheets(SH_IMP)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 3 Then ErrorReportRows = Empty: Exit Function
    ErrorReportRows = ws.Range("A3:D" & lastR).Value
End Function

Public Function ErrorReportTitle() As String
    ErrorReportTitle = NzS(ThisWorkbook.Worksheets(SH_IMP).Cells(1, 1).Value)
End Function

'==============================================================================
' Export
'==============================================================================
Public Function AskExportPath(ByVal suggestedName As String, ByVal ext As String) As String
    Dim res As Variant, folder As String, filt As String
    folder = ThisWorkbook.path
    Select Case ext
        Case ".xlsx": filt = "Excel-Arbeitsmappe (*.xlsx),*.xlsx"
        Case ".csv": filt = "CSV UTF-8 (*.csv),*.csv"
        Case ".pdf": filt = "PDF (*.pdf),*.pdf"
        Case ".zip": filt = "ZIP-Archiv (*.zip),*.zip"
        Case ".html": filt = "HTML (*.html),*.html"
        Case ".md": filt = "Markdown (*.md),*.md"
        Case ".txt": filt = "Text (*.txt),*.txt"
    End Select
    res = Application.GetSaveAsFilename(InitialFileName:=UniquePath(folder, SafeFileName(suggestedName), ext), FileFilter:=filt, Title:="Export speichern unter")
    If VarType(res) <> vbString Then Exit Function
    If LCase$(Right$(res, Len(ext))) <> ext Then res = res & ext
    If FileExists(CStr(res)) Then
        If Not mUI.Confirm("Die Datei existiert bereits:" & vbLf & res & vbLf & vbLf & "Überschreiben?", "Export", "Überschreiben", "Abbrechen", True) Then Exit Function
    End If
    AskExportPath = CStr(res)
End Function

' Aktuelle (gefilterte) Ansicht exportieren
Public Sub ExportView(ByVal vid As String, ByVal fmt As String)
    Dim tid As String, rows() As Long, n As Long, cols As Collection, c As Variant, flds As New Collection, f As cField, t As cTableData
    Dim path As String, st As cViewState, filt As String, title As String
    tid = mCfg.ViewGet(vid, "Tabelle")
    Set st = mNav.State(vid)
    rows = mViewTable.ViewRows(vid, n)
    For Each c In mViewTable.ViewColumns(vid)
        flds.Add mCfg.Field(CStr(c(0)))
    Next
    filt = mQuery.CondsText(mQuery.ParseConds(st.Conds))
    If Len(st.Search) > 0 Then filt = filt & IIf(Len(filt) > 0, "; ", "") & "Suche: " & st.Search
    If st.ArchMode = mQuery.ARCH_ONLY Then filt = filt & IIf(Len(filt) > 0, "; ", "") & "nur archivierte Einträge"
    title = mCfg.TableName(tid) & ": " & mCfg.ViewGet(vid, "Name")
    path = AskExportPath(title & "_" & StampForFile(), IIf(fmt = "CSV", ".csv", ".xlsx"))
    If Len(path) = 0 Then Exit Sub
    If fmt = "CSV" Then
        WriteCsvExport path, tid, rows, n, flds
    Else
        WriteXlsxExport path, title, filt, tid, rows, n, flds
    End If
    mLog.Add "Export", tid, vid, "", fmt & ", " & n & " Einträge"
    mUI.Alert "Export erstellt (" & Plural(n, "Eintrag", "Einträge") & "):" & vbLf & path & vbLf & vbLf & _
              IIf(fmt = "CSV", "CSV-Hinweis: Texte, die mit =, +, - oder @ beginnen, wurden mit einem vorangestellten Apostroph geschützt, damit Tabellenprogramme sie nicht als Formel ausführen.", _
                  "Die Datei enthält keine Makros und keine Verbindung zu dieser Arbeitsdatei."), "Export"
End Sub

'--- Excel-Export mit Auswahl: mehrere Ansichten oder Tabellen, schlicht oder gestaltet -----
Public Sub ExportFlow(ByVal vid As String)
    Dim fp As frmProps, tid As String, vids As Variant, vNames() As Variant, tids As Variant, tNames() As Variant, i As Long, preV As String, preT As String
    Dim scope As String, jobs As New Collection, a As Variant, styled As Boolean, colorNm As String, fontNm As String, band As String, path As String, nAll As Long
    tid = mCfg.ViewGet(vid, "Tabelle")
    vids = mCfg.ViewIds(tid)
    ReDim vNames(0 To Application.Max(0, ArrLen(vids) - 1))
    For i = 0 To ArrLen(vids) - 1
        vNames(i) = mCfg.ViewGet(CStr(vids(i)), "Name") & "  (" & mCfg.ViewGet(CStr(vids(i)), "Typ") & ")"
        If CStr(vids(i)) = vid Then preV = CStr(i)
    Next
    tids = mCfg.TableIds()
    ReDim tNames(0 To Application.Max(0, ArrLen(tids) - 1))
    For i = 0 To ArrLen(tids) - 1
        tNames(i) = mCfg.TableName(CStr(tids(i)))
        If CStr(tids(i)) = tid Then preT = CStr(i)
    Next
    Set fp = New frmProps
    fp.Init "Als Excel-Datei exportieren", "Jede Ansicht bzw. Tabelle wird ein eigenes Blatt derselben Datei. Exportiert wird, was die Ansicht gerade zeigt (Filter, Suche, Sortierung, Spalten). Die Datei enthält keine Makros.", 560
    fp.AddCombo "umfang", "Was exportieren?", Array("Diese Ansicht", "Mehrere Ansichten dieser Tabelle", "Mehrere Tabellen"), "Diese Ansicht"
    fp.AddMulti "views", "Ansichten von „" & mCfg.TableName(tid) & "“", vNames, preV
    fp.ShowWhen "views", "umfang", "Mehrere Ansichten dieser Tabelle"
    fp.AddMulti "tables", "Tabellen (je Tabelle die zuletzt genutzte bzw. erste Ansicht)", tNames, preT
    fp.ShowWhen "tables", "umfang", "Mehrere Tabellen"
    fp.AddCombo "stil", "Darstellung", Array("Gestaltet (Titel, farbige Kopfzeile, Linien)", "Schlichte Liste ohne Formatierung"), _
                mCfg.SysGet("ExportStil", "Gestaltet (Titel, farbige Kopfzeile, Linien)"), "Schlicht: nur Spaltenköpfe und Werte, gut zum Weiterrechnen."
    fp.AddStyle "farbe", "Farbstil", mCfg.SysGet("ExportFarbe", mDocStyle.DefaultStyle())
    fp.ShowWhen "farbe", "stil", "Gestaltet (Titel, farbige Kopfzeile, Linien)"
    fp.AddCombo "schrift", "Schriftart", mDocStyle.DocFonts(), mCfg.SysGet("ExportSchrift", mDocStyle.DefaultFont())
    fp.ShowWhen "schrift", "stil", "Gestaltet (Titel, farbige Kopfzeile, Linien)"
    fp.AddCombo "band", "Zeilen", Array("Abwechselnd leicht getönt", "Nur feine Linien"), mCfg.SysGet("ExportZeilen", "Abwechselnd leicht getönt")
    fp.ShowWhen "band", "stil", "Gestaltet (Titel, farbige Kopfzeile, Linien)"
    Do
        If Not fp.ShowDialog("Exportieren") Then Unload fp: Exit Sub
        fp.ClearErrors
        scope = fp.GetText("umfang")
        If scope = "Mehrere Ansichten dieser Tabelle" And Len(fp.GetText("views")) = 0 Then
            fp.SetError "views", "Bitte mindestens eine Ansicht wählen."
        ElseIf scope = "Mehrere Tabellen" And Len(fp.GetText("tables")) = 0 Then
            fp.SetError "tables", "Bitte mindestens eine Tabelle wählen."
        Else
            Exit Do
        End If
    Loop
    Select Case scope
        Case "Mehrere Ansichten dieser Tabelle"
            For Each a In Split(fp.GetText("views"), ";")
                jobs.Add CStr(vids(CLng(a)))
            Next
        Case "Mehrere Tabellen"
            For Each a In Split(fp.GetText("tables"), ";")
                If CStr(tids(CLng(a))) = tid Then jobs.Add vid Else jobs.Add mNav.FirstViewOf(CStr(tids(CLng(a))))
            Next
        Case Else
            jobs.Add vid
    End Select
    styled = (fp.GetIndex("stil") <> 1)
    colorNm = fp.GetText("farbe"): fontNm = fp.GetText("schrift"): band = fp.GetText("band")
    Unload fp
    ' Wahl merken (Komfort für den nächsten Export)
    If Not mApp.gReadOnly Then
        mCfg.SysSet "ExportStil", IIf(styled, "Gestaltet (Titel, farbige Kopfzeile, Linien)", "Schlichte Liste ohne Formatierung")
        mCfg.SysSet "ExportFarbe", colorNm: mCfg.SysSet "ExportSchrift", fontNm: mCfg.SysSet "ExportZeilen", band
    End If
    If jobs.Count = 1 Then
        path = AskExportPath(mCfg.TableName(mCfg.ViewGet(jobs(1), "Tabelle")) & ", " & mCfg.ViewGet(jobs(1), "Name") & "_" & StampForFile(), ".xlsx")
    Else
        path = AskExportPath(IIf(scope = "Mehrere Tabellen", "Tabellen", mCfg.TableName(tid)) & "_Export_" & StampForFile(), ".xlsx")
    End If
    If Len(path) = 0 Then Exit Sub
    nAll = WriteXlsxMulti(path, jobs, styled, mDocStyle.DocColor(colorNm), mDocStyle.DocFont(fontNm), (band = "Abwechselnd leicht getönt"))
    mLog.Add "Export", tid, vid, "", "XLSX, " & jobs.Count & " Blatt/Blätter, " & nAll & " Einträge"
    mUI.Alert "Export erstellt (" & Plural(jobs.Count, "Blatt", "Blätter") & ", " & Plural(nAll, "Eintrag", "Einträge") & "):" & vbLf & path & vbLf & vbLf & _
              "Die Datei enthält keine Makros und keine Verbindung zu dieser Arbeitsdatei.", "Export"
End Sub

' Je Ansicht ein Blatt; liefert die Gesamtzahl der Einträge
Private Function WriteXlsxMulti(ByVal path As String, ByVal jobs As Collection, ByVal styled As Boolean, ByVal accent As Long, _
                                ByVal fontNm As String, ByVal banded As Boolean) As Long
    Dim wb As Workbook, ws As Worksheet, j As Variant, k As Long, vid As String, tid As String, rows() As Long, n As Long, flds As Collection
    Dim c As Variant, st As cViewState, filt As String, used As New cMap, nm As String
    mApp.BeginWork "Export wird erstellt …"
    On Error GoTo failed
    Set wb = Workbooks.Add(xlWBATWorksheet)
    For Each j In jobs
        k = k + 1
        vid = CStr(j)
        tid = mCfg.ViewGet(vid, "Tabelle")
        Set st = mNav.State(vid)
        rows = mViewTable.ViewRows(vid, n)
        Set flds = New Collection
        For Each c In mViewTable.ViewColumns(vid)
            flds.Add mCfg.Field(CStr(c(0)))
        Next
        filt = mQuery.CondsText(mQuery.ParseConds(st.Conds))
        If Len(st.Search) > 0 Then filt = filt & IIf(Len(filt) > 0, "; ", "") & "Suche: " & st.Search
        If st.ArchMode = mQuery.ARCH_ONLY Then filt = filt & IIf(Len(filt) > 0, "; ", "") & "nur archivierte Einträge"
        If k = 1 Then Set ws = wb.Worksheets(1) Else Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ' Blattname: Tabelle bzw. „Tabelle, Ansicht“, eindeutig, höchstens 31 Zeichen
        nm = mCfg.TableName(tid)
        If jobs.Count > 1 Then
            For Each c In jobs
                If CStr(c) <> vid And mCfg.ViewGet(CStr(c), "Tabelle") = tid Then nm = nm & ", " & mCfg.ViewGet(vid, "Name"): Exit For
            Next
        End If
        ws.Name = UniqueSheetName(nm, used)
        WriteExportSheet ws, mCfg.TableName(tid) & ": " & mCfg.ViewGet(vid, "Name"), filt, tid, rows, n, flds, styled, accent, fontNm, banded
        WriteXlsxMulti = WriteXlsxMulti + n
    Next
    wb.Worksheets(1).Activate
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=51
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    mApp.EndWork
    Exit Function
failed:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    mApp.EndWork
    On Error GoTo 0
    Err.Raise en, , ed
End Function

' Nur für automatisierte Prüfungen: gestalteter Export ohne Dialog
Public Function ExportForTest(ByVal path As String, ByVal vids As Collection) As Long
    ExportForTest = WriteXlsxMulti(path, vids, True, mDocStyle.DocColor("Petrol"), "Garamond", True)
End Function

Private Function UniqueSheetName(ByVal s As String, ByVal used As cMap) As String
    Dim ch As Variant, stem As String, i As Long
    For Each ch In Array("[", "]", ":", "*", "?", "/", "\")
        s = Replace(s, CStr(ch), " ")
    Next
    s = Trim$(s)
    If Len(s) = 0 Then s = "Export"
    stem = Left$(s, 31): s = stem: i = 1
    Do While used.Exists(LCase$(s))
        i = i + 1
        s = Left$(stem, 31 - Len(CStr(i)) - 1) & " " & i
    Loop
    used.SetVal LCase$(s), True
    UniqueSheetName = s
End Function

' Ein Exportblatt. Schlicht: Kopfzeile in Zeile 1, Werte darunter, ohne Gestaltung.
' Gestaltet: Titel, Infozeile, farbige Kopfzeile, feine Linien, optional getönte Zeilen.
Private Sub WriteExportSheet(ByVal ws As Worksheet, ByVal title As String, ByVal filt As String, ByVal tid As String, rows() As Long, ByVal n As Long, _
                             ByVal flds As Collection, ByVal styled As Boolean, ByVal accent As Long, ByVal fontNm As String, ByVal banded As Boolean)
    Dim t As cTableData, a() As Variant, k As Long, i As Long, f As cField, nc As Long, r0 As Long, body As Range, hdr As Range, fc As FormatCondition
    Set t = mData.TData(tid)
    nc = flds.Count
    r0 = IIf(styled, 5, 1)
    ReDim a(1 To n + 1, 1 To nc)
    i = 0
    For Each f In flds
        i = i + 1
        a(1, i) = f.Name
        If f.IsTextLike Or f.IsLinkType Or f.IsOptionType Then ws.Range(ws.Cells(r0 + 1, i), ws.Cells(r0 + n + 1, i)).NumberFormat = "@"
    Next
    For k = 1 To n
        i = 0
        For Each f In flds
            i = i + 1
            If f.IsLinkType Then
                a(k + 1, i) = mFmt.DisplayText(f, Empty, t.IdAt(rows(k)))
            Else
                a(k + 1, i) = mFmt.CellValue(f, t.Val_(rows(k), f.ID), t.IdAt(rows(k)))
            End If
        Next
    Next
    ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + n, nc)).Value = CellTexts(a)
    i = 0
    For Each f In flds
        i = i + 1
        If Not (f.IsTextLike Or f.IsLinkType Or f.IsOptionType) And n > 0 Then
            ws.Range(ws.Cells(r0 + 1, i), ws.Cells(r0 + n, i)).NumberFormat = mFmt.NumberFormatFor(f, ColWhole(a, n, i))
        End If
    Next
    Set hdr = ws.Range(ws.Cells(r0, 1), ws.Cells(r0, nc))
    If Not styled Then
        hdr.Font.Bold = True
        ws.Columns.AutoFit
        For i = 1 To nc
            If ws.Columns(i).ColumnWidth > 60 Then ws.Columns(i).ColumnWidth = 60
        Next
        Exit Sub
    End If
    ws.Cells.Font.Name = fontNm
    ws.Cells.Font.Size = mDocStyle.DocFontSize(fontNm, 10)
    ws.Columns(1).ColumnWidth = 12
    ws.Cells(1, 1).Value = title
    ws.Cells(1, 1).Font.Size = mDocStyle.DocFontSize(fontNm, 15): ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Color = mDocStyle.DocInk(accent)
    ws.Rows(1).RowHeight = 26
    ws.Cells(2, 1).Value = "Erstellt am " & FmtDateTime(Now) & " · " & Plural(n, "Eintrag", "Einträge")
    ws.Cells(3, 1).Value = IIf(Len(filt) > 0, "Filter: " & filt, "Ohne Filter (alle aktiven Einträge)")
    ws.Range("A2:A3").Font.Color = RGB(107, 114, 128)
    ws.Range("A2:A3").Font.Size = mDocStyle.DocFontSize(fontNm, 9)
    ' Kopfzeile: Akzentfarbe, Schrift passend hell/dunkel
    hdr.Interior.Color = accent
    hdr.Font.Color = mDocStyle.DocOnColor(accent)
    hdr.Font.Bold = True
    hdr.VerticalAlignment = xlCenter
    ws.Rows(r0).RowHeight = 24
    If n > 0 Then
        Set body = ws.Range(ws.Cells(r0 + 1, 1), ws.Cells(r0 + n, nc))
        body.VerticalAlignment = xlCenter
        body.RowHeight = 18
        With body.Borders(xlInsideHorizontal): .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(229, 231, 235): End With
        With body.Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlThin: .Color = Blend(accent, RGB(255, 255, 255), 0.4): End With
        If banded Then
            ' bedingte Formate erwarten die Formel in der Sprache der Excel-Oberfläche:
            ' über eine Hilfszelle übersetzen (=MOD(ROW(),2)=… wird z. B. zu =REST(ZEILE();2)=…)
            Dim fx As String
            With ws.Cells(1, nc + 3)
                .Formula = "=MOD(ROW(),2)=" & IIf((r0 + 1) Mod 2 = 0, 1, 0)
                fx = .FormulaLocal
                .ClearContents
            End With
            Set fc = body.FormatConditions.Add(Type:=xlExpression, Formula1:=fx)
            fc.Interior.Color = Blend(accent, RGB(255, 255, 255), 0.92)
        End If
    End If
    ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + Application.Max(n, 1), nc)).AutoFilter
    ' Breite nur nach Kopf und Daten (der Titel in A1 darf überlaufen)
    ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + n, nc)).Columns.AutoFit
    For i = 1 To nc
        If ws.Columns(i).ColumnWidth > 60 Then ws.Columns(i).ColumnWidth = 60
        If ws.Columns(i).ColumnWidth < 10 Then ws.Columns(i).ColumnWidth = 10
    Next
    ' Titelzeilen nicht in die Spaltenbreite der ersten Spalte zwingen
    ws.Range("A1:A3").WrapText = False
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ws.Cells(r0 + 1, 1).Select
    ActiveWindow.FreezePanes = True
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = ws.Rows(r0).Address
    End With
    Application.PrintCommunication = True
End Sub

Private Function CsvCell(ByVal s As String) As String
    If Len(s) > 0 Then
        Select Case Left$(s, 1)
            Case "=", "+", "-", "@", vbTab, vbCr
                s = "'" & s
        End Select
    End If
    If InStr(s, ";") > 0 Or InStr(s, """") > 0 Or InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Or Left$(s, 1) = " " Then
        s = """" & Replace(s, """", """""") & """"
    End If
    CsvCell = s
End Function

Private Sub WriteCsvExport(ByVal path As String, ByVal tid As String, rows() As Long, ByVal n As Long, ByVal flds As Collection)
    Dim sb As String, f As cField, k As Long, t As cTableData, lnTxt As String, parts() As String, i As Long
    Set t = mData.TData(tid)
    ReDim parts(0 To n)
    lnTxt = ""
    For Each f In flds
        lnTxt = lnTxt & IIf(Len(lnTxt) > 0, ";", "") & CsvCell(f.Name)
    Next
    parts(0) = lnTxt
    For k = 1 To n
        lnTxt = ""
        i = 0
        For Each f In flds
            lnTxt = lnTxt & IIf(i > 0, ";", "") & CsvCell(mFmt.DisplayText(f, t.Val_(rows(k), f.ID), t.IdAt(rows(k))))
            i = i + 1
        Next
        parts(k) = lnTxt
    Next
    WriteUtf8File path, Join(parts, vbCrLf) & vbCrLf, True
End Sub

Private Sub WriteXlsxExport(ByVal path As String, ByVal title As String, ByVal filt As String, ByVal tid As String, rows() As Long, ByVal n As Long, ByVal flds As Collection)
    Dim wb As Workbook, ws As Worksheet, t As cTableData, a() As Variant, k As Long, i As Long, f As cField, lo_ As ListObject, nc As Long
    Set t = mData.TData(tid)
    nc = flds.Count
    mApp.BeginWork "Export wird erstellt …"
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Export"
    ws.Cells(1, 1).Value = title
    ws.Cells(1, 1).Font.Size = 14: ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Value = "Erstellt am " & FmtDateTime(Now) & " · " & Plural(n, "Eintrag", "Einträge")
    ws.Cells(3, 1).Value = IIf(Len(filt) > 0, "Filter: " & filt, "Ohne Filter (alle aktiven Einträge)")
    ws.Range("A2:A3").Font.Color = RGB(100, 100, 100)
    ReDim a(1 To n + 1, 1 To nc)
    i = 0
    For Each f In flds
        i = i + 1
        a(1, i) = f.Name
        If f.IsTextLike Or f.IsLinkType Or f.IsOptionType Then ws.Range(ws.Cells(6, i), ws.Cells(5 + n + 1, i)).NumberFormat = "@"
    Next
    For k = 1 To n
        i = 0
        For Each f In flds
            i = i + 1
            If f.IsLinkType Then
                a(k + 1, i) = mFmt.DisplayText(f, Empty, t.IdAt(rows(k)))
            Else
                a(k + 1, i) = mFmt.CellValue(f, t.Val_(rows(k), f.ID), t.IdAt(rows(k)))
            End If
        Next
    Next
    ws.Range(ws.Cells(5, 1), ws.Cells(5 + n, nc)).Value = a
    i = 0
    For Each f In flds
        i = i + 1
        If Not (f.IsTextLike Or f.IsLinkType Or f.IsOptionType) Then
            ws.Range(ws.Cells(6, i), ws.Cells(5 + n, i)).NumberFormat = mFmt.NumberFormatFor(f, ColWhole(a, n, i))
        End If
    Next
    Set lo_ = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(5, 1), ws.Cells(5 + Application.Max(n, 1), nc)), , xlYes)
    lo_.TableStyle = "TableStyleLight9"
    ws.Columns.AutoFit
    For i = 1 To nc
        If ws.Columns(i).ColumnWidth > 60 Then ws.Columns(i).ColumnWidth = 60
    Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = True
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=51
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    mApp.EndWork
End Sub

' Dashboard als XLSX (Werte je Element + neu erzeugte Diagramme) oder PDF
Public Sub ExportDashboard(ByVal did As String, ByVal fmt As String)
    Dim path As String, ws As Worksheet, wb As Workbook, xs As Worksheet, r As Long, e As Variant, res As cMap, n As Long, i As Long
    Dim labels As Variant, values As Variant, ch As ChartObject, title As String
    title = "Dashboard " & mCfg.GetS(SH_DSB, did, "Name")
    If fmt = "PDF" Then
        path = AskExportPath(title & "_" & StampForFile(), ".pdf")
        If Len(path) = 0 Then Exit Sub
        Set ws = mRender.Sheet(SH_DASH)
        mApp.BeginWork "PDF wird erstellt …"
        Application.PrintCommunication = False
        With ws.PageSetup
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHeader = "&""Segoe UI,Fett""" & title
            .LeftFooter = "Erstellt am " & FmtDateTime(Now)
            .RightFooter = "Seite &P von &N"
            .LeftMargin = Application.CentimetersToPoints(1)
            .RightMargin = Application.CentimetersToPoints(1)
            .PrintGridlines = False
        End With
        Application.PrintCommunication = True
        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path, Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        mApp.EndWork
    Else
        path = AskExportPath(title & "_" & StampForFile(), ".xlsx")
        If Len(path) = 0 Then Exit Sub
        mApp.BeginWork "Excel-Export wird erstellt …"
        Set wb = Workbooks.Add(xlWBATWorksheet)
        Set xs = wb.Worksheets(1)
        xs.Name = "Dashboard"
        xs.Cells(1, 1).Value = title: xs.Cells(1, 1).Font.Size = 14: xs.Cells(1, 1).Font.Bold = True
        xs.Cells(2, 1).Value = "Erstellt am " & FmtDateTime(Now) & IIf(Len(mDash.DashFilterText(did)) > 0, " · Filter: " & mDash.DashFilterText(did), "")
        r = 4
        For Each e In mCfg.ElementIds(did)
            Set res = mDash.Compute(CStr(e), did)
            xs.Cells(r, 1).Value = mDash.ElemGet(CStr(e), "Titel"): xs.Cells(r, 1).Font.Bold = True
            xs.Cells(r + 1, 1).Value = NzS(res.GetVal("notes", NzS(res.GetVal("error", ""))))
            xs.Cells(r + 1, 1).Font.Color = RGB(110, 110, 110): xs.Cells(r + 1, 1).Font.Size = 8
            r = r + 2
            n = CLng(res.GetVal("n", 0))
            xs.Cells(r, 1).Value = "Gruppe": xs.Cells(r, 2).Value = mDash.MeasureLabel(CStr(res.GetVal("measure", "ANZAHL")))
            xs.Range(xs.Cells(r, 1), xs.Cells(r, 2)).Font.Bold = True
            If n > 0 Then
                labels = res.GetVal("labels"): values = res.GetVal("values")
                For i = 0 To n - 1
                    xs.Cells(r + 1 + i, 1).NumberFormat = "@"
                    xs.Cells(r + 1 + i, 1).Value = labels(i)
                    If Not IsBlank(values(i)) Then xs.Cells(r + 1 + i, 2).Value = values(i)
                Next
                If res.GetVal("measure") = "ANTEIL" Then xs.Range(xs.Cells(r + 1, 2), xs.Cells(r + n, 2)).NumberFormat = "0.0%"
                If mDash.ElemGet(CStr(e), "Typ") <> "Tabelle" And mDash.ElemGet(CStr(e), "Typ") <> "Kachel" And mDash.ElemGet(CStr(e), "Typ") <> "Rangliste" _
                   And Not (mDash.ElemGet(CStr(e), "Typ") = "Netz" And n < 3) Then
                    Set ch = mDash.NewChart(xs, mDash.ElemGet(CStr(e), "Typ"), xs.Cells(r, 5).Left, xs.Cells(r, 5).Top, 360, 200, _
                                            xs.Range(xs.Cells(r + 1, 1), xs.Cells(r + n, 1)), xs.Range(xs.Cells(r + 1, 2), xs.Cells(r + n, 2)), res)
                    r = r + Application.Max(n + 2, 16)
                Else
                    r = r + n + 2
                End If
                xs.Cells(r - 1, 1).Value = "Gesamt (eindeutig)": xs.Cells(r - 1, 2).Value = res.GetVal("totalValue")
                xs.Cells(r - 1, 1).Font.Italic = True
            Else
                xs.Cells(r + 1, 1).Value = "Gesamt": xs.Cells(r + 1, 2).Value = res.GetVal("totalValue")
                If res.GetVal("measure") = "ANTEIL" Then xs.Cells(r + 1, 2).NumberFormat = "0.0%"
                r = r + 3
            End If
            r = r + 1
        Next
        xs.Columns(1).ColumnWidth = 40
        xs.Columns(2).ColumnWidth = 22
        Application.DisplayAlerts = False
        wb.SaveAs Filename:=path, FileFormat:=51
        wb.Close SaveChanges:=False
        Application.DisplayAlerts = True
        ThisWorkbook.Activate
        mApp.EndWork
    End If
    mLog.Add "Export", "", did, "", "Dashboard " & fmt
    mUI.Alert "Export erstellt:" & vbLf & path, "Export"
End Sub

' Fehlerbericht als CSV
Public Sub ExportErrorReport()
    Dim rows As Variant, path As String, parts() As String, i As Long
    rows = ErrorReportRows()
    If IsEmpty(rows) Then mUI.Alert "Es liegt kein Fehlerbericht vor.", "Importbericht": Exit Sub
    path = AskExportPath("Importfehler_" & StampForFile(), ".csv")
    If Len(path) = 0 Then Exit Sub
    ReDim parts(0 To UBound(rows, 1))
    parts(0) = "Quellzeile;Spalte;Wert;Grund"
    For i = 1 To UBound(rows, 1)
        parts(i) = CsvCell(NzS(rows(i, 1))) & ";" & CsvCell(NzS(rows(i, 2))) & ";" & CsvCell(NzS(rows(i, 3))) & ";" & CsvCell(NzS(rows(i, 4)))
    Next
    WriteUtf8File path, Join(parts, vbCrLf) & vbCrLf, True
    mUI.Alert "Fehlerbericht gespeichert:" & vbLf & path, "Importbericht"
End Sub
