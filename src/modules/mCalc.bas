Attribute VB_Name = "mCalc"
'==============================================================================
' mCalc - Berechnungen. Zwei Formen:
'   Excel-Formel   "XL|=RUNDEN({F0003}*{F0004};2)"  (Eingabe: =RUNDEN([Preis]*[Menge];2))
'   Baukasten      "FUNKTION|F:F0003|K:+|N:100"     (Assistent, Rollups)
' Excel-Formeln laufen nur auf dem Hilfsblatt _Rechnen mit gesperrten Funktionen
' (kein Netzwerk-, Datei- oder Blattzugriff); niemals als VBA-Code.
' Zirkelbezuege werden beim Einrichten verhindert.
'==============================================================================
Option Explicit

Private mTopoRev As Long
Private mTopo As Collection              ' FeldIDs in Berechnungsreihenfolge
Private mDepRev As Long
Private mRollDeps As cMap                ' TabellenID -> Collection(Array(FormelFeld, LinkFeld))

' --- Funktionskatalog -----------------------------------------------------------
Public Function FnCodes() As Variant
    FnCodes = Array("RECHNEN", "RUNDEN", "ANTEIL", "WENN", "PRUEFEN", "DATDIFF", "TAGEBIS", "UEBERFAELLIG", _
                    "DATPLUS", "TEXT", "ROLLUP")
End Function

Public Function FnLabel(ByVal code As String) As String
    Select Case code
        Case "RECHNEN": FnLabel = "Rechnen (+ " & ChrW(&H2212) & " × ÷)"
        Case "RUNDEN": FnLabel = "Runden"
        Case "ANTEIL": FnLabel = "Anteil in Prozent (Teil ÷ Ganzes)"
        Case "WENN": FnLabel = "Wenn … dann … sonst"
        Case "PRUEFEN": FnLabel = "Bedingung prüfen (Ja/Nein)"
        Case "DATDIFF": FnLabel = "Zeitraum zwischen zwei Daten"
        Case "TAGEBIS": FnLabel = "Tage bis Frist (ab heute)"
        Case "UEBERFAELLIG": FnLabel = "Überfällig? (Frist vor heute)"
        Case "DATPLUS": FnLabel = "Datum plus Tage"
        Case "TEXT": FnLabel = "Texte verbinden"
        Case "ROLLUP": FnLabel = "Verknüpfte Einträge auswerten (Anzahl, Summe …)"
    End Select
End Function

Public Function FnHelp(ByVal code As String) As String
    Select Case code
        Case "RECHNEN": FnHelp = "Verbindet zwei Werte mit einer Grundrechenart. Ist ein Wert leer, bleibt das Ergebnis leer (nicht 0). Division durch 0 ergibt einen erkennbaren Fehler."
        Case "RUNDEN": FnHelp = "Rundet einen Zahlenwert kaufmännisch auf die angegebene Zahl von Nachkommastellen."
        Case "ANTEIL": FnHelp = "Berechnet Teil ÷ Ganzes als Prozentwert. Ist das Ganze 0, entsteht ein Fehlerhinweis."
        Case "WENN": FnHelp = "Prüft eine Bedingung (optional zweite Bedingung mit UND/ODER) und liefert je nach Ergebnis einen von zwei Werten."
        Case "PRUEFEN": FnHelp = "Prüft eine Bedingung (optional zweite mit UND/ODER) und liefert Ja oder Nein."
        Case "DATDIFF": FnHelp = "Berechnet den Abstand von Datum 1 bis Datum 2 in Tagen, Wochen, Monaten oder Jahren (ganze Einheiten)."
        Case "TAGEBIS": FnHelp = "Tage von heute bis zum Datum. Negative Werte bedeuten: Frist liegt in der Vergangenheit. Wird täglich beim Öffnen aktualisiert."
        Case "UEBERFAELLIG": FnHelp = "Ja, wenn das Datum vor heute liegt und der Status (optional) nicht als abgeschlossen markiert ist."
        Case "DATPLUS": FnHelp = "Addiert eine Anzahl Tage (Feld oder fester Wert) zu einem Datum."
        Case "TEXT": FnHelp = "Verbindet bis zu drei Werte zu einem Text, getrennt durch das angegebene Trennzeichen. Leere Werte werden ausgelassen."
        Case "ROLLUP": FnHelp = "Wertet die über ein Verknüpfungsfeld zugeordneten Einträge aus. Jeder verknüpfte Eintrag wird genau einmal gezählt; leere Werte werden nicht als 0 gezählt."
    End Select
End Function

' Argumentbeschreibung: "Art:Beschriftung"; Arten: val num date text op cmp logic unit agg link tfield status digits opt
Public Function FnArgs(ByVal code As String) As Variant
    Select Case code
        Case "RECHNEN": FnArgs = Array("num:Wert 1", "op:Rechenart", "num:Wert 2")
        Case "RUNDEN": FnArgs = Array("num:Wert", "digits:Nachkommastellen")
        Case "ANTEIL": FnArgs = Array("num:Teil", "num:Ganzes")
        Case "WENN": FnArgs = Array("val:Wenn Wert", "cmp:Vergleich", "val:mit", "logic:Verknüpfung", "val:und/oder Wert", "cmp:Vergleich 2", "val:mit", "val:Dann", "val:Sonst")
        Case "PRUEFEN": FnArgs = Array("val:Wert", "cmp:Vergleich", "val:mit", "logic:Verknüpfung", "val:und/oder Wert", "cmp:Vergleich 2", "val:mit")
        Case "DATDIFF": FnArgs = Array("date:Datum 1 (von)", "date:Datum 2 (bis)", "unit:Einheit")
        Case "TAGEBIS": FnArgs = Array("date:Frist")
        Case "UEBERFAELLIG": FnArgs = Array("date:Frist", "status:Status (optional)")
        Case "DATPLUS": FnArgs = Array("date:Datum", "num:Tage")
        Case "TEXT": FnArgs = Array("val:Wert 1", "text:Trennzeichen", "val:Wert 2", "val:Wert 3 (optional)")
        Case "ROLLUP": FnArgs = Array("link:Verknüpfungsfeld", "agg:Auswertung", "tfield:Feld der verknüpften Einträge", "tfield:Nur zählen, wenn Feld (optional)", "cmp:Vergleich", "val:Wert")
        Case Else: FnArgs = Array()
    End Select
End Function

Public Function DefaultResultType(ByVal def As String, ByVal tid As String) As String
    Dim p As Variant, fn As String, f As cField
    If IsXl(def) Then DefaultResultType = FT_ZAHL: Exit Function
    p = Split(def, "|")
    If UBound(p) < 0 Then DefaultResultType = FT_ZAHL: Exit Function
    fn = p(0)
    Select Case fn
        Case "ANTEIL": DefaultResultType = FT_PROZENT
        Case "PRUEFEN", "UEBERFAELLIG": DefaultResultType = FT_JANEIN
        Case "DATPLUS": DefaultResultType = FT_DATUM
        Case "TEXT": DefaultResultType = FT_TEXT
        Case "WENN"
            DefaultResultType = ArgType(ArgAt(p, 8))
            If Len(DefaultResultType) = 0 Then DefaultResultType = ArgType(ArgAt(p, 9))
            If Len(DefaultResultType) = 0 Then DefaultResultType = FT_TEXT
        Case "RECHNEN", "RUNDEN"
            DefaultResultType = ArgType(ArgAt(p, 1))
            If DefaultResultType <> FT_GELD And DefaultResultType <> FT_PROZENT Then DefaultResultType = FT_ZAHL
        Case "ROLLUP"
            Select Case UCase$(Mid$(ArgAt(p, 2), 3))
                Case "SUMME", "MIN", "MAX", "MITTEL"
                    DefaultResultType = ArgType(ArgAt(p, 3))
                    If Len(DefaultResultType) = 0 Then DefaultResultType = FT_ZAHL
                    If DefaultResultType <> FT_GELD And DefaultResultType <> FT_PROZENT And DefaultResultType <> FT_DATUM And DefaultResultType <> FT_DATUMZEIT Then DefaultResultType = FT_ZAHL
                Case Else: DefaultResultType = FT_ZAHL
            End Select
        Case Else: DefaultResultType = FT_ZAHL
    End Select
End Function

Private Function ArgType(ByVal a As String) As String
    Dim f As cField
    Select Case Left$(a, 2)
        Case "F:"
            Set f = mCfg.Field(Mid$(a, 3))
            If Not f Is Nothing Then
                ArgType = f.ValueType
                If ArgType = FT_AUSWAHL Or ArgType = FT_STATUS Or ArgType = FT_MEHRFACH Then ArgType = FT_TEXT
            End If
        Case "N:": ArgType = FT_ZAHL
        Case "T:", "O:": ArgType = FT_TEXT
        Case "D:": ArgType = FT_DATUM
        Case "B:": ArgType = FT_JANEIN
    End Select
End Function

Public Function ArgAt(ByVal p As Variant, ByVal i As Long) As String
    If i <= UBound(p) Then ArgAt = CStr(p(i))
End Function

'==============================================================================
' Excel-Formeln (seit 1.3.0): "XL|=[Preis]*[Menge]" mit gespeicherten Feld-IDs
' "XL|={F0003}*{F0004}". Die Formel wird in deutscher Schreibweise eingegeben
' (WENN, RUNDEN, Semikolon). Ausgewertet wird spaltenweise auf dem ausgeblendeten
' Blatt _Rechnen: Eingabewerte je Datensatz in Hilfsspalten, die Formel daneben,
' nur dieser Bereich wird berechnet. Bezuege auf Blaetter/Zellen, externe Dateien
' und Funktionen mit Netzwerk- oder Systemzugriff sind gesperrt.
'==============================================================================
Public Function IsXl(ByVal def As String) As Boolean
    IsXl = (Left$(def, 3) = "XL|")
End Function

' gespeicherte Form -> Anzeige "=[Preis]*[Menge]"
Public Function XlDisplay(ByVal def As String) As String
    Dim s As String, p As Long, q As Long, fid As String, nm As String
    s = Mid$(def, 4)
    p = InStr(s, "{")
    Do While p > 0
        q = InStr(p, s, "}")
        If q = 0 Then Exit Do
        fid = Mid$(s, p + 1, q - p - 1)
        If mCfg.FieldExists(fid) Then nm = mCfg.FieldName(fid) Else nm = "?"
        s = Left$(s, p - 1) & "[" & nm & "]" & Mid$(s, q + 1)
        p = InStr(p + Len(nm) + 2, s, "{")
    Loop
    XlDisplay = s
End Function

' Anzeige -> gespeicherte Form; msg bei unbekannter Spalte
Public Function XlFromDisplay(ByVal tid As String, ByVal txt As String, ByRef msg As String) As String
    Dim s As String, p As Long, q As Long, nm As String, fid As String, out As String
    s = Trim$(Replace(Replace(txt, vbCr, " "), vbLf, " "))
    If Len(s) = 0 Then msg = "Bitte eine Formel eingeben, z. B. =[Preis]*[Menge].": Exit Function
    If Left$(s, 1) <> "=" Then s = "=" & s
    p = 1
    Do
        q = InStr(p, s, "[")
        If q = 0 Then out = out & Mid$(s, p): Exit Do
        out = out & Mid$(s, p, q - p)
        p = InStr(q, s, "]")
        If p = 0 Then msg = "Eine eckige Klammer „[“ wird nicht geschlossen.": Exit Function
        nm = Trim$(Mid$(s, q + 1, p - q - 1))
        fid = mCfg.FieldIdByName(tid, nm)
        If Len(fid) = 0 Then msg = "Die Spalte [" & nm & "] gibt es in dieser Tabelle nicht.": Exit Function
        out = out & "{" & fid & "}"
        p = p + 1
    Loop
    If InStr(out, "{") = 0 And InStr(out, "]") > 0 Then msg = "Eine eckige Klammer „]“ ist zu viel.": Exit Function
    XlFromDisplay = "XL|" & out
End Function

Private Function XlDeps(ByVal def As String) As Collection
    Dim res As New Collection, s As String, p As Long, q As Long
    s = Mid$(def, 4)
    p = InStr(s, "{")
    Do While p > 0
        q = InStr(p, s, "}")
        If q = 0 Then Exit Do
        On Error Resume Next
        res.Add Mid$(s, p + 1, q - p - 1), Mid$(s, p + 1, q - p - 1)
        On Error GoTo 0
        p = InStr(q, s, "{")
    Loop
    Set XlDeps = res
End Function

Private Function CalcSheet() As Worksheet
    Dim ws As Worksheet, prev As Object, ev As Boolean, su As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CALC)
    On Error GoTo 0
    If ws Is Nothing Then
        Set prev = ActiveSheet
        ev = Application.EnableEvents: su = Application.ScreenUpdating
        Application.EnableEvents = False: Application.ScreenUpdating = False
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_CALC
        ws.Visible = xlSheetVeryHidden
        If Not prev Is Nothing Then prev.Activate
        Application.EnableEvents = ev: Application.ScreenUpdating = su
    End If
    If ws.ProtectContents Then ws.Unprotect Password:=""
    Set CalcSheet = ws
End Function

' Uebersetzt die (deutsche) Formel in die interne Form mit Hilfsspaltenbezuegen
' (Spalte B.. = Eingaben, Zeile 1 relativ). Liefert "" und msg bei Fehlern.
Private Function XlCompile(ByVal def As String, ByRef refs As Collection, ByRef msg As String) As String
    ' Zwischenspeicher (gross/klein unterscheidend, daher keine Collection)
    Static cK() As String, cV() As String, cN As Long
    Dim s As String, k As Long, ws As Worksheet, eng As String, chk As String
    Dim bad As Variant, ev As Boolean, fid As Variant
    Set refs = XlDeps(def)
    For k = 1 To cN
        If StrComp(cK(k), def, vbBinaryCompare) = 0 Then
            If Left$(cV(k), 1) = "!" Then msg = Mid$(cV(k), 2) Else XlCompile = cV(k)
            Exit Function
        End If
    Next
    k = 0
    s = Mid$(def, 4)
    For Each fid In refs
        k = k + 1
        s = Replace(s, "{" & fid & "}", ColLetter(k + 1) & "1")
    Next
    If InStr(s, "{") > 0 Or InStr(s, "[") > 0 Then msg = "Ungültiger Spaltenbezug.": GoTo store
    Set ws = CalcSheet()
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    ws.Range("A1").FormulaLocal = s
    If Err.Number <> 0 Then
        Err.Clear
        Application.EnableEvents = ev
        msg = "Die Formel ist so nicht gültig. Bitte Klammern, Semikolons (;) und Funktionsnamen prüfen."
        GoTo store
    End If
    eng = ws.Range("A1").Formula
    ws.Range("A1").ClearContents
    On Error GoTo 0
    Application.EnableEvents = ev
    ' Sperren: Bezuege auf andere Blaetter/Dateien, feste Zellbezuege ausserhalb der Spalten,
    ' Funktionen mit Netzwerk-/Systemzugriff
    chk = UCase$(StripStrings(eng))
    If InStr(chk, "!") > 0 Or InStr(chk, "[") > 0 Then msg = "Bezüge auf andere Blätter oder Dateien sind nicht erlaubt. Bitte nur Spalten in [ ] verwenden.": GoTo store
    For Each bad In Array("WEBSERVICE(", "FILTERXML(", "RTD(", "CALL(", "REGISTER.ID(", "IMAGE(", "HYPERLINK(", "INDIRECT(", "OFFSET(", "INFO(", "CELL(")
        If InStr(chk, CStr(bad)) > 0 Then msg = "Die Funktion " & Left$(bad, Len(bad) - 1) & " ist in Formelspalten nicht erlaubt.": GoTo store
    Next
    If HasCellRef(chk, refs.Count + 1) Then msg = "Zellbezüge wie A1 sind nicht erlaubt. Spalten bitte in eckigen Klammern angeben, z. B. [Preis].": GoTo store
    XlCompile = eng
store:
    If cN >= 200 Then cN = 0
    cN = cN + 1
    ReDim Preserve cK(1 To cN): ReDim Preserve cV(1 To cN)
    cK(cN) = def
    If Len(msg) > 0 Then cV(cN) = "!" & msg Else cV(cN) = eng
End Function

' Texte in Anfuehrungszeichen entfernen (fuer Pruefungen)
Private Function StripStrings(ByVal s As String) As String
    Dim i As Long, inQ As Boolean, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            inQ = Not inQ
        ElseIf Not inQ Then
            r = r & ch
        End If
    Next
    StripStrings = r
End Function

' Zellbezug ausserhalb der erlaubten Hilfsspalten B1..(maxCol)1 vorhanden?
Private Function HasCellRef(ByVal s As String, ByVal maxCol As Long) As Boolean
    Dim i As Long, j As Long, letters As String, digits As String, prevCh As String, ch As String, colN As Long, k As Long
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Z$]" And Not (prevCh Like "[A-Z0-9_.]") Then
            j = i: letters = "": digits = ""
            Do While j <= Len(s) And Mid$(s, j, 1) Like "[A-Z$]"
                If Mid$(s, j, 1) <> "$" Then letters = letters & Mid$(s, j, 1)
                j = j + 1
            Loop
            Do While j <= Len(s) And Mid$(s, j, 1) Like "[0-9$]"
                If Mid$(s, j, 1) <> "$" Then digits = digits & Mid$(s, j, 1)
                j = j + 1
            Loop
            If Len(letters) >= 1 And Len(letters) <= 3 And Len(digits) > 0 Then
                If j > Len(s) Or Not (Mid$(s, j, 1) Like "[A-Z0-9_(.]") Then
                    colN = 0
                    For k = 1 To Len(letters)
                        colN = colN * 26 + (Asc(Mid$(letters, k, 1)) - 64)
                    Next
                    If Not (digits = "1" And colN >= 2 And colN <= maxCol) Then HasCellRef = True: Exit Function
                End If
            End If
            If j > i Then i = j - 1
        End If
        prevCh = Mid$(s, i, 1)
        i = i + 1
    Loop
End Function

Private Function ColLetter(ByVal n As Long) As String
    Dim s As String
    Do While n > 0
        s = Chr$(65 + (n - 1) Mod 26) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

' Pruefen und Vorschau: liefert msg ("" = gueltig)
Public Function XlCheck(ByVal def As String) As String
    Dim refs As Collection, msg As String
    If Not IsXl(def) Or Len(Trim$(Mid$(def, 4))) <= 1 Then XlCheck = "Bitte eine Formel eingeben.": Exit Function
    XlCompile def, refs, msg
    XlCheck = msg
End Function

' Eingabewert eines Feldes fuer die Excel-Auswertung
Private Function XlInput(ByVal f As cField, ByVal t As cTableData, ByVal r As Long) As Variant
    Dim v As Variant, c As Long
    If f.IsLinkType Then
        v = mFmt.LinkLabels(f, mData.LinkTargets(f.ID, t.IdAt(r)))
        If Len(v) = 0 Then XlInput = Empty Else XlInput = "'" & v
        Exit Function
    End If
    c = t.ColOf(f.ID)
    If c = 0 Then XlInput = CVErr(xlErrRef): Exit Function
    v = t.Cell(r, c)
    If IsBlank(v) Then XlInput = Empty: Exit Function
    If IsErrVal(v) Then XlInput = CVErr(xlErrValue): Exit Function
    Select Case f.ValueType
        Case FT_ZAHL, FT_GELD, FT_PROZENT
            If IsNumeric(v) And VarType(v) <> vbString Then XlInput = CDbl(v) Else XlInput = "'" & CStr(v)
        Case FT_DATUM, FT_DATUMZEIT
            If IsDate(v) Or (IsNumeric(v) And VarType(v) <> vbString) Then XlInput = CDbl(CDate(v)) Else XlInput = "'" & CStr(v)
        Case FT_JANEIN
            XlInput = CBool(v)
        Case FT_AUSWAHL, FT_STATUS, FT_MEHRFACH
            XlInput = "'" & mFmt.DisplayText(f, v)
        Case Else
            XlInput = "'" & CStr(v)
    End Select
End Function

Private Function XlErrText(ByVal code As Long) As String
    Select Case code
        Case 2007: XlErrText = "#DIV/0! (Division durch 0)"
        Case 2015: XlErrText = "#WERT! (falscher Werttyp)"
        Case 2042: XlErrText = "#NV (kein Wert gefunden)"
        Case 2023: XlErrText = "#BEZUG!"
        Case 2029: XlErrText = "#NAME? (unbekannte Funktion?)"
        Case 2036: XlErrText = "#ZAHL!"
        Case 2000: XlErrText = "#NULL!"
        Case Else: XlErrText = "Excel-Fehler " & code
    End Select
End Function

' Ergebnis in den Ergebnistyp des Feldes umwandeln
Private Function XlResult(ByVal v As Variant, ByVal rt As String) As Variant
    If IsError(v) Then XlResult = ErrV(XlErrText(CLng(v))): Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(v) = 0 Then Exit Function
    End If
    On Error GoTo asIs
    Select Case rt
        Case FT_DATUM, FT_DATUMZEIT
            If VarType(v) = vbDouble Then XlResult = CDate(v) Else XlResult = v
        Case FT_JANEIN
            If VarType(v) = vbDouble Then
                XlResult = (v <> 0)
            Else
                XlResult = v
            End If
        Case FT_TEXT
            If VarType(v) = vbBoolean Then XlResult = IIf(v, "WAHR", "FALSCH") Else XlResult = CStr(v)
        Case Else
            XlResult = v
    End Select
    Exit Function
asIs:
    XlResult = v
End Function

' Wertet eine Excel-Formel fuer die Zeilen rows(1..n) der Tabelle aus (ein Rechenvorgang)
Public Function XlEvalRows(ByVal def As String, ByVal t As cTableData, rows() As Long, ByVal n As Long, ByVal rt As String) As Variant()
    Dim refs As Collection, msg As String, eng As String, k As Long, i As Long, flds() As cField, inp() As Variant
    Dim ws As Worksheet, res As Variant, out() As Variant, fid As Variant, ev As Boolean, nR As Long, missing As String
    ReDim out(1 To Application.Max(1, n))
    If n = 0 Then XlEvalRows = out: Exit Function
    eng = XlCompile(def, refs, msg)
    If Len(eng) = 0 Then
        For i = 1 To n: out(i) = ErrV(msg): Next
        XlEvalRows = out: Exit Function
    End If
    nR = refs.Count
    If nR > 0 Then
        ReDim flds(1 To nR)
        For Each fid In refs
            k = k + 1
            Set flds(k) = mCfg.Field(CStr(fid))
            If flds(k) Is Nothing Then missing = "Spalte entfernt"
        Next
        If Len(missing) > 0 Then
            For i = 1 To n: out(i) = ErrV(missing): Next
            XlEvalRows = out: Exit Function
        End If
        ReDim inp(1 To n, 1 To nR)
        For i = 1 To n
            For k = 1 To nR
                inp(i, k) = XlInput(flds(k), t, rows(i))
            Next
        Next
    End If
    Set ws = CalcSheet()
    ev = Application.EnableEvents
    Application.EnableEvents = False
    If nR > 0 Then ws.Range(ws.Cells(1, 2), ws.Cells(n, nR + 1)).Value = inp
    With ws.Range(ws.Cells(1, 1), ws.Cells(n, 1))
        .Formula = eng
        .Calculate
        res = .Value2
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(n, nR + 1)).ClearContents
    Application.EnableEvents = ev
    If n = 1 Then
        out(1) = XlResult(res, rt)
    Else
        For i = 1 To n
            out(i) = XlResult(res(i, 1), rt)
        Next
    End If
    XlEvalRows = out
End Function

Private Function XlEvalOne(ByVal def As String, ByVal t As cTableData, ByVal r As Long, ByVal rt As String) As Variant
    Dim rows(1 To 1) As Long, res() As Variant
    rows(1) = r
    res = XlEvalRows(def, t, rows, 1, rt)
    XlEvalOne = res(1)
End Function

'==============================================================================
' Auswertung
'==============================================================================
Private Function ErrV(ByVal msg As String) As String
    ErrV = ERR_PREFIX & msg
End Function

' Wert eines Arguments fuer einen Datensatz (Zeile r der Tabelle t)
Private Function ArgVal(ByVal a As String, ByVal t As cTableData, ByVal r As Long) As Variant
    Dim f As cField, c As Long
    Select Case Left$(a, 2)
        Case "F:"
            Set f = mCfg.Field(Mid$(a, 3))
            If f Is Nothing Then ArgVal = ErrV("Feld entfernt"): Exit Function
            If f.IsLinkType Then ArgVal = ErrV("Verknüpfung nur in Auswertungen"): Exit Function
            c = t.ColOf(f.ID)
            If c = 0 Then ArgVal = ErrV("Feld fehlt"): Exit Function
            ArgVal = t.Cell(r, c)
            If VarType(ArgVal) = vbString And f.ValueType = FT_ZAHL Then
                If Not IsErrVal(ArgVal) Then ArgVal = ErrV("Typ")
            End If
        Case "N:": ArgVal = InvNum(Mid$(a, 3))
        Case "T:": ArgVal = Unesc(Mid$(a, 3))
        Case "D:": ArgVal = InvDate(Mid$(a, 3))
        Case "B:": ArgVal = (Mid$(a, 3) = "1")
        Case "O:": ArgVal = Mid$(a, 3)
        Case "K:": ArgVal = Mid$(a, 3)
        Case Else: ArgVal = Empty
    End Select
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    If IsBlank(v) Or IsErrVal(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    If VarType(v) = vbString Then Exit Function
    IsNum = IsNumeric(v) Or VarType(v) = vbDate
End Function

Public Function EvalDef(ByVal def As String, ByVal t As cTableData, ByVal r As Long, Optional ByVal rt As String = "") As Variant
    Dim p As Variant, a As Variant, b As Variant, c As Variant, d As Variant, fn As String, n As Double, op As String
    Dim c1 As Variant, c2 As Variant, logic As String
    On Error GoTo failed
    If IsXl(def) Then EvalDef = XlEvalOne(def, t, r, rt): Exit Function
    p = Split(def, "|")
    If UBound(p) < 0 Then EvalDef = Empty: Exit Function
    fn = p(0)
    Select Case fn
        Case "RECHNEN"
            a = ArgVal(ArgAt(p, 1), t, r): op = Mid$(ArgAt(p, 2), 3): b = ArgVal(ArgAt(p, 3), t, r)
            If IsErrVal(a) Then EvalDef = a: Exit Function
            If IsErrVal(b) Then EvalDef = b: Exit Function
            If IsBlank(a) Or IsBlank(b) Then EvalDef = Empty: Exit Function
            If Not IsNum(a) Or Not IsNum(b) Then EvalDef = ErrV("Kein Zahlenwert"): Exit Function
            Select Case op
                Case "+": EvalDef = CDbl(a) + CDbl(b)
                Case "-": EvalDef = CDbl(a) - CDbl(b)
                Case "*": EvalDef = CDbl(a) * CDbl(b)
                Case "/"
                    If CDbl(b) = 0 Then EvalDef = ErrV("Division durch 0") Else EvalDef = CDbl(a) / CDbl(b)
                Case Else: EvalDef = ErrV("Rechenart")
            End Select
        Case "RUNDEN"
            a = ArgVal(ArgAt(p, 1), t, r)
            If IsErrVal(a) Or IsBlank(a) Then EvalDef = a: Exit Function
            If Not IsNum(a) Then EvalDef = ErrV("Kein Zahlenwert"): Exit Function
            n = InvNum(Mid$(ArgAt(p, 2), 3))
            EvalDef = RoundHalfUp(CDbl(a), CLng(n))
        Case "ANTEIL"
            a = ArgVal(ArgAt(p, 1), t, r): b = ArgVal(ArgAt(p, 2), t, r)
            If IsErrVal(a) Then EvalDef = a: Exit Function
            If IsErrVal(b) Then EvalDef = b: Exit Function
            If IsBlank(a) Or IsBlank(b) Then EvalDef = Empty: Exit Function
            If Not IsNum(a) Or Not IsNum(b) Then EvalDef = ErrV("Kein Zahlenwert"): Exit Function
            If CDbl(b) = 0 Then EvalDef = ErrV("Division durch 0") Else EvalDef = CDbl(a) / CDbl(b)
        Case "WENN", "PRUEFEN"
            c1 = Cond(ArgVal(ArgAt(p, 1), t, r), Mid$(ArgAt(p, 2), 3), ArgVal(ArgAt(p, 3), t, r))
            logic = Mid$(ArgAt(p, 4), 3)
            If logic = "UND" Or logic = "ODER" Then
                c2 = Cond(ArgVal(ArgAt(p, 5), t, r), Mid$(ArgAt(p, 6), 3), ArgVal(ArgAt(p, 7), t, r))
                If IsErrVal(c1) Then
                    EvalDef = c1: Exit Function
                ElseIf IsErrVal(c2) Then
                    EvalDef = c2: Exit Function
                End If
                If logic = "UND" Then c1 = (c1 And c2) Else c1 = (c1 Or c2)
            End If
            If IsErrVal(c1) Then EvalDef = c1: Exit Function
            If fn = "PRUEFEN" Then
                EvalDef = CBool(c1)
            ElseIf CBool(c1) Then
                EvalDef = ArgVal(ArgAt(p, 8), t, r)
            Else
                EvalDef = ArgVal(ArgAt(p, 9), t, r)
            End If
        Case "DATDIFF"
            a = ArgVal(ArgAt(p, 1), t, r): b = ArgVal(ArgAt(p, 2), t, r)
            If IsErrVal(a) Then EvalDef = a: Exit Function
            If IsErrVal(b) Then EvalDef = b: Exit Function
            If IsBlank(a) Or IsBlank(b) Then EvalDef = Empty: Exit Function
            If Not IsNum(a) Or Not IsNum(b) Then EvalDef = ErrV("Kein Datum"): Exit Function
            Select Case Mid$(ArgAt(p, 3), 3)
                Case "Wochen": EvalDef = Fix((Int(CDbl(b)) - Int(CDbl(a))) / 7)
                Case "Monate": EvalDef = FullMonths(CDate(a), CDate(b))
                Case "Jahre": EvalDef = Fix(FullMonths(CDate(a), CDate(b)) / 12)
                Case Else: EvalDef = Int(CDbl(b)) - Int(CDbl(a))
            End Select
        Case "TAGEBIS"
            a = ArgVal(ArgAt(p, 1), t, r)
            If IsErrVal(a) Or IsBlank(a) Then EvalDef = a: Exit Function
            If Not IsNum(a) Then EvalDef = ErrV("Kein Datum"): Exit Function
            EvalDef = Int(CDbl(a)) - CDbl(Date)
        Case "UEBERFAELLIG"
            a = ArgVal(ArgAt(p, 1), t, r)
            If IsErrVal(a) Then EvalDef = a: Exit Function
            If IsBlank(a) Then EvalDef = Empty: Exit Function
            If Not IsNum(a) Then EvalDef = ErrV("Kein Datum"): Exit Function
            EvalDef = (Int(CDbl(a)) < CDbl(Date))
            If EvalDef And Len(ArgAt(p, 2)) > 2 Then
                b = ArgVal(ArgAt(p, 2), t, r)
                If Not IsBlank(b) And Not IsErrVal(b) Then
                    If mCfg.OptDone(CStr(b)) Then EvalDef = False
                End If
            End If
        Case "DATPLUS"
            a = ArgVal(ArgAt(p, 1), t, r): b = ArgVal(ArgAt(p, 2), t, r)
            If IsErrVal(a) Then EvalDef = a: Exit Function
            If IsErrVal(b) Then EvalDef = b: Exit Function
            If IsBlank(a) Or IsBlank(b) Then EvalDef = Empty: Exit Function
            If Not IsNum(a) Or Not IsNum(b) Then EvalDef = ErrV("Kein Datum/keine Zahl"): Exit Function
            EvalDef = CDate(CDbl(a) + CDbl(b))
        Case "TEXT"
            EvalDef = JoinText(Array(ArgAt(p, 1), ArgAt(p, 3), ArgAt(p, 4)), Unesc(Mid$(ArgAt(p, 2), 3)), t, r)
        Case "ROLLUP"
            EvalDef = Rollup(p, t, r)
        Case Else
            EvalDef = ErrV("Unbekannte Funktion")
    End Select
    Exit Function
failed:
    EvalDef = ErrV("Ungültige Operation")
End Function

Private Function RoundHalfUp(ByVal x As Double, ByVal digits As Long) As Double
    Dim f As Double
    f = 10 ^ digits
    RoundHalfUp = Sgn(x) * Int(Abs(x) * f + 0.5 + 0.0000000001) / f
End Function

Private Function FullMonths(ByVal d1 As Date, ByVal d2 As Date) As Long
    Dim m As Long
    m = (Year(d2) - Year(d1)) * 12 + Month(d2) - Month(d1)
    If m > 0 And Day(d2) < Day(d1) Then m = m - 1
    If m < 0 And Day(d2) > Day(d1) Then m = m + 1
    FullMonths = m
End Function

Private Function JoinText(ByVal args As Variant, ByVal sep As String, ByVal t As cTableData, ByVal r As Long) As Variant
    Dim i As Long, s As String, v As Variant, f As cField, piece As String
    For i = 0 To UBound(args)
        If Len(args(i)) > 0 Then
            v = ArgVal(CStr(args(i)), t, r)
            If IsErrVal(v) Then JoinText = v: Exit Function
            If Left$(args(i), 2) = "F:" Then
                Set f = mCfg.Field(Mid$(args(i), 3))
                piece = mFmt.DisplayText(f, v, t.IdAt(r))
            ElseIf Left$(args(i), 2) = "O:" Then
                piece = mCfg.OptName(CStr(v))
            Else
                piece = NzS(v)
            End If
            If Len(piece) > 0 Then s = s & IIf(Len(s) > 0, sep, "") & piece
        End If
    Next
    If Len(s) = 0 Then JoinText = Empty Else JoinText = s
End Function

' Vergleich; liefert Boolean oder Fehlerwert
Private Function Cond(ByVal a As Variant, ByVal op As String, ByVal b As Variant) As Variant
    If IsErrVal(a) Then Cond = a: Exit Function
    Select Case op
        Case "LEER": Cond = IsBlank(a): Exit Function
        Case "NICHTLEER": Cond = Not IsBlank(a): Exit Function
    End Select
    If IsErrVal(b) Then Cond = b: Exit Function
    If op = "ENTHAELT" Then
        If IsBlank(a) Then Cond = False: Exit Function
        If VarType(a) = vbString And Left$(CStr(b), 1) = "O" And IsAllDigits(Mid$(CStr(b), 2)) Then
            Cond = ListHas(CStr(a), CStr(b))
        Else
            Cond = (InStr(1, CStr(a), NzS(b), vbTextCompare) > 0)
        End If
        Exit Function
    End If
    If IsBlank(a) Or IsBlank(b) Then
        ' Vergleich mit leerem Wert: nur Gleichheit/Ungleichheit sinnvoll
        Select Case op
            Case "=": Cond = (IsBlank(a) And IsBlank(b))
            Case "<>": Cond = Not (IsBlank(a) And IsBlank(b))
            Case Else: Cond = False
        End Select
        Exit Function
    End If
    Dim c As Long
    If VarType(a) = vbBoolean Or VarType(b) = vbBoolean Then
        If VarType(a) = vbBoolean And VarType(b) = vbBoolean Then
            c = IIf(a = b, 0, 1)
        Else
            c = 1
        End If
    ElseIf IsNum(a) And IsNum(b) Then
        If CDbl(a) < CDbl(b) Then
            c = -1
        ElseIf CDbl(a) > CDbl(b) Then
            c = 1
        Else
            c = 0
        End If
    Else
        c = StrComp(CStr(a), CStr(b), vbTextCompare)
    End If
    Select Case op
        Case "=": Cond = (c = 0)
        Case "<>": Cond = (c <> 0)
        Case ">": Cond = (c > 0)
        Case ">=": Cond = (c >= 0)
        Case "<": Cond = (c < 0)
        Case "<=": Cond = (c <= 0)
        Case Else: Cond = ErrV("Vergleich")
    End Select
End Function

Private Function Rollup(ByVal p As Variant, ByVal t As cTableData, ByVal r As Long) As Variant
    Dim lf As cField, agg As String, tf As cField, ff As cField, targets As Variant, tt As String
    Dim td As cTableData, i As Long, rr As Long, v As Variant, n As Long, cnt As Long, s As Double
    Dim mn As Variant, mx As Variant, fv As Variant, okF As Variant, cmpOp As String, cmpV As Variant, hasF As Boolean
    Dim tc As Long, fc As Long
    Set lf = mCfg.Field(Mid$(ArgAt(p, 1), 3))
    If lf Is Nothing Then Rollup = ErrV("Verknüpfung entfernt"): Exit Function
    If Not lf.IsLinkType Then Rollup = ErrV("Keine Verknüpfung"): Exit Function
    agg = Mid$(ArgAt(p, 2), 3)
    If lf.FType = FT_LINK Then tt = lf.TargetTable Else tt = mCfg.Field(lf.PairField).TableID
    Set td = mData.TData(tt)
    If Len(ArgAt(p, 3)) > 2 Then
        Set tf = mCfg.Field(Mid$(ArgAt(p, 3), 3))
        If tf Is Nothing Then Rollup = ErrV("Feld entfernt"): Exit Function
        tc = td.ColOf(tf.ID)
    End If
    If Len(ArgAt(p, 4)) > 2 Then
        Set ff = mCfg.Field(Mid$(ArgAt(p, 4), 3))
        If Not ff Is Nothing Then
            hasF = True
            fc = td.ColOf(ff.ID)
            cmpOp = Mid$(ArgAt(p, 5), 3)
            If Left$(ArgAt(p, 6), 2) = "F:" Then cmpV = Empty Else cmpV = ArgVal(ArgAt(p, 6), td, 1)
        End If
    End If
    targets = Split(mData.LinkTargets(lf.ID, t.IdAt(r)), ";")
    mn = Empty: mx = Empty
    For i = 0 To UBound(targets)
        rr = td.RowOf(CStr(targets(i)))
        If rr > 0 Then
            If Not td.IsArchived(rr) Then
                okF = True
                If hasF And fc > 0 Then okF = Cond(td.Cell(rr, fc), cmpOp, cmpV)
                If IsErrVal(okF) Then okF = False
                If okF Then
                    cnt = cnt + 1
                    If tc > 0 Then
                        v = td.Cell(rr, tc)
                        If Not IsBlank(v) And Not IsErrVal(v) And IsNum(v) Then
                            n = n + 1
                            s = s + CDbl(v)
                            If IsEmpty(mn) Then
                                mn = v
                            ElseIf v < mn Then
                                mn = v
                            End If
                            If IsEmpty(mx) Then
                                mx = v
                            ElseIf v > mx Then
                                mx = v
                            End If
                        ElseIf Not IsBlank(v) Then
                            If agg = "GEFUELLT" Then n = n + 1
                        End If
                    End If
                End If
            End If
        End If
    Next
    Select Case agg
        Case "ANZAHL", "ANZAHLWENN": Rollup = cnt
        Case "GEFUELLT": Rollup = n
        Case "SUMME": If n = 0 Then Rollup = IIf(cnt = 0, 0, Empty) Else Rollup = s
        Case "MITTEL": If n = 0 Then Rollup = Empty Else Rollup = s / n
        Case "MIN": Rollup = mn
        Case "MAX": Rollup = mx
        Case Else: Rollup = ErrV("Auswertung")
    End Select
End Function

'==============================================================================
' Abhaengigkeiten, Zirkelbezuege, Reihenfolge
'==============================================================================
Public Function Deps(ByVal def As String) As Collection
    Dim p As Variant, i As Long, res As New Collection
    If IsXl(def) Then Set Deps = XlDeps(def): Exit Function
    p = Split(def, "|")
    For i = 1 To UBound(p)
        If Left$(p(i), 2) = "F:" Then
            On Error Resume Next
            res.Add Mid$(p(i), 3), Mid$(p(i), 3)
            On Error GoTo 0
        End If
    Next
    Set Deps = res
End Function

' Prueft, ob fid mit Definition def einen (in-)direkten Zirkelbezug erzeugen wuerde
Public Function HasCycle(ByVal fid As String, ByVal def As String, ByRef path As String) As Boolean
    Dim visited As New cMap
    HasCycle = Visit(fid, def, fid, visited, path)
End Function

Private Function Visit(ByVal startFid As String, ByVal def As String, ByVal cur As String, ByVal visited As cMap, ByRef path As String) As Boolean
    Dim d As Variant, f As cField, ddef As String, lf As cField
    For Each d In Deps(def)
        If CStr(d) = startFid Then
            path = mCfg.FieldName(cur) & " " & ChrW(&H2192) & " " & mCfg.FieldName(startFid)
            Visit = True: Exit Function
        End If
        If Not visited.Exists(CStr(d)) Then
            visited.SetVal CStr(d), True
            Set f = mCfg.Field(CStr(d))
            If Not f Is Nothing Then
                If f.FType = FT_FORMEL Then
                    If Visit(startFid, f.Formula, f.ID, visited, path) Then
                        path = mCfg.FieldName(cur) & " " & ChrW(&H2192) & " " & path
                        Visit = True: Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

' Berechnungsreihenfolge aller Formelfelder (Abhaengigkeiten zuerst)
Public Function TopoOrder() As Collection
    Dim tids As Variant, i As Long, f As cField, done As New cMap, res As New Collection
    If Not mTopo Is Nothing And mTopoRev = mApp.gCfgRev Then Set TopoOrder = mTopo: Exit Function
    tids = mCfg.TableIds(True)
    For i = 0 To ArrLen(tids) - 1
        For Each f In mCfg.Fields(CStr(tids(i)), False)
            If f.FType = FT_FORMEL Then TopoVisit f.ID, done, res, 0
        Next
    Next
    Set mTopo = res
    mTopoRev = mApp.gCfgRev
    Set TopoOrder = res
End Function

Private Sub TopoVisit(ByVal fid As String, ByVal done As cMap, ByVal res As Collection, ByVal depth As Long)
    Dim f As cField, d As Variant
    If done.Exists(fid) Or depth > 50 Then Exit Sub
    done.SetVal fid, True
    Set f = mCfg.Field(fid)
    If f Is Nothing Then Exit Sub
    If f.FType <> FT_FORMEL Then Exit Sub
    For Each d In Deps(f.Formula)
        TopoVisit CStr(d), done, res, depth + 1
    Next
    res.Add fid
End Sub

' Formelfelder je Tabelle in Reihenfolge
Private Function FormulaFieldsOf(ByVal tid As String) As Collection
    Dim res As New Collection, v As Variant, f As cField
    For Each v In TopoOrder()
        Set f = mCfg.Field(CStr(v))
        If Not f Is Nothing Then
            If f.TableID = tid And f.IsActive Then res.Add f
        End If
    Next
    Set FormulaFieldsOf = res
End Function

' Welche Rollups haengen von Datensaetzen einer Tabelle ab: tid -> Collection(Array(Formelfeld, Linkfeld))
Private Function RollupDeps(ByVal tid As String) As Collection
    Dim v As Variant, f As cField, p As Variant, lf As cField, srcTid As String, col As Collection
    If mRollDeps Is Nothing Or mDepRev <> mApp.gCfgRev Then
        Set mRollDeps = New cMap
        For Each v In TopoOrder()
            Set f = mCfg.Field(CStr(v))
            If Not f Is Nothing Then
                p = Split(f.Formula, "|")
                If UBound(p) >= 1 Then
                    If p(0) = "ROLLUP" Then
                        Set lf = mCfg.Field(Mid$(p(1), 3))
                        If Not lf Is Nothing Then
                            If lf.FType = FT_LINK Then srcTid = lf.TargetTable Else srcTid = mCfg.Field(lf.PairField).TableID
                            If mRollDeps.Exists(srcTid) Then
                                Set col = mRollDeps.GetVal(srcTid)
                            Else
                                Set col = New Collection
                                mRollDeps.SetVal srcTid, col
                            End If
                            col.Add Array(f.ID, lf.ID)
                        End If
                    End If
                End If
            End If
        Next
        mDepRev = mApp.gCfgRev
    End If
    If mRollDeps.Exists(tid) Then Set RollupDeps = mRollDeps.GetVal(tid) Else Set RollupDeps = New Collection
End Function

'==============================================================================
' Neuberechnung
'==============================================================================
' Berechnet die Formelfelder eines Datensatzes; True, wenn sich ein Wert geaendert hat
Public Function RecalcRecord(ByVal tid As String, ByVal rid As String) As Boolean
    Dim t As cTableData, r As Long, f As cField, v As Variant, ch As New cMap, c As Long, oldA As Variant
    Set t = mData.TData(tid)
    r = t.RowOf(rid)
    If r = 0 Then Exit Function
    oldA = t.RowArray(r)
    For Each f In FormulaFieldsOf(tid)
        c = t.ColOf(f.ID)
        If c > 0 Then
            v = EvalDef(f.Formula, t, r, f.ValueType)
            If Not mData.SameValue(t.Cell(r, c), v) Then
                ch.SetVal f.ID, v
                ' sofort im Abbild sichtbar, damit abhaengige Felder desselben Datensatzes stimmen
                t.SetCell r, c, v
            End If
        End If
    Next
    If ch.Count > 0 Then
        ' Abbild auf Ausgangsstand und ueber den gemeinsamen Schreibweg speichern
        t.SetRow r, oldA
        mData.UpdateRecord tid, rid, ch, False
        RecalcRecord = True
    End If
End Function

' Abarbeitung aller betroffenen Datensaetze inkl. Weitergabe ueber Rollups
Public Sub RecalcTouched(ByVal touched As cMap)
    Dim queue As New Collection, visits As New cMap, item As Variant, tid As String, rid As String
    Dim changed As Boolean, dep As Variant, lf As cField, src As Variant, i As Long, k As Variant, initial As New cMap
    If touched Is Nothing Then Exit Sub
    If TopoOrder().Count = 0 Then Exit Sub
    If touched.Count > 150 Then
        ' Massenaenderung: spaltenweise Neuberechnung ist deutlich schneller
        RecalcAll
        Exit Sub
    End If
    For Each k In touched.Keys
        queue.Add touched.GetVal(CStr(k))
        initial.SetVal CStr(k), True
    Next
    Do While queue.Count > 0
        item = queue(1)
        queue.Remove 1
        tid = item(0): rid = item(1)
        If visits.GetVal(tid & "|" & rid, 0) < 4 Then
            visits.AddNum tid & "|" & rid, 1
            changed = False
            If mData.RecordExists(tid, rid) Then changed = RecalcRecord(tid, rid)
            If changed Or initial.Exists(tid & "|" & rid) Then
                initial.Remove tid & "|" & rid
                For Each dep In RollupDeps(tid)
                    Set lf = mCfg.Field(CStr(dep(1)))
                    If Not lf Is Nothing Then
                        src = Split(mData.LinkTargets(lf.PairField, rid), ";")
                        For i = 0 To UBound(src)
                            If Len(src(i)) > 0 Then queue.Add Array(lf.TableID, CStr(src(i)))
                        Next
                    End If
                Next
            End If
        End If
    Loop
End Sub

' Vollstaendige Neuberechnung aller Formelfelder (spaltenweise)
Public Sub RecalcAll()
    Dim v As Variant, f As cField, t As cTableData, r As Long, vals() As Variant, c As Long, allRows() As Long
    mApp.BeginWork "Berechnete Felder werden aktualisiert …"
    For Each v In TopoOrder()
        Set f = mCfg.Field(CStr(v))
        If Not f Is Nothing Then
            If f.IsActive Then
                Set t = mData.TData(f.TableID)
                c = t.ColOf(f.ID)
                If t.RowCount > 0 And c > 0 Then
                    If IsXl(f.Formula) Then
                        ' Excel-Formel: alle Datensaetze in einem Rechenvorgang
                        ReDim allRows(1 To t.RowCount)
                        For r = 1 To t.RowCount: allRows(r) = r: Next
                        vals = XlEvalRows(f.Formula, t, allRows, t.RowCount, f.ValueType)
                        For r = 1 To t.RowCount
                            t.SetCell r, c, vals(r)
                        Next
                    Else
                        ReDim vals(1 To t.RowCount)
                        For r = 1 To t.RowCount
                            vals(r) = EvalDef(f.Formula, t, r)
                            t.SetCell r, c, vals(r)
                        Next
                    End If
                    mData.WriteColumnRaw f.TableID, f.ID, vals
                End If
            End If
        End If
    Next
    mCfg.SysSet "LetzteBerechnung", DateInv(Date)
    mApp.EndWork
End Sub

Public Sub RecalcIfNewDay()
    If mCfg.SysGet("LetzteBerechnung", "") <> DateInv(Date) Then
        If TopoOrder().Count > 0 Then RecalcAll Else mCfg.SysSet "LetzteBerechnung", DateInv(Date)
    End If
End Sub

'==============================================================================
' Lesbare Beschreibung
'==============================================================================
Public Function Describe(ByVal def As String) As String
    Dim p As Variant, fn As String, s As String, lf As cField, tt As String
    If IsXl(def) Then Describe = XlDisplay(def): Exit Function
    p = Split(def, "|")
    If UBound(p) < 0 Then Exit Function
    fn = p(0)
    Select Case fn
        Case "RECHNEN": s = DArg(ArgAt(p, 1)) & " " & OpText(Mid$(ArgAt(p, 2), 3)) & " " & DArg(ArgAt(p, 3))
        Case "RUNDEN": s = DArg(ArgAt(p, 1)) & " gerundet auf " & Mid$(ArgAt(p, 2), 3) & " Nachkommastellen"
        Case "ANTEIL": s = DArg(ArgAt(p, 1)) & " als Anteil von " & DArg(ArgAt(p, 2)) & " in %"
        Case "WENN", "PRUEFEN"
            s = "Wenn " & DArg(ArgAt(p, 1)) & " " & CmpText(Mid$(ArgAt(p, 2), 3)) & DCmpArg(ArgAt(p, 2), ArgAt(p, 3))
            If Mid$(ArgAt(p, 4), 3) = "UND" Or Mid$(ArgAt(p, 4), 3) = "ODER" Then
                s = s & " " & LCase$(Mid$(ArgAt(p, 4), 3)) & " " & DArg(ArgAt(p, 5)) & " " & CmpText(Mid$(ArgAt(p, 6), 3)) & DCmpArg(ArgAt(p, 6), ArgAt(p, 7))
            End If
            If fn = "WENN" Then
                s = s & ", dann " & DArg(ArgAt(p, 8)) & ", sonst " & DArg(ArgAt(p, 9))
            Else
                s = s & " " & ChrW(&H2192) & " Ja, sonst Nein"
            End If
        Case "DATDIFF": s = Mid$(ArgAt(p, 3), 3) & " von " & DArg(ArgAt(p, 1)) & " bis " & DArg(ArgAt(p, 2))
        Case "TAGEBIS": s = "Tage von heute bis " & DArg(ArgAt(p, 1)) & " (negativ = vergangen)"
        Case "UEBERFAELLIG"
            s = "Ja, wenn " & DArg(ArgAt(p, 1)) & " vor heute liegt"
            If Len(ArgAt(p, 2)) > 2 Then s = s & " und " & DArg(ArgAt(p, 2)) & " nicht abgeschlossen ist"
        Case "DATPLUS": s = DArg(ArgAt(p, 1)) & " plus " & DArg(ArgAt(p, 2)) & " Tage"
        Case "TEXT": s = "Text aus " & DArg(ArgAt(p, 1)) & ", " & DArg(ArgAt(p, 3)) & IIf(Len(ArgAt(p, 4)) > 0, ", " & DArg(ArgAt(p, 4)), "") & _
                         " (Trennzeichen „" & Unesc(Mid$(ArgAt(p, 2), 3)) & "“)"
        Case "ROLLUP"
            Set lf = mCfg.Field(Mid$(ArgAt(p, 1), 3))
            If Not lf Is Nothing Then
                If lf.FType = FT_LINK Then tt = lf.TargetTable Else tt = mCfg.Field(lf.PairField).TableID
            End If
            Select Case Mid$(ArgAt(p, 2), 3)
                Case "ANZAHL", "ANZAHLWENN": s = "Anzahl"
                Case "GEFUELLT": s = "Anzahl ausgefüllter Werte von " & DArg(ArgAt(p, 3)) & " bei"
                Case "SUMME": s = "Summe von " & DArg(ArgAt(p, 3)) & " über"
                Case "MITTEL": s = "Durchschnitt von " & DArg(ArgAt(p, 3)) & " über"
                Case "MIN": s = "Kleinster Wert von " & DArg(ArgAt(p, 3)) & " bei"
                Case "MAX": s = "Größter Wert von " & DArg(ArgAt(p, 3)) & " bei"
            End Select
            s = s & " den über " & DArg(ArgAt(p, 1)) & " verknüpften Einträgen"
            If Len(tt) > 0 Then s = s & " der Tabelle „" & mCfg.TableName(tt) & "“"
            If Len(ArgAt(p, 4)) > 2 Then s = s & ", nur wenn " & DArg(ArgAt(p, 4)) & " " & CmpText(Mid$(ArgAt(p, 5), 3)) & DCmpArg(ArgAt(p, 5), ArgAt(p, 6))
            s = s & " (Zähleinheit: je verknüpftem Eintrag einmal; archivierte ausgenommen)"
    End Select
    Describe = s
End Function

Private Function DCmpArg(ByVal cmpArg As String, ByVal valArg As String) As String
    Select Case Mid$(cmpArg, 3)
        Case "LEER", "NICHTLEER": DCmpArg = ""
        Case Else: DCmpArg = " " & DArg(valArg)
    End Select
End Function

Public Function DArg(ByVal a As String) As String
    Select Case Left$(a, 2)
        Case "F:": DArg = "[" & mCfg.FieldName(Mid$(a, 3)) & "]"
        Case "N:": DArg = FmtNum(InvNum(Mid$(a, 3)), 6, True)
        Case "T:": DArg = "„" & Unesc(Mid$(a, 3)) & "“"
        Case "D:": DArg = FmtDate(InvDate(Mid$(a, 3)))
        Case "B:": DArg = IIf(Mid$(a, 3) = "1", "Ja", "Nein")
        Case "O:": DArg = "„" & mCfg.OptName(Mid$(a, 3)) & "“"
        Case "K:": DArg = Mid$(a, 3)
        Case Else: DArg = "(leer)"
    End Select
End Function

Public Function OpText(ByVal op As String) As String
    Select Case op
        Case "+": OpText = "plus"
        Case "-": OpText = "minus"
        Case "*": OpText = "mal"
        Case "/": OpText = "geteilt durch"
        Case Else: OpText = op
    End Select
End Function

Public Function CmpText(ByVal op As String) As String
    Select Case op
        Case "=": CmpText = "gleich"
        Case "<>": CmpText = "ungleich"
        Case ">": CmpText = "größer als"
        Case ">=": CmpText = "mindestens"
        Case "<": CmpText = "kleiner als"
        Case "<=": CmpText = "höchstens"
        Case "LEER": CmpText = "ist leer"
        Case "NICHTLEER": CmpText = "ist nicht leer"
        Case "ENTHAELT": CmpText = "enthält"
        Case Else: CmpText = op
    End Select
End Function

Public Function CmpCodes() As Variant
    CmpCodes = Array("=", "<>", ">", ">=", "<", "<=", "LEER", "NICHTLEER", "ENTHAELT")
End Function

Public Function UsesToday(ByVal def As String) As Boolean
    UsesToday = (Left$(def, 7) = "TAGEBIS" Or Left$(def, 12) = "UEBERFAELLIG")
    If IsXl(def) Then UsesToday = (InStr(1, def, "HEUTE(", vbTextCompare) > 0 Or InStr(1, def, "JETZT(", vbTextCompare) > 0)
End Function
