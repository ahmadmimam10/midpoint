Attribute VB_Name = "mListImport"
'==============================================================================
' mListImport - "Liste uebernehmen": vorhandene Listen lesen (beliebige Excel-/
' CSV-Listen oder 2ask-Befragungsexporte mit Codeplan), Tabellen, Kopfzeilen,
' Summen-/Gruppenzeilen und Spaltentypen erkennen, Rueckfragen mit Vorschlag
' erzeugen und daraus einen Importplan fuer mIO.CheckImport/ExecuteImport bauen.
' Die Quelldatei wird nur gelesen (schreibgeschuetzt, Makros deaktiviert).
'==============================================================================
Option Explicit

Private Const MAX_ROWS As Long = 20000
Private Const MAX_COLS As Long = 200
Private Const SUM_WORDS As String = "summe|gesamt|zwischensumme|insgesamt|total|endsumme|gesamtsumme|gesamtbetrag|summen"

' --- Sitzung (der Assistent liest diese Werte) ---
Public liPath As String
Public liMode As String                ' "2ask" | "liste"
Public liLooksAsk As Boolean           ' Modus "liste", Datei sieht aber wie ein 2ask-Export aus
Public liSheets As Collection          ' cLiSheet
Public liTables As Collection          ' cLiTab
Public liTab As cLiTab                 ' gewaehlte Tabelle
Public liCols As Collection            ' cLiCol
Public liQs As Collection              ' cLiQ
Public liNotes As Collection           ' Array(Quellzeile, Feld, Wert, Grund)
Public liAns As cMap                   ' Frage-ID -> Antwortindex
Public liOvr As cMap                   ' Feldschluessel -> Array(Name, Typ, Uebernehmen)
Public liValOvr As cMap                ' Feldschluessel|Rohwert -> Zielwert
Public liAskRows As Long               ' 2ask: Frageboegen
Public liAskQuestions As Collection    ' 2ask: Array(Variable, Fragetext, Art, Felder)
Public liRowSrc As Variant             ' Quellangabe je Zeile des Importplans

' --- intern ---
Private mCsv As Boolean
Private mMembers As Collection         ' cLiTab (eine oder mehrere Tabellen)
Private mUCols As Collection           ' Array(Name, cMap Mitglied->Spalte, obere Kopfzeile, Schluessel, Hilfetext)
Private mRecs As Collection            ' Array(Mitglied, Zeile, Gruppe)
Private mPivot As Collection           ' Indizes (mUCols) der Zeitraumspalten
Private mPivotKind As String           ' "monat" | "quartal" | "jahr"
Private mPivotYear As Long
Private mEuroHint As Boolean           ' Titel nennt "Euro"
Private mAskSheet As cLiSheet
Private mAskHdr As Long
Private mAskLast As Long
Private mAskKeyCol As cMap
Private mAskMissing As Long
Private mAskPretest As Long
Private mCp As Collection              ' Codeplan: Array(Schluessel, type, sub_type, data_type, coding, Text, Fragetext)
Private mQKind As cMap                 ' Fragenummer -> sub_type der Frage
Private mUiForm As frmListImport       ' nur Sichtpruefung (LiUi)
Private mRvMap As cMap                 ' Puffer von RecVal: zuletzt benutzte Spaltenzuordnung, Mitglied,
Private mRvMem As Long                 ' Spalte und Blatt (RecValReset, sobald sich Zuordnungen aendern)
Private mRvCol As Long
Private mRvSh As cLiSheet
Private mProfOn As Boolean             ' nur Laufzeitmessung (LiProfile)
Private mProfT As Double
Private mProfLog As String
Private mProfAgg As cMap
Private mDitto As cMap                 ' "dito": letzter Wert je Feld und Blatt

Private Sub RecValReset()
    Set mRvMap = Nothing: Set mRvSh = Nothing: mRvMem = 0: mRvCol = 0
End Sub

' Laufzeitmessung (Entwicklung): Zeit seit der letzten Marke protokollieren bzw. je Bezeichnung aufsummieren
Private Sub Prof(ByVal label As String)
    If Not mProfOn Then Exit Sub
    mProfLog = mProfLog & label & ": " & Format$(Timer - mProfT, "0.00") & " s" & vbLf
    mProfT = Timer
End Sub

Private Sub ProfAdd(ByVal label As String)
    If Not mProfOn Then Exit Sub
    mProfAgg.AddNum label, Timer - mProfT
    mProfT = Timer
End Sub

'==============================================================================
' Sitzung, Lesen
'==============================================================================
Public Sub LiReset()
    liPath = "": liMode = "": liLooksAsk = False
    Set liSheets = New Collection
    Set liTables = New Collection
    Set liTab = Nothing
    Set liCols = New Collection
    Set liQs = New Collection
    Set liNotes = New Collection
    Set liAns = New cMap
    Set liOvr = New cMap
    Set liValOvr = New cMap
    Set liAskQuestions = New Collection
    liAskRows = 0
    liRowSrc = Empty
    mCsv = False
    Set mMembers = Nothing: Set mUCols = Nothing: Set mRecs = Nothing: Set mPivot = Nothing
    RecValReset
    mPivotKind = "": mPivotYear = 0: mEuroHint = False
    Set mAskSheet = Nothing: Set mAskKeyCol = Nothing: Set mCp = Nothing: Set mQKind = Nothing
    mAskHdr = 0: mAskLast = 0: mAskMissing = 0: mAskPretest = 0
End Sub

' Liest die Datei. mode: "2ask" | "liste" | "" (automatisch). Liefert den erkannten Modus.
Public Function LiReadFile(ByVal path As String, ByVal mode As String) As String
    Dim ext As String, wb As Workbook, ws As Worksheet, data As Variant, isAsk As Boolean, msg As String, em As String
    LiReset
    liPath = path
    ext = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    If ext = "csv" Or ext = "txt" Then
        If mode = "2ask" Then
            RaiseUser "Für eine 2ask-Befragung bitte den kodierten Excel-Export mit Codeplan wählen. In 2ask: Ergebnisse, " & _
                      "Ergebnisse herunterladen, „Kodierte Ergebnisse/Antworten“, Format Excel (XLSX) und Haken bei „Codeplan hinzufügen“. " & _
                      "Eine CSV-Datei enthält keinen Codeplan."
        End If
        data = mIO.ReadCsv(path, "auto", "")
        If IsEmpty(data) Then RaiseUser "Die Datei enthält keine Daten."
        mCsv = True
        liSheets.Add LiSheetFromArray(FileNameOf(path), data, True)
        liMode = "liste"
        DetectTables
        LiReadFile = liMode
        Exit Function
    End If
    Set wb = mIO.OpenSourceWorkbook(path)
    Prof "Quelldatei öffnen"
    On Error GoTo failed
    isAsk = SheetExists(wb, "results") And SheetExists(wb, "code_plan")
    If mode = "2ask" Or (mode = "" And isAsk) Then
        If Not isAsk Then
            If SheetExists(wb, "results") Then
                msg = "Dieser 2ask-Export enthält keinen Codeplan. Bitte in 2ask erneut exportieren: „Kodierte Ergebnisse/Antworten“, " & _
                      "Format Excel (XLSX) und Haken bei „Codeplan hinzufügen“."
            Else
                msg = "Die Datei ist kein 2ask-Export (die Blätter „results“ und „code_plan“ fehlen). Für andere Listen bitte „Sonstige Liste“ wählen."
            End If
            GoTo userErr
        End If
        liMode = "2ask"
        ReadTwoAsk wb
    Else
        liMode = "liste"
        liLooksAsk = isAsk
        For Each ws In wb.Worksheets
            If ws.Visible = xlSheetVisible Then liSheets.Add LoadSheet(ws)
        Next
        Prof "Blätter lesen"
        DetectTables
        Prof "Tabellen finden"
    End If
    mIO.CloseSource wb
    Prof "Quelldatei schließen"
    LiReadFile = liMode
    Exit Function
userErr:
    mIO.CloseSource wb
    RaiseUser msg
    Exit Function
failed:
    em = Err.Description
    mIO.CloseSource wb
    RaiseUser em
End Function

Private Function FileNameOf(ByVal path As String) As String
    FileNameOf = Mid$(path, InStrRev(path, "\") + 1)
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next
End Function

Private Function LoadSheet(ByVal ws As Worksheet) As cLiSheet
    Dim sh As New cLiSheet, ur As Range, lastR As Long, lastC As Long, v As Variant, r As Long, c As Long
    Dim fr As Range, cl As Range, lo As ListObject, mc As Variant, one(1 To 1, 1 To 1) As Variant, nf As Variant
    sh.Name = ws.Name
    Set LoadSheet = sh
    Set ur = ws.UsedRange
    If Application.WorksheetFunction.CountA(ur) = 0 Then Exit Function
    lastR = ur.Row + ur.Rows.Count - 1
    lastC = ur.Column + ur.Columns.Count - 1
    If lastR > MAX_ROWS Then lastR = MAX_ROWS: sh.Truncated = True
    If lastC > MAX_COLS Then lastC = MAX_COLS: sh.Truncated = True
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    If Not IsArray(v) Then one(1, 1) = v: v = one
    For r = 1 To UBound(v, 1)
        For c = 1 To UBound(v, 2)
            If IsError(v(r, c)) Then v(r, c) = Empty: sh.ErrCells = sh.ErrCells + 1
        Next
    Next
    sh.SetValues v
    ' Formeln (Text der Formel, englisch)
    On Error Resume Next
    Set fr = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not fr Is Nothing Then
        If fr.Cells.CountLarge <= 50000 Then
            For Each cl In fr.Cells
                If cl.Row <= lastR And cl.Column <= lastC Then sh.AddFormula cl.Row, cl.Column, CStr(cl.Formula)
            Next
        End If
    End If
    ' verbundene Zellen (nur Spalten, in denen es welche gibt) und ausgeblendete Spalten
    For c = 1 To lastC
        mc = ws.Range(ws.Cells(1, c), ws.Cells(lastR, c)).MergeCells
        If IsNull(mc) Then
            ScanMerges ws, sh, c, lastR
        ElseIf CBool(mc) Then
            ScanMerges ws, sh, c, lastR
        End If
        If ws.Columns(c).Hidden Then sh.HiddenCols.SetVal CStr(c), True
    Next
    ' Zahlenformate: einheitlich je Spalte, sonst je Zahlenzelle
    For c = 1 To sh.C
        nf = ws.Range(ws.Cells(1, c), ws.Cells(lastR, c)).NumberFormat
        If IsNull(nf) Then
            For r = 1 To Application.Min(sh.R, 5000)
                If IsNumVal(v(r, c)) Then sh.AddCellFmt r, c, CStr(ws.Cells(r, c).NumberFormat)
            Next
            For r = sh.R To 1 Step -1
                If IsNumVal(v(r, c)) Then sh.ColFmt.SetVal CStr(c), CStr(ws.Cells(r, c).NumberFormat): Exit For
            Next
        Else
            sh.ColFmt.SetVal CStr(c), CStr(nf)
        End If
    Next
    ' Excel-Tabellen
    For Each lo In ws.ListObjects
        If lo.ShowHeaders Then
            sh.Tables.Add Array(lo.Range.Row, lo.Range.Column, lo.Range.Row + lo.Range.Rows.Count - 1, lo.Range.Column + lo.Range.Columns.Count - 1, lo.Name)
        End If
    Next
End Function

Private Sub ScanMerges(ByVal ws As Worksheet, ByVal sh As cLiSheet, ByVal c As Long, ByVal lastR As Long)
    Dim r As Long, rr As Long, cl As Range, ma As Range
    r = 1
    Do While r <= lastR
        Set cl = ws.Cells(r, c)
        If cl.MergeCells Then
            Set ma = cl.MergeArea
            For rr = ma.Row To ma.Row + ma.Rows.Count - 1
                If rr <= lastR Then sh.AddMerge rr, c, ma.Row & "|" & ma.Column
            Next
            r = ma.Row + ma.Rows.Count
        Else
            r = r + 1
        End If
    Loop
End Sub

' Blatt aus einem Werte-Array (CSV, Selbsttest)
Public Function LiSheetFromArray(ByVal nm As String, ByVal data As Variant, Optional ByVal isCsv As Boolean = False) As cLiSheet
    Dim sh As New cLiSheet, r As Long, c As Long, v() As Variant, lb1 As Long, lb2 As Long
    sh.Name = nm
    sh.IsCsv = isCsv
    lb1 = LBound(data, 1): lb2 = LBound(data, 2)
    ReDim v(1 To UBound(data, 1) - lb1 + 1, 1 To UBound(data, 2) - lb2 + 1)
    For r = 1 To UBound(v, 1)
        For c = 1 To UBound(v, 2)
            v(r, c) = data(r + lb1 - 1, c + lb2 - 1)
            If IsError(v(r, c)) Then v(r, c) = Empty
        Next
    Next
    sh.SetValues v
    Set LiSheetFromArray = sh
End Function

' Selbsttest/Testlauf: Blaetter direkt setzen und Tabellen suchen
Public Sub LiLoadSheets(ByVal sheets As Collection)
    Dim sh As cLiSheet
    LiReset
    liMode = "liste"
    For Each sh In sheets
        liSheets.Add sh
        If sh.IsCsv Then mCsv = True
    Next
    DetectTables
End Sub

'==============================================================================
' Hilfen fuer Zellwerte
'==============================================================================
Private Function IsNumVal(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbLong, vbInteger, vbByte, vbDecimal: IsNumVal = True
    End Select
End Function

' Zahl als Text ohne Tausenderpunkte (Postleitzahl 55116 bleibt 55116)
Private Function NumText(ByVal d As Double) As String
    If d = Fix(d) And Abs(d) < 1E+15 Then
        NumText = Format$(d, "0")
    Else
        NumText = Replace(NumInv(Round(d, 10)), ".", ",")
    End If
End Function

' Zellwert als Text (Datum deutsch, Zahl ohne Tausenderpunkte)
Private Function T_(ByVal v As Variant) As String
    If IsBlank(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            If CDbl(v) = Int(CDbl(v)) Then T_ = FmtDate(v) Else T_ = FmtDateTime(v)
        Case vbBoolean
            T_ = IIf(v, "Ja", "Nein")
        Case vbString
            ' Tabulatoren (z. B. "FSP L" + Tab aus Fachverfahren) wie Leerzeichen behandeln
            T_ = Trim$(Replace(Replace(Replace(CStr(v), vbCr, ""), ChrW(160), " "), vbTab, " "))
        Case vbDouble, vbSingle, vbCurrency, vbLong, vbInteger, vbByte, vbDecimal
            T_ = NumText(CDbl(v))
        Case Else
            T_ = Trim$(CStr(v))
    End Select
End Function

' Feldname aus Ueberschrift oder Fragetext: bereinigt und bei Bedarf sinnvoll gekuerzt
Private Function CleanName(ByVal s As String) As String
    CleanName = ShortName(FullName(s))
End Function

' Ueberschrift bereinigt, aber ungekuerzt (fuer den Hilfetext des Felds)
Private Function FullName(ByVal s As String) As String
    s = JoinBreaks(Replace(s, ChrW(173), ""))
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " ")
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, ";", ",")
    s = Replace(Replace(s, "[", "("), "]", ")")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Trim$(s)
    If Right$(s, 1) = ":" Then s = Trim$(Left$(s, Len(s) - 1))
    FullName = StripFootnote(s)
End Function

' Zeilenumbruch in Ueberschriften: nach Trennstrich zusammenziehen ("Ganztags-" + "betreuung" = "Ganztagsbetreuung",
' "Kfz-" + "Zulassung" = "Kfz-Zulassung", "Kinder-" + "und Jugendhilfe" bleibt getrennt), sonst Leerzeichen
Private Function JoinBreaks(ByVal s As String) As String
    Dim p As Long, nxt As String, w As String, q As Long, ch As String
    s = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
    p = InStr(s, vbLf)
    Do While p > 0
        nxt = Mid$(s, p + 1)
        ch = ""
        If p > 1 Then ch = Mid$(s, p - 1, 1)
        If ch = "-" Then
            q = InStr(nxt & " ", " ")
            w = LCase$(Left$(nxt, q - 1))
            If w = "und" Or w = "oder" Or w = "bzw." Or w = "sowie" Or w = "als" Then
                s = Left$(s, p - 1) & " " & nxt
            ElseIf Left$(nxt, 1) Like "[a-z]" Or InStr(ChrW(228) & ChrW(246) & ChrW(252) & ChrW(223), Left$(nxt, 1)) > 0 And Len(nxt) > 0 Then
                s = Left$(s, p - 2) & nxt
            Else
                s = Left$(s, p - 1) & nxt
            End If
        ElseIf ch = "/" Then
            s = Left$(s, p - 1) & nxt
        Else
            s = Left$(s, p - 1) & " " & nxt
        End If
        p = InStr(s, vbLf)
    Loop
    JoinBreaks = s
End Function

' Fussnotenzeichen am Ende: "Fläche km2 1)" -> "Fläche km2", "Postleitzahl3)" -> "Postleitzahl",
' "Zensus 20222)" -> "Zensus 2022", "Betrag*" -> "Betrag"; ausgeglichene Klammern ("Kinder (6-12)") bleiben
Private Function StripFootnote(ByVal s As String) As String
    Dim t As String, i As Long, nOpen As Long, nClose As Long, run As Long
    t = s
    Do While Len(t) > 1 And (Right$(t, 1) = "*" Or InStr(ChrW(&HB9) & ChrW(&HB2) & ChrW(&HB3), Right$(t, 1)) > 0)
        t = RTrim$(Left$(t, Len(t) - 1))
    Loop
    If Right$(t, 1) = ")" Then
        nOpen = Len(t) - Len(Replace(t, "(", ""))
        nClose = Len(t) - Len(Replace(t, ")", ""))
        If nClose > nOpen Then
            i = Len(t) - 1
            Do While i > 0
                If Not (Mid$(t, i, 1) Like "#") Then Exit Do
                i = i - 1
            Loop
            run = Len(t) - 1 - i
            If run >= 1 And run <= 2 Then
                t = RTrim$(Left$(t, i))
            ElseIf run > 2 Then
                t = Left$(t, Len(t) - 2)
            End If
        End If
    End If
    If Len(t) = 0 Then t = s
    StripFootnote = t
End Function

' Lange Ueberschriften und Fragetexte kuerzen; der vollstaendige Text kommt in den Hilfetext
Private Function ShortName(ByVal s As String) As String
    Dim p As Long, t As String, sep As Variant
    If Len(s) <= 70 Then ShortName = s: Exit Function
    ' zwei Saetze, Frage am Ende: die Frage ("... genutzt. Wie gut kamen Sie damit zurecht?")
    If Right$(s, 1) = "?" Then
        p = InStrRev(Left$(s, Len(s) - 1), ". ")
        If p > 0 Then
            t = Trim$(Mid$(s, p + 2))
            If Len(t) >= 15 And Len(t) <= 70 Then ShortName = t: Exit Function
        End If
    End If
    ' vor Klammer, Strich mit Leerzeichen, Komma oder Doppelpunkt
    For Each sep In Array(" (", " - ", " " & ChrW(&H2013) & " ", ", ", ": ")
        p = InStr(s, CStr(sep))
        If p > IIf(CStr(sep) = " (", 4, 15) And p <= 71 Then ShortName = Trim$(Left$(s, p - 1)): Exit Function
    Next
    ' am Wortende vor 70 Zeichen
    p = InStrRev(Left$(s, 70), " ")
    If p > 20 Then ShortName = Trim$(Left$(s, p - 1)) Else ShortName = Trim$(Left$(s, 70))
End Function

' Optionsnamen duerfen kein Semikolon und kein Komma enthalten
Public Function LiOptName(ByVal s As String) As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), ChrW(160), " ")
    s = Replace(Replace(s, ";", " "), ",", " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Trim$(s)
    If Len(s) > 100 Then s = Trim$(Left$(s, 100))
    LiOptName = s
End Function

' Technische Spaltenkoepfe (FALLNR, KENNZ_AKTIV) lesbar machen
Private Function NiceHeader(ByVal s As String) As String
    Dim parts As Variant, i As Long, w As String, res As String
    parts = Split(Replace(s, "_", " "), " ")
    For i = 0 To UBound(parts)
        w = parts(i)
        If Len(w) > 0 Then
            Select Case w
                Case "NR": w = "Nr."
                Case "KZ", "KENNZ": w = "Kennzeichen"
                Case "DAT", "DATUM": w = "Datum"
                Case "SACHB", "SACHBEARB", "SB": w = "Sachbearbeitung"
                Case "BEZ", "BEZEICHN": w = "Bezeichnung"
                Case "ANZ": w = "Anzahl"
                Case "BEM": w = "Bemerkung"
                Case "ID", "PLZ", "IBAN", "BIC", "URL"
                    ' Abkuerzung bleibt
                Case Else
                    If Len(w) > 3 And Right$(w, 2) = "NR" Then
                        w = Left$(w, 1) & LCase$(Mid$(w, 2, Len(w) - 3)) & "nummer"
                    Else
                        w = Left$(w, 1) & LCase$(Mid$(w, 2))
                    End If
            End Select
            res = res & IIf(Len(res) > 0, " ", "") & w
        End If
    Next
    NiceHeader = res
End Function

Private Function IsMarkText(ByVal s As String) As Boolean
    Select Case MarkNorm(s)
        Case "x", "ja", "j", "1", "wahr", "yes", "y", ChrW(&H2713), ChrW(&H2714), ChrW(&H2611), ChrW(&H2612)
            IsMarkText = True
    End Select
End Function

Private Function IsNoText(ByVal s As String) As Boolean
    Select Case MarkNorm(s)
        Case "nein", "n", "0", "falsch", "no", "-", ChrW(&H2610)
            IsNoText = True
    End Select
End Function

' "ja!", "(ja)", "Ja." wie "ja"
Private Function MarkNorm(ByVal s As String) As String
    s = LCase$(Trim$(s))
    If Len(s) > 2 And Left$(s, 1) = "(" And Right$(s, 1) = ")" Then s = Trim$(Mid$(s, 2, Len(s) - 2))
    Do While Len(s) > 1 And (Right$(s, 1) = "!" Or Right$(s, 1) = ".")
        s = Left$(s, Len(s) - 1)
    Loop
    MarkNorm = s
End Function

' Platzhalter fuer "kein Wert" in Zahlen- und Datumsspalten (amtliche Statistik: "-" nichts vorhanden, "." unbekannt)
Private Function IsPlaceholder(ByVal s As String) As Boolean
    Select Case LCase$(Trim$(s))
        Case "-", "--", ".", "...", "/", "?", "k.a.", "k. a.", "n.v.", "n. v.", "entfällt", ChrW(&H2013), ChrW(&H2014), ChrW(&H2026)
            IsPlaceholder = True
    End Select
End Function

' "dito", "siehe oben": gemeint ist der Wert der Zeile darueber
Private Function IsDittoText(ByVal s As String) As Boolean
    Select Case LCase$(Trim$(s))
        Case "dito", "ditto", "do.", "dto.", "s.o.", "s. o.", "siehe oben", "wie oben", """", ChrW(&H3003), ChrW(&H201E), ChrW(&H201C)
            IsDittoText = True
    End Select
End Function

Private Function IsMarkVal(ByVal v As Variant) As Boolean
    If IsBlank(v) Then Exit Function
    If VarType(v) = vbBoolean Then IsMarkVal = CBool(v): Exit Function
    If IsNumVal(v) Then IsMarkVal = (CDbl(v) = 1): Exit Function
    IsMarkVal = IsMarkText(CStr(v))
End Function

' Datum als Text: TT.MM.JJ(JJ) oder JJJJ-MM-TT (mit Jahr), optional Uhrzeit
Private Function TextDate(ByVal s As String, ByRef d As Date, ByRef hasTime As Boolean) As Boolean
    Dim dp As String, p As Long
    s = Trim$(s)
    If Len(s) < 6 Then Exit Function
    p = InStr(s, " ")
    If p > 0 Then dp = Left$(s, p - 1) Else dp = s
    If Len(dp) - Len(Replace(dp, ".", "")) <> 2 And Len(dp) - Len(Replace(dp, "-", "")) <> 2 Then Exit Function
    If Right$(dp, 1) = "." Then Exit Function
    TextDate = ParseDateDE(s, d, hasTime)
End Function

' Zahl als Text. In Excel als Text gespeicherte Werte mit Punkt ("1.12" Raum, "5.1" TOP)
' bleiben Text; in CSV-Dateien gilt ein Punkt mit Nachkommastellen als Dezimalpunkt.
Private Function TextNumber(ByVal s As String, ByRef d As Double, ByVal csv As Boolean) As Boolean
    Dim t As String
    t = Trim$(Replace(s, ChrW(160), " "))
    If Len(t) = 0 Then Exit Function
    ' Ungefaehrangaben und Schreibweisen aus Verwaltungslisten: "ca. 1.200 €", "rd. 300", "500,-"
    Select Case True
        Case LCase$(Left$(t, 4)) = "ca. ", LCase$(Left$(t, 4)) = "rd. ": t = Trim$(Mid$(t, 5))
        Case LCase$(Left$(t, 3)) = "ca.", LCase$(Left$(t, 3)) = "ca ": t = Trim$(Mid$(t, 4))
        Case LCase$(Left$(t, 5)) = "etwa ": t = Trim$(Mid$(t, 6))
        Case LCase$(Left$(t, 6)) = "circa ": t = Trim$(Mid$(t, 7))
        Case Left$(t, 1) = "~" Or Left$(t, 1) = ChrW(&H2248): t = Trim$(Mid$(t, 2))
    End Select
    If Right$(t, 3) = ",--" Then t = Left$(t, Len(t) - 3) Else If Right$(t, 2) = ",-" Then t = Left$(t, Len(t) - 2)
    If Len(t) = 0 Then Exit Function
    If t Like "*# #*" Then Exit Function          ' Zifferngruppen mit Leerzeichen: Telefon, IBAN
    If Left$(t, 1) = "+" Then Exit Function
    If Not csv Then
        ' "1.12" (Raum) bleibt Text; mit Euro ist der Punkt ein Tausenderpunkt ("ca. 1.200 €")
        If InStr(t, ".") > 0 And InStr(t, ",") = 0 And InStr(t, "€") = 0 And InStr(LCase$(t), "eur") = 0 Then Exit Function
    End If
    TextNumber = ParseNumberDE(t, d)
End Function

Private Function HasLeadingZero(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) >= 2 Then HasLeadingZero = (Left$(s, 1) = "0" And IsAllDigits(s))
End Function

Private Function IsSumText(ByVal s As String) As Boolean
    Dim w As Variant, t As String
    t = LCase$(Trim$(s))
    If Len(t) = 0 Then Exit Function
    For Each w In Split(SUM_WORDS, "|")
        If t = w Or Left$(t, Len(w) + 1) = w & " " Or Left$(t, Len(w) + 1) = w & ":" Then IsSumText = True: Exit Function
    Next
End Function

' Summenformel ueber mehrere Zeilen derselben Spalte (=SUM(C6:C25)); Zeilensummen (=SUM(B5:D5)) zaehlen nicht
Private Function IsVerticalSum(ByVal f As String) As Boolean
    Dim u As String, p As Long, a As String, b As String, i As Long
    u = UCase$(f)
    If InStr(u, "SUM(") = 0 And InStr(u, "SUBTOTAL(") = 0 And InStr(u, "COUNTA(") = 0 And InStr(u, "COUNT(") = 0 Then Exit Function
    p = InStr(u, ":")
    Do While p > 0
        a = "": b = ""
        i = p - 1
        Do While i >= 1
            If Mid$(u, i, 1) Like "#" Then a = Mid$(u, i, 1) & a Else Exit Do
            i = i - 1
        Loop
        i = p + 1
        Do While i <= Len(u)
            If Mid$(u, i, 1) Like "[A-Z$]" Then i = i + 1 Else Exit Do
        Loop
        Do While i <= Len(u)
            If Mid$(u, i, 1) Like "#" Then b = b & Mid$(u, i, 1): i = i + 1 Else Exit Do
        Loop
        If Len(a) > 0 And Len(b) > 0 Then
            If a <> b Then IsVerticalSum = True: Exit Function
        End If
        p = InStr(p + 1, u, ":")
    Loop
End Function

Private Function IsTextOnly(ByVal v As Variant, ByVal csv As Boolean) As Boolean
    Dim d As Double, dt As Date, ht As Boolean
    If IsBlank(v) Then Exit Function
    If VarType(v) <> vbString Then Exit Function
    If TextNumber(CStr(v), d, csv) Then Exit Function
    If TextDate(CStr(v), dt, ht) Then Exit Function
    IsTextOnly = True
End Function

Private Function ColLetter(ByVal c As Long) As String
    Dim s As String, m As Long
    Do While c > 0
        m = (c - 1) Mod 26
        s = Chr$(65 + m) & s
        c = (c - m - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Function FirstText(ByVal sh As cLiSheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long) As String
    Dim c As Long
    For c = c1 To c2
        If Not IsBlank(sh.Cell(r, c)) Then FirstText = T_(sh.Cell(r, c)): Exit Function
    Next
End Function

'--- Spaltenkoepfe deuten --------------------------------------------------------
Private Function HasWord(ByVal hdr As String, ByVal words As String) As Boolean
    Dim w As Variant
    For Each w In Split(words, "|")
        If InStr(hdr, CStr(w)) > 0 Then HasWord = True: Exit Function
    Next
End Function

Private Function IdLikeHeader(ByVal hdr As String) As Boolean
    Dim t As String
    t = " " & Replace(Replace(Replace(Replace(LCase$(hdr), ".", " "), "-", " "), "_", " "), "/", " ") & " "
    IdLikeHeader = (InStr(t, " nr ") > 0 Or InStr(t, "nummer") > 0 Or InStr(t, " plz ") > 0 Or InStr(t, "postleitzahl") > 0 _
        Or InStr(t, "konto") > 0 Or InStr(t, "kennz") > 0 Or InStr(t, " az ") > 0 Or InStr(t, "aktenzeichen") > 0 _
        Or InStr(t, "kostenstelle") > 0 Or InStr(t, " id ") > 0 Or InStr(t, "schlüssel") > 0 Or InStr(t, "telefon") > 0 _
        Or InStr(t, " top ") > 0 Or InStr(t, " iban ") > 0 Or InStr(t, " raum ") > 0 Or Right$(Trim$(t), 2) = "nr")
End Function

Private Function MoneyHeader(ByVal hdr As String) As Boolean
    hdr = LCase$(hdr)
    MoneyHeader = HasWord(hdr, "betrag|kosten|preis|budget|euro|€|zuschuss|honorar|gebühr|entgelt|ansatz|ausgaben|einnahmen") _
                  Or hdr = "wert" Or Left$(hdr, 5) = "wert " Or HasWord(hdr, "anschaffungswert|restwert|kaufpreis")
End Function

Private Function CountLikeHeader(ByVal hdr As String) As Boolean
    CountLikeHeader = HasWord(LCase$(hdr), "anzahl|menge|stück|plätze|anz.|teilnehm|personen|kinder")
End Function

Private Function PhoneHeader(ByVal hdr As String) As Boolean
    PhoneHeader = HasWord(LCase$(hdr), "tel|fon|mobil|handy|fax")
End Function

Private Function LongTextHeader(ByVal hdr As String) As Boolean
    LongTextHeader = HasWord(LCase$(hdr), "bemerkung|beschreibung|notiz|kommentar|anmerkung|begründung|erläuterung|sachstand|freitext|hinweise")
End Function

Private Function LabelHeader(ByVal hdr As String) As Boolean
    LabelHeader = HasWord(LCase$(hdr), "name|bezeichnung|titel|betreff|maßnahme|projekt|aufgabe|artikel|veranstaltung|thema|vorgang|antragsteller|schule|einrichtung")
End Function

Private Function NameHeader(ByVal hdr As String) As Boolean
    NameHeader = HasWord(LCase$(hdr), "name|person|ansprech|leitung|kontakt")
End Function

'==============================================================================
' Tabellen finden
'==============================================================================
Private Sub DetectTables()
    Dim i As Long, sh As cLiSheet, t As cLiTab, lo As Variant
    For i = 1 To liSheets.Count
        Set sh = liSheets(i)
        If sh.R > 0 Then
            If sh.Tables.Count > 0 Then
                For Each lo In sh.Tables
                    Set t = AnalyzeBlock(sh, i, CLng(lo(0)), CLng(lo(2)), True, CLng(lo(1)), CLng(lo(3)), 0)
                    If Not t Is Nothing Then liTables.Add t
                Next
            Else
                DetectInSheet sh, i
            End If
        End If
    Next
    AddMergedTables
End Sub

Private Sub DetectInSheet(ByVal sh As cLiSheet, ByVal si As Long)
    Dim nz() As Long, r As Long, segS As New Collection, segE As New Collection, i As Long, b1 As Long, b2 As Long, t As cLiTab
    Dim titlesOnly As Boolean, rr As Long, gap As Long, joinIt As Boolean
    ReDim nz(1 To sh.R)
    For r = 1 To sh.R: nz(r) = sh.FilledInRow(r): Next
    ' Abschnitte aus nicht leeren Zeilen
    r = 1
    Do While r <= sh.R
        If nz(r) > 0 Then
            segS.Add r
            Do While r <= sh.R
                If nz(r) = 0 Then Exit Do
                r = r + 1
            Loop
            segE.Add r - 1
        Else
            r = r + 1
        End If
    Loop
    If segS.Count = 0 Then Exit Sub
    b1 = segS(1): b2 = segE(1)
    For i = 2 To segS.Count + 1
        joinIt = False
        If i <= segS.Count Then
            titlesOnly = True
            For rr = b1 To b2
                If nz(rr) > 1 Then titlesOnly = False: Exit For
            Next
            gap = segS(i) - b2 - 1
            If titlesOnly And gap <= 3 Then
                joinIt = True                                   ' Titel ueber der Tabelle
            ElseIf gap = 1 Then
                joinIt = Not StartsNewTable(sh, nz, CLng(segS(i)), CLng(segE(i)))   ' eine Leerzeile in der Tabelle
            End If
        End If
        If joinIt Then
            b2 = segE(i)
        Else
            Set t = AnalyzeBlock(sh, si, b1, b2, False, 0, 0, 0)
            If Not t Is Nothing Then liTables.Add t
            If i <= segS.Count Then b1 = segS(i): b2 = segE(i)
        End If
    Next
End Sub

' Abschnitt beginnt mit einer Titelzeile und einer Kopfzeile aus Text
Private Function StartsNewTable(ByVal sh As cLiSheet, ByRef nz() As Long, ByVal s As Long, ByVal e As Long) As Boolean
    Dim c As Long, n As Long
    If nz(s) <> 1 Or s >= e Then Exit Function
    For c = 1 To sh.C
        If Not IsBlank(sh.Cell(s + 1, c)) Then
            If Not IsTextOnly(sh.Cell(s + 1, c), sh.IsCsv) Then Exit Function
            n = n + 1
        End If
    Next
    StartsNewTable = (n >= 2)
End Function

' forceHdr: 0 = automatisch, 1/2 = Kopf in Zeile r1 ueber 1 bzw. 2 Zeilen
Private Function AnalyzeBlock(ByVal sh As cLiSheet, ByVal si As Long, ByVal r1 As Long, ByVal r2 As Long, ByVal fromLO As Boolean, _
                              ByVal cMin As Long, ByVal cMax As Long, ByVal forceHdr As Long) As cLiTab
    Dim t As New cLiTab, r As Long, c As Long, c1 As Long, c2 As Long, nzr As Long, h As Long, lastData As Long, dense As Long
    Dim top As String, bot As String, nm As String, used As New cMap, k As Long, cls As String, lo As Long, hi As Long
    Dim botCount As New cMap, nCaps As Long, nHdr As Long, nForm As Long
    Dim ann As String, annCount As New cMap, grpTop As String, full As String
    t.SheetIdx = si: t.R1 = r1: t.R2 = r2: t.FromListObject = fromLO
    lo = IIf(cMin > 0, cMin, 1)
    hi = IIf(cMax > 0, cMax, sh.C)
    If hi > sh.C Then hi = sh.C
    For r = r1 To r2
        For c = lo To hi
            If Not IsBlank(sh.Cell(r, c)) Then
                If c1 = 0 Or c < c1 Then c1 = c
                If c > c2 Then c2 = c
            End If
        Next
    Next
    If c1 = 0 Then Exit Function
    t.C1 = c1: t.C2 = c2
    ' Titel und Kopfzeile
    If fromLO Or forceHdr > 0 Then
        h = r1
    Else
        For r = r1 To r2
            nzr = sh.FilledInRow(r, c1, c2)
            If nzr >= 2 Then h = r: Exit For
            If nzr = 1 Then
                t.RowClass.SetVal CStr(r), "titel"
                t.Title = Trim$(t.Title & " " & FirstText(sh, r, c1, c2))
            End If
        Next
    End If
    If h = 0 Then Exit Function
    ' Formular: Bezeichnung mit Doppelpunkt, Wert daneben
    If Not fromLO And forceHdr = 0 Then
        If IsFormBlock(sh, h, r2, c1, c2) Then
            t.Kind = "formular"
            t.HdrRow = h: t.HdrRows = 0
            For r = h To r2
                If Right$(T_(sh.Cell(r, c1)), 1) = ":" And FormValueCol(sh, r, c1, c2) > 0 Then
                    t.RowClass.SetVal CStr(r), "daten"
                ElseIf sh.FilledInRow(r, c1, c2) > 0 Then
                    t.RowClass.SetVal CStr(r), "notiz"
                End If
            Next
            t.Note = "Formular mit " & t.CountClass("daten") & " Angaben, wird ein Eintrag"
            t.Score = t.CountClass("daten")
            Set AnalyzeBlock = t
            Exit Function
        End If
    End If
    t.Kind = "liste"
    t.HdrRow = h: t.HdrRows = 1
    If forceHdr = 2 Then
        t.HdrRows = 2
    ElseIf forceHdr = 0 And Not fromLO And h + 2 <= r2 Then
        If IsSecondHeaderRow(sh, h, c1, c2) Then t.HdrRows = 2
    End If
    If h + t.HdrRows > r2 Then Exit Function
    ' Hinweiszeile unter dem Kopf (Stichtag, englische Uebersetzung, "Schlüssel"/"Bezeichnung")
    If forceHdr = 0 And Not fromLO Then
        If IsAnnotationRow(sh, h + t.HdrRows, c1, c2, r2) Then t.AnnRow = h + t.HdrRows
    End If
    If t.FirstDataRow() > r2 Then Exit Function
    ' Spaltennamen
    If t.HdrRows = 2 Then
        For c = c1 To c2
            bot = SecondRowText(sh, h, c)
            If Len(bot) > 0 Then botCount.AddNum bot, 1
        Next
    End If
    If t.AnnRow > 0 Then
        For c = c1 To c2
            ann = T_(sh.Cell(t.AnnRow, c))
            If Len(ann) > 0 Then annCount.AddNum ann, 1
        Next
    End If
    For c = c1 To c2
        top = HeaderText(sh, h, c)
        bot = ""
        If t.HdrRows = 2 Then bot = SecondRowText(sh, h, c)
        If Len(bot) > 0 Then
            If (CDbl(botCount.GetVal(bot, 0)) > 1 Or GenericSubHeader(bot)) And top <> bot Then
                grpTop = top
                If Len(grpTop) = 0 And GenericSubHeader(bot) Then grpTop = GroupTop(sh, h, c, c1)
                If Len(grpTop) > 0 Then nm = GroupWord(grpTop) & " " & bot Else nm = bot
            Else
                nm = bot
            End If
            If Len(top) > 0 And top <> bot Then t.ColTop.SetVal CStr(c), CleanName(top)
        Else
            nm = top
        End If
        full = FullName(nm)
        ' Hinweiszeile: wiederholte Unterbegriffe ergaenzen den Namen, Einzeltexte kommen in den Hilfetext
        If t.AnnRow > 0 Then
            ann = T_(sh.Cell(t.AnnRow, c))
            If Len(ann) > 0 Then
                If (Len(full) = 0 Or ((CDbl(annCount.GetVal(ann, 0)) > 1 Or GenericSubHeader(ann)) And Len(ann) <= 20)) And Not IsStampText(ann) Then
                    grpTop = full
                    If Len(grpTop) = 0 Then grpTop = GroupTop(sh, h, c, c1)
                    nm = Trim$(GroupWord(grpTop) & " " & ann)
                    full = FullName(nm)
                Else
                    t.ColHelp.SetVal CStr(c), FullName(ann)
                End If
            End If
        End If
        nm = ShortName(full)
        If Len(full) > Len(nm) Then t.ColHelp.SetVal CStr(c), full & IIf(t.ColHelp.Exists(CStr(c)), " (" & CStr(t.ColHelp.GetVal(CStr(c), "")) & ")", "")
        If Len(nm) = 0 Then nm = "Spalte " & ColLetter(c)
        t.ColName.SetVal CStr(c), nm
        nHdr = nHdr + 1
        If Len(nm) >= 3 And nm = UCase$(nm) And nm Like "*[A-Z]*" Then nCaps = nCaps + 1
    Next
    If nHdr > 1 And nCaps >= nHdr * 0.8 Then
        For c = c1 To c2
            t.ColName.SetVal CStr(c), CleanName(NiceHeader(CStr(t.ColName.GetVal(CStr(c)))))
        Next
    End If
    For c = c1 To c2
        nm = CStr(t.ColName.GetVal(CStr(c)))
        If used.Exists(nm) Then
            k = 2
            Do While used.Exists(nm & " (" & k & ")"): k = k + 1: Loop
            nm = nm & " (" & k & ")"
            t.ColName.SetVal CStr(c), nm
        End If
        used.SetVal nm, True
    Next
    ' Zeilen einordnen
    For r = h To t.FirstDataRow() - 1: t.RowClass.SetVal CStr(r), "kopf": Next
    For r = t.FirstDataRow() To r2
        nzr = sh.FilledInRow(r, c1, c2)
        nForm = 0
        If nzr > 0 Then
            For c = c1 To c2
                If sh.IsFormula(r, c) Then
                    If IsVerticalSum(sh.FormulaText(r, c)) Then nForm = nForm + 1
                End If
            Next
        End If
        If nzr = 0 Then
            cls = "leer"
        ElseIf IsSumText(FirstText(sh, r, c1, IIf(c1 + 1 <= c2, c1 + 1, c2))) Or nForm > 0 Then
            cls = "summe"
        ElseIf nzr = 1 And c2 > c1 And IsTextOnly(sh.Cell(r, c1), sh.IsCsv) Then
            cls = "gruppe"
        Else
            cls = "daten": dense = dense + 1: lastData = r
        End If
        t.RowClass.SetVal CStr(r), cls
    Next
    For r = t.FirstDataRow() To r2
        If t.RowClass.GetVal(CStr(r)) = "gruppe" Then
            If dense < 2 Then
                t.RowClass.SetVal CStr(r), "daten"
            ElseIf r > lastData Then
                t.RowClass.SetVal CStr(r), "notiz"
            End If
        End If
    Next
    If IsWeekGrid(t) Then t.Kind = "raster"
    If t.DataRowCount() = 0 Then Exit Function
    t.Score = t.DataRowCount() * (c2 - c1 + 1)
    If t.Kind = "raster" Then t.Score = 0
    t.Note = DescribeTable(t)
    Set AnalyzeBlock = t
End Function

' Kopftext einer Zelle; leer in einem waagerecht verbundenen Bereich -> Text der Ursprungszelle
Private Function HeaderText(ByVal sh As cLiSheet, ByVal r As Long, ByVal c As Long) As String
    Dim o As String, a As Variant
    HeaderText = T_(sh.Cell(r, c))
    If Len(HeaderText) > 0 Then Exit Function
    o = sh.MergeOrigin(r, c)
    If Len(o) > 0 Then
        a = Split(o, "|")
        If CLng(a(0)) = r Then HeaderText = T_(sh.Cell(r, CLng(a(1))))
    End If
End Function

' Text der zweiten Kopfzeile; "" wenn die Zelle senkrecht mit der ersten verbunden ist
Private Function SecondRowText(ByVal sh As cLiSheet, ByVal h As Long, ByVal c As Long) As String
    Dim o As String, a As Variant
    o = sh.MergeOrigin(h + 1, c)
    If Len(o) > 0 Then
        a = Split(o, "|")
        If CLng(a(0)) <= h Then Exit Function
        If CLng(a(1)) <> c Then SecondRowText = T_(sh.Cell(CLng(a(0)), CLng(a(1)))): Exit Function
    End If
    SecondRowText = T_(sh.Cell(h + 1, c))
End Function

' Zeile enthaelt typische Datenwerte (Zahl, Datum, Kreuz)
Private Function RowHasDataSigns(ByVal sh As cLiSheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long) As Boolean
    Dim c As Long, v As Variant
    For c = c1 To c2
        v = sh.Cell(r, c)
        If Not IsBlank(v) Then
            If Not IsTextOnly(v, sh.IsCsv) Then RowHasDataSigns = True: Exit Function
            If IsMarkText(T_(v)) Then RowHasDataSigns = True: Exit Function
        End If
    Next
End Function

Private Function IsSecondHeaderRow(ByVal sh As cLiSheet, ByVal h As Long, ByVal c1 As Long, ByVal c2 As Long) As Boolean
    Dim c As Long, gapFilled As Long, nz2 As Long, o As String
    If Not RowAllText(sh, h + 1, c1, c2) Then Exit Function
    If RowHasDataSigns(sh, h + 1, c1, c2) Then Exit Function
    For c = c1 To c2
        If Not IsBlank(sh.Cell(h + 1, c)) Then
            nz2 = nz2 + 1
            If IsBlank(sh.Cell(h, c)) Then
                gapFilled = gapFilled + 1
            Else
                o = sh.MergeOrigin(h, c)
                If Len(o) > 0 And o <> (h & "|" & c) Then gapFilled = gapFilled + 1
            End If
        End If
    Next
    If nz2 = 0 Or gapFilled = 0 Then Exit Function
    ' die Zeile danach muss nach Daten aussehen
    IsSecondHeaderRow = RowHasDataSigns(sh, h + 2, c1, c2) Or Not RowAllText(sh, h + 2, c1, c2) _
                        Or sh.FilledInRow(h + 2, c1, c2) <> sh.FilledInRow(h + 1, c1, c2)
End Function

' Zeile direkt unter dem Kopf, die nur Text enthaelt, waehrend darunter in mindestens zwei dieser Spalten
' Zahlen oder Daten stehen: Stichtag ("31.12.2025 (Jahr)"), Uebersetzung ("start date"), "Schlüssel"/"Bezeichnung".
' Sie gehoert zum Kopf und ist kein Datensatz.
Private Function IsAnnotationRow(ByVal sh As cLiSheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal r2 As Long) As Boolean
    Dim c As Long, rr As Long, nTyped As Long, nTot As Long, conflicts As Long, filled As Long
    If r >= r2 Or r > sh.R Then Exit Function
    If Not RowAllText(sh, r, c1, c2) Then Exit Function
    For c = c1 To c2
        If Not IsBlank(sh.Cell(r, c)) Then
            filled = filled + 1
            nTyped = 0: nTot = 0
            For rr = r + 1 To Application.Min(r2, r + 80)
                If Not IsBlank(sh.Cell(rr, c)) Then
                    nTot = nTot + 1
                    If Not IsTextOnly(sh.Cell(rr, c), sh.IsCsv) Then nTyped = nTyped + 1
                    If nTot >= 20 Then Exit For
                End If
            Next
            If nTot >= 3 And nTyped >= nTot * 0.8 Then conflicts = conflicts + 1
        End If
    Next
    IsAnnotationRow = (conflicts >= 2 And filled >= 2)
End Function

' Unterbegriffe, die nur zusammen mit der Gruppenueberschrift verstaendlich sind ("insgesamt", "männlich", "je km2")
Private Function GenericSubHeader(ByVal s As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(s))
    Select Case t
        Case "insgesamt", "gesamt", "zusammen", "männlich", "weiblich", "divers", "darunter", "davon", "anzahl", "summe", _
             "absolut", "prozent", "%", "in %", "schlüssel", "bezeichnung", "name", "nr.", "nummer", "code", "betrag", "wert", _
             "ist", "soll", "plan", "vorjahr", "veränderung", "anteil"
            GenericSubHeader = True
        Case Else
            GenericSubHeader = (Left$(t, 3) = "je " Or Left$(t, 9) = "darunter ")
    End Select
End Function

' Stichtag oder Stand ("31.12.2025 (Jahr)", "Gebietsstand am 30.09.2026"): Hinweis, kein Namensteil
Private Function IsStampText(ByVal s As String) As Boolean
    IsStampText = (s Like "*##.##.####*" Or s Like "*##.##.##*" Or HasWord(LCase$(s), "stand|stichtag|(jahr)|quartal"))
End Function

' Gruppenueberschrift links in der oberen Kopfzeile (nicht verbundene Kopfzellen: Text steht nur ueber der ersten Spalte)
Private Function GroupTop(ByVal sh As cLiSheet, ByVal h As Long, ByVal c As Long, ByVal c1 As Long) As String
    Dim cc As Long
    For cc = c To c1 Step -1
        GroupTop = HeaderText(sh, h, cc)
        If Len(GroupTop) > 0 Then Exit Function
    Next
End Function

' Kern einer Gruppenueberschrift: "Bevölkerung auf Grundlage des Zensus 2022" -> "Bevölkerung"
Private Function GroupWord(ByVal s As String) As String
    Dim w As Variant, p As Long
    s = FullName(s)
    For Each w In Array(" auf ", " nach ", " gemäß ", " laut ", " (", " am ", " zum ", " für ")
        p = InStr(s, CStr(w))
        If p > 4 Then s = Left$(s, p - 1)
    Next
    GroupWord = Trim$(s)
End Function

Private Function RowAllText(ByVal sh As cLiSheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long) As Boolean
    Dim c As Long, n As Long
    If r > sh.R Then Exit Function
    For c = c1 To c2
        If Not IsBlank(sh.Cell(r, c)) Then
            If Not IsTextOnly(sh.Cell(r, c), sh.IsCsv) Then Exit Function
            n = n + 1
        End If
    Next
    RowAllText = (n >= 1)
End Function

Private Function FormValueCol(ByVal sh As cLiSheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long) As Long
    Dim c As Long
    For c = c1 + 1 To c2
        If Not IsBlank(sh.Cell(r, c)) Then FormValueCol = c: Exit Function
    Next
End Function

Private Function IsFormBlock(ByVal sh As cLiSheet, ByVal r1 As Long, ByVal r2 As Long, ByVal c1 As Long, ByVal c2 As Long) As Boolean
    Dim r As Long, nRows As Long, nColon As Long, nzr As Long
    For r = r1 To r2
        nzr = sh.FilledInRow(r, c1, c2)
        If nzr > 2 Then Exit Function
        If nzr > 0 Then
            nRows = nRows + 1
            If Right$(T_(sh.Cell(r, c1)), 1) = ":" Then nColon = nColon + 1
        End If
    Next
    IsFormBlock = (nRows >= 3 And nColon >= nRows * 0.6)
End Function

' Wochenraster: mindestens drei Wochentage als Spalten UND Uhrzeiten in der ersten Spalte
' (eine Anwesenheitsliste mit Spalten Mo, Di, Mi und Namen in der ersten Spalte ist kein Raster)
Private Function IsWeekGrid(ByVal t As cLiTab) As Boolean
    Dim k As Variant, n As Long, sh As cLiSheet, r As Long, nV As Long, nT As Long, cv As Variant
    For Each k In t.ColName.Items
        Select Case LCase$(Replace(CStr(k), ".", ""))
            Case "montag", "dienstag", "mittwoch", "donnerstag", "freitag", "samstag", "sonntag", "mo", "di", "mi", "do", "fr", "sa", "so"
                n = n + 1
        End Select
    Next
    If n < 3 Then Exit Function
    Set sh = liSheets(t.SheetIdx)
    For r = t.FirstDataRow() To t.R2
        cv = sh.Cell(r, t.C1)
        If Not IsBlank(cv) Then
            nV = nV + 1
            If VarType(cv) = vbDate Then
                If CDbl(cv) < 1 Then nT = nT + 1
            ElseIf IsNumVal(cv) Then
                If CDbl(cv) > 0 And CDbl(cv) < 1 Then nT = nT + 1
            ElseIf IsTimeText(T_(cv)) Then
                nT = nT + 1
            End If
        End If
    Next
    IsWeekGrid = (nV > 0 And nT >= nV * 0.6)
End Function

Private Function IsTimeText(ByVal s As String) As Boolean
    s = LCase$(Trim$(s))
    IsTimeText = (s Like "#:##*" Or s Like "##:##*" Or s Like "# uhr*" Or s Like "## uhr*" Or s Like "#.## uhr*" Or s Like "##.## uhr*")
End Function

Private Function DescribeTable(ByVal t As cLiTab) As String
    Dim s As String
    If t.Kind = "raster" Then DescribeTable = "Wochenraster (Uhrzeit × Tag), wird nicht unterstützt": Exit Function
    If t.HdrRows = 2 Then s = "Kopf in Zeile " & t.HdrRow & " und " & t.HdrRow + 1 Else s = "Kopf in Zeile " & t.HdrRow
    If t.AnnRow > 0 Then s = s & ", Hinweiszeile " & t.AnnRow
    If t.FromListObject Then s = s & ", Excel-Tabelle"
    If t.CountClass("gruppe") > 0 Then s = s & ", " & Plural(t.CountClass("gruppe"), "Gruppenzeile", "Gruppenzeilen")
    If t.CountClass("summe") > 0 Then s = s & ", " & Plural(t.CountClass("summe"), "Summenzeile", "Summenzeilen")
    DescribeTable = s
End Function

' Gleich aufgebaute Tabellen auf mehreren Blaettern -> zusaetzlicher Eintrag "zusammen" (zuerst)
Private Sub AddMergedTables()
    Dim i As Long, j As Long, a As cLiTab, b As cLiTab, used As New cMap, grp As Collection, t As cLiTab, names As String
    Dim merged As New Collection
    For i = 1 To liTables.Count
        Set a = liTables(i)
        If a.Kind = "liste" And Not used.Exists(CStr(i)) Then
            Set grp = New Collection
            grp.Add a
            For j = i + 1 To liTables.Count
                Set b = liTables(j)
                If b.Kind = "liste" And b.SheetIdx <> a.SheetIdx And Not used.Exists(CStr(j)) Then
                    If HeaderSimilarity(a, b) >= 0.6 Then grp.Add b: used.SetVal CStr(j), True
                End If
            Next
            If grp.Count >= 2 Then
                Set t = New cLiTab
                t.Kind = "zusammen"
                t.SheetIdx = a.SheetIdx
                Set t.Members = grp
                names = ""
                For Each b In grp
                    names = names & IIf(Len(names) > 0, ", ", "") & liSheets(b.SheetIdx).Name
                Next
                t.Note = grp.Count & " gleich aufgebaute Blätter zusammen: " & Truncate(names, 120)
                t.Score = 1E+30
                merged.Add t
                used.SetVal CStr(i), True
            End If
        End If
    Next
    For Each t In merged
        If liTables.Count = 0 Then liTables.Add t Else liTables.Add t, , 1
    Next
End Sub

' Anteil gleicher Spaltenkoepfe; "Kosten" und "Kosten ca." oder "dringend?" und "dringend" gelten als gleich
Private Function HeaderSimilarity(ByVal a As cLiTab, ByVal b As cLiTab) As Double
    Dim k As Variant, sa As New Collection, inter As Long, uni As Long, x As Variant, hit As Boolean
    For Each k In a.ColName.Items: sa.Add HeadNorm(CStr(k)): Next
    uni = sa.Count
    For Each k In b.ColName.Items
        hit = False
        For Each x In sa
            If SameHead(CStr(x), HeadNorm(CStr(k))) Then hit = True: Exit For
        Next
        If hit Then inter = inter + 1 Else uni = uni + 1
    Next
    If uni > 0 Then HeaderSimilarity = inter / uni
End Function

Private Function HeadNorm(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(Replace(Replace(Replace(s, "?", ""), ":", ""), "*", ""), ".", "")
    HeadNorm = Trim$(s)
End Function

Private Function SameHead(ByVal a As String, ByVal b As String) As Boolean
    If a = b Then SameHead = True: Exit Function
    If Len(a) >= 4 And Left$(b, Len(a) + 1) = a & " " Then SameHead = True: Exit Function
    If Len(b) >= 4 And Left$(a, Len(b) + 1) = b & " " Then SameHead = True
End Function

' Beste Tabelle als Vorschlag (hoechste Bewertung)
Public Function LiBestTable() As Long
    Dim i As Long, best As Double, t As cLiTab
    best = -1
    For i = 1 To liTables.Count
        Set t = liTables(i)
        If t.Score > best Then best = t.Score: LiBestTable = i
    Next
End Function

' Beschreibung fuer die Auswahlliste: Array(Blatt, Bereich, Eintraege, Spalten, Erkannt)
Public Function LiTableRow(ByVal i As Long) As Variant
    Dim t As cLiTab, sh As cLiSheet
    Set t = liTables(i)
    If t.Kind = "zusammen" Then
        LiTableRow = Array(t.Members.Count & " Blätter", "", FmtInt(t.DataRowCount()), CStr(t.Members(1).ColCount()), t.Note)
        Exit Function
    End If
    Set sh = liSheets(t.SheetIdx)
    LiTableRow = Array(sh.Name, ColLetter(t.C1) & t.R1 & ":" & ColLetter(t.C2) & t.R2, FmtInt(t.DataRowCount()), _
                       IIf(t.Kind = "formular", CStr(t.CountClass("daten")), CStr(t.ColCount())), t.Note)
End Function

' Vorschau der Tabelle: Array 0-basiert (Zeilen, 1 + Spalten); erste Spalte = Zeile bzw. Art
Public Function LiPreview(ByVal i As Long, ByVal maxRows As Long, ByVal maxCols As Long, Optional ByVal maxLen As Long = 28) As Variant
    Dim t As cLiTab, sh As cLiSheet, r As Long, c As Long, out() As Variant, n As Long, nc As Long, cls As String
    Set t = liTables(i)
    If t.Kind = "zusammen" Then Set t = t.Members(1)
    Set sh = liSheets(t.SheetIdx)
    nc = LiPreviewColCount(i, maxCols)
    ReDim out(0 To maxRows - 1, 0 To nc)
    For r = t.R1 To t.R2
        If n >= maxRows Then Exit For
        cls = CStr(t.RowClass.GetVal(CStr(r), ""))
        If cls <> "leer" And sh.FilledInRow(r, t.C1, t.C2) > 0 Then
            Select Case cls
                Case "kopf": out(n, 0) = "Kopf"
                Case "summe": out(n, 0) = "Summe"
                Case "gruppe": out(n, 0) = "Gruppe"
                Case "titel": out(n, 0) = "Titel"
                Case "notiz": out(n, 0) = "Notiz"
                Case Else: out(n, 0) = "Zeile " & r
            End Select
            For c = 0 To nc - 1
                out(n, c + 1) = Truncate(T_(sh.Cell(r, t.C1 + c)), maxLen)
            Next
            n = n + 1
        End If
    Next
    If n = 0 Then n = 1
    LiPreview = SubRows(out, n, nc + 1)
End Function

' Anzahl der Spalten in der Vorschau (fuer die Spaltenbreiten im Assistenten)
Public Function LiPreviewColCount(ByVal i As Long, ByVal maxCols As Long) As Long
    Dim t As cLiTab
    Set t = liTables(i)
    If t.Kind = "zusammen" Then Set t = t.Members(1)
    LiPreviewColCount = t.ColCount()
    If LiPreviewColCount > maxCols Then LiPreviewColCount = maxCols
End Function

Private Function SubRows(ByRef a() As Variant, ByVal n As Long, ByVal cols As Long) As Variant
    Dim b() As Variant, r As Long, c As Long
    ReDim b(0 To n - 1, 0 To cols - 1)
    For r = 0 To n - 1
        For c = 0 To cols - 1: b(r, c) = a(r, c): Next
    Next
    SubRows = b
End Function

' Kopfzeile von Hand setzen (Zeile im Blatt, 1 oder 2 Zeilen)
Public Sub LiSetHeader(ByVal i As Long, ByVal hdrRow As Long, ByVal nRows As Long)
    Dim t As cLiTab, sh As cLiSheet, nt As cLiTab, r As Long, lastR As Long
    Set t = liTables(i)
    If t.Kind = "zusammen" Then RaiseUser "Bei zusammengeführten Blättern lässt sich die Kopfzeile nicht ändern. Bitte ein einzelnes Blatt wählen."
    Set sh = liSheets(t.SheetIdx)
    If hdrRow < 1 Or hdrRow > sh.R Then RaiseUser "Bitte eine Zeile zwischen 1 und " & sh.R & " angeben."
    lastR = t.R2
    If hdrRow + nRows > lastR Then lastR = sh.R
    Set nt = AnalyzeBlock(sh, t.SheetIdx, hdrRow, lastR, False, IIf(t.FromListObject, t.C1, 0), IIf(t.FromListObject, t.C2, 0), IIf(nRows = 2, 2, 1))
    If nt Is Nothing Then RaiseUser "Unter Zeile " & hdrRow & " wurden keine Datenzeilen gefunden."
    For r = t.R1 To hdrRow - 1
        If sh.FilledInRow(r) > 0 Then nt.RowClass.SetVal CStr(r), "titel"
    Next
    If t.R1 < nt.R1 Then nt.R1 = t.R1
    nt.Title = t.Title
    nt.Note = DescribeTable(nt)
    liTables.Add nt, , , i
    liTables.Remove i
End Sub

'==============================================================================
' Planung der Felder (sonstige Liste)
'==============================================================================
Public Sub LiChooseTable(ByVal i As Long)
    Dim t As cLiTab
    Set t = liTables(i)
    If t.Kind = "raster" Then
        RaiseUser "Diese Tabelle ist ein Wochenraster (Uhrzeiten in den Zeilen, Wochentage in den Spalten). Daraus kann MidPoint keine " & _
                  "Einträge bilden. Bitte die Termine als Liste anlegen (je Zeile ein Termin mit Beginn, Ende und Raum) und diese Liste übernehmen."
    End If
    Set liTab = t
    Set liAns = New cMap
    Set liOvr = New cMap
    Set liValOvr = New cMap
    LiPlan
End Sub

Private Function Ans(ByVal id As String, ByVal dflt As Long) As Long
    If liAns.Exists(id) Then Ans = CLng(liAns.GetVal(id)) Else Ans = dflt
End Function

Private Sub AddQ(ByVal id As String, ByVal topic As String, ByVal txt As String, ByVal choices As Variant, ByVal dflt As Long)
    Dim q As New cLiQ
    q.ID = id: q.Topic = topic: q.Text = txt: q.Choices = choices: q.Default_ = dflt
    q.Answer = Ans(id, dflt)
    If q.Answer < 0 Or q.Answer > UBound(choices) Then q.Answer = dflt
    liQs.Add q
End Sub

Private Function HasQuestion(ByVal id As String) As Boolean
    Dim q As cLiQ
    For Each q In liQs
        If q.ID = id Then HasQuestion = True: Exit Function
    Next
End Function

' Antwort auf eine Rueckfrage setzen (der Plan wird danach neu gebaut)
Public Sub LiSetAnswer(ByVal id As String, ByVal idx As Long)
    liAns.SetVal id, idx
End Sub

' Baut Felder und Rueckfragen aus der gewaehlten Tabelle und den bisherigen Antworten
Public Sub LiPlan()
    Set liCols = New Collection
    Set liQs = New Collection
    Set liNotes = New Collection
    If liMode = "2ask" Then PlanTwoAsk: Exit Sub
    BuildMembers
    Set mPivot = Nothing: mPivotKind = "": mPivotYear = 0
    If liTab.Kind = "formular" Then PlanForm: Exit Sub
    BuildUnifiedCols
    BuildRecords
    Prof "Datensätze bilden"
    QuestionMergeCols
    PlanPivot
    Prof "Zusammenführen und Zeitraumspalten"
    PlanColumns
    Prof "Felder planen"
    ApplyOverrides
End Sub

Private Sub BuildMembers()
    Dim m As cLiTab, s As String
    Set mMembers = New Collection
    If liTab.Kind = "zusammen" Then
        For Each m In liTab.Members: mMembers.Add m: Next
    Else
        mMembers.Add liTab
    End If
    mEuroHint = False
    For Each m In mMembers
        s = LCase$(m.Title)
        If InStr(s, "euro") > 0 Or InStr(s, "€") > 0 Then mEuroHint = True
    Next
End Sub

' Spalten aller Mitglieder nach Namen vereinheitlichen
Private Sub BuildUnifiedCols()
    Dim mi As Long, t As cLiTab, c As Long, nm As String, idx As New cMap, u As Variant, m As cMap
    RecValReset
    Set mUCols = New Collection
    For mi = 1 To mMembers.Count
        Set t = mMembers(mi)
        For c = t.C1 To t.C2
            nm = CStr(t.ColName.GetVal(CStr(c), ""))
            If idx.Exists(nm) Then
                u = mUCols(CLng(idx.GetVal(nm)))
                Set m = u(1)
                m.SetVal CStr(mi), c
            Else
                Set m = New cMap
                m.SetVal CStr(mi), c
                mUCols.Add Array(nm, m, CStr(t.ColTop.GetVal(CStr(c), "")), "S:" & LCase$(nm), CStr(t.ColHelp.GetVal(CStr(c), "")))
                idx.SetVal nm, mUCols.Count
            End If
        Next
    Next
End Sub

' Datensaetze (Datenzeilen aller Mitglieder) mit Gruppenueberschrift
Private Sub BuildRecords()
    Dim mi As Long, t As cLiTab, r As Long, grp As String, cls As String, sumAsData As Boolean
    Set mRecs = New Collection
    If AnyClass("summe") Then sumAsData = (Ans("summe", 0) = 1)
    For mi = 1 To mMembers.Count
        Set t = mMembers(mi)
        grp = ""
        For r = t.FirstDataRow() To t.R2
            cls = CStr(t.RowClass.GetVal(CStr(r), ""))
            If cls = "gruppe" Then
                grp = T_(liSheets(t.SheetIdx).Cell(r, t.C1))
            ElseIf cls = "daten" Or (cls = "summe" And sumAsData) Then
                mRecs.Add Array(mi, r, grp)
            End If
        Next
    Next
End Sub

Private Function AnyClass(ByVal cls As String) As Boolean
    Dim t As cLiTab
    For Each t In mMembers
        If t.CountClass(cls) > 0 Then AnyClass = True: Exit Function
    Next
End Function

Private Function CountAll(ByVal cls As String) As Long
    Dim t As cLiTab
    For Each t In mMembers: CountAll = CountAll + t.CountClass(cls): Next
End Function

Private Function FirstClassText(ByVal cls As String) As String
    Dim t As cLiTab, r As Long
    For Each t In mMembers
        For r = t.R1 To t.R2
            If t.RowClass.GetVal(CStr(r), "") = cls Then FirstClassText = Truncate(FirstText(liSheets(t.SheetIdx), r, t.C1, t.C2), 50): Exit Function
        Next
    Next
End Function

' Rohwert eines Datensatzes in einer vereinheitlichten Spalte (verbundene Zellen nach unten aufgefuellt).
' Spalte und Blatt des Mitglieds werden fuer die zuletzt benutzte Spalte gemerkt (Schleifen ueber alle
' Datensaetze fragen dieselbe Spalte ab; das spart zwei Nachschlagevorgaenge je Zelle).
Private Function RecVal(ByRef rec As Variant, ByRef u As Variant) As Variant
    Dim m As cMap, c As Long, sh As cLiSheet, r As Long, o As String, a As Variant, mi As Long
    Set m = u(1)
    mi = CLng(rec(0))
    If Not (m Is mRvMap And mi = mRvMem) Or mRvSh Is Nothing Then
        Set mRvMap = m: mRvMem = mi
        Set mRvSh = liSheets(mMembers(mi).SheetIdx)
        mRvCol = 0
        If m.Exists(CStr(mi)) Then mRvCol = CLng(m.GetVal(CStr(mi)))
    End If
    If mRvCol = 0 Then Exit Function
    c = mRvCol: Set sh = mRvSh
    r = CLng(rec(1))
    RecVal = sh.Cell(r, c)
    If IsBlank(RecVal) Then
        o = sh.MergeOrigin(r, c)
        If Len(o) > 0 Then
            a = Split(o, "|")
            If CLng(a(1)) = c And CLng(a(0)) < r Then RecVal = sh.Cell(CLng(a(0)), c)
        End If
    End If
End Function

Private Function UHidden(ByVal u As Variant) As Boolean
    Dim m As cMap, k As Variant
    Set m = u(1)
    For Each k In m.Keys
        If liSheets(mMembers(CLng(k)).SheetIdx).HiddenCols.Exists(CStr(m.GetVal(CStr(k)))) Then UHidden = True: Exit Function
    Next
End Function

Private Function UFmt(ByVal u As Variant) As String
    Dim m As cMap, k As Variant
    Set m = u(1)
    For Each k In m.Keys
        UFmt = UFmt & CStr(liSheets(mMembers(CLng(k)).SheetIdx).ColFmt.GetVal(CStr(m.GetVal(CStr(k))), "")) & "|"
    Next
End Function

' 0 = keine Formeln, 1 = Formeln aus derselben Zeile, 2 = Formeln mit Bezug auf andere Blaetter
Private Function FormulaState(ByVal u As Variant) As Long
    Dim rec As Variant, m As cMap, c As Long, sh As cLiSheet, n As Long, nForm As Long, nOther As Long, k As Variant, anyF As Boolean
    Set m = u(1)
    ' ohne Formeln in dieser Spalte (der Normalfall) ist nichts zu zaehlen
    For Each k In m.Keys
        If liSheets(mMembers(CLng(k)).SheetIdx).HasFormulaCol(CLng(m.GetVal(CStr(k)))) Then anyF = True: Exit For
    Next
    If Not anyF Then Exit Function
    For Each rec In mRecs
        If m.Exists(CStr(rec(0))) Then
            c = CLng(m.GetVal(CStr(rec(0))))
            Set sh = liSheets(mMembers(CLng(rec(0))).SheetIdx)
            If Not IsBlank(sh.Cell(CLng(rec(1)), c)) Or sh.IsFormula(CLng(rec(1)), c) Then
                n = n + 1
                If sh.IsFormula(CLng(rec(1)), c) Then
                    nForm = nForm + 1
                    If InStr(sh.FormulaText(CLng(rec(1)), c), "!") > 0 Then nOther = nOther + 1
                End If
            End If
        End If
    Next
    If n > 0 And nForm >= n * 0.8 Then FormulaState = IIf(nOther * 2 > nForm, 2, 1)
End Function

' Spalten, die nie zusammen auf einem Blatt vorkommen (z. B. Menge/Anzahl), zusammenlegen
Private Sub QuestionMergeCols()
    Dim i As Long, j As Long, a As Variant, b As Variant, ma As cMap, mb As cMap, k As Variant, overlap As Boolean, id As String
    Dim keep As New Collection, gone As New cMap
    If mMembers.Count < 2 Then Exit Sub
    For i = 1 To mUCols.Count
        If Not gone.Exists(CStr(i)) Then
            a = mUCols(i)
            Set ma = a(1)
            If ma.Count < mMembers.Count Then
                For j = i + 1 To mUCols.Count
                    If Not gone.Exists(CStr(j)) Then
                        b = mUCols(j)
                        Set mb = b(1)
                        If ma.Count + mb.Count = mMembers.Count Then
                            overlap = False
                            For Each k In mb.Keys
                                If ma.Exists(CStr(k)) Then overlap = True: Exit For
                            Next
                            If Not overlap Then
                                If ProfileU(a).GetVal("typ") = ProfileU(b).GetVal("typ") Then
                                    id = "merge:" & a(3) & ":" & b(3)
                                    AddQ id, CStr(a(0)), "Die Spalte „" & b(0) & "“ gibt es nur auf " & MemberSheetNames(mb) & ", die Spalte „" & a(0) & _
                                         "“ nur auf den anderen Blättern. Meinen beide dasselbe?", Array("Ja, zu einem Feld „" & a(0) & "“ zusammenlegen", "Nein, getrennt lassen"), 0
                                    If Ans(id, 0) = 0 Then
                                        For Each k In mb.Keys: ma.SetVal CStr(k), mb.GetVal(CStr(k)): Next
                                        gone.SetVal CStr(j), True
                                        RecValReset
                                    End If
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next
    If gone.Count = 0 Then Exit Sub
    For i = 1 To mUCols.Count
        If Not gone.Exists(CStr(i)) Then keep.Add mUCols(i)
    Next
    Set mUCols = keep
End Sub

Private Function MemberSheetNames(ByVal m As cMap) As String
    Dim k As Variant, s As String
    For Each k In m.Keys
        s = s & IIf(Len(s) > 0, ", ", "") & "„" & liSheets(mMembers(CLng(k)).SheetIdx).Name & "“"
    Next
    MemberSheetNames = IIf(m.Count = 1, "dem Blatt ", "den Blättern ") & s
End Function

'------------------------------------------------------------------------------
' Spaltenprofil: Typvorschlag
'------------------------------------------------------------------------------
Private Function ProfileU(ByVal u As Variant) As cMap
    Dim vals As New Collection, rec As Variant
    For Each rec In mRecs
        vals.Add RecVal(rec, u)
    Next
    Set ProfileU = ProfileValues(vals, CStr(u(0)), UFmt(u), mCsv)
End Function

' Profil einer Werteliste. Schluessel: typ, n, ndist, distinct (cMap Wert -> Anzahl), ungueltig (Collection),
' sep, marks, decimals, beispiele, hinweis
Public Function ProfileValues(ByVal vals As Collection, ByVal header As String, ByVal fmt As String, ByVal csv As Boolean) As cMap
    Dim p As New cMap, v As Variant, s As String, n As Long, nNum As Long, nDate As Long, nTime As Long, nBool As Long, nMark As Long, nNo As Long
    Dim nCur As Long, nEuro As Long, nPct As Long, nLead As Long, nLong As Long, nSep As Long, nMail As Long, nPhone As Long, nUrl As Long
    Dim d As Double, dt As Date, ht As Boolean, dist As New cMap, msg As String, sepSet As String, ex As String, nEx As Long
    Dim bad As New Collection, typ As String, hdr As String, nInt As Long, maxDec As Long, dc As Long, totLen As Double
    Dim maxCnt As Double, k As Variant, consec As Boolean, prevNum As Double, hint As String, isNum As Boolean, avgLen As Double
    Dim nPh As Long, nDit As Long, nEff As Long, sv As String
    hdr = LCase$(header)
    consec = True
    For Each v In vals
        If Not IsBlank(v) Then
            n = n + 1
            s = T_(v)
            totLen = totLen + Len(s)
            dist.AddNum s, 1
            If nEx < 3 Then
                ' Beispiele ohne Doppel in anderer Schreibweise („erledigt“, „Erledigt“): Werte werden ebenso zusammengefuehrt
                If InStr(1, ex & "|", "|" & s & "|", vbTextCompare) = 0 Then ex = ex & "|" & s: nEx = nEx + 1
            End If
            isNum = False
            Select Case VarType(v)
                Case vbDate
                    nDate = nDate + 1
                    If CDbl(v) <> Int(CDbl(v)) Then nTime = nTime + 1
                Case vbBoolean
                    nBool = nBool + 1
                Case vbCurrency, vbDouble, vbSingle, vbLong, vbInteger, vbByte, vbDecimal
                    nNum = nNum + 1
                    If VarType(v) = vbCurrency Then nCur = nCur + 1
                    d = CDbl(v): isNum = True
                    If d = 1 Then nMark = nMark + 1
                    If d = 0 Then nNo = nNo + 1
                Case Else
                    If IsPlaceholder(s) Then nPh = nPh + 1
                    If IsDittoText(s) Then nDit = nDit + 1
                    If IsMarkText(s) Then nMark = nMark + 1
                    If IsNoText(s) Then nNo = nNo + 1
                    If HasLeadingZero(s) Then
                        nLead = nLead + 1
                    ElseIf TextDate(s, dt, ht) Then
                        nDate = nDate + 1
                        If ht Then nTime = nTime + 1
                    ElseIf TextNumber(s, d, csv) Then
                        nNum = nNum + 1: isNum = True
                        If InStr(s, "€") > 0 Then nEuro = nEuro + 1
                        If InStr(s, "%") > 0 Then nPct = nPct + 1
                    Else
                        If InStr(s, "@") > 0 Then
                            If mValid.KindFormatOk(FT_EMAIL, s, msg) Then nMail = nMail + 1
                        End If
                        If s Like "*#*" Then
                            If mValid.KindFormatOk(FT_TELEFON, s, msg) Then nPhone = nPhone + 1
                        End If
                        If LCase$(Left$(s, 7)) = "http://" Or LCase$(Left$(s, 8)) = "https://" Or LCase$(Left$(s, 4)) = "www." Then nUrl = nUrl + 1
                        If Len(s) > 120 Or InStr(s, vbLf) > 0 Then nLong = nLong + 1
                        If HasSeparator(s) Then nSep = nSep + 1: sepSet = sepSet & SepsIn(s)
                    End If
            End Select
            If isNum Then
                If d = Fix(d) Then nInt = nInt + 1
                dc = Decs(d)
                If dc > maxDec Then maxDec = dc
                If consec Then
                    If n = 1 Then consec = (d = 1) Else consec = (d = prevNum + 1)
                End If
                prevNum = d
            Else
                consec = False
            End If
        End If
    Next
    For Each k In dist.Keys
        If CDbl(dist.GetVal(CStr(k))) > maxCnt Then maxCnt = CDbl(dist.GetVal(CStr(k)))
    Next
    If n > 0 Then avgLen = totLen / n
    p.SetVal "n", n
    p.SetVal "ndist", dist.Count
    p.SetVal "distinct", dist
    p.SetVal "beispiele", Mid$(ex, 2)
    p.SetVal "decimals", IIf(maxDec > 4, 4, maxDec)
    p.SetVal "ndito", nDit
    ' Platzhalter ("-", "k. A.") und "dito" zaehlen fuer Zahl und Datum nicht mit
    nEff = n - nPh - nDit
    If nEff < 1 Then nEff = n
    typ = FT_TEXT
    If n = 0 Then
        typ = FT_TEXT
    ElseIf (nMark + nNo + nBool) = n And dist.Count <= 8 And (nMark > 0 Or nBool > 0) And nNum < n Then
        typ = FT_JANEIN
        p.SetVal "marks", (nNo = 0 And nBool = 0)
        If nNo = 0 And nBool = 0 Then hint = "leer = Nein"
    ElseIf nDate > 0 And nDate >= nEff * 0.9 Then
        typ = IIf(nTime > 0, FT_DATUMZEIT, FT_DATUM)
    ElseIf nLead > 0 And nLead + nInt >= nEff * 0.9 Then
        typ = FT_TEXT: hint = "führende Nullen bleiben erhalten"
    ElseIf IdLikeHeader(hdr) And nNum + nLead >= nEff * 0.9 And nInt + nLead >= nEff * 0.9 And nEuro = 0 And Not (consec And n >= 3) Then
        typ = FT_TEXT: hint = "Nummer als Text"
    ElseIf nNum > 0 And nNum >= nEff * 0.9 Then
        typ = NumKind(nCur, nEuro, nPct, fmt, hdr)
        If consec And n >= 3 Then hint = "laufende Nummer"
        If nPh > 0 Then hint = Trim$(hint & IIf(Len(hint) > 0, ", ", "") & "„-“ und ähnliche Zeichen bleiben leer")
    ElseIf nDate >= nEff * 0.6 And nDate >= 2 Then
        typ = IIf(nTime > 0, FT_DATUMZEIT, FT_DATUM)
    ElseIf nNum >= nEff * 0.6 And nNum >= 2 And nLead = 0 Then
        typ = NumKind(nCur, nEuro, nPct, fmt, hdr)
    ElseIf nMail = n Then
        typ = FT_EMAIL
    ElseIf nUrl = n Then
        typ = FT_URL
    ElseIf nPhone = n And PhoneHeader(hdr) Then
        typ = FT_TELEFON
    ElseIf nLong > 0 Or LongTextHeader(hdr) Or avgLen > 60 Then
        typ = FT_LTEXT
    ElseIf nSep > 0 And Not NameHeader(hdr) And MultiPlausible(dist, n, nSep) Then
        typ = FT_MEHRFACH
        p.SetVal "sep", UniqueChars(sepSet)
    ElseIf IsChoiceLike(dist.Count, n, maxCnt, avgLen) And Not IdLikeHeader(hdr) Then
        typ = FT_AUSWAHL
    Else
        typ = FT_TEXT
    End If
    ' nicht passende Werte
    If typ = FT_DATUM Or typ = FT_DATUMZEIT Or typ = FT_ZAHL Or typ = FT_GELD Or typ = FT_PROZENT Then
        For Each v In vals
            If Not IsBlank(v) Then
                sv = T_(v)
                If Not IsPlaceholder(sv) And Not IsDittoText(sv) Then
                    If Not FitsType(v, typ, csv) Then
                        If bad.Count < 200 Then bad.Add sv
                    End If
                End If
            End If
        Next
    End If
    p.SetVal "typ", typ
    p.SetVal "hinweis", hint
    p.SetVal "ungueltig", bad
    Set ProfileValues = p
End Function

Private Function NumKind(ByVal nCur As Long, ByVal nEuro As Long, ByVal nPct As Long, ByVal fmt As String, ByVal hdr As String) As String
    If nCur > 0 Or nEuro > 0 Or InStr(fmt, "€") > 0 Or InStr(LCase$(fmt), "eur") > 0 Then NumKind = FT_GELD: Exit Function
    If nPct > 0 Or InStr(fmt, "%") > 0 Then NumKind = FT_PROZENT: Exit Function
    If (MoneyHeader(hdr) Or mEuroHint) And Not CountLikeHeader(hdr) Then NumKind = FT_GELD: Exit Function
    NumKind = FT_ZAHL
End Function

Private Function IsChoiceLike(ByVal ndist As Long, ByVal n As Long, ByVal maxCnt As Double, ByVal avgLen As Double) As Boolean
    If n < 4 Or ndist < 1 Or avgLen > 40 Then Exit Function
    If ndist <= 15 And n >= ndist * 2 Then IsChoiceLike = True: Exit Function
    If ndist <= 25 And ndist <= n * 0.6 And maxCnt >= 3 Then IsChoiceLike = True
End Function

Private Function Decs(ByVal d As Double) As Long
    Dim s As String, p As Long
    s = Trim$(Str$(Round(d, 6)))
    If InStr(s, "E") > 0 Then Decs = 6: Exit Function
    p = InStr(s, ".")
    If p > 0 Then Decs = Len(s) - p
End Function

Private Function HasSeparator(ByVal s As String) As Boolean
    HasSeparator = (InStr(s, ", ") > 0 Or InStr(s, ";") > 0 Or InStr(s, " / ") > 0 Or (InStr(s, ",") > 0 And Not s Like "*#,#*"))
End Function

Private Function SepsIn(ByVal s As String) As String
    If InStr(s, ",") > 0 Then SepsIn = SepsIn & ","
    If InStr(s, ";") > 0 Then SepsIn = SepsIn & ";"
    If InStr(s, " / ") > 0 Then SepsIn = SepsIn & "/"
End Function

Private Function UniqueChars(ByVal s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr(UniqueChars, ch) = 0 Then UniqueChars = UniqueChars & ch
    Next
End Function

' Teile kommen auch einzeln vor ("Amt 40, Amt 61" neben "Amt 40")
Private Function PartsKnown(ByVal dist As cMap) As Boolean
    Dim k As Variant, parts As Variant, i As Long, nKnown As Long, nParts As Long
    For Each k In dist.Keys
        If HasSeparator(CStr(k)) Then
            parts = SplitMulti(CStr(k), ",;/")
            For i = 0 To ArrLen(parts) - 1
                nParts = nParts + 1
                If dist.Exists(CStr(parts(i))) Then nKnown = nKnown + 1
            Next
        End If
    Next
    PartsKnown = (nParts > 0 And nKnown * 2 >= nParts)
End Function

' Mehrere kurze Werte je Zelle, die sich wiederholen (Themenlisten), nicht Saetze oder Namen
Private Function MultiPlausible(ByVal dist As cMap, ByVal n As Long, ByVal nSep As Long) As Boolean
    Dim k As Variant, parts As Variant, i As Long, nParts As Double, nShort As Double, lenSum As Double, cnt As Double, words As Long
    Dim seen As New cMap, allParts As New cMap, occ As Double
    ' Auswahlwerte muessen sich wiederholen: viele verschiedene Teile ("Adenau, Stadt") oder Teile, die kaum
    ' mehrfach vorkommen ("Bürgerhaus, gr. Saal") sprechen fuer Text, nicht fuer eine Auswahl
    For Each k In dist.Keys
        parts = SplitMulti(CStr(k), ",;/")
        For i = 0 To ArrLen(parts) - 1
            allParts.SetVal LCase$(CStr(parts(i))), True
            occ = occ + CDbl(dist.GetVal(CStr(k)))
        Next
    Next
    If allParts.Count > 40 Then Exit Function
    If dist.Count >= 6 And allParts.Count > 0 Then
        If occ / allParts.Count < 1.25 Then Exit Function
    End If
    ' ein oder zwei Zellwerte ("DE300 - Berlin, Stadt") sind keine Auswahl aus mehreren Werten
    If dist.Count < 3 Then Exit Function
    If PartsKnown(dist) Then MultiPlausible = True: Exit Function
    If nSep < n * 0.3 Then Exit Function
    For Each k In dist.Keys
        cnt = CDbl(dist.GetVal(CStr(k)))
        parts = SplitMulti(CStr(k), ",;/")
        For i = 0 To ArrLen(parts) - 1
            nParts = nParts + cnt
            lenSum = lenSum + Len(parts(i)) * cnt
            words = UBound(Split(Trim$(parts(i)), " ")) + 1
            If words <= 3 Then nShort = nShort + cnt
            seen.SetVal CStr(parts(i)), True
        Next
    Next
    If nParts = 0 Then Exit Function
    MultiPlausible = (lenSum / nParts <= 25 And nShort >= nParts * 0.8 And seen.Count <= nParts * 0.85)
End Function

Public Function SplitMulti(ByVal s As String, ByVal seps As String) As Variant
    Dim i As Long, a As Variant, out() As String, n As Long, t As String
    ' Zeilenumbrüche in einer Zelle (Alt+Eingabe) trennen immer
    s = Replace(Replace(Replace(s, vbCrLf, ";"), vbLf, ";"), vbCr, ";")
    If InStr(seps, "/") > 0 Then s = Replace(s, " / ", ";")
    If InStr(seps, ",") > 0 Then s = Replace(s, ",", ";")
    a = Split(s, ";")
    If UBound(a) < 0 Then SplitMulti = Array(): Exit Function
    ReDim out(0 To UBound(a))
    For i = 0 To UBound(a)
        t = Trim$(a(i))
        If Len(t) > 0 Then out(n) = t: n = n + 1
    Next
    If n = 0 Then SplitMulti = Array(): Exit Function
    ReDim Preserve out(0 To n - 1)
    SplitMulti = out
End Function

Private Function FitsType(ByVal v As Variant, ByVal typ As String, ByVal csv As Boolean) As Boolean
    Dim outV As Variant, why As String
    FitsType = ConvertTyped(v, typ, csv, outV, why)
End Function

'------------------------------------------------------------------------------
' Zeitraeume als Spalten (Kreuztabelle)
'------------------------------------------------------------------------------
Private Function MonthIndex(ByVal s As String) As Long
    Dim t As String, names As Variant, i As Long
    t = LCase$(Trim$(Replace(s, ".", "")))
    If t = "mrz" Or t = "märz" Or t = "maerz" Or t = "mär" Then MonthIndex = 3: Exit Function
    names = Array("januar", "februar", "märz", "april", "mai", "juni", "juli", "august", "september", "oktober", "november", "dezember")
    For i = 0 To 11
        If t = names(i) Then MonthIndex = i + 1: Exit Function
        If Len(t) >= 3 And Left$(names(i), Len(t)) = t Then MonthIndex = i + 1: Exit Function
    Next
End Function

' Zeitraum aus einem Spaltenkopf: Art "monat" (1..12), "quartal" (1..4) oder "jahr"; Jahr, falls angegeben
Private Function PeriodOf(ByVal s As String, ByRef kind As String, ByRef idx As Long, ByRef yr As Long) As Boolean
    Dim t As String, p As Long, tail As String
    kind = "": idx = 0: yr = 0
    t = LCase$(Trim$(s))
    If Len(t) = 0 Then Exit Function
    If Len(t) = 4 And IsAllDigits(t) Then
        If CLng(t) >= 1990 And CLng(t) <= 2100 Then
            kind = "jahr": idx = CLng(t): yr = idx: PeriodOf = True
        End If
        Exit Function
    End If
    p = InStrRev(t, " ")
    If p = 0 Then p = InStrRev(t, "/")
    If p > 0 Then
        tail = Mid$(t, p + 1)
        If IsAllDigits(tail) And (Len(tail) = 2 Or Len(tail) = 4) Then
            yr = CLng(tail)
            If yr < 100 Then yr = 2000 + yr
            t = Trim$(Left$(t, p - 1))
        End If
    End If
    If t Like "q#" Then
        kind = "quartal": idx = CLng(Mid$(t, 2, 1))
    ElseIf t Like "#. quartal" Or t Like "#.quartal" Then
        kind = "quartal": idx = CLng(Left$(t, 1))
    ElseIf t Like "quartal #" Then
        kind = "quartal": idx = CLng(Right$(t, 1))
    End If
    If kind = "quartal" Then
        PeriodOf = (idx >= 1 And idx <= 4)
        If Not PeriodOf Then kind = "": idx = 0
        Exit Function
    End If
    idx = MonthIndex(t)
    If idx > 0 Then kind = "monat": PeriodOf = True
End Function

Private Function YearFromText(ByVal s As String) As Long
    Dim i As Long, y As Long
    For i = 1 To Len(s) - 3
        If Mid$(s, i, 4) Like "####" Then
            y = CLng(Mid$(s, i, 4))
            If y >= 1990 And y <= 2100 Then
                If i + 4 > Len(s) Then
                    YearFromText = y: Exit Function
                ElseIf Not Mid$(s, i + 4, 1) Like "#" Then
                    If i = 1 Then
                        YearFromText = y: Exit Function
                    ElseIf Not Mid$(s, i - 1, 1) Like "#" Then
                        YearFromText = y: Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

Private Function PeriodWord(ByVal kind As String, ByVal plural As Boolean) As String
    Select Case kind
        Case "monat": PeriodWord = IIf(plural, "Monate", "Monat")
        Case "quartal": PeriodWord = IIf(plural, "Quartale", "Quartal")
        Case Else: PeriodWord = IIf(plural, "Jahre", "Jahr")
    End Select
End Function

Private Sub PlanPivot()
    Dim i As Long, run As New Collection, best As Collection, bestKind As String, kind As String, idx As Long, yr As Long
    Dim prevKind As String, prevIdx As Long, minLen As Long, firstN As String, lastN As String
    For i = 1 To mUCols.Count + 1
        kind = ""
        If i <= mUCols.Count Then
            If Not PeriodOf(CStr(mUCols(i)(0)), kind, idx, yr) Then kind = ""
        End If
        If Len(kind) > 0 And kind = prevKind And idx = prevIdx + 1 Then
            run.Add i
        Else
            If run.Count > 0 Then
                minLen = IIf(prevKind = "monat", 4, 3)
                If run.Count >= minLen Then
                    If best Is Nothing Then
                        Set best = run: bestKind = prevKind
                    ElseIf run.Count > best.Count Then
                        Set best = run: bestKind = prevKind
                    End If
                End If
            End If
            Set run = New Collection
            If Len(kind) > 0 Then run.Add i
        End If
        prevKind = kind: prevIdx = idx
    Next
    If best Is Nothing Then Exit Sub
    firstN = CStr(mUCols(best(1))(0)): lastN = CStr(mUCols(best(best.Count))(0))
    AddQ "pivot", "Zeitraumspalten", "Die Spalten „" & firstN & "“ bis „" & lastN & "“ sind " & PeriodWord(bestKind, True) & _
         " (Kreuztabelle). Wie sollen sie übernommen werden?", _
         Array("Je Zeile und " & PeriodWord(bestKind, False) & " einen Eintrag anlegen (empfohlen, gut auszuwerten)", _
               "Je " & PeriodWord(bestKind, False) & " ein eigenes Feld"), 0
    If Ans("pivot", 0) = 0 Then
        Set mPivot = best
        mPivotKind = bestKind
        PeriodOf firstN, kind, idx, yr
        If yr = 0 Then yr = YearFromText(liTab.Title & " " & mMembers(1).Title & " " & liSheets(mMembers(1).SheetIdx).Name)
        mPivotYear = yr
    End If
End Sub

Private Function IsPivotCol(ByVal i As Long) As Boolean
    Dim k As Variant
    If mPivot Is Nothing Then Exit Function
    For Each k In mPivot
        If CLng(k) = i Then IsPivotCol = True: Exit Function
    Next
End Function

' Name fuer die Werte der Zeitraumspalten: aus dem Titel ("Belegte Plätze je Einrichtung 2025" -> "Belegte Plätze")
Private Function PivotValueName() As String
    Dim s As String, w As Variant, p As Long
    s = Trim$(mMembers(1).Title)
    For Each w In Array(" je ", " pro ", " nach ", " im ", " in ", " (", " für ", ":")
        p = InStr(1, s, CStr(w), vbTextCompare)
        If p > 1 Then s = Trim$(Left$(s, p - 1))
    Next
    s = CleanName(s)
    If Len(s) = 0 Or Len(s) > 40 Or YearFromText(s) > 0 Then s = "Wert"
    PivotValueName = s
End Function

'------------------------------------------------------------------------------
' Felder planen
'------------------------------------------------------------------------------
Private Sub PlanColumns()
    Dim i As Long, u As Variant, p As cMap, col As cLiCol, inCross As New cMap, grpName As String, rec As Variant, dist As New cMap
    Dim nFormCols As Long, mostlyFormula As Boolean
    ' Gruppenzeilen
    If AnyClass("gruppe") Then
        grpName = GroupFieldName()
        AddQ "gruppe", "Gruppenzeilen", "Einige Zeilen enthalten nur einen Text in der ersten Spalte, z. B. „" & FirstClassText("gruppe") & _
             "“. Das sind vermutlich Überschriften für die Zeilen darunter.", _
             Array("Als Feld „" & grpName & "“ an jeden Eintrag darunter übernehmen", "Weglassen"), 0
        If Ans("gruppe", 0) = 0 Then
            Set col = New cLiCol
            col.Key = "G:gruppe": col.Mode = "gruppe": col.Name = grpName
            For Each rec In mRecs
                If Len(rec(2)) > 0 Then dist.AddNum CStr(rec(2)), 1
            Next
            col.FType = IIf(dist.Count <= 40, FT_AUSWAHL, FT_TEXT)
            col.Hint = "aus den Gruppenzeilen"
            col.Sample = GroupSamples()
            EarlyType col
            If col.IsOptionType Then FillOptions col
            liCols.Add col
        End If
    End If
    ' Summenzeilen
    If AnyClass("summe") Then
        AddQ "summe", "Summenzeilen", "Die Tabelle enthält " & Plural(CountAll("summe"), "Summen- oder Gesamtzeile", "Summen- oder Gesamtzeilen") & _
             ", z. B. „" & FirstClassText("summe") & "“.", Array("Weglassen (empfohlen)", "Wie normale Zeilen übernehmen"), 0
    End If
    ' Blattname bei zusammengefuehrten Blaettern
    If mMembers.Count > 1 Then
        Set col = New cLiCol
        col.Key = "B:blatt": col.Mode = "blatt": col.Name = "Blatt": col.FType = FT_AUSWAHL
        col.Hint = "Name des Tabellenblatts"
        For i = 1 To mMembers.Count
            If Not InColl(col.Opts, LiOptName(liSheets(mMembers(i).SheetIdx).Name)) Then col.Opts.Add LiOptName(liSheets(mMembers(i).SheetIdx).Name)
        Next
        col.Sample = CStr(col.Opts(1))
        If col.Opts.Count > 1 Then col.Sample = Truncate(col.Sample & ", " & col.Opts(2), 60)
        liCols.Add col
    End If
    ' Kreuzspalten
    PlanCrossRuns inCross
    ' Formelspalten
    For i = 1 To mUCols.Count
        If Not inCross.Exists(CStr(i)) And Not IsPivotCol(i) Then
            If FormulaState(mUCols(i)) > 0 Then nFormCols = nFormCols + 1
        End If
    Next
    mostlyFormula = (nFormCols * 2 > mUCols.Count)
    ' uebrige Spalten
    For i = 1 To mUCols.Count
        If Not inCross.Exists(CStr(i)) And Not IsPivotCol(i) Then
            u = mUCols(i)
            ProfAdd "(übrige)"
            Set p = ProfileU(u)
            ProfAdd "Werte prüfen (ProfileU)"
            Set col = New cLiCol
            col.Key = CStr(u(3)): col.Mode = "wert": col.Name = CStr(u(0))
            If UBound(u) >= 4 Then col.Help = CStr(u(4))
            col.Src.Add i
            col.FType = CStr(p.GetVal("typ"))
            col.Hint = CStr(p.GetVal("hinweis", ""))
            col.Sample = Truncate(Replace(CStr(p.GetVal("beispiele", "")), "|", ", "), 60)
            If col.FType = FT_JANEIN Then col.BlankFalse = CBool(p.GetVal("marks", False))
            col.Decimals = CLng(p.GetVal("decimals", 0))
            col.AllUnique = (CLng(p.GetVal("ndist", 0)) = CLng(p.GetVal("n", 0)) And CLng(p.GetVal("n", 0)) > 1)
            If CLng(p.GetVal("n", 0)) > 0 Then col.DistRatio = CLng(p.GetVal("ndist", 0)) / CLng(p.GetVal("n", 0))
            If mRecs.Count > 0 Then col.FillRatio = CLng(p.GetVal("n", 0)) / mRecs.Count
            If CLng(p.GetVal("n", 0)) = 0 Then col.Include = False: col.Hint = "Spalte ist leer"
            If Not mPivot Is Nothing And col.FType = FT_TEXT Then
                If CLng(p.GetVal("ndist", 99)) <= 40 Then col.FType = FT_AUSWAHL
            End If
            PlanColumnQuestions col, u, p, mostlyFormula
            ProfAdd "Rückfragen je Spalte"
            EarlyType col
            If col.IsOptionType Then FillOptions col
            ProfAdd "Optionen sammeln"
            liCols.Add col
        End If
    Next
    SortColsBySource
    If Not mPivot Is Nothing Then PlanPivotCols
    MarkLabelCandidates
End Sub

' Felder in der Reihenfolge der Quellspalten (zusammengefasste Kreuzspalten stuenden sonst vor allen anderen);
' Gruppen- und Blattfeld bleiben vorn. Sortiert werden alle Felder mit Quellspalte (wert, liste, kreuz).
Private Sub SortColsBySource()
    Dim fixed As New Collection, rest As New Collection, col As cLiCol, i As Long, placed As Boolean
    For Each col In liCols
        If col.Mode <> "gruppe" And col.Mode <> "blatt" And col.Src.Count > 0 Then
            placed = False
            For i = 1 To rest.Count
                If col.FirstSrc < rest(i).FirstSrc Then rest.Add col, Before:=i: placed = True: Exit For
            Next
            If Not placed Then rest.Add col
        Else
            fixed.Add col
        End If
    Next
    Set liCols = New Collection
    For Each col In fixed
        liCols.Add col
    Next
    For Each col In rest
        liCols.Add col
    Next
End Sub

' Vom Nutzer gewaehlten Typ schon vor dem Sammeln der Optionen anwenden (keine Rueckfragen zu einem Typ,
' den das Feld gar nicht mehr hat, z. B. Abkuerzungen bei einem Textfeld)
Private Sub EarlyType(ByVal col As cLiCol)
    Dim ov As String
    ov = OvType(col.Key)
    If Len(ov) = 0 Or ov = col.FType Then Exit Sub
    col.FType = ov
    If ov = FT_MEHRFACH And col.Mode = "wert" Then
        col.Mode = "liste"
        If Len(col.Sep) = 0 Then col.Sep = ",;/"
    ElseIf ov <> FT_MEHRFACH And col.Mode = "liste" Then
        col.Mode = "wert"
    End If
End Sub

Private Sub PlanColumnQuestions(ByVal col As cLiCol, ByVal u As Variant, ByVal p As cMap, ByVal mostlyFormula As Boolean)
    Dim bad As Collection, id As String, ex As String, i As Long, fs As Long, dflt As Long
    ' ausgeblendete Spalte
    If UHidden(u) Then
        col.IsHidden = True
        id = "hidden:" & col.Key
        AddQ id, col.Name, "Die Spalte „" & col.Name & "“ ist in der Quelle ausgeblendet.", Array("Nicht übernehmen", "Übernehmen"), 0
        If Ans(id, 0) = 0 Then col.Include = False
    End If
    ' Formelspalte
    fs = FormulaState(u)
    If fs > 0 Then
        col.IsFormula = True
        id = "formel:" & col.Key
        dflt = IIf(fs = 2 Or mostlyFormula, 1, 0)
        AddQ id, col.Name, "Die Spalte „" & col.Name & "“ wird in der Quelle mit Formeln berechnet.", _
             Array("Weglassen (bei Bedarf in MidPoint als Formelfeld neu anlegen)", "Die berechneten Werte übernehmen"), dflt
        If Ans(id, dflt) = 0 Then col.Include = False
    End If
    ' "dito" / "siehe oben"
    If CLng(p.GetVal("ndito", 0)) > 0 Then
        id = "dito:" & col.Key
        AddQ id, col.Name, "In der Spalte „" & col.Name & "“ steht " & Plural(CLng(p.GetVal("ndito")), "Mal", "Mal") & " „dito“ oder „siehe oben“. " & _
             "Gemeint ist der Wert der Zeile darüber.", Array("Wert der Zeile darüber übernehmen", "So übernehmen, wie es dasteht"), 0
        col.Ditto = (Ans(id, 0) = 0)
    End If
    ' Typ vom Nutzer festgelegt: keine Rueckfragen zu Typ und Trennzeichen
    If Len(OvType(col.Key)) > 0 Then
        If col.FType = FT_MEHRFACH Then
            col.Sep = CStr(p.GetVal("sep", ",;/"))
            If Len(col.Sep) = 0 Then col.Sep = ",;/"
            col.Mode = "liste"
        End If
        Exit Sub
    End If
    ' mehrere Werte je Zelle
    If col.FType = FT_MEHRFACH Then
        col.Sep = CStr(p.GetVal("sep", ",;/"))
        If Len(col.Sep) = 0 Then col.Sep = ",;/"
        col.Mode = "liste"
        id = "multi:" & col.Key
        AddQ id, col.Name, "In der Spalte „" & col.Name & "“ stehen teils mehrere Werte in einer Zelle, z. B. „" & FirstWithSep(p) & "“.", _
             Array("Als Mehrfachauswahl übernehmen (getrennt bei " & SepWords(col.Sep) & ")", "Als Text übernehmen"), 0
        If Ans(id, 0) = 1 Then col.FType = FT_TEXT: col.Mode = "wert"
    End If
    ' nicht passende Werte
    Set bad = p.GetVal("ungueltig")
    If bad.Count > 0 And (col.FType = FT_DATUM Or col.FType = FT_DATUMZEIT Or col.FType = FT_ZAHL Or col.FType = FT_GELD Or col.FType = FT_PROZENT) Then
        For i = 1 To bad.Count
            If i > 3 Then Exit For
            ex = ex & IIf(Len(ex) > 0, ", ", "") & "„" & Truncate(CStr(bad(i)), 30) & "“"
        Next
        id = "typ:" & col.Key
        AddQ id, col.Name, "Die Spalte „" & col.Name & "“ enthält überwiegend Werte vom Typ " & FieldTypeName(col.FType) & ", aber " & _
             IIf(bad.Count = 1, "1 Wert passt", bad.Count & " Werte passen") & " nicht: " & ex & ".", _
             Array("Als " & FieldTypeName(col.FType) & " übernehmen, diese Werte leer lassen und im Bericht melden", "Alles als Text übernehmen"), 0
        If Ans(id, 0) = 1 Then col.FType = FT_TEXT
        col.Unsure = True
    End If
End Sub

Private Function FirstWithSep(ByVal p As cMap) As String
    Dim k As Variant, dist As cMap
    Set dist = p.GetVal("distinct")
    For Each k In dist.Keys
        If HasSeparator(CStr(k)) Then FirstWithSep = Truncate(CStr(k), 40): Exit Function
    Next
End Function

Private Function SepWords(ByVal seps As String) As String
    Dim s As String
    If InStr(seps, ",") > 0 Then s = "Komma"
    If InStr(seps, ";") > 0 Then s = s & IIf(Len(s) > 0, ", ", "") & "Semikolon"
    If InStr(seps, "/") > 0 Then s = s & IIf(Len(s) > 0, ", ", "") & "Schrägstrich"
    SepWords = s
End Function

' 1 = nur Kreuze, 2 = leer, 0 = sonst
Private Function MarkState(ByVal u As Variant) As Long
    Dim rec As Variant, v As Variant, n As Long, nMark As Long, nOther As Long, maxOther As Double
    If CountLikeHeader(CStr(u(0))) Or MoneyHeader(CStr(u(0))) Then Exit Function
    ' mehr andere Werte als 10 % aller Datensaetze: kann keine Kreuzspalte mehr werden (Abbruch spart Zeit)
    maxOther = mRecs.Count * 0.1
    If maxOther < 1 Then maxOther = 1
    For Each rec In mRecs
        v = RecVal(rec, u)
        If Not IsBlank(v) Then
            n = n + 1
            If IsMarkVal(v) Then nMark = nMark + 1 Else nOther = nOther + 1
            If nOther > maxOther Then Exit Function
        End If
    Next
    If n = 0 Then MarkState = 2: Exit Function
    If nMark >= 1 And nMark > nOther And nOther <= Application.Max(1, n * 0.1) Then MarkState = 1
End Function

' Nebeneinanderliegende Spalten nur mit Kreuzen -> eine Mehrfachauswahl (nach Rueckfrage)
Private Sub PlanCrossRuns(ByVal inCross As cMap)
    Dim i As Long, j As Long, run As Collection, st As Long, nMark As Long, top As String, t2 As String, id As String
    Dim col As cLiCol, k As Variant, nm As String, others As String, dflt As Long
    i = 1
    Do While i <= mUCols.Count
        st = 0
        If Not IsPivotCol(i) Then st = MarkState(mUCols(i))
        If st = 1 Then
            Set run = New Collection
            run.Add i
            top = CStr(mUCols(i)(2))
            nMark = 1
            j = i + 1
            Do While j <= mUCols.Count
                If IsPivotCol(j) Then Exit Do
                t2 = CStr(mUCols(j)(2))
                If Len(t2) > 0 And Len(top) > 0 And StrComp(t2, top, vbTextCompare) <> 0 Then Exit Do
                st = MarkState(mUCols(j))
                If st = 0 Then Exit Do
                run.Add j
                If st = 1 Then nMark = nMark + 1
                j = j + 1
            Loop
            Do While run.Count > 1
                If MarkState(mUCols(run(run.Count))) = 2 Then run.Remove run.Count Else Exit Do
            Loop
            If nMark >= 2 And run.Count >= 2 Then
                nm = top
                If Len(nm) = 0 Then nm = CommonPrefix(run)
                If Len(nm) = 0 Then nm = "Auswahl"
                dflt = IIf(Len(top) > 0 Or run.Count >= 3, 0, 1)
                id = "kreuz:" & mUCols(run(1))(3)
                others = OtherMarkValues(run)
                AddQ id, nm, "Die Spalten „" & mUCols(run(1))(0) & "“ bis „" & mUCols(run(run.Count))(0) & "“ enthalten nur Kreuze." & _
                     IIf(Len(others) > 0, " Andere Einträge (" & others & ") gelten als nicht angekreuzt und werden im Bericht gemeldet.", ""), _
                     Array("Zu einer Mehrfachauswahl „" & nm & "“ zusammenfassen", "Je Spalte ein Ja/Nein-Feld"), dflt
                If Ans(id, dflt) = 0 Then
                    Set col = New cLiCol
                    col.Key = "K:" & mUCols(run(1))(3): col.Mode = "kreuz": col.Name = nm: col.FType = FT_MEHRFACH
                    For Each k In run
                        col.Src.Add CLng(k)
                        col.SrcNames.Add LiOptName(CStr(mUCols(CLng(k))(0)))
                        col.Opts.Add LiOptName(CStr(mUCols(CLng(k))(0)))
                        inCross.SetVal CStr(k), True
                    Next
                    col.Hint = "aus " & run.Count & " Kreuzspalten"
                    col.Sample = Truncate(col.Opts(1) & ", " & col.Opts(2), 60)
                    liCols.Add col
                End If
                i = CLng(run(run.Count)) + 1
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub

' gemeinsamer Anfang der Spaltennamen ("Modul 1", "Modul 2" -> "Modul")
Private Function CommonPrefix(ByVal run As Collection) As String
    Dim s As String, k As Variant, nm As String, i As Long
    s = CStr(mUCols(run(1))(0))
    For Each k In run
        nm = CStr(mUCols(CLng(k))(0))
        i = 0
        Do While i < Len(s) And i < Len(nm)
            If LCase$(Mid$(s, i + 1, 1)) <> LCase$(Mid$(nm, i + 1, 1)) Then Exit Do
            i = i + 1
        Loop
        s = Left$(s, i)
    Next
    s = Trim$(s)
    Do While Len(s) > 0
        If Right$(s, 1) Like "#" Or Right$(s, 1) = "(" Or Right$(s, 1) = "-" Or Right$(s, 1) = ":" Then s = Trim$(Left$(s, Len(s) - 1)) Else Exit Do
    Loop
    If Len(s) < 3 Then s = ""
    CommonPrefix = s
End Function

Private Function OtherMarkValues(ByVal run As Collection) As String
    Dim k As Variant, rec As Variant, v As Variant, seen As New cMap, s As String
    For Each k In run
        For Each rec In mRecs
            v = RecVal(rec, mUCols(CLng(k)))
            If Not IsBlank(v) Then
                If Not IsMarkVal(v) And Not seen.Exists(T_(v)) Then
                    seen.SetVal T_(v), True
                    s = s & IIf(Len(s) > 0, ", ", "") & "„" & Truncate(T_(v), 25) & "“"
                End If
            End If
        Next
    Next
    OtherMarkValues = s
End Function

' Feldname fuer Gruppenzeilen: gemeinsames erstes Wort ("Produkt 36.1 ...", "Produkt 36.2 ...") sonst "Gruppe"
Private Function GroupFieldName() As String
    Dim t As cLiTab, r As Long, w As String, firstW As String, n As Long, same As Boolean
    same = True
    For Each t In mMembers
        For r = t.R1 To t.R2
            If t.RowClass.GetVal(CStr(r), "") = "gruppe" Then
                w = FirstText(liSheets(t.SheetIdx), r, t.C1, t.C2)
                If InStr(w, " ") > 0 Then w = Left$(w, InStr(w, " ") - 1) Else w = ""
                n = n + 1
                If n = 1 Then
                    firstW = w
                ElseIf StrComp(w, firstW, vbTextCompare) <> 0 Then
                    same = False
                End If
            End If
        Next
    Next
    If same And n >= 2 And Len(firstW) >= 3 And Not IsNumeric(firstW) Then GroupFieldName = CleanName(firstW) Else GroupFieldName = "Gruppe"
End Function

Private Function GroupSamples() As String
    Dim rec As Variant, seen As New cMap, s As String
    For Each rec In mRecs
        If Len(rec(2)) > 0 And Not seen.Exists(CStr(rec(2))) Then
            seen.SetVal CStr(rec(2)), True
            s = s & IIf(Len(s) > 0, ", ", "") & CStr(rec(2))
            If seen.Count >= 2 Then Exit For
        End If
    Next
    GroupSamples = Truncate(s, 60)
End Function

' Quellwert eines Felds fuer einen Datensatz (vor der Umwandlung)
Private Function SrcValue(ByVal col As cLiCol, ByVal rec As Variant) As Variant
    Select Case col.Mode
        Case "gruppe": SrcValue = rec(2)
        Case "blatt": SrcValue = liSheets(mMembers(CLng(rec(0))).SheetIdx).Name
        Case "formular": SrcValue = liSheets(liTab.SheetIdx).Cell(col.FirstSrc, col.SrcCol)
        Case "wert", "liste": SrcValue = RecVal(rec, mUCols(col.FirstSrc))
    End Select
End Function

' Optionen (Reihenfolge des ersten Auftretens), Abkuerzungsfragen, Zuordnung Rohwert -> Option
Private Sub FillOptions(ByVal col As cLiCol)
    Dim rec As Variant, v As Variant, s As String, parts As Variant, i As Long, cnt As New cMap, k As Variant
    Dim order As New Collection, cands As Collection, id As String, choices() As String, j As Long, best As Long, target As String, dflt As Long
    Set col.Opts = New Collection
    Set col.ValMap = New cMap
    For Each rec In mRecs
        v = SrcValue(col, rec)
        If Not IsBlank(v) Then
            If col.Mode = "liste" Then parts = SplitMulti(T_(v), col.Sep) Else parts = Array(T_(v))
            For i = 0 To ArrLen(parts) - 1
                s = LiOptName(CStr(parts(i)))
                If col.Ditto And IsDittoText(s) Then s = ""
                If Len(s) > 0 Then
                    If Not cnt.Exists(s) Then order.Add s
                    cnt.AddNum s, 1
                End If
            Next
        End If
    Next
    ' Abkuerzungen ("erl." -> "erledigt", "JHA" -> "Jugendhilfeausschuss")
    For Each k In order
        Set cands = AbbrevCandidates(CStr(k), order)
        If cands.Count > 0 Then
            id = "abk:" & col.Key & ":" & LCase$(CStr(k))
            ReDim choices(0 To cands.Count)
            best = 1
            For j = 1 To cands.Count
                choices(j - 1) = "„" & cands(j) & "“ ist gemeint"
                If CDbl(cnt.GetVal(cands(j))) > CDbl(cnt.GetVal(cands(best))) Then best = j
            Next
            choices(cands.Count) = "„" & k & "“ als eigenen Wert behalten"
            If CDbl(cnt.GetVal(CStr(k))) <= CDbl(cnt.GetVal(cands(best))) Then dflt = best - 1 Else dflt = cands.Count
            AddQ id, col.Name, "Spalte „" & col.Name & "“: „" & k & "“ sieht wie eine Abkürzung aus.", choices, dflt
            j = Ans(id, dflt)
            If j >= 0 And j < cands.Count Then col.ValMap.SetVal CStr(k), cands(j + 1)
        End If
    Next
    ' Zuordnung und Optionsliste
    For Each k In order
        target = CStr(col.ValMap.GetVal(CStr(k), k))
        If liValOvr.Exists(col.Key & "|" & CStr(k)) Then target = CStr(liValOvr.GetVal(col.Key & "|" & CStr(k)))
        col.ValMap.SetVal CStr(k), target
        AddTargetOpts col, target
    Next
End Sub

Private Sub AddTargetOpts(ByVal col As cLiCol, ByVal target As String)
    Dim a As Variant, i As Long, t As String
    If Len(target) = 0 Then Exit Sub
    If col.FType = FT_MEHRFACH Then a = Split(target, ";") Else a = Array(target)
    For i = 0 To UBound(a)
        t = LiOptName(CStr(a(i)))
        If Len(t) > 0 Then
            If Not InColl(col.Opts, t) Then col.Opts.Add t
        End If
    Next
End Sub

Private Function InColl(ByVal c As Collection, ByVal s As String) As Boolean
    Dim k As Variant
    For Each k In c
        If StrComp(CStr(k), s, vbTextCompare) = 0 Then InColl = True: Exit Function
    Next
End Function

' Werte, fuer die s eine Abkuerzung sein kann (gleicher Anfangsbuchstabe, Buchstaben in Reihenfolge enthalten)
Private Function AbbrevCandidates(ByVal s As String, ByVal order As Collection) As Collection
    Dim res As New Collection, k As Variant, core As String, isAbbr As Boolean
    Set AbbrevCandidates = res
    core = Trim$(s)
    isAbbr = (Right$(core, 1) = ".") Or (Len(core) <= 4 And UCase$(core) = core And core Like "*[A-ZÄÖÜ]*")
    If Not isAbbr Then Exit Function
    core = Replace(Replace(core, ".", ""), " ", "")
    If Len(core) < 2 Then Exit Function
    For Each k In order
        If StrComp(CStr(k), s, vbTextCompare) <> 0 And Len(CStr(k)) > Len(core) + 1 And Right$(CStr(k), 1) <> "." Then
            If LCase$(Left$(CStr(k), 1)) = LCase$(Left$(core, 1)) Then
                If IsSubsequence(LCase$(core), LCase$(Replace(CStr(k), " ", ""))) Then res.Add CStr(k)
            End If
        End If
    Next
End Function

Private Function IsSubsequence(ByVal a As String, ByVal b As String) As Boolean
    Dim i As Long, j As Long
    j = 1
    For i = 1 To Len(a)
        Do While j <= Len(b)
            If Mid$(b, j, 1) = Mid$(a, i, 1) Then Exit Do
            j = j + 1
        Loop
        If j > Len(b) Then Exit Function
        j = j + 1
    Next
    IsSubsequence = True
End Function

Private Sub PlanPivotCols()
    Dim col As cLiCol, k As Variant, vals As New Collection, rec As Variant, v As Variant, p As cMap, bad As Collection
    Set col = New cLiCol
    col.Key = "P:zeit": col.Mode = "monat": col.Name = PeriodWord(mPivotKind, False)
    col.Year_ = mPivotYear
    Select Case mPivotKind
        Case "monat"
            If mPivotYear > 0 Then
                col.FType = FT_DATUM: col.Hint = "Monatserster " & mPivotYear
            Else
                col.FType = FT_AUSWAHL: col.Hint = "aus den Spaltenköpfen"
            End If
        Case "quartal"
            col.FType = FT_AUSWAHL: col.Hint = "aus den Spaltenköpfen"
        Case Else
            col.FType = FT_ZAHL: col.Decimals = 0: col.Hint = "aus den Spaltenköpfen"
    End Select
    For Each k In mPivot
        col.Src.Add CLng(k)
        If col.FType = FT_AUSWAHL Then col.Opts.Add LiOptName(CStr(mUCols(CLng(k))(0)))
    Next
    col.Sample = CStr(mUCols(mPivot(1))(0)) & ", " & CStr(mUCols(mPivot(2))(0))
    liCols.Add col
    For Each k In mPivot
        For Each rec In mRecs
            v = RecVal(rec, mUCols(CLng(k)))
            If Not IsBlank(v) Then vals.Add v
        Next
    Next
    Set p = ProfileValues(vals, "", UFmt(mUCols(mPivot(1))), mCsv)
    Set col = New cLiCol
    col.Key = "P:wert": col.Mode = "monatswert": col.Name = PivotValueName()
    For Each k In mPivot: col.Src.Add CLng(k): Next
    col.FType = CStr(p.GetVal("typ"))
    If Not (col.FType = FT_ZAHL Or col.FType = FT_GELD Or col.FType = FT_PROZENT) Then col.FType = FT_TEXT
    col.Decimals = CLng(p.GetVal("decimals", 0))
    col.Sample = Truncate(Replace(CStr(p.GetVal("beispiele", "")), "|", ", "), 60)
    col.Hint = "Werte aus den Zeitraumspalten"
    Set bad = p.GetVal("ungueltig")
    If bad.Count > 0 And col.FType <> FT_TEXT Then
        AddQ "typ:P:wert", col.Name, "In den Zeitraumspalten " & IIf(bad.Count = 1, "passt 1 Wert", "passen " & bad.Count & " Werte") & _
             " nicht zum Typ " & FieldTypeName(col.FType) & ", z. B. „" & Truncate(CStr(bad(1)), 30) & "“.", _
             Array("Diese Werte leer lassen und im Bericht melden", "Alles als Text übernehmen"), 0
        If Ans("typ:P:wert", 0) = 1 Then col.FType = FT_TEXT
        col.Unsure = True
    End If
    liCols.Add col
End Sub

Private Function PivotPeriodValue(ByVal col As cLiCol, ByVal pivotIdx As Long) As Variant
    Dim nm As String, kind As String, idx As Long, yr As Long
    nm = CStr(mUCols(pivotIdx)(0))
    PeriodOf nm, kind, idx, yr
    If col.FType = FT_DATUM Or col.FType = FT_DATUMZEIT Then
        If yr = 0 Then yr = col.Year_
        If yr > 0 And kind = "monat" Then PivotPeriodValue = DateSerial(yr, idx, 1)
    ElseIf col.FType = FT_ZAHL Then
        If kind = "jahr" Then PivotPeriodValue = CDbl(idx)
    Else
        PivotPeriodValue = LiOptName(nm)
    End If
End Function

Private Sub MarkLabelCandidates()
    Dim col As cLiCol
    For Each col In liCols
        col.CanBeLabel = (col.FType = FT_TEXT Or col.FType = FT_EMAIL Or col.FType = FT_LTEXT)
    Next
End Sub

' Formular: eine Angabe je Zeile -> ein Eintrag
Private Sub PlanForm()
    Dim sh As cLiSheet, r As Long, col As cLiCol, nm As String, vals As Collection, p As cMap, used As New cMap, k As Long, vc As Long
    Set sh = liSheets(liTab.SheetIdx)
    Set mRecs = New Collection
    mRecs.Add Array(1, liTab.HdrRow, "")
    Set mUCols = New Collection
    For r = liTab.HdrRow To liTab.R2
        If liTab.RowClass.GetVal(CStr(r), "") = "daten" Then
            vc = FormValueCol(sh, r, liTab.C1, liTab.C2)
            nm = CleanName(T_(sh.Cell(r, liTab.C1)))
            If Len(nm) = 0 Then nm = "Angabe " & r
            If used.Exists(nm) Then
                k = 2
                Do While used.Exists(nm & " (" & k & ")"): k = k + 1: Loop
                nm = nm & " (" & k & ")"
            End If
            used.SetVal nm, True
            Set col = New cLiCol
            col.Key = "F:" & r: col.Mode = "formular": col.Name = nm
            col.Src.Add r
            col.SrcCol = vc
            Set vals = New Collection
            vals.Add sh.Cell(r, vc)
            Set p = ProfileValues(vals, nm, sh.CellFormat(r, vc), sh.IsCsv)
            col.FType = CStr(p.GetVal("typ"))
            If col.FType = FT_AUSWAHL Or col.FType = FT_MEHRFACH Or col.FType = FT_JANEIN Then col.FType = FT_TEXT
            col.Decimals = CLng(p.GetVal("decimals", 0))
            col.Sample = Truncate(T_(sh.Cell(r, vc)), 60)
            liCols.Add col
        End If
    Next
    MarkLabelCandidates
    ApplyOverrides
End Sub

' Aenderungen des Nutzers (Name, Typ, Uebernehmen) anwenden
Private Sub ApplyOverrides()
    Dim col As cLiCol, a As Variant, oldType As String
    For Each col In liCols
        If liOvr.Exists(col.Key) Then
            a = liOvr.GetVal(col.Key)
            If Len(a(0)) > 0 Then col.Name = CleanName(CStr(a(0)))
            oldType = col.FType
            If Len(a(1)) > 0 Then col.FType = CStr(a(1))
            If Not IsEmpty(a(2)) Then col.Include = CBool(a(2))
            If col.FType <> oldType Then AfterTypeChange col
        End If
    Next
    EnsureUniqueNames
    MarkLabelCandidates
End Sub

' vom Nutzer gewaehlter Typ eines Felds ("" = keiner)
Private Function OvType(ByVal key As String) As String
    Dim a As Variant
    If liOvr Is Nothing Then Exit Function
    If liOvr.Exists(key) Then
        a = liOvr.GetVal(key)
        OvType = CStr(a(1))
    End If
End Function

Private Sub AfterTypeChange(ByVal col As cLiCol)
    If liMode = "2ask" Then Exit Sub
    If col.FType = FT_MEHRFACH And col.Mode = "wert" Then
        col.Mode = "liste"
        If Len(col.Sep) = 0 Then col.Sep = ",;/"
    ElseIf col.FType <> FT_MEHRFACH And col.Mode = "liste" Then
        col.Mode = "wert"
    End If
    If col.IsOptionType Then
        Select Case col.Mode
            Case "wert", "liste", "gruppe", "formular": FillOptions col
        End Select
    End If
End Sub

Private Sub EnsureUniqueNames()
    Dim col As cLiCol, used As New cMap, nm As String, k As Long
    For Each col In liCols
        If col.Include Then
            nm = col.Name
            If used.Exists(nm) Then
                k = 2
                Do While used.Exists(nm & " (" & k & ")"): k = k + 1: Loop
                nm = nm & " (" & k & ")"
                col.Name = nm
            End If
            used.SetVal nm, True
        End If
    Next
End Sub

' Aenderung an einem Feld merken (Name, Typ-Code, Uebernehmen); danach LiPlan
Public Sub LiSetField(ByVal key As String, ByVal nm As String, ByVal ftype As String, ByVal incl As Boolean)
    liOvr.SetVal key, Array(nm, ftype, incl)
End Sub

' Teil einer Feldaenderung merken: part = "name" | "type" | "incl"
Public Sub LiSetFieldPart(ByVal key As String, ByVal part As String, ByVal value As Variant)
    Dim a As Variant
    If liOvr.Exists(key) Then a = liOvr.GetVal(key) Else a = Array("", "", Empty)
    Select Case part
        Case "name": a(0) = CStr(value)
        Case "type": a(1) = CStr(value)
        Case "incl": a(2) = CBool(value)
    End Select
    liOvr.SetVal key, a
End Sub

Public Function LiCleanName(ByVal s As String) As String
    LiCleanName = CleanName(s)
End Function

' Feldtypen, die der Assistent anbietet
Public Function LiTypeCodes() As Variant
    LiTypeCodes = Array(FT_TEXT, FT_LTEXT, FT_ZAHL, FT_GELD, FT_PROZENT, FT_DATUM, FT_DATUMZEIT, FT_JANEIN, FT_AUSWAHL, FT_MEHRFACH, _
                        FT_EMAIL, FT_TELEFON, FT_URL)
End Function

' Vorschlag fuer den Namen der neuen Tabelle (eindeutig)
Public Function LiSuggestTableName() As String
    Dim s As String, stem As String, k As Long, sh As cLiSheet
    If liMode = "2ask" Then
        s = "Befragung"
    ElseIf liTab Is Nothing Then
        s = FileBase()
    ElseIf liTab.Kind = "zusammen" Then
        s = CommonTitle()
        If Len(s) < 4 Then s = FileBase()
    Else
        Set sh = liSheets(liTab.SheetIdx)
        s = sh.Name
        If sh.IsCsv Or IsDefaultSheetName(s) Then s = FileBase()
    End If
    s = CleanName(s)
    If Len(s) > 60 Then s = Trim$(Left$(s, 60))
    If Len(s) = 0 Then s = "Übernommene Liste"
    stem = s: k = 2
    Do While Len(mCfg.TableIdByName(s)) > 0
        s = stem & " (" & k & ")": k = k + 1
    Loop
    LiSuggestTableName = s
End Function

Private Function FileBase() As String
    Dim s As String
    s = FileNameOf(liPath)
    If InStrRev(s, ".") > 1 Then s = Left$(s, InStrRev(s, ".") - 1)
    FileBase = Trim$(Replace(s, "_", " "))
End Function

Private Function IsDefaultSheetName(ByVal s As String) As Boolean
    s = LCase$(s)
    IsDefaultSheetName = (s Like "tabelle#*" Or s Like "sheet#*" Or s Like "blatt#*")
End Function

' gemeinsamer Anfang der Titel zusammengefuehrter Blaetter ("Bedarfsmeldung 2026: Kita ..." -> "Bedarfsmeldung 2026")
Private Function CommonTitle() As String
    Dim m As cLiTab, s As String, t As String, i As Long, first As Boolean
    first = True
    For Each m In liTab.Members
        t = m.Title
        If first Then
            s = t: first = False
        Else
            i = 0
            Do While i < Len(s) And i < Len(t)
                If Mid$(s, i + 1, 1) <> Mid$(t, i + 1, 1) Then Exit Do
                i = i + 1
            Loop
            s = Left$(s, i)
        End If
    Next
    If InStr(s, ":") > 0 Then s = Left$(s, InStr(s, ":") - 1)
    s = Trim$(s)
    Do While Len(s) > 0
        If InStr(" -,;(/", Right$(s, 1)) > 0 Then s = Trim$(Left$(s, Len(s) - 1)) Else Exit Do
    Loop
    CommonTitle = s
End Function

' Zielwert fuer einen Rohwert eines Auswahlfelds ("" = leer lassen); danach LiPlan
Public Sub LiSetValueMap(ByVal key As String, ByVal rawValue As String, ByVal target As String)
    liValOvr.SetVal key & "|" & rawValue, target
End Sub

Public Function LiFieldByKey(ByVal key As String) As cLiCol
    Dim col As cLiCol
    For Each col In liCols
        If col.Key = key Then Set LiFieldByKey = col: Exit Function
    Next
End Function

' Lassen sich die Werte dieses Felds umbenennen/zusammenlegen?
Public Function LiValuesEditable(ByVal key As String) As Boolean
    Dim col As cLiCol
    Set col = LiFieldByKey(key)
    If col Is Nothing Then Exit Function
    If liMode = "2ask" Or Not col.IsOptionType Then Exit Function
    Select Case col.Mode
        Case "wert", "liste", "gruppe", "formular": LiValuesEditable = True
    End Select
End Function

' Werte eines Auswahlfelds mit Anzahl: Array(Rohwert, Anzahl, Zielwert)
Public Function LiFieldValues(ByVal key As String) As Collection
    Dim col As cLiCol, res As New Collection, cnt As New cMap, order As New Collection, rec As Variant, v As Variant
    Dim parts As Variant, i As Long, s As String, k As Variant, r As Long, why As String, raw As String, pv As Variant
    Set LiFieldValues = res
    Set col = LiFieldByKey(key)
    If col Is Nothing Then Exit Function
    If Not col.IsOptionType Then Exit Function
    If liMode = "2ask" Then
        For r = mAskHdr + 1 To mAskLast
            v = AskFieldValue(col, r, why, raw)
            If Not IsBlank(v) Then
                For Each k In Split(CStr(v), ";")
                    cnt.AddNum CStr(k), 1
                Next
            End If
        Next
        For Each k In col.Opts
            res.Add Array(CStr(k), FmtInt(CDbl(cnt.GetVal(CStr(k), 0))), CStr(k))
        Next
        Exit Function
    End If
    Select Case col.Mode
        Case "kreuz"
            For i = 1 To col.Src.Count
                For Each rec In mRecs
                    If IsMarkVal(RecVal(rec, mUCols(CLng(col.Src(i))))) Then cnt.AddNum CStr(col.SrcNames(i)), 1
                Next
                res.Add Array(CStr(col.SrcNames(i)), FmtInt(CDbl(cnt.GetVal(CStr(col.SrcNames(i)), 0))), CStr(col.SrcNames(i)))
            Next
        Case "monat"
            For Each pv In mPivot
                For Each rec In mRecs
                    If Not IsBlank(RecVal(rec, mUCols(CLng(pv)))) Then cnt.AddNum CStr(mUCols(CLng(pv))(0)), 1
                Next
                res.Add Array(CStr(mUCols(CLng(pv))(0)), FmtInt(CDbl(cnt.GetVal(CStr(mUCols(CLng(pv))(0)), 0))), LiOptName(CStr(mUCols(CLng(pv))(0))))
            Next
        Case "blatt"
            For Each rec In mRecs
                cnt.AddNum liSheets(mMembers(CLng(rec(0))).SheetIdx).Name, 1
            Next
            For Each k In cnt.Keys
                res.Add Array(CStr(k), FmtInt(CDbl(cnt.GetVal(CStr(k)))), LiOptName(CStr(k)))
            Next
        Case Else
            For Each rec In mRecs
                v = SrcValue(col, rec)
                If Not IsBlank(v) Then
                    If col.Mode = "liste" Then parts = SplitMulti(T_(v), col.Sep) Else parts = Array(T_(v))
                    For i = 0 To ArrLen(parts) - 1
                        s = LiOptName(CStr(parts(i)))
                        If Len(s) > 0 Then
                            If Not cnt.Exists(s) Then order.Add s
                            cnt.AddNum s, 1
                        End If
                    Next
                End If
            Next
            For Each k In order
                res.Add Array(CStr(k), FmtInt(CDbl(cnt.GetVal(CStr(k)))), CStr(col.ValMap.GetVal(CStr(k), k)))
            Next
    End Select
End Function

' Vorschlag fuer das Bezeichnungsfeld (Feldschluessel)
Public Function LiSuggestLabel() As String
    Dim col As cLiCol, best As String, bestScore As Double, sc As Double, first As String, isId As Boolean
    If liMode = "2ask" Then
        For Each col In liCols
            If col.Key = "survey_session_id" And col.Include Then LiSuggestLabel = col.Key: Exit Function
        Next
    End If
    bestScore = -1E+30
    For Each col In liCols
        If col.Include And col.CanBeLabel And col.Mode <> "blatt" And col.Mode <> "gruppe" Then
            If Len(first) = 0 Then first = col.Key
            sc = 0
            ' Kennungen ("Projektcode", "Schulnr.") nur, wenn es keine Namensspalte gibt
            isId = IdLikeHeader(col.Name) Or HasWord(LCase$(col.Name), "code|kennung")
            If LabelHeader(col.Name) And Not isId Then sc = sc + 5
            If isId And Not col.AllUnique Then sc = sc - 6
            If col.AllUnique Then sc = sc + 2
            ' viele gleiche Werte (Kreis, Amt, Status) taugen nicht als Bezeichnung
            If col.DistRatio > 0 And col.DistRatio < 0.5 Then
                sc = sc - 6
            ElseIf col.DistRatio > 0 And col.DistRatio < 0.8 Then
                sc = sc - 2
            End If
            ' Bemerkungen und selten ausgefuellte Spalten ebenso wenig
            If LongTextHeader(col.Name) Then sc = sc - 6
            If col.FillRatio > 0 And col.FillRatio < 0.8 Then sc = sc - 6
            If col.FType = FT_LTEXT And Not LabelHeader(col.Name) Then sc = sc - 2
            If col.FType = FT_EMAIL Then sc = sc - 3
            If sc > bestScore Then bestScore = sc: best = col.Key
        End If
    Next
    If Len(best) = 0 Then best = first
    LiSuggestLabel = best
End Function

'==============================================================================
' Werte umwandeln
'==============================================================================
Private Function ConvertTyped(ByVal v As Variant, ByVal typ As String, ByVal csv As Boolean, ByRef outV As Variant, ByRef why As String) As Boolean
    Dim s As String, d As Double, dt As Date, ht As Boolean, msg As String
    outV = Empty: why = ""
    If IsBlank(v) Then ConvertTyped = True: Exit Function
    s = T_(v)
    ' "-" und aehnliche Platzhalter in Zahlen- und Datumsfeldern: leer, ohne Meldung
    If VarType(v) = vbString And IsPlaceholder(s) Then
        Select Case typ
            Case FT_ZAHL, FT_GELD, FT_PROZENT, FT_DATUM, FT_DATUMZEIT: ConvertTyped = True: Exit Function
        End Select
    End If
    Select Case typ
        Case FT_ZAHL, FT_GELD, FT_PROZENT
            If IsNumVal(v) Then
                outV = CDbl(v)
            ElseIf VarType(v) = vbString And TextNumber(s, d, csv) Then
                If typ = FT_PROZENT And InStr(s, "%") > 0 Then d = d / 100
                outV = d
            Else
                why = IIf(typ = FT_GELD, "kein Geldbetrag", "keine Zahl"): Exit Function
            End If
        Case FT_DATUM, FT_DATUMZEIT
            If VarType(v) = vbDate Then
                dt = CDate(v)
            ElseIf IsNumVal(v) Then
                If CDbl(v) < 3000 Or CDbl(v) > 2958465 Then why = "kein Datum": Exit Function
                dt = CDate(CDbl(v))
            ElseIf Not TextDate(s, dt, ht) Then
                why = "kein Datum": Exit Function
            End If
            If Year(dt) < 1900 Then why = "Datum außerhalb des gültigen Bereichs": Exit Function
            If typ = FT_DATUM Then dt = Int(CDbl(dt))
            outV = dt
        Case FT_JANEIN, FT_KONTROLL
            If VarType(v) = vbBoolean Then
                outV = CBool(v)
            ElseIf IsMarkVal(v) Then
                outV = True
            ElseIf IsNoText(s) Or (IsNumVal(v) And s = "0") Then
                outV = False
            Else
                why = "kein Ja/Nein": Exit Function
            End If
        Case FT_EMAIL, FT_TELEFON, FT_URL
            If Not mValid.KindFormatOk(typ, s, msg) Then why = "Format passt nicht: " & msg: Exit Function
            outV = s
        Case Else
            outV = s
    End Select
    ConvertTyped = True
End Function

' Wert eines geplanten Felds fuer einen Datensatz (sonstige Liste)
Private Function FieldValue(ByVal col As cLiCol, ByVal rec As Variant, ByVal pivotIdx As Long, ByRef why As String, ByRef raw As String) As Variant
    Dim v As Variant, s As String, i As Long, parts As Variant, t As String, outV As Variant, tg As Variant, j As Long
    why = "": raw = ""
    Select Case col.Mode
        Case "kreuz"
            For i = 1 To col.Src.Count
                v = RecVal(rec, mUCols(CLng(col.Src(i))))
                If IsMarkVal(v) Then
                    s = ListAdd(s, CStr(col.SrcNames(i)))
                ElseIf Not IsBlank(v) Then
                    raw = raw & IIf(Len(raw) > 0, ", ", "") & mUCols(CLng(col.Src(i)))(0) & ": " & T_(v)
                    why = "übernommen; kein Kreuz, gilt als nicht angekreuzt"
                End If
            Next
            If Len(s) > 0 Then FieldValue = s
            Exit Function
        Case "monat"
            FieldValue = PivotPeriodValue(col, pivotIdx)
            Exit Function
        Case "monatswert"
            v = RecVal(rec, mUCols(pivotIdx))
        Case Else
            v = SrcValue(col, rec)
    End Select
    ' "dito": Wert der Zeile darueber (je Blatt)
    If col.Ditto Then
        If mDitto Is Nothing Then Set mDitto = New cMap
        If VarType(v) = vbString Then
            If IsDittoText(T_(v)) Then v = mDitto.GetVal(col.Key & "|" & rec(0), Empty)
        End If
        If Not IsBlank(v) Then mDitto.SetVal col.Key & "|" & rec(0), v
    End If
    raw = T_(v)
    If IsBlank(v) Then
        If (col.FType = FT_JANEIN Or col.FType = FT_KONTROLL) And col.BlankFalse Then FieldValue = False
        Exit Function
    End If
    Select Case col.FType
        Case FT_AUSWAHL, FT_STATUS
            t = LiOptName(T_(v))
            t = LiOptName(Replace(CStr(col.ValMap.GetVal(t, t)), ";", " "))
            If Len(t) > 0 Then FieldValue = t
        Case FT_MEHRFACH
            If col.Mode = "liste" Then parts = SplitMulti(T_(v), col.Sep) Else parts = Array(T_(v))
            For i = 0 To ArrLen(parts) - 1
                t = LiOptName(CStr(parts(i)))
                tg = Split(CStr(col.ValMap.GetVal(t, t)), ";")
                For j = 0 To UBound(tg)
                    s = ListAdd(s, LiOptName(CStr(tg(j))))
                Next
            Next
            If Len(s) > 0 Then FieldValue = s
        Case FT_TEXT, FT_LTEXT
            FieldValue = T_(v)
        Case Else
            If ConvertTyped(v, col.FType, mCsv, outV, why) Then FieldValue = outV
    End Select
End Function

'==============================================================================
' Importplan bauen (fuer mIO.CheckImport / ExecuteImport)
'==============================================================================
' tid = "" -> neue Tabelle newName; keyKey/labelKey = Feldschluessel (optional)
Public Function LiBuildPlan(ByVal newName As String, ByVal tid As String, ByVal keyKey As String, ByVal labelKey As String) As cMap
    Dim plan As New cMap, cols As New Collection, col As cLiCol, data() As Variant, nOut As Long, j As Long, r As Long, total As Long
    Dim map As New Collection, ord As New cMap, rec As Variant, why As String, raw As String, v As Variant, props As cMap
    Dim target As String, labelCol As Long, keyFid As String, pv As Variant, pi As Long, src() As String, f As cField, skipPre As Boolean
    Set liNotes = New Collection
    Set mDitto = New cMap
    For Each col In liCols
        If col.Include Then cols.Add col
    Next
    If cols.Count = 0 Then RaiseUser "Bitte mindestens ein Feld zum Übernehmen auswählen."
    nOut = cols.Count
    total = LiRecordCount()
    ReDim data(1 To total + 1, 1 To nOut)
    ReDim src(1 To total + 1)
    j = 0
    For Each col In cols
        j = j + 1
        data(1, j) = col.Name
        target = ""
        If Len(tid) > 0 Then
            target = mCfg.FieldIdByName(tid, col.Name)
            If Len(target) > 0 Then
                Set f = mCfg.Field(target)
                If f.FType = FT_FORMEL Or f.FType = FT_BACKLINK Or f.FType = FT_LINK Or IsAutoKind(f.Kind) Then
                    liNotes.Add Array("", col.Name, "", "Spalte wird nicht übernommen: Das Feld wird in der Zieltabelle berechnet oder verknüpft.")
                    target = "-"
                End If
            End If
        End If
        If target = "-" Then
            target = ""
        ElseIf Len(target) = 0 Then
            target = "NEW:" & j
        End If
        Set props = New cMap
        If col.FType = FT_ZAHL Then props.SetVal "Dezimalstellen", IIf(col.Decimals >= 0, col.Decimals, 2)
        If Len(col.Help) > 0 Then props.SetVal "Hilfetext", Truncate(col.Help, 250)
        map.Add Array(j, target, col.FType, col.Name, props)
        If col.IsOptionType And col.Opts.Count > 0 And Len(target) > 0 Then ord.SetVal target, col.Opts
        If col.Key = labelKey And Left$(target, 4) = "NEW:" Then labelCol = j
        If col.Key = keyKey And Len(tid) > 0 And Len(target) > 0 And Left$(target, 4) <> "NEW:" Then keyFid = target
    Next
    ' Datensaetze
    r = 1
    If liMode = "2ask" Then
        skipPre = SkipPretest()
        For pi = mAskHdr + 1 To mAskLast
            If Not (skipPre And IsPretestRow(pi)) And r <= total Then
                r = r + 1
                src(r) = "Fragebogen " & (pi - mAskHdr) & " (Zeile " & pi & ")"
                j = 0
                For Each col In cols
                    j = j + 1
                    v = AskFieldValue(col, pi, why, raw)
                    If Len(why) > 0 Then liNotes.Add Array(src(r), col.Name, raw, NoteText(why))
                    If Not IsBlank(v) Then data(r, j) = v
                Next
            End If
        Next
    ElseIf liTab.Kind = "formular" Then
        r = 2
        FillRecord data, src, r, cols, mRecs(1), 0
    Else
        For Each rec In mRecs
            If mPivot Is Nothing Then
                If r <= total Then
                    r = r + 1
                    FillRecord data, src, r, cols, rec, 0
                End If
            Else
                For Each pv In mPivot
                    If Not IsBlank(RecVal(rec, mUCols(CLng(pv)))) And r <= total Then
                        r = r + 1
                        FillRecord data, src, r, cols, rec, CLng(pv)
                    End If
                Next
            End If
        Next
    End If
    liRowSrc = src
    plan.SetVal "data", mData.SliceRows(data, r, nOut)
    plan.SetVal "tid", tid
    plan.SetVal "newTable", newName
    plan.SetVal "map", map
    plan.SetVal "decPoint", False
    plan.SetVal "createOptions", True
    plan.SetVal "labelCol", labelCol
    plan.SetVal "optionOrder", ord
    plan.SetVal "tableDesc", "Übernommen aus „" & FileNameOf(liPath) & "“ am " & FmtDate(Date)
    If Len(keyFid) > 0 Then plan.SetVal "key", keyFid
    Set LiBuildPlan = plan
End Function

Private Sub FillRecord(ByRef data() As Variant, ByRef src() As String, ByVal r As Long, ByVal cols As Collection, ByVal rec As Variant, ByVal pivotIdx As Long)
    Dim col As cLiCol, j As Long, v As Variant, why As String, raw As String, lbl As String
    If liTab.Kind = "formular" Then
        src(r) = "Formular"
    ElseIf mMembers.Count > 1 Then
        src(r) = "Blatt „" & liSheets(mMembers(CLng(rec(0))).SheetIdx).Name & "“, Zeile " & rec(1)
    Else
        src(r) = "Zeile " & rec(1)
    End If
    If pivotIdx > 0 Then src(r) = src(r) & ", " & mUCols(pivotIdx)(0)
    For Each col In cols
        j = j + 1
        v = FieldValue(col, rec, pivotIdx, why, raw)
        If Len(why) > 0 Then
            lbl = src(r)
            If col.Mode = "formular" Then lbl = "Formular, Zeile " & col.FirstSrc
            liNotes.Add Array(lbl, col.Name, raw, NoteText(why))
        End If
        If Not IsBlank(v) Then data(r, j) = v
    Next
End Sub

' Berichtstext eines Hinweises: Gruende ohne eigenes "übernommen ..." bedeuten, dass der Wert leer bleibt
Private Function NoteText(ByVal why As String) As String
    If Left$(why, 10) = "übernommen" Then NoteText = why Else NoteText = "übernommen, Wert bleibt leer: " & why
End Function

' Quellangabe zu einer Zeile des Importplans (2 = erster Datensatz)
Public Function LiSourceRow(ByVal planRow As Long) As String
    If Not IsArray(liRowSrc) Then LiSourceRow = CStr(planRow): Exit Function
    If planRow < LBound(liRowSrc) Or planRow > UBound(liRowSrc) Then LiSourceRow = CStr(planRow): Exit Function
    LiSourceRow = CStr(liRowSrc(planRow))
    If Len(LiSourceRow) = 0 Then LiSourceRow = CStr(planRow)
End Function

' Anzahl Datensaetze, die uebernommen werden
Public Function LiRecordCount() As Long
    Dim rec As Variant, pv As Variant, r As Long, skipPre As Boolean
    If liMode = "2ask" Then
        skipPre = SkipPretest()
        For r = mAskHdr + 1 To mAskLast
            If Not (skipPre And IsPretestRow(r)) Then LiRecordCount = LiRecordCount + 1
        Next
        Exit Function
    End If
    If liTab Is Nothing Then Exit Function
    If liTab.Kind = "formular" Then LiRecordCount = 1: Exit Function
    If mRecs Is Nothing Then Exit Function
    If mPivot Is Nothing Then LiRecordCount = mRecs.Count: Exit Function
    For Each rec In mRecs
        For Each pv In mPivot
            If Not IsBlank(RecVal(rec, mUCols(CLng(pv)))) Then LiRecordCount = LiRecordCount + 1
        Next
    Next
End Function

'==============================================================================
' 2ask: kodierter Excel-Export mit Codeplan
'==============================================================================
Private Sub ReadTwoAsk(ByVal wb As Workbook)
    Dim r As Long, c As Long, sh As cLiSheet, tc As Long
    Set mAskSheet = LoadSheet(wb.Worksheets("results"))
    Set sh = mAskSheet
    For r = 1 To Application.Min(10, sh.R)
        If LCase$(T_(sh.Cell(r, 1))) = "survey_session_id" Then mAskHdr = r: Exit For
    Next
    If mAskHdr = 0 Then RaiseUser "Im Blatt „results“ fehlt die Kopfzeile mit „survey_session_id“. Ist das ein unveränderter 2ask-Export?"
    mAskLast = mAskHdr
    For r = mAskHdr + 1 To sh.R
        If Not IsBlank(sh.Cell(r, 1)) Then mAskLast = r
    Next
    liAskRows = mAskLast - mAskHdr
    If liAskRows = 0 Then RaiseUser "Der 2ask-Export enthält keine Fragebögen."
    Set mAskKeyCol = New cMap
    For c = 1 To sh.C
        If Len(T_(sh.Cell(mAskHdr, c))) > 0 Then mAskKeyCol.SetVal T_(sh.Cell(mAskHdr, c)), c
    Next
    ReadCodePlan wb.Worksheets("code_plan")
    tc = AskKeyCol("survey_session_type")
    If tc > 0 Then
        For r = mAskHdr + 1 To mAskLast
            If LCase$(T_(sh.Cell(r, tc))) = "pretest" Then mAskPretest = mAskPretest + 1
        Next
    End If
    liSheets.Add mAskSheet
End Sub

Private Sub ReadCodePlan(ByVal ws As Worksheet)
    Dim v As Variant, r As Long, lastR As Long, q As String
    Set mCp = New Collection
    Set mQKind = New cMap
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then RaiseUser "Der Codeplan im 2ask-Export ist leer."
    v = ws.Range("A1:F" & lastR).Value
    If LCase$(T_(v(1, 1))) <> "export_key" Then RaiseUser "Das Blatt „code_plan“ hat nicht den erwarteten Aufbau (Spalte A „export_key“)."
    For r = 2 To lastR
        Select Case T_(v(r, 2))
            Case "question"
                q = T_(v(r, 6))
                mQKind.SetVal AskQuestionKey(T_(v(r, 1))), T_(v(r, 3))
            Case "standard_field"
                q = ""
        End Select
        mCp.Add Array(T_(v(r, 1)), T_(v(r, 2)), T_(v(r, 3)), T_(v(r, 4)), v(r, 5), T_(v(r, 6)), q)
    Next
End Sub

Private Function AskKeyCol(ByVal key As String) As Long
    If mAskKeyCol Is Nothing Then Exit Function
    If mAskKeyCol.Exists(key) Then AskKeyCol = CLng(mAskKeyCol.GetVal(key))
End Function

Private Function IsPretestRow(ByVal r As Long) As Boolean
    Dim tc As Long
    tc = AskKeyCol("survey_session_type")
    If tc > 0 Then IsPretestRow = (LCase$(T_(mAskSheet.Cell(r, tc))) = "pretest")
End Function

Private Function SkipPretest() As Boolean
    SkipPretest = HasQuestion("pretest") And Ans("pretest", 0) = 0
End Function

Private Function MetaName(ByVal key As String) As String
    Select Case key
        Case "survey_session_id": MetaName = "Teilnahme-ID"
        Case "survey_session_type": MetaName = "Art der Teilnahme"
        Case "started_at": MetaName = "Begonnen am"
        Case "ended_at": MetaName = "Beendet am"
        Case "last_interaction_at": MetaName = "Letzte Bearbeitung"
        Case "progress_percent": MetaName = "Fortschritt in %"
        Case "duration_in_seconds": MetaName = "Dauer in Sekunden"
        Case "items_answered_count": MetaName = "Beantwortete Angaben"
        Case "items_answered_percent": MetaName = "Beantwortete Angaben in %"
        Case "locale": MetaName = "Sprache der Umfrage"
        Case "source": MetaName = "Zugangsweg"
        Case "country": MetaName = "Land"
        Case "device_type": MetaName = "Gerät"
        Case "browser": MetaName = "Browser"
        Case "browser_locale": MetaName = "Browsersprache"
        Case "browser_version": MetaName = "Browserversion"
        Case "operating_system": MetaName = "Betriebssystem"
        Case "contact.id": MetaName = "Kontakt-ID"
        Case "contact.external_key": MetaName = "Kontakt: externer Schlüssel"
        Case "contact.email": MetaName = "Kontakt: E-Mail"
        Case "contact.first_name": MetaName = "Kontakt: Vorname"
        Case "contact.last_name": MetaName = "Kontakt: Nachname"
        Case "contact.language": MetaName = "Kontakt: Sprache"
        Case "contact.gender": MetaName = "Kontakt: Anrede"
        Case Else: MetaName = key
    End Select
End Function

Private Function AskColHasValue(ByVal c As Long) As Boolean
    Dim r As Long
    For r = mAskHdr + 1 To mAskLast
        If Not IsBlank(mAskSheet.Cell(r, c)) Then AskColHasValue = True: Exit Function
    Next
End Function

' Unterbeschriftung (Zeile ueber dem Import-JSON): "Answer: X; Answer-Option:Y", "Answer: X; Group: Y"
Private Function AskSub(ByVal c As Long, ByVal part As String) As String
    Dim s As String, p As Long, q As Long, tag As String
    If mAskHdr < 3 Then Exit Function
    s = T_(mAskSheet.Cell(mAskHdr - 2, c))
    Select Case part
        Case "answer": tag = "Answer:"
        Case "option": tag = "Answer-Option:"
        Case "group": tag = "Group:"
    End Select
    p = InStr(s, tag)
    If p = 0 Then Exit Function
    s = Mid$(s, p + Len(tag))
    q = InStr(s, ";")
    If q > 0 Then s = Left$(s, q - 1)
    s = Trim$(s)
    If part = "option" And Right$(s, 8) = " (Label)" Then s = Trim$(Left$(s, Len(s) - 8))
    AskSub = s
End Function

' Fragenummer: "Q2t_05" -> "Q2", "Q12_1_01" -> "Q12"
Private Function AskQuestionKey(ByVal key As String) As String
    Dim p As Long, s As String
    p = InStr(key, "_")
    If p > 0 Then s = Left$(key, p - 1) Else s = key
    If Len(s) > 2 And Right$(s, 1) = "t" Then
        If IsAllDigits(Mid$(s, 2, Len(s) - 2)) Then s = Left$(s, Len(s) - 1)
    End If
    AskQuestionKey = s
End Function

' Matrixzeile: "Q5_1_01" -> "Q5_1"
Private Function AskGroupKey(ByVal key As String) As String
    Dim p As Long
    p = InStrRev(key, "_")
    If p > 0 Then AskGroupKey = Left$(key, p - 1) Else AskGroupKey = key
End Function

Private Function AskKindName(ByVal sb As String, ByVal qsub As String) As String
    Select Case qsub
        Case "multiple_choice"
            Select Case sb
                Case "select": AskKindName = "Dropdown"
                Case "checkbox": AskKindName = "Mehrfachauswahl"
                Case Else: AskKindName = "Einfachauswahl"
            End Select
        Case "matrix_table"
            Select Case sb
                Case "checkbox": AskKindName = "Matrix Mehrfachauswahl"
                Case "input": AskKindName = "Matrix Textfeld"
                Case "select": AskKindName = "Matrix Dropdown"
                Case Else: AskKindName = "Matrix"
            End Select
        Case "slider": AskKindName = "Sterne/Schieberegler"
        Case "form": AskKindName = "Formular"
        Case "side_by_side": AskKindName = "Gruppenmatrix"
        Case "rank_order": AskKindName = "Rangfolge"
        Case "constant_sum": AskKindName = "Summe"
        Case "text_entry": AskKindName = "Eingabefeld"
        Case Else: AskKindName = IIf(Len(qsub) > 0, qsub, sb)
    End Select
End Function

Private Sub RegisterAskQuestion(ByVal qList As cMap, ByVal qKey As String, ByVal qText As String, ByVal kind As String, ByVal nFields As Long)
    Dim a As Variant
    If qList.Exists(qKey) Then
        a = qList.GetVal(qKey)
        a(3) = a(3) + nFields
        If Len(kind) > 0 And (a(2) = "Zusatztext" Or Len(a(2)) = 0) Then a(2) = kind
        qList.SetVal qKey, a
    Else
        qList.SetVal qKey, Array(qKey, qText, kind, nFields)
    End If
End Sub

Private Function NewAskCol(ByVal key As String, ByVal nm As String, ByVal ftype As String, ByVal mode As String, ByVal c As Long) As cLiCol
    Dim col As New cLiCol
    col.Key = key: col.Name = CleanName(nm): col.FType = ftype: col.Mode = mode
    If Len(col.Name) = 0 Then col.Name = key
    ' gekuerzter Fragetext: vollstaendige Frage als Hilfetext des Felds
    If FullName(nm) <> col.Name And Len(FullName(nm)) > 0 Then col.Help = FullName(nm)
    If c > 0 Then col.Src.Add c
    Set NewAskCol = col
End Function

Private Function AskChoiceCol(ByVal key As String, ByVal nm As String, ByVal opts As cMap, ByVal c As Long, ByRef isBip As Boolean) As cLiCol
    Dim col As cLiCol, k As Variant, allBracket As Boolean, lbl As String
    Set col = NewAskCol(key, nm, FT_AUSWAHL, "code", c)
    allBracket = (opts.Count > 0)
    For Each k In opts.Keys
        lbl = CStr(opts.GetVal(CStr(k)))
        If Not (Left$(lbl, 1) = "[" And Right$(lbl, 1) = "]") Then allBracket = False
        col.CodeMap.SetVal CStr(k), LiOptName(lbl)
        If Not InColl(col.Opts, LiOptName(lbl)) Then col.Opts.Add LiOptName(lbl)
    Next
    isBip = allBracket
    If isBip Then
        col.Tag = "bip:" & AskQuestionKey(key)
        If Ans(col.Tag, 0) = 0 Then col.FType = FT_ZAHL: col.Decimals = 0
    End If
    Set AskChoiceCol = col
End Function

Private Function AskLongText(ByVal c As Long) As Boolean
    Dim r As Long, s As String
    For r = mAskHdr + 1 To mAskLast
        s = T_(mAskSheet.Cell(r, c))
        If InStr(s, vbLf) > 0 Or Len(s) > 200 Then AskLongText = True: Exit Function
    Next
End Function

Private Function AskAllKind(ByVal c As Long, ByVal kind As String) As Boolean
    Dim r As Long, s As String, n As Long, msg As String
    For r = mAskHdr + 1 To mAskLast
        If Not IsBlank(mAskSheet.Cell(r, c)) And Not IsAskMissing(mAskSheet.Cell(r, c)) Then
            s = T_(mAskSheet.Cell(r, c))
            If kind = FT_EMAIL And InStr(s, "@") = 0 Then Exit Function
            If Not mValid.KindFormatOk(kind, s, msg) Then Exit Function
            n = n + 1
        End If
    Next
    AskAllKind = (n > 0)
End Function

Private Function AskMaxDecimals(ByVal c As Long) As Long
    Dim r As Long, v As Variant, d As Double, dc As Long
    For r = mAskHdr + 1 To mAskLast
        v = mAskSheet.Cell(r, c)
        dc = 0
        If IsNumVal(v) Then
            dc = Decs(CDbl(v))
        ElseIf VarType(v) = vbString Then
            If TextNumber(CStr(v), d, False) Then dc = Decs(d)
        End If
        If dc > AskMaxDecimals Then AskMaxDecimals = dc
    Next
    If AskMaxDecimals > 4 Then AskMaxDecimals = 4
End Function

' Umkodierte fehlende Antworten (-99, -98, -97 usw.), auch als Text
Private Function IsAskMissing(ByVal v As Variant) As Boolean
    Dim d As Double
    If IsBlank(v) Then Exit Function
    If IsNumVal(v) Then
        d = CDbl(v)
    ElseIf VarType(v) = vbString Then
        If Not TextNumber(CStr(v), d, False) Then Exit Function
    Else
        Exit Function
    End If
    IsAskMissing = (d <= -90 And d >= -999 And d = Fix(d))
End Function

Private Function AskCodeText(ByVal v As Variant) As String
    Dim d As Double
    If IsNumVal(v) Then
        AskCodeText = CStr(CLng(CDbl(v)))
    ElseIf VarType(v) = vbString And TextNumber(CStr(v), d, False) Then
        AskCodeText = CStr(CLng(d))
    Else
        AskCodeText = T_(v)
    End If
End Function

Private Sub PlanTwoAsk()
    Dim c As Long, key As String, col As cLiCol, cp As Variant, info As cMap, k As Variant, qText As String
    Dim grp As New cMap, gcol As cLiCol, gk As String, nm As String, qList As New cMap, isBip As Boolean
    Dim keyInfo As New cMap, opts As cMap, rowLbl As String, optLbl As String, groupLbl As String, dtype As String, sb As String
    Dim qKey As String, qKind As String
    Set liAskQuestions = New Collection
    ' Codeplan je Variable zusammenfassen
    For Each cp In mCp
        key = CStr(cp(0))
        If keyInfo.Exists(key) Then
            Set info = keyInfo.GetVal(key)
        Else
            Set info = New cMap
            info.SetVal "type", "": info.SetVal "sub", "": info.SetVal "dtype", "": info.SetVal "label", ""
            info.SetVal "frage", CStr(cp(6))
            Set opts = New cMap
            info.SetVal "opts", opts
            keyInfo.SetVal key, info
        End If
        Select Case CStr(cp(1))
            Case "standard_field"
                info.SetVal "type", "meta": info.SetVal "dtype", CStr(cp(3))
            Case "question"
                info.SetVal "frage", CStr(cp(5))
                If Len(CStr(info.GetVal("dtype"))) = 0 Then info.SetVal "dtype", CStr(cp(3))
            Case "answer"
                info.SetVal "type", "answer": info.SetVal "sub", CStr(cp(2)): info.SetVal "dtype", CStr(cp(3)): info.SetVal "label", CStr(cp(5))
            Case "answer_option"
                If Len(CStr(info.GetVal("type"))) = 0 Then info.SetVal "type", "option"
                If CStr(info.GetVal("type")) = "option" Then info.SetVal "sub", CStr(cp(2))
                If Len(CStr(info.GetVal("dtype"))) = 0 Then info.SetVal "dtype", CStr(cp(3))
                Set opts = info.GetVal("opts")
                If Not IsBlank(cp(4)) Then opts.SetVal AskCodeText(cp(4)), CStr(cp(5))
            Case "text_entry"
                info.SetVal "type", "textopt": info.SetVal "dtype", "string"
        End Select
    Next
    ' Spalten in der Reihenfolge des Exports
    For c = 1 To mAskSheet.C
        key = T_(mAskSheet.Cell(mAskHdr, c))
        If Len(key) = 0 Then GoTo nextCol
        qKey = AskQuestionKey(key)
        qKind = CStr(mQKind.GetVal(qKey, ""))
        If Not keyInfo.Exists(key) Then
            nm = T_(mAskSheet.Cell(Application.Max(1, mAskHdr - 3), c))
            If Len(nm) = 0 Then nm = key
            Set col = NewAskCol(key, nm, FT_TEXT, "code", c)
            col.Hint = "nicht im Codeplan": col.Unsure = True
            liCols.Add col
            GoTo nextCol
        End If
        Set info = keyInfo.GetVal(key)
        Set opts = info.GetVal("opts")
        qText = CStr(info.GetVal("frage"))
        sb = CStr(info.GetVal("sub"))
        dtype = CStr(info.GetVal("dtype"))
        Select Case CStr(info.GetVal("type"))
            Case "meta"
                Set col = NewAskCol(key, MetaName(key), FT_TEXT, "code", c)
                If dtype = "date" Then col.FType = FT_DATUMZEIT
                If dtype = "number" Then col.FType = FT_ZAHL: col.Decimals = AskMaxDecimals(c)
                col.Include = (key = "survey_session_id" Or key = "ended_at" Or (Left$(key, 8) = "contact." And AskColHasValue(c)))
                col.Hint = "Angabe zur Teilnahme"
                liCols.Add col
            Case "textopt"
                optLbl = AskSub(c, "option")
                If Len(optLbl) = 0 Then optLbl = "Sonstiges"
                Set col = NewAskCol(key, optLbl & " (Text)", FT_TEXT, "code", c)
                If AskLongText(c) Then col.FType = FT_LTEXT
                col.QKey = qKey
                col.Hint = "Zusatztext zu " & qKey
                liCols.Add col
                RegisterAskQuestion qList, qKey, qText, "Zusatztext", 1
            Case "option"
                Select Case sb
                    Case "checkbox"
                        rowLbl = AskSub(c, "answer")
                        If Len(rowLbl) > 0 Then gk = AskGroupKey(key) Else gk = qKey
                        If grp.Exists(gk) Then
                            Set gcol = grp.GetVal(gk)
                        Else
                            Set gcol = NewAskCol("M:" & gk, IIf(Len(rowLbl) > 0, rowLbl, qText), FT_MEHRFACH, "codemulti", 0)
                            gcol.QKey = qKey
                            gcol.Hint = AskKindName(sb, qKind) & " (" & qKey & ")"
                            grp.SetVal gk, gcol
                            liCols.Add gcol
                            RegisterAskQuestion qList, qKey, qText, AskKindName(sb, qKind), 1
                        End If
                        gcol.Src.Add c
                        optLbl = ""
                        For Each k In opts.Keys
                            optLbl = LiOptName(CStr(opts.GetVal(CStr(k))))
                            gcol.CodeMap.SetVal CStr(k) & "@" & c, optLbl
                            Exit For
                        Next
                        If Len(optLbl) = 0 Then optLbl = LiOptName(AskSub(c, "option"))
                        gcol.SrcNames.Add optLbl
                        If Not InColl(gcol.Opts, optLbl) Then gcol.Opts.Add optLbl
                    Case "input"
                        Set col = NewAskCol(key, AskSub(c, "answer") & ": " & AskSub(c, "option"), FT_TEXT, "code", c)
                        If AskLongText(c) Then col.FType = FT_LTEXT
                        col.QKey = qKey
                        col.Hint = AskKindName(sb, qKind) & " (" & qKey & ")"
                        liCols.Add col
                        RegisterAskQuestion qList, qKey, qText, AskKindName(sb, qKind), 1
                    Case "drag_drop"
                        optLbl = AskSub(c, "option")
                        If Len(optLbl) = 0 Then
                            For Each k In opts.Keys
                                optLbl = CStr(opts.GetVal(CStr(k))): Exit For
                            Next
                        End If
                        Set col = NewAskCol(key, "Rang: " & optLbl, FT_ZAHL, "code", c)
                        col.Decimals = 0
                        col.QKey = qKey
                        col.Hint = "Rangfolge (" & qKey & ")"
                        liCols.Add col
                        RegisterAskQuestion qList, qKey, qText, "Rangfolge", 1
                    Case Else
                        Set col = AskChoiceCol(key, qText, opts, c, isBip)
                        col.QKey = qKey
                        col.Hint = AskKindName(sb, qKind) & " (" & qKey & ")"
                        liCols.Add col
                        RegisterAskQuestion qList, qKey, qText, AskKindName(sb, qKind), 1
                End Select
            Case "answer"
                rowLbl = CStr(info.GetVal("label"))
                groupLbl = AskSub(c, "group")
                If Len(groupLbl) > 0 Then rowLbl = rowLbl & ": " & groupLbl
                Select Case sb
                    Case "radio", "select", "select_coded"
                        Set col = AskChoiceCol(key, rowLbl, opts, c, isBip)
                        If isBip Then
                            nm = AskSub(c, "answer")
                            If Len(nm) > 0 Then col.Name = CleanName(nm)
                            col.Hint = "Skala zwischen zwei Gegensätzen (" & qKey & ")"
                        End If
                    Case "input"
                        Set col = NewAskCol(key, rowLbl, FT_TEXT, "code", c)
                        If AskLongText(c) Then col.FType = FT_LTEXT
                        If AskAllKind(c, FT_EMAIL) Then col.FType = FT_EMAIL
                        If PhoneHeader(rowLbl) Then
                            If AskAllKind(c, FT_TELEFON) Then col.FType = FT_TELEFON
                        End If
                        If col.FType = FT_TEXT Then AskInferType col, c
                    Case "range", "number"
                        Set col = NewAskCol(key, rowLbl, FT_ZAHL, "code", c)
                        col.Decimals = AskMaxDecimals(c)
                    Case Else
                        Set col = NewAskCol(key, rowLbl, IIf(dtype = "number", FT_ZAHL, FT_TEXT), "code", c)
                        If col.FType = FT_ZAHL Then col.Decimals = AskMaxDecimals(c)
                End Select
                col.QKey = qKey
                If Len(col.Hint) = 0 Then col.Hint = AskKindName(sb, qKind) & " (" & qKey & ")"
                liCols.Add col
                RegisterAskQuestion qList, qKey, qText, AskKindName(sb, qKind), 1
            Case Else
                ' Eingabefelder (text_entry) und Fragen ohne Antwortzeilen
                Select Case dtype
                    Case "number"
                        Set col = NewAskCol(key, qText, FT_ZAHL, "code", c)
                        col.Decimals = AskMaxDecimals(c)
                    Case "localdate"
                        Set col = NewAskCol(key, qText, FT_DATUM, "code", c)
                    Case "date"
                        Set col = NewAskCol(key, qText, FT_DATUMZEIT, "code", c)
                    Case "localtime"
                        Set col = NewAskCol(key, qText, FT_TEXT, "code", c)
                        col.Hint = "Uhrzeit als Text (" & qKey & ")"
                    Case Else
                        Set col = NewAskCol(key, qText, FT_TEXT, "code", c)
                        If AskLongText(c) Then col.FType = FT_LTEXT
                        If AskAllKind(c, FT_EMAIL) Then col.FType = FT_EMAIL
                        If col.FType = FT_TEXT Then AskInferType col, c
                End Select
                col.QKey = qKey
                If Len(col.Hint) = 0 Then col.Hint = "Eingabefeld (" & qKey & ")"
                liCols.Add col
                RegisterAskQuestion qList, qKey, qText, AskKindName("", "text_entry"), 1
        End Select
nextCol:
    Next
    PlanAskQuestions
    ' gleiche Namen (z. B. Matrixzeilen) mit der Fragenummer unterscheiden
    AskUniqueNames
    For Each k In qList.Keys
        liAskQuestions.Add qList.GetVal(CStr(k))
    Next
    ApplyOverrides
    For Each col In liCols
        col.Sample = AskSampleCol(col)
    Next
End Sub

Private Sub AskUniqueNames()
    Dim col As cLiCol, cnt As New cMap
    For Each col In liCols: cnt.AddNum col.Name, 1: Next
    For Each col In liCols
        If CDbl(cnt.GetVal(col.Name)) > 1 And Len(col.QKey) > 0 Then col.Name = col.Name & " (" & col.QKey & ")"
    Next
End Sub

Private Function AskSampleCol(ByVal col As cLiCol) As String
    Dim r As Long, v As Variant, s As String, n As Long, why As String, raw As String
    For r = mAskHdr + 1 To mAskLast
        v = AskFieldValue(col, r, why, raw)
        If Not IsBlank(v) Then
            s = s & IIf(Len(s) > 0, ", ", "") & Truncate(Replace(T_(v), ";", ", "), 25)
            n = n + 1
            If n >= 3 Then Exit For
        End If
    Next
    AskSampleCol = s
End Function

Private Sub PlanAskQuestions()
    Dim col As cLiCol, bips As New cMap, k As Variant, c As Long, r As Long
    For Each col In liCols
        If Left$(col.Tag, 4) = "bip:" Then
            If Not bips.Exists(col.Tag) Then bips.SetVal col.Tag, col.Name
        End If
    Next
    For Each k In bips.Keys
        AddQ CStr(k), "Skala " & Mid$(CStr(k), 5), "Frage " & Mid$(CStr(k), 5) & " ist eine Skala zwischen zwei Gegensätzen, z. B. „" & bips.GetVal(CStr(k)) & _
             "“. Die Stufen haben keinen eigenen Text.", _
             Array("Als Nummer übernehmen (Stufe 1, 2, 3 …), gut für Mittelwerte", "Als Auswahl mit den Stufen [1], [2], [3] … übernehmen"), 0
    Next
    If mAskPretest > 0 And mAskPretest < liAskRows Then
        AddQ "pretest", "Pretest", Plural(mAskPretest, "Fragebogen stammt", "Fragebögen stammen") & " aus dem Pretest (Probedurchläufe vor dem Start der Befragung).", _
             Array("Pretest-Fragebögen weglassen (empfohlen)", "Mit übernehmen"), 0
    End If
    mAskMissing = 0
    For c = 1 To mAskSheet.C
        For r = mAskHdr + 1 To mAskLast
            If IsAskMissing(mAskSheet.Cell(r, c)) Then mAskMissing = mAskMissing + 1
        Next
    Next
End Sub

' Wert eines 2ask-Felds fuer eine Ergebniszeile
Private Function AskFieldValue(ByVal col As cLiCol, ByVal r As Long, ByRef why As String, ByRef raw As String) As Variant
    Dim v As Variant, c As Long, i As Long, code As String, s As String, nm As String, outV As Variant
    why = "": raw = ""
    If col.Mode = "codemulti" Then
        For i = 1 To col.Src.Count
            c = CLng(col.Src(i))
            v = mAskSheet.Cell(r, c)
            If Not IsBlank(v) And Not IsAskMissing(v) Then
                code = AskCodeText(v)
                nm = CStr(col.CodeMap.GetVal(code & "@" & c, ""))
                If Len(nm) = 0 Then nm = CStr(col.SrcNames(i))
                s = ListAdd(s, nm)
            End If
        Next
        If Len(s) = 0 Then Exit Function
        Select Case col.FType
            Case FT_MEHRFACH: AskFieldValue = s
            Case FT_AUSWAHL, FT_STATUS
                AskFieldValue = Split(s, ";")(0)
                If InStr(s, ";") > 0 Then why = "übernommen; von mehreren Antworten nur die erste": raw = Replace(s, ";", ", ")
            Case Else: AskFieldValue = Replace(s, ";", ", ")
        End Select
        Exit Function
    End If
    c = col.FirstSrc
    If c = 0 Then Exit Function
    v = mAskSheet.Cell(r, c)
    If IsBlank(v) Or IsAskMissing(v) Then Exit Function
    raw = T_(v)
    If col.CodeMap.Count > 0 Then
        code = AskCodeText(v)
        ' unkodierter Export: Antworttext statt Code ("sehr zufrieden" statt 1)
        If Not col.CodeMap.Exists(code) Then code = AskCodeOfLabel(col, T_(v))
        If Len(code) = 0 Then why = "Code " & AskCodeText(v) & " steht nicht im Codeplan": Exit Function
        Select Case col.FType
            Case FT_ZAHL, FT_GELD, FT_PROZENT: AskFieldValue = CDbl(Val(code))
            Case Else: AskFieldValue = CStr(col.CodeMap.GetVal(code))
        End Select
        Exit Function
    End If
    Select Case col.FType
        Case FT_TEXT, FT_LTEXT: AskFieldValue = T_(v)
        Case FT_AUSWAHL, FT_STATUS, FT_MEHRFACH: AskFieldValue = LiOptName(T_(v))
        Case Else
            If ConvertTyped(v, col.FType, False, outV, why) Then AskFieldValue = outV
    End Select
End Function

' Code zu einem Antworttext des Codeplans ("" = unbekannt)
Private Function AskCodeOfLabel(ByVal col As cLiCol, ByVal lbl As String) As String
    Dim k As Variant, t As String
    t = LiOptName(lbl)
    If Len(t) = 0 Then Exit Function
    For Each k In col.CodeMap.Keys
        If StrComp(CStr(col.CodeMap.GetVal(CStr(k))), t, vbTextCompare) = 0 Then AskCodeOfLabel = CStr(k): Exit Function
    Next
End Function

' Eingabefeld ohne Typangabe im Codeplan (Formular, Textfeld): Datum oder Zahl aus den Antworten erkennen
Private Sub AskInferType(ByVal col As cLiCol, ByVal c As Long)
    Dim vals As New Collection, r As Long, p As cMap, typ As String
    For r = mAskHdr + 1 To mAskLast
        If Not IsAskMissing(mAskSheet.Cell(r, c)) Then vals.Add mAskSheet.Cell(r, c)
    Next
    Set p = ProfileValues(vals, col.Name, "", False)
    If CLng(p.GetVal("n", 0)) < 2 Then Exit Sub
    typ = CStr(p.GetVal("typ"))
    Select Case typ
        Case FT_DATUM, FT_DATUMZEIT, FT_ZAHL, FT_GELD
            col.FType = typ
            If typ = FT_ZAHL Then col.Decimals = CLng(p.GetVal("decimals", 0))
            If p.GetVal("ungueltig").Count > 0 Then col.Unsure = True: col.Hint = "Typ aus den Antworten erkannt; einzelne Werte passen nicht"
    End Select
End Sub

' Zusammenfassung fuer den Assistenten
Public Function LiAskSummary() As String
    Dim nF As Long, col As cLiCol, s As String
    For Each col In liCols
        If col.Include Then nF = nF + 1
    Next
    s = "2ask-Export mit Codeplan: " & Plural(liAskQuestions.Count, "Frage", "Fragen") & ", " & Plural(liAskRows, "Fragebogen", "Fragebögen")
    If mAskPretest > 0 Then s = s & " (davon " & mAskPretest & " aus dem Pretest)"
    s = s & ". Daraus entstehen " & Plural(nF, "Feld", "Felder") & "; von den Angaben zur Teilnahme sind nur Teilnahme-ID und Ende vorgewählt."
    If mAskMissing > 0 Then s = s & " " & Plural(mAskMissing, "umkodierte fehlende Antwort", "umkodierte fehlende Antworten") & " (z. B. -99) " & _
                                IIf(mAskMissing = 1, "wird", "werden") & " leer übernommen."
    LiAskSummary = s
End Function

'==============================================================================
' Testlauf (Entwicklung): Datei lesen, Vorschlag ausgeben, optional uebernehmen
'==============================================================================
' Sichtpruefung (Entwicklung): Assistent nicht-modal oeffnen und per Befehl bedienen, damit Fensteraufnahmen
' jedes Schritts entstehen koennen. Befehle: "open" | "set:<Steuerelement>=<Wert>" | "idx:<Steuerelement>=<n>" |
' "click:<Schluessel>" | "quiet:1|0" | "close". Liefert den sichtbaren Schritttext.
Public Function LiUi(ByVal cmd As String) As String
    Dim p As Long, nm As String, v As String, c As Object
    On Error GoTo failed
    Select Case True
        Case cmd = "open"
            If Not mUiForm Is Nothing Then Unload mUiForm
            Set mUiForm = New frmListImport
            mUiForm.Setup
            mUiForm.Show vbModeless
        Case cmd = "close"
            If Not mUiForm Is Nothing Then Unload mUiForm
            Set mUiForm = Nothing
            LiUi = "geschlossen": Exit Function
        Case Left$(cmd, 6) = "quiet:"
            mUI.gQuiet = (Mid$(cmd, 7) = "1")
            mUI.gQuietLog = ""
        Case Left$(cmd, 4) = "set:", Left$(cmd, 4) = "idx:", Left$(cmd, 4) = "sel:"
            p = InStr(cmd, "=")
            nm = Mid$(cmd, 5, p - 5): v = Mid$(cmd, p + 1)
            Set c = mUiForm.Controls(nm)
            If Left$(cmd, 4) = "idx:" Then
                c.ListIndex = CLng(v)
            ElseIf Left$(cmd, 4) = "sel:" Then
                c.Selected(CLng(v)) = Not c.Selected(CLng(v))
            ElseIf TypeName(c) = "TextBox" Then
                c.Text = v
            Else
                c.Value = (v = "1" Or LCase$(v) = "wahr" Or LCase$(v) = "true")
            End If
        Case Left$(cmd, 6) = "click:"
            mUiForm.HandleEvent Mid$(cmd, 7), "Click"
    End Select
    ' sofort neu zeichnen, damit die Aufnahme nicht einen halb gezeichneten Zustand zeigt
    mUiForm.Repaint
    DoEvents
    LiUi = mUiForm.Controls("step").Caption & IIf(Len(mUI.gQuietLog) > 0, " | " & mUI.gQuietLog, "")
    Exit Function
failed:
    LiUi = "FEHLER: " & Err.Description
End Function

' Laufzeitmessung (Entwicklung): Lesen, Auswerten, Plan und Pruefung einer Liste je Arbeitsschritt
Public Function LiProfile(ByVal path As String) As String
    Dim plan As cMap, errs As Collection, valid As Collection, upd As Long, k As Variant, lbl As String, t00 As Double
    On Error GoTo failed
    mUI.gQuiet = True
    mProfOn = True: mProfLog = "": Set mProfAgg = New cMap: mProfT = Timer: t00 = Timer
    LiReadFile path, ""
    LiChooseTable LiBestTable()
    For Each k In mProfAgg.Keys
        mProfLog = mProfLog & "  davon " & k & ": " & Format$(CDbl(mProfAgg.GetVal(k)), "0.00") & " s" & vbLf
    Next
    mProfT = Timer
    lbl = LiSuggestLabel()
    Prof "Bezeichnungsfeld vorschlagen"
    Set plan = LiBuildPlan("Messung", "", "", lbl)
    Prof "Importplan bauen"
    Set errs = mIO.CheckImport(plan, valid, upd)
    Prof "Prüfung (CheckImport)"
    Call mIO.ExecuteImport(plan, valid)
    Prof "Übernahme (ExecuteImport mit Sicherung)"
    mProfLog = mProfLog & "gesamt: " & Format$(Timer - t00, "0.00") & " s"
    mProfOn = False
    mUI.gQuiet = False
    LiProfile = mProfLog
    Exit Function
failed:
    mProfOn = False
    mUI.gQuiet = False
    LiProfile = mProfLog & "FEHLER: " & Err.Description
End Function

' Aufruf mit einem Argument "Pfad|Tabelle|Uebernehmen|Ergebnisdatei" (fuer build\listimport_test.ps1);
' das Ergebnis wird als UTF-8-Datei geschrieben
Public Function LiTestCase(ByVal spec As String) As String
    Dim a As Variant, res As String
    a = Split(spec, "|")
    If UBound(a) < 3 Then LiTestCase = "FEHLER: erwartet Pfad|Tabelle|Uebernehmen|Ergebnisdatei": Exit Function
    res = LiTestFile(CStr(a(0)), a(1), a(2))
    WriteUtf8File CStr(a(3)), res, True
    LiTestCase = IIf(InStr(res, "FEHLER:") > 0, "FEHLER", "OK")
End Function

' Entwicklung: alle Tabellen "Test ..." mit Feldern (Typ, Optionen) und ersten Eintraegen in eine UTF-8-Datei schreiben
Public Function LiDumpTests(ByVal outFile As String) As String
    Dim ids As Variant, i As Long, s As String
    ids = mCfg.TableIds(True)
    For i = 0 To ArrLen(ids) - 1
        If Left$(mCfg.TableName(CStr(ids(i))), 5) = "Test " Then s = s & LiDumpTable(CStr(ids(i)), 4) & vbLf
    Next
    WriteUtf8File outFile, s, True
    LiDumpTests = "OK"
End Function

Public Function LiDumpTable(ByVal tid As String, ByVal maxRows As Long) As String
    Dim s As String, f As cField, t As cTableData, r As Long, v As Variant, ln As String, oid As Variant, ops As String
    Set t = mData.TData(tid)
    s = "Tabelle " & mCfg.TableName(tid) & " (" & t.RowCount & " Einträge, Bezeichnungsfeld " & mCfg.FieldName(mCfg.LabelField(tid)) & ")" & vbLf
    For Each f In mCfg.Fields(tid)
        ops = ""
        If f.IsOptionType Then
            For Each oid In mCfg.OptionIds(f.ID)
                ops = ops & IIf(Len(ops) > 0, ", ", "") & mCfg.OptName(CStr(oid))
            Next
        End If
        s = s & "  " & f.Name & " [" & FieldTypeName(f.Kind) & IIf(f.FType = FT_ZAHL, ", " & f.Decimals & " Nachkommastellen", "") & "]" & _
            IIf(Len(ops) > 0, " Optionen: " & ops, "") & vbLf
    Next
    For r = 1 To t.RowCount
        If r > maxRows Then Exit For
        ln = ""
        For Each f In mCfg.Fields(tid)
            v = t.Cell(r, t.ColOf(f.ID))
            ln = ln & IIf(Len(ln) > 0, " | ", "") & DumpVal(f, v)
        Next
        s = s & "  " & r & ": " & ln & vbLf
    Next
    LiDumpTable = s
End Function

Private Function DumpVal(ByVal f As cField, ByVal v As Variant) As String
    Dim a As Variant, i As Long, s As String
    If IsBlank(v) Then Exit Function
    If f.IsOptionType Then
        a = Split(CStr(v), ";")
        For i = 0 To UBound(a)
            s = s & IIf(Len(s) > 0, " + ", "") & mCfg.OptName(CStr(a(i)))
        Next
        DumpVal = s
    Else
        DumpVal = T_(v)
    End If
End Function

Public Function LiTestFile(ByVal path As String, ByVal tableArg As Variant, ByVal importArg As Variant) As String
    Dim s As String, i As Long, a As Variant, col As cLiCol, q As cLiQ, plan As cMap, errs As Collection, valid As Collection, upd As Long, tid As String
    Dim nm As String, lbl As String, tableIdx As Long, doImport As Boolean, t0 As Double, tPlan As Double, tCheck As Double, tImp As Double
    tableIdx = CLng(Val(CStr(tableArg)))
    doImport = (CStr(importArg) = "1" Or LCase$(CStr(importArg)) = "true" Or LCase$(CStr(importArg)) = "wahr")
    On Error GoTo failed
    mUI.gQuiet = True
    t0 = Timer
    LiReadFile path, ""
    s = "Modus=" & liMode & vbLf
    If liMode = "2ask" Then
        LiPlan
        s = s & LiAskSummary() & vbLf
    Else
        For i = 1 To liTables.Count
            a = LiTableRow(i)
            s = s & "Tabelle " & i & ": " & a(0) & " " & a(1) & " Einträge=" & a(2) & " Spalten=" & a(3) & " (" & a(4) & ")" & vbLf
        Next
        If liTables.Count = 0 Then LiTestFile = s & "keine Tabelle": Exit Function
        If tableIdx < 1 Or tableIdx > liTables.Count Then tableIdx = LiBestTable()
        LiChooseTable tableIdx
        s = s & "gewählt: Tabelle " & tableIdx & vbLf
    End If
    For Each q In liQs
        s = s & "Frage " & q.ID & ": " & q.Text & " -> " & q.Choices(q.Answer) & vbLf
    Next
    For Each col In liCols
        s = s & IIf(col.Include, "Feld ", "(weg) ") & col.Name & " [" & FieldTypeName(col.FType) & "]" & _
            IIf(col.Opts.Count > 0 And col.IsOptionType, " Optionen=" & col.Opts.Count, "") & IIf(Len(col.Sample) > 0, " z. B. " & col.Sample, "") & vbLf
    Next
    lbl = LiSuggestLabel()
    tPlan = Timer - t0
    s = s & "Datensätze=" & LiRecordCount() & " Bezeichnung=" & lbl & vbLf
    s = s & "Zeit: lesen und auswerten " & Format$(tPlan, "0.0") & " s" & vbLf
    If doImport Then
        nm = "Test " & Format$(Now, "hhmmss") & " " & Left$(Replace(FileNameOf(path), ".", " "), 30)
        t0 = Timer
        Set plan = LiBuildPlan(nm, "", "", lbl)
        Set errs = mIO.CheckImport(plan, valid, upd)
        tCheck = Timer - t0
        s = s & "Prüfung: gültig=" & valid.Count & " Fehler=" & errs.Count & " Hinweise=" & liNotes.Count & " (" & Format$(tCheck, "0.0") & " s)" & vbLf
        For i = 1 To errs.Count
            If i > 5 Then Exit For
            s = s & "  Fehler " & LiSourceRow(CLng(errs(i)(0))) & " " & errs(i)(1) & ": " & errs(i)(2) & " (" & errs(i)(3) & ")" & vbLf
        Next
        For i = 1 To liNotes.Count
            If i > 5 Then Exit For
            s = s & "  Hinweis " & liNotes(i)(0) & " " & liNotes(i)(1) & ": " & liNotes(i)(2) & " (" & liNotes(i)(3) & ")" & vbLf
        Next
        t0 = Timer
        tid = mIO.ExecuteImport(plan, valid)
        tImp = Timer - t0
        If Len(tid) > 0 Then s = s & "Übernommen: Tabelle " & mCfg.TableName(tid) & " mit " & mData.TData(tid).RowCount & " Einträgen (" & _
                                 Format$(tImp, "0.0") & " s inklusive Sicherung)" & vbLf
    End If
    mUI.gQuiet = False
    LiTestFile = s
    Exit Function
failed:
    mUI.gQuiet = False
    mApp.ResetWork
    LiTestFile = s & "FEHLER: " & Err.Description
End Function
