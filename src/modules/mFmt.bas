Attribute VB_Name = "mFmt"
'==============================================================================
' mFmt - Darstellung von Feldwerten (deutsche Konventionen), Zellformate,
' Eingabetexte fuer Formulare.
'==============================================================================
Option Explicit

Public Function DisplayText(ByVal f As cField, ByVal v As Variant, Optional ByVal rid As String = "") As String
    Dim vt As String, a As Variant, i As Long, s As String, tt As String
    If f Is Nothing Then DisplayText = NzS(v): Exit Function
    If f.IsLinkType Then
        If Len(rid) = 0 Then Exit Function
        DisplayText = LinkLabels(f, mData.LinkTargets(f.ID, rid))
        Exit Function
    End If
    If IsError(v) Then DisplayText = "Fehler": Exit Function
    If IsBlank(v) Then Exit Function
    If IsErrVal(v) Then DisplayText = "Fehler: " & Mid$(CStr(v), Len(ERR_PREFIX) + 1): Exit Function
    vt = f.ValueType
    On Error GoTo fallback
    Select Case vt
        Case FT_TEXT, FT_LTEXT, FT_PFAD
            DisplayText = CStr(v)
            If f.Kind = FT_DATEIEN Then DisplayText = FileNames(CStr(v))
        Case FT_ZAHL
            If IsYearLike(f) Then
                DisplayText = Format$(CDbl(v), "0")
            ElseIf f.FType = FT_FORMEL And f.Decimals = 0 Then
                ' wie in der Tabelle: ganze Zahlen ohne, sonst mit zwei Nachkommastellen
                DisplayText = FmtNum(CDbl(v), IIf(CDbl(v) = Fix(CDbl(v)), 0, 2))
            Else
                DisplayText = FmtNum(CDbl(v), f.Decimals)
            End If
        Case FT_GELD
            DisplayText = FmtNum(CDbl(v), 2) & " €"
        Case FT_PROZENT
            DisplayText = FmtNum(CDbl(v) * 100, IIf(f.FType = FT_FORMEL And f.Decimals = 0, 1, f.Decimals), f.FType = FT_FORMEL) & " %"
        Case FT_DATUM
            DisplayText = FmtDate(v)
        Case FT_DATUMZEIT
            DisplayText = FmtDateTime(v)
        Case FT_JANEIN
            If f.Kind = FT_KONTROLL Then
                DisplayText = IIf(CBool(v), ChrW(&H2611), ChrW(&H2610))
            Else
                DisplayText = IIf(CBool(v), "Ja", "Nein")
            End If
        Case FT_AUSWAHL, FT_STATUS
            DisplayText = mCfg.OptName(CStr(v))
            If Len(DisplayText) = 0 Then DisplayText = "(unbekannte Option)"
        Case FT_MEHRFACH
            a = Split(CStr(v), ";")
            For i = 0 To UBound(a)
                If Len(a(i)) > 0 Then s = s & IIf(Len(s) > 0, ", ", "") & mCfg.OptName(CStr(a(i)))
            Next
            DisplayText = s
        Case Else
            DisplayText = CStr(v)
    End Select
    Exit Function
fallback:
    DisplayText = CStr(v)
End Function

Public Function LinkLabels(ByVal f As cField, ByVal targets As String, Optional ByVal maxN As Long = 0) As String
    Dim a As Variant, i As Long, s As String, tt As String, n As Long
    If Len(targets) = 0 Then Exit Function
    If f.FType = FT_LINK Then tt = f.TargetTable Else tt = mCfg.Field(f.PairField).TableID
    a = Split(targets, ";")
    For i = 0 To UBound(a)
        If Len(a(i)) > 0 Then
            n = n + 1
            If maxN > 0 And n > maxN Then
                s = s & " +" & (UBound(a) - i + 1)
                Exit For
            End If
            s = s & IIf(Len(s) > 0, ", ", "") & LinkLabel(tt, CStr(a(i)))
        End If
    Next
    LinkLabels = s
End Function

' Bezeichnung eines verknuepften Eintrags; zwischengespeichert, solange sich weder Daten
' noch Struktur aendern (Tabellen mit vielen Verknuepfungen bauen so deutlich schneller auf)
Private Function LinkLabel(ByVal tt As String, ByVal rid As String) As String
    Static cache As Collection, rev As String
    Dim cur As String
    cur = mApp.gDataRev & "|" & mApp.gCfgRev
    If cache Is Nothing Or rev <> cur Then Set cache = New Collection: rev = cur
    On Error Resume Next
    LinkLabel = cache(tt & "|" & rid)
    If Err.Number = 0 Then Exit Function
    Err.Clear
    On Error GoTo 0
    LinkLabel = mData.RecordLabel(tt, rid)
    If mData.IsArchivedRec(tt, rid) Then LinkLabel = LinkLabel & " (archiviert)"
    cache.Add LinkLabel, tt & "|" & rid
End Function

Public Function FileNames(ByVal s As String) As String
    Dim a As Variant, i As Long, r As String
    a = Split(s, "|")
    For i = 0 To UBound(a)
        If Len(Trim$(a(i))) > 0 Then r = r & IIf(Len(r) > 0, ", ", "") & Mid$(a(i), InStrRev(a(i), "\") + 1)
    Next
    FileNames = r
End Function

' Wert fuer eine Tabellenzelle (typisiert, damit Excel Zahlen/Daten korrekt zeigt)
Public Function CellValue(ByVal f As cField, ByVal v As Variant, ByVal rid As String) As Variant
    Dim vt As String
    If f.IsLinkType Then CellValue = DisplayText(f, Empty, rid): Exit Function
    If f.Kind = FT_KONTROLL And IsBlank(v) Then CellValue = ChrW(&H2610): Exit Function
    If f.Kind = FT_BUTTON Then CellValue = ChrW(&H25B6) & " " & f.Name: Exit Function
    If IsBlank(v) Then CellValue = Empty: Exit Function
    If IsErrVal(v) Or IsError(v) Then CellValue = DisplayText(f, v, rid): Exit Function
    vt = f.ValueType
    Select Case vt
        Case FT_ZAHL, FT_GELD, FT_PROZENT
            If IsNumeric(v) Then CellValue = CDbl(v) Else CellValue = CStr(v)
        Case FT_DATUM, FT_DATUMZEIT
            If IsDate(v) Or IsNumeric(v) Then CellValue = CDate(v) Else CellValue = CStr(v)
        Case FT_TEXT, FT_LTEXT, FT_PFAD
            CellValue = CStr(v)
        Case Else
            CellValue = DisplayText(f, v, rid)
    End Select
End Function

' allWhole: alle angezeigten Werte sind ganze Zahlen. Berechnungen ohne feste
' Kommastellen zeigen dann keine Nachkommastellen, sonst einheitlich zwei
' (Excel kann ein Format "#,##0.##" nicht ohne haengendes Komma darstellen).
' Jahreszahlen ohne Tausenderpunkt: Ganzzahl mit Grenzen im Bereich 1000-9999 (z. B. Baujahr mit Regel)
' oder ein Feldname aus einem Wort auf "jahr" (Baujahr, Pflanzjahr, Jahr), etwa nach "Liste übernehmen"
Public Function IsYearLike(ByVal f As cField) As Boolean
    If f.FType <> FT_ZAHL Or f.Decimals <> 0 Then Exit Function
    If YearName(f.Name) Then IsYearLike = True: Exit Function
    If IsEmpty(f.MinVal) Or IsEmpty(f.MaxVal) Then Exit Function
    On Error Resume Next
    IsYearLike = (CDbl(f.MinVal) >= 1000 And CDbl(f.MaxVal) <= 9999)
End Function

Private Function YearName(ByVal nm As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(nm))
    YearName = (InStr(s, " ") = 0 And s Like "*jahr")
End Function

' Hinweis zu einem Jahresfeld: "Jahreszahl von 1980 bis 2030" (Grenzen nur, soweit gesetzt)
Public Function YearHint(ByVal f As cField, Optional ByVal between As Boolean = False) As String
    Dim lo As Boolean, hi As Boolean
    lo = Not IsEmpty(f.MinVal): hi = Not IsEmpty(f.MaxVal)
    YearHint = "Jahreszahl"
    If lo And hi Then
        YearHint = YearHint & IIf(between, " zwischen ", " von ") & Format$(f.MinVal, "0") & IIf(between, " und ", " bis ") & Format$(f.MaxVal, "0")
    ElseIf lo Then
        YearHint = YearHint & " ab " & Format$(f.MinVal, "0")
    ElseIf hi Then
        YearHint = YearHint & " bis " & Format$(f.MaxVal, "0")
    End If
End Function

Public Function NumberFormatFor(ByVal f As cField, Optional ByVal allWhole As Boolean = True) As String
    Select Case f.ValueType
        Case FT_ZAHL
            If IsYearLike(f) Then
                NumberFormatFor = "0"
            ElseIf f.Decimals > 0 Then
                NumberFormatFor = "#,##0." & String$(f.Decimals, "0")
            ElseIf f.FType = FT_FORMEL Then
                NumberFormatFor = IIf(allWhole, "#,##0", "#,##0.00")
            Else
                NumberFormatFor = "#,##0"
            End If
        Case FT_GELD: NumberFormatFor = "#,##0.00 [$€-407]"
        Case FT_PROZENT
            If f.Decimals > 0 Then NumberFormatFor = "0." & String$(f.Decimals, "0") & "%" Else NumberFormatFor = "0%"
            If f.FType = FT_FORMEL And f.Decimals = 0 Then NumberFormatFor = "0.0%"
        Case FT_DATUM: NumberFormatFor = "dd.mm.yyyy"
        Case FT_DATUMZEIT: NumberFormatFor = "dd.mm.yyyy hh:mm"
        Case Else: NumberFormatFor = "@"
    End Select
End Function

' Text fuer Formular-Eingabefelder
Public Function InputText(ByVal f As cField, ByVal v As Variant) As String
    If IsBlank(v) Then Exit Function
    If IsErrVal(v) Then InputText = DisplayText(f, v): Exit Function
    On Error GoTo fallback
    Select Case f.ValueType
        Case FT_ZAHL
            If IsYearLike(f) Then InputText = Format$(CDbl(v), "0"): Exit Function
            InputText = FmtNum(CDbl(v), IIf(f.Decimals > 0, f.Decimals, 6), True)
        Case FT_GELD
            InputText = FmtNum(CDbl(v), 2)
        Case FT_PROZENT
            InputText = FmtNum(CDbl(v) * 100, 6, True)
        Case FT_DATUM
            InputText = FmtDate(v)
        Case FT_DATUMZEIT
            InputText = FmtDateTime(v)
        Case Else
            InputText = DisplayText(f, v)
    End Select
    Exit Function
fallback:
    InputText = CStr(v)
End Function

' Ab 1.5 stehen alle Werte linksbündig (Listen, Formulare, Erfassungsdateien), einheitliches Bild
Public Function IsRightAligned(ByVal f As cField) As Boolean
    IsRightAligned = False
End Function

' Kurze, verstaendliche Typbeschreibung fuer Formulare
Public Function InputHint(ByVal f As cField) As String
    Select Case f.ValueType
        Case FT_ZAHL
            InputHint = IIf(f.Decimals = 0, "Ganze Zahl", "Zahl mit bis zu " & f.Decimals & " Nachkommastellen")
            If IsYearLike(f) Then InputHint = YearHint(f): Exit Function
            If Not IsEmpty(f.MinVal) Then InputHint = InputHint & ", mindestens " & FmtNum(f.MinVal, 6, True)
            If Not IsEmpty(f.MaxVal) Then InputHint = InputHint & ", höchstens " & FmtNum(f.MaxVal, 6, True)
        Case FT_GELD: InputHint = "Betrag in Euro, z. B. 1.250,00"
        Case FT_PROZENT: InputHint = "Prozentwert, z. B. 25 für 25 %"
        Case FT_DATUM: InputHint = "TT.MM.JJJJ"
        Case FT_DATUMZEIT: InputHint = "TT.MM.JJJJ hh:mm"
    End Select
End Function
