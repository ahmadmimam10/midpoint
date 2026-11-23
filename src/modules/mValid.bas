Attribute VB_Name = "mValid"
'==============================================================================
' mValid - Gemeinsame Pruef- und Umwandlungsregeln fuer ALLE Schreibwege:
' Formular, direkte Zellbearbeitung, Mehrzellen-Einfuegen, Sammelaenderung, Import.
'==============================================================================
Option Explicit

' Text einer Eingabe in den gespeicherten Wert umwandeln.
' allowInactiveOid: bisheriger Wert darf eine deaktivierte Option sein.
Public Function ParseInput(ByVal f As cField, ByVal txt As String, ByRef outV As Variant, ByRef msg As String, _
                           Optional ByVal allowInactive As String = "") As Boolean
    Dim n As Double, d As Date, ht As Boolean, a As Variant, i As Long, oid As String, res As String, t As String
    msg = ""
    outV = Empty
    If IsAutoKind(f.Kind) Then msg = "Dieses Feld wird automatisch gesetzt.": Exit Function
    If f.FType <> FT_LTEXT Then txt = Trim$(txt)
    If Len(Trim$(txt)) = 0 Then
        If f.Kind = FT_KONTROLL Then outV = False
        ParseInput = True
        Exit Function
    End If
    Select Case f.FType
        Case FT_TEXT, FT_PFAD
            If InStr(txt, vbLf) > 0 Or InStr(txt, vbCr) > 0 Then txt = Replace(Replace(Replace(txt, vbCrLf, " "), vbLf, " "), vbCr, " ")
            outV = txt
            If Not KindFormatOk(f.Kind, CStr(txt), msg) Then Exit Function
        Case FT_LTEXT
            outV = txt
        Case FT_ZAHL, FT_GELD
            If Not ParseNumberDE(txt, n) Then
                msg = "Bitte eine Zahl eingeben (z. B. 1.234,5)."
                Exit Function
            End If
            outV = n
        Case FT_PROZENT
            If Not ParseNumberDE(txt, n) Then
                msg = "Bitte einen Prozentwert eingeben, z. B. 25 für 25 %."
                Exit Function
            End If
            outV = n / 100
        Case FT_DATUM, FT_DATUMZEIT
            If Not ParseDateDE(txt, d, ht) Then
                msg = "Bitte ein gültiges Datum im Format TT.MM.JJJJ eingeben" & IIf(f.FType = FT_DATUMZEIT, " (optional mit Uhrzeit hh:mm).", ".")
                Exit Function
            End If
            If f.FType = FT_DATUM And ht Then
                If d - Int(d) <> 0 Then
                    msg = "Dieses Feld speichert nur ein Datum ohne Uhrzeit."
                    Exit Function
                End If
            End If
            outV = d
        Case FT_JANEIN
            Select Case LCase$(txt)
                Case "ja", "j", "x", "wahr", "true", "1", "yes", ChrW(&H2611): outV = True
                Case "nein", "n", "falsch", "false", "0", "no", ChrW(&H2610): outV = False
                Case Else
                    msg = "Bitte Ja oder Nein angeben."
                    Exit Function
            End Select
        Case FT_AUSWAHL, FT_STATUS
            oid = mCfg.OptIdByName(f.ID, txt)
            If Len(oid) = 0 Then
                msg = """" & txt & """ ist keine Auswahlmöglichkeit dieses Feldes."
                Exit Function
            End If
            If Not mCfg.OptActive(oid) And oid <> allowInactive Then
                msg = """" & txt & """ ist für neue Eingaben deaktiviert."
                Exit Function
            End If
            outV = oid
        Case FT_MEHRFACH
            t = Replace(txt, ",", ";")
            a = Split(t, ";")
            For i = 0 To UBound(a)
                If Len(Trim$(a(i))) > 0 Then
                    oid = mCfg.OptIdByName(f.ID, Trim$(a(i)))
                    If Len(oid) = 0 Then
                        msg = """" & Trim$(a(i)) & """ ist keine Auswahlmöglichkeit dieses Feldes."
                        Exit Function
                    End If
                    If Not mCfg.OptActive(oid) And Not ListHas(allowInactive, oid) Then
                        msg = """" & Trim$(a(i)) & """ ist für neue Eingaben deaktiviert."
                        Exit Function
                    End If
                    res = ListAdd(res, oid)
                End If
            Next
            If Len(res) > 0 Then outV = res
        Case FT_LINK, FT_BACKLINK
            msg = "Verknüpfungen bitte über die Auswahl im Formular festlegen."
            Exit Function
        Case FT_FORMEL
            msg = "Berechnete Felder können nicht manuell geändert werden."
            Exit Function
    End Select
    ParseInput = CheckValue(f, outV, msg)
End Function

' Formatpruefung fuer URL, E-Mail und Telefon
Public Function KindFormatOk(ByVal kind As String, ByVal s As String, ByRef msg As String) As Boolean
    Dim at As Long, i As Long, ch As String, digits As Long
    KindFormatOk = True
    Select Case kind
        Case FT_URL
            If Not (LCase$(Left$(s, 7)) = "http://" Or LCase$(Left$(s, 8)) = "https://" Or LCase$(Left$(s, 4)) = "www.") Or InStr(s, " ") > 0 Then
                msg = "Bitte eine Webadresse eingeben, z. B. https://www.beispiel.de": KindFormatOk = False
            End If
        Case FT_EMAIL
            at = InStr(s, "@")
            If at < 2 Then
                KindFormatOk = False
            ElseIf InStr(at + 1, s, "@") > 0 Or InStr(at, s, ".") = 0 Or Right$(s, 1) = "." Or InStr(s, " ") > 0 Then
                KindFormatOk = False
            End If
            If Not KindFormatOk Then
                msg = "Bitte eine gültige E-Mail-Adresse eingeben, z. B. name@beispiel.de"
            End If
        Case FT_TELEFON
            For i = 1 To Len(s)
                ch = Mid$(s, i, 1)
                If ch Like "#" Then
                    digits = digits + 1
                ElseIf InStr(" +-/()", ch) = 0 Then
                    msg = "Telefonnummern dürfen nur Ziffern, Leerzeichen und + - / ( ) enthalten.": KindFormatOk = False: Exit Function
                End If
            Next
            If digits < 3 Then msg = "Bitte eine vollständige Telefonnummer eingeben.": KindFormatOk = False
    End Select
End Function

' Wert aus einer Excel-Zelle (kann bereits Zahl/Datum sein)
Public Function ParseCell(ByVal f As cField, ByVal cellV As Variant, ByRef outV As Variant, ByRef msg As String, _
                          Optional ByVal allowInactive As String = "") As Boolean
    msg = ""
    If IsError(cellV) Then msg = "Ungültiger Zellwert.": Exit Function
    If IsBlank(cellV) Then
        outV = Empty
        If f.Kind = FT_KONTROLL Then outV = False
        ParseCell = True
        Exit Function
    End If
    If VarType(cellV) = vbString Then
        ParseCell = ParseInput(f, CStr(cellV), outV, msg, allowInactive)
        Exit Function
    End If
    Select Case f.FType
        Case FT_ZAHL, FT_GELD
            If IsNumeric(cellV) And VarType(cellV) <> vbBoolean Then
                outV = CDbl(cellV)
            Else
                msg = "Bitte eine Zahl eingeben.": Exit Function
            End If
        Case FT_PROZENT
            ' Prozentformatierte Zelle liefert bereits den Anteil (25 % = 0,25)
            If IsNumeric(cellV) And VarType(cellV) <> vbBoolean Then
                outV = CDbl(cellV)
            Else
                msg = "Bitte einen Prozentwert eingeben.": Exit Function
            End If
        Case FT_DATUM, FT_DATUMZEIT
            If VarType(cellV) = vbDate Or (IsNumeric(cellV) And VarType(cellV) <> vbBoolean) Then
                outV = CDate(cellV)
                If f.FType = FT_DATUM And CDbl(outV) - Int(CDbl(outV)) <> 0 Then
                    msg = "Dieses Feld speichert nur ein Datum ohne Uhrzeit.": Exit Function
                End If
                If Year(outV) < 1900 Then msg = "Das Datum liegt außerhalb des gültigen Bereichs.": Exit Function
            Else
                msg = "Bitte ein gültiges Datum eingeben.": Exit Function
            End If
        Case FT_JANEIN
            If VarType(cellV) = vbBoolean Then outV = CBool(cellV) Else ParseCell = ParseInput(f, CStr(cellV), outV, msg): Exit Function
        Case Else
            ParseCell = ParseInput(f, CStr(cellV), outV, msg, allowInactive)
            Exit Function
    End Select
    ParseCell = CheckValue(f, outV, msg)
End Function

' Bereichs-/Laengenpruefungen eines bereits typisierten Wertes
Public Function CheckValue(ByVal f As cField, ByVal v As Variant, ByRef msg As String) As Boolean
    Dim d As Double
    If IsBlank(v) Then CheckValue = True: Exit Function
    Select Case f.FType
        Case FT_TEXT, FT_PFAD
            If f.MaxLen > 0 And Len(CStr(v)) > f.MaxLen Then
                msg = "Höchstens " & f.MaxLen & " Zeichen erlaubt (aktuell " & Len(CStr(v)) & ")."
                Exit Function
            End If
        Case FT_LTEXT
            If Len(CStr(v)) > 32000 Then msg = "Der Text ist zu lang (höchstens 32.000 Zeichen).": Exit Function
        Case FT_ZAHL
            d = CDbl(v)
            If f.Decimals = 0 And d <> Fix(d) Then
                msg = "Bitte eine ganze Zahl eingeben.": Exit Function
            End If
            If f.Decimals > 0 Then
                If Abs(Round(d, f.Decimals) - d) > 0.0000000001 Then
                    msg = "Höchstens " & f.Decimals & " Nachkommastellen erlaubt.": Exit Function
                End If
            End If
            If Not RangeOk(f, d, msg) Then Exit Function
        Case FT_GELD
            d = CDbl(v)
            If Abs(Round(d, 2) - d) > 0.0000000001 Then msg = "Geldbeträge haben höchstens 2 Nachkommastellen.": Exit Function
            If Not RangeOk(f, d, msg) Then Exit Function
        Case FT_PROZENT
            d = CDbl(v)
            If Not RangeOk(f, d * 100, msg) Then Exit Function
        Case FT_DATUM, FT_DATUMZEIT
            If Not IsDate(v) And Not IsNumeric(v) Then msg = "Ungültiges Datum.": Exit Function
        Case FT_MEHRFACH
            ' Höchstzahl der Auswahl (bei Mehrfachauswahl im Feld "Maximum" gespeichert)
            If MultiCount(CStr(v)) > MaxPicks(f) And MaxPicks(f) > 0 Then
                msg = "Höchstens " & MaxPicks(f) & " Auswahlmöglichkeiten erlaubt (gewählt: " & MultiCount(CStr(v)) & ")."
                Exit Function
            End If
    End Select
    CheckValue = True
End Function

' Höchstzahl wählbarer Optionen einer Mehrfachauswahl (0 = unbegrenzt)
Public Function MaxPicks(ByVal f As cField) As Long
    If f.FType <> FT_MEHRFACH Then Exit Function
    If IsEmpty(f.MaxVal) Then Exit Function
    If IsNumeric(f.MaxVal) Then MaxPicks = CLng(f.MaxVal)
    If MaxPicks < 0 Then MaxPicks = 0
End Function

Public Function MultiCount(ByVal s As String) As Long
    Dim a As Variant, i As Long
    a = Split(s, ";")
    For i = 0 To UBound(a)
        If Len(Trim$(a(i))) > 0 Then MultiCount = MultiCount + 1
    Next
End Function

Private Function RangeOk(ByVal f As cField, ByVal d As Double, ByRef msg As String) As Boolean
    If Not IsEmpty(f.MinVal) Then
        If d < CDbl(f.MinVal) Then
            msg = "Der Wert muss mindestens " & FmtNum(CDbl(f.MinVal), 6, True) & IIf(f.FType = FT_PROZENT, " %", "") & " betragen."
            Exit Function
        End If
    End If
    If Not IsEmpty(f.MaxVal) Then
        If d > CDbl(f.MaxVal) Then
            msg = "Der Wert darf höchstens " & FmtNum(CDbl(f.MaxVal), 6, True) & IIf(f.FType = FT_PROZENT, " %", "") & " betragen."
            Exit Function
        End If
    End If
    RangeOk = True
End Function

' Pruefung eines vollstaendigen Datensatzes.
' vals: FeldID -> neuer Wert (nur geaenderte oder alle); links: FeldID -> "R1;R2"
' errs erhaelt FeldID -> Meldung. rid leer = neuer Datensatz.
Public Function ValidateRecord(ByVal tid As String, ByVal rid As String, ByVal vals As cMap, ByVal links As cMap, ByRef errs As cMap) As Boolean
    Dim f As cField, v As Variant, msg As String, other As String, nb As cField, v2 As Variant
    Set errs = New cMap
    For Each f In mCfg.Fields(tid)
        If f.FType <> FT_FORMEL And Not IsAutoKind(f.Kind) Then
            If f.IsLinkType Then
                If links.Exists(f.ID) Then
                    v = links.GetVal(f.ID)
                ElseIf Len(rid) > 0 Then
                    v = mData.LinkTargets(f.ID, rid)
                Else
                    v = ""
                End If
                If f.Required And Len(CStr(v)) = 0 Then errs.SetVal f.ID, "Bitte mindestens einen Eintrag auswählen."
                If f.FType = FT_LINK And Not f.Multi And ListCount(CStr(v)) > 1 Then errs.SetVal f.ID, "Hier ist nur ein Eintrag erlaubt."
            Else
                v = CurrentOrNew(tid, rid, f.ID, vals)
                If f.Required And IsBlank(v) Then
                    errs.SetVal f.ID, "Pflichtangabe: bitte ausfüllen."
                ElseIf f.FType = FT_JANEIN And Not f.AllowEmpty And f.Kind <> FT_KONTROLL And IsBlank(v) Then
                    errs.SetVal f.ID, "Bitte Ja oder Nein wählen."
                ElseIf Not IsBlank(v) Then
                    msg = ""
                    If vals.Exists(f.ID) Then
                        If Not CheckValue(f, v, msg) Then errs.SetVal f.ID, msg
                    End If
                    If Not errs.Exists(f.ID) And f.Unique Then
                        other = FindDuplicateValue(tid, rid, f.ID, v)
                        If Len(other) > 0 Then errs.SetVal f.ID, "Dieser Wert muss eindeutig sein und ist bereits bei """ & _
                                                  mData.RecordLabel(tid, other) & """ vergeben."
                    End If
                    If Not errs.Exists(f.ID) And Len(f.NotBefore) > 0 And f.IsDate_ Then
                        v2 = CurrentOrNew(tid, rid, f.NotBefore, vals)
                        If Not IsBlank(v2) And (IsDate(v2) Or IsNumeric(v2)) Then
                            If CDbl(CDate(v)) < CDbl(CDate(v2)) Then
                                errs.SetVal f.ID, "Darf nicht vor """ & mCfg.FieldName(f.NotBefore) & """ (" & FmtDate(v2) & ") liegen."
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    ValidateRecord = (errs.Count = 0)
End Function

Public Function CurrentOrNew(ByVal tid As String, ByVal rid As String, ByVal fid As String, ByVal vals As cMap) As Variant
    If vals.Exists(fid) Then
        CurrentOrNew = vals.GetVal(fid)
    ElseIf Len(rid) > 0 Then
        CurrentOrNew = mData.GetValue(tid, rid, fid)
    Else
        CurrentOrNew = Empty
    End If
End Function

' Anderer Datensatz mit gleichem (normalisiertem) Wert
Public Function FindDuplicateValue(ByVal tid As String, ByVal rid As String, ByVal fid As String, ByVal v As Variant) As String
    Dim t As cTableData, c As Long, r As Long, key As String
    key = NormKey(v)
    If Len(key) = 0 Then Exit Function
    Set t = mData.TData(tid)
    c = t.ColOf(fid)
    If c = 0 Then Exit Function
    For r = 1 To t.RowCount
        If CStr(t.Cell(r, 1)) <> rid Then
            If NormKey(t.Cell(r, c)) = key Then FindDuplicateValue = CStr(t.Cell(r, 1)): Exit Function
        End If
    Next
End Function

' Dublettenpruefung anhand der konfigurierten Felder (Warnung, keine Sperre)
Public Function FindDuplicateRecord(ByVal tid As String, ByVal rid As String, ByVal vals As cMap) As String
    Dim flds As Variant, i As Long, t As cTableData, r As Long, ok As Boolean, keys() As String, cols() As Long, n As Long
    flds = Split(mCfg.GetS(SH_TAB, tid, "Dublettenfelder"), ";")
    If UBound(flds) < 0 Then Exit Function
    Set t = mData.TData(tid)
    ReDim keys(0 To UBound(flds)): ReDim cols(0 To UBound(flds))
    For i = 0 To UBound(flds)
        If Not mCfg.FieldExists(CStr(flds(i))) Then Exit Function
        keys(i) = NormKey(CurrentOrNew(tid, rid, CStr(flds(i)), vals))
        If Len(keys(i)) = 0 Then Exit Function
        cols(i) = t.ColOf(CStr(flds(i)))
        If cols(i) = 0 Then Exit Function
    Next
    For r = 1 To t.RowCount
        If CStr(t.Cell(r, 1)) <> rid Then
            ok = True
            For i = 0 To UBound(flds)
                If NormKey(t.Cell(r, cols(i))) <> keys(i) Then ok = False: Exit For
            Next
            If ok Then FindDuplicateRecord = CStr(t.Cell(r, 1)): Exit Function
        End If
    Next
End Function

' Standardwert eines Feldes fuer neue Datensaetze
Public Function DefaultValue(ByVal f As cField) As Variant
    Dim s As String, n As Double, msg As String, v As Variant, p As Long
    s = Trim$(f.DefaultVal)
    DefaultValue = Empty
    If f.FType = FT_JANEIN And Not f.AllowEmpty And Len(s) = 0 Then DefaultValue = False: Exit Function
    If Len(s) = 0 Then Exit Function
    Select Case f.FType
        Case FT_DATUM, FT_DATUMZEIT
            If LCase$(Left$(s, 5)) = "heute" Then
                n = 0
                If Len(s) > 5 Then n = Val(Replace(Mid$(s, 6), " ", ""))
                If f.FType = FT_DATUMZEIT And LCase$(s) = "heute" Then DefaultValue = Now Else DefaultValue = Date + n
                Exit Function
            End If
            v = InvDate(s)
            If Not IsEmpty(v) Then DefaultValue = v
        Case FT_ZAHL, FT_GELD, FT_PROZENT
            DefaultValue = InvNum(s)
        Case FT_JANEIN
            DefaultValue = (s = "Ja")
        Case FT_AUSWAHL, FT_STATUS, FT_MEHRFACH
            DefaultValue = s
        Case Else
            DefaultValue = s
    End Select
End Function
