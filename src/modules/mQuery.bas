Attribute VB_Name = "mQuery"
'==============================================================================
' mQuery - Suche, Filter (UND bindet staerker als ODER), mehrstufige Sortierung
' und Gruppierung auf dem Arbeitsspeicherabbild einer Tabelle.
' Filterformat: "Verknuepfung|FeldID|Operator|Wert1|Wert2;..." (Werte maskiert)
'==============================================================================
Option Explicit

Private mGrpCache As cMap
Private mGrpRev As Long
Private mSrchKey As String
Private mSrchRows() As String
Public gSrchProfile As String

Public Const ARCH_ACTIVE As Long = 0
Public Const ARCH_ONLY As Long = 1
Public Const ARCH_ALL As Long = 2

'--- Operatoren ------------------------------------------------------------------
Public Function OpKind(ByVal f As cField) As String
    If f Is Nothing Then OpKind = "text": Exit Function
    If f.IsLinkType Then OpKind = "link": Exit Function
    Select Case f.ValueType
        Case FT_ZAHL, FT_GELD, FT_PROZENT: OpKind = "num"
        Case FT_DATUM, FT_DATUMZEIT: OpKind = "date"
        Case FT_JANEIN: OpKind = "bool"
        Case FT_AUSWAHL: OpKind = "opt"
        Case FT_STATUS: OpKind = "status"
        Case FT_MEHRFACH: OpKind = "multi"
        Case Else: OpKind = "text"
    End Select
End Function

Public Function OpsFor(ByVal kind As String) As Variant
    Select Case kind
        Case "num": OpsFor = Array("=", "<>", ">", ">=", "<", "<=", "ZWISCHEN", "LEER", "NICHTLEER")
        Case "date": OpsFor = Array("AM", "VOR", "NACH", "ZWISCHEN", "HEUTE", "DIESEWOCHE", "DIESERMONAT", "LETZTE7", "NAECHSTE7", "NAECHSTE30", "VERGANGEN", "LEER", "NICHTLEER")
        Case "bool": OpsFor = Array("JA", "NEIN", "LEER")
        Case "opt": OpsFor = Array("IST", "ISTNICHT", "LEER", "NICHTLEER")
        Case "status": OpsFor = Array("IST", "ISTNICHT", "OFFEN", "ABGESCHLOSSEN", "LEER", "NICHTLEER")
        Case "multi": OpsFor = Array("ENTHAELT", "ENTHAELTALLE", "ENTHAELTNICHT", "LEER", "NICHTLEER")
        Case "link": OpsFor = Array("VERKNUEPFT", "NICHTVERKNUEPFT", "LEER", "NICHTLEER")
        Case Else: OpsFor = Array("ENTHAELT", "NICHTENTHAELT", "GLEICH", "BEGINNT", "LEER", "NICHTLEER")
    End Select
End Function

Public Function OpLabel(ByVal op As String) As String
    Select Case op
        Case "ENTHAELT": OpLabel = "enthält"
        Case "NICHTENTHAELT": OpLabel = "enthält nicht"
        Case "GLEICH": OpLabel = "ist genau"
        Case "BEGINNT": OpLabel = "beginnt mit"
        Case "LEER": OpLabel = "ist leer"
        Case "NICHTLEER": OpLabel = "ist nicht leer"
        Case "=": OpLabel = "gleich"
        Case "<>": OpLabel = "ungleich"
        Case ">": OpLabel = "größer als"
        Case ">=": OpLabel = "mindestens"
        Case "<": OpLabel = "kleiner als"
        Case "<=": OpLabel = "höchstens"
        Case "ZWISCHEN": OpLabel = "zwischen"
        Case "AM": OpLabel = "am"
        Case "VOR": OpLabel = "vor"
        Case "NACH": OpLabel = "nach"
        Case "HEUTE": OpLabel = "ist heute"
        Case "DIESEWOCHE": OpLabel = "in dieser Woche"
        Case "DIESERMONAT": OpLabel = "in diesem Monat"
        Case "LETZTE7": OpLabel = "in den letzten 7 Tagen"
        Case "NAECHSTE7": OpLabel = "in den nächsten 7 Tagen"
        Case "NAECHSTE30": OpLabel = "in den nächsten 30 Tagen"
        Case "VERGANGEN": OpLabel = "liegt vor heute"
        Case "JA": OpLabel = "ist Ja"
        Case "NEIN": OpLabel = "ist Nein"
        Case "IST": OpLabel = "ist"
        Case "ISTNICHT": OpLabel = "ist nicht"
        Case "OFFEN": OpLabel = "ist offen (nicht abgeschlossen)"
        Case "ABGESCHLOSSEN": OpLabel = "ist abgeschlossen"
        Case "ENTHAELTALLE": OpLabel = "enthält alle von"
        Case "ENTHAELTNICHT": OpLabel = "enthält keine von"
        Case "VERKNUEPFT": OpLabel = "ist verknüpft mit"
        Case "NICHTVERKNUEPFT": OpLabel = "ist nicht verknüpft mit"
        Case Else: OpLabel = op
    End Select
End Function

Public Function OpLabelToCode(ByVal kind As String, ByVal lbl As String) As String
    Dim v As Variant
    For Each v In OpsFor(kind)
        If OpLabel(CStr(v)) = lbl Then OpLabelToCode = CStr(v): Exit Function
    Next
End Function

' Wie viele Werte braucht ein Operator (0, 1, 2) und welcher Art
Public Function OpArity(ByVal op As String) As Long
    Select Case op
        Case "LEER", "NICHTLEER", "HEUTE", "DIESEWOCHE", "DIESERMONAT", "LETZTE7", "NAECHSTE7", "NAECHSTE30", "VERGANGEN", "JA", "NEIN", "OFFEN", "ABGESCHLOSSEN"
            OpArity = 0
        Case "ZWISCHEN": OpArity = 2
        Case Else: OpArity = 1
    End Select
End Function

'--- Bedingungen lesen/schreiben ------------------------------------------------
' Liefert Collection von Array(conn, fid, op, v1, v2)
Public Function ParseConds(ByVal s As String) As Collection
    Dim res As New Collection, parts As Variant, i As Long, p As Variant
    If Len(s) > 0 Then
        parts = Split(s, ";")
        For i = 0 To UBound(parts)
            p = Split(parts(i), "|")
            If UBound(p) >= 2 Then
                res.Add Array(CStr(p(0)), CStr(p(1)), CStr(p(2)), Unesc(ArgAtV(p, 3)), Unesc(ArgAtV(p, 4)))
            End If
        Next
    End If
    Set ParseConds = res
End Function

Private Function ArgAtV(ByVal p As Variant, ByVal i As Long) As String
    If i <= UBound(p) Then ArgAtV = CStr(p(i))
End Function

Public Function SerializeConds(ByVal conds As Collection) As String
    Dim c As Variant, s As String
    For Each c In conds
        s = s & IIf(Len(s) > 0, ";", "") & c(0) & "|" & c(1) & "|" & c(2) & "|" & Esc(CStr(c(3))) & "|" & Esc(CStr(c(4)))
    Next
    SerializeConds = s
End Function

' Lesbare Zusammenfassung einer Bedingung
Public Function CondText(ByVal c As Variant) As String
    Dim f As cField, s As String, a As Variant, i As Long, vals As String, tt As String
    Set f = mCfg.Field(CStr(c(1)))
    If f Is Nothing Then CondText = "(entferntes Feld)": Exit Function
    s = f.Name & " " & OpLabel(CStr(c(2)))
    If OpArity(CStr(c(2))) >= 1 Then
        Select Case OpKind(f)
            Case "opt", "status", "multi"
                a = Split(CStr(c(3)), ",")
                For i = 0 To UBound(a)
                    vals = vals & IIf(Len(vals) > 0, ", ", "") & mCfg.OptName(CStr(a(i)))
                Next
                s = s & " „" & vals & "“"
            Case "link"
                If f.FType = FT_LINK Then tt = f.TargetTable Else tt = mCfg.Field(f.PairField).TableID
                a = Split(CStr(c(3)), ",")
                For i = 0 To UBound(a)
                    vals = vals & IIf(Len(vals) > 0, ", ", "") & mData.RecordLabel(tt, CStr(a(i)))
                Next
                s = s & " „" & vals & "“"
            Case "num"
                s = s & " " & FmtNum(InvNum(CStr(c(3))), 6, True)
                If OpArity(CStr(c(2))) = 2 Then s = s & " und " & FmtNum(InvNum(CStr(c(4))), 6, True)
            Case "date"
                s = s & " " & FmtDate(InvDate(CStr(c(3))))
                If OpArity(CStr(c(2))) = 2 Then s = s & " und " & FmtDate(InvDate(CStr(c(4))))
            Case Else
                s = s & " „" & c(3) & "“"
        End Select
    End If
    CondText = s
End Function

Public Function CondsText(ByVal conds As Collection) As String
    Dim c As Variant, s As String, first As Boolean
    first = True
    For Each c In conds
        If Not first Then s = s & IIf(c(0) = "ODER", " oder ", " und ")
        s = s & CondText(c)
        first = False
    Next
    CondsText = s
End Function

'--- Auswertung einer Bedingung --------------------------------------------------
Private Function DateOnly(ByVal v As Variant) As Double
    DateOnly = Int(CDbl(CDate(v)))
End Function

Public Function MatchCond(ByVal f As cField, ByVal v As Variant, ByVal rid As String, ByVal op As String, ByVal v1 As String, ByVal v2 As String) As Boolean
    Dim kind As String, d As Double, x As Double, a As Variant, i As Long, s As String, blank As Boolean, lst As String
    kind = OpKind(f)
    If kind = "link" Then
        lst = mData.LinkTargets(f.ID, rid)
        Select Case op
            Case "LEER": MatchCond = (Len(lst) = 0)
            Case "NICHTLEER": MatchCond = (Len(lst) > 0)
            Case "VERKNUEPFT", "NICHTVERKNUEPFT"
                a = Split(v1, ",")
                For i = 0 To UBound(a)
                    If ListHas(lst, CStr(a(i))) Then MatchCond = True: Exit For
                Next
                If op = "NICHTVERKNUEPFT" Then MatchCond = Not MatchCond
        End Select
        Exit Function
    End If
    blank = IsBlank(v)
    If op = "LEER" Then MatchCond = blank: Exit Function
    If op = "NICHTLEER" Then MatchCond = Not blank: Exit Function
    If IsErrVal(v) Then Exit Function
    Select Case kind
        Case "text"
            s = LCase$(NzS(v))
            Select Case op
                Case "ENTHAELT": MatchCond = (InStr(1, s, LCase$(v1)) > 0)
                Case "NICHTENTHAELT": MatchCond = (InStr(1, s, LCase$(v1)) = 0)
                Case "GLEICH": MatchCond = (Trim$(s) = Trim$(LCase$(v1)))
                Case "BEGINNT": MatchCond = (Left$(s, Len(v1)) = LCase$(v1))
            End Select
        Case "num"
            If blank Or Not IsNumeric(v) Then Exit Function
            x = CDbl(v)
            If f.ValueType = FT_PROZENT Then x = Round(x * 100, 9)
            d = InvNum(v1)
            Select Case op
                Case "=": MatchCond = (Abs(x - d) < 0.0000001)
                Case "<>": MatchCond = (Abs(x - d) >= 0.0000001)
                Case ">": MatchCond = (x > d)
                Case ">=": MatchCond = (x >= d)
                Case "<": MatchCond = (x < d)
                Case "<=": MatchCond = (x <= d)
                Case "ZWISCHEN": MatchCond = (x >= d And x <= InvNum(v2))
            End Select
        Case "date"
            If blank Or Not (IsDate(v) Or IsNumeric(v)) Then Exit Function
            x = DateOnly(v)
            Select Case op
                Case "AM": MatchCond = (x = DateOnly(InvDate(v1)))
                Case "VOR": MatchCond = (x < DateOnly(InvDate(v1)))
                Case "NACH": MatchCond = (x > DateOnly(InvDate(v1)))
                Case "ZWISCHEN": MatchCond = (x >= DateOnly(InvDate(v1)) And x <= DateOnly(InvDate(v2)))
                Case "HEUTE": MatchCond = (x = CDbl(Date))
                Case "DIESEWOCHE"
                    d = CDbl(Date) - Weekday(Date, vbMonday) + 1
                    MatchCond = (x >= d And x <= d + 6)
                Case "DIESERMONAT": MatchCond = (Year(CDate(x)) = Year(Date) And Month(CDate(x)) = Month(Date))
                Case "LETZTE7": MatchCond = (x >= CDbl(Date) - 7 And x <= CDbl(Date))
                Case "NAECHSTE7": MatchCond = (x >= CDbl(Date) And x <= CDbl(Date) + 7)
                Case "NAECHSTE30": MatchCond = (x >= CDbl(Date) And x <= CDbl(Date) + 30)
                Case "VERGANGEN": MatchCond = (x < CDbl(Date))
            End Select
        Case "bool"
            If blank Then Exit Function
            If op = "JA" Then MatchCond = CBool(v) Else MatchCond = Not CBool(v)
        Case "opt", "status"
            Select Case op
                Case "IST", "ISTNICHT"
                    MatchCond = ListHas(v1, NzS(v), ",")
                    If op = "ISTNICHT" Then MatchCond = Not MatchCond
                Case "OFFEN": If Not blank Then MatchCond = Not mCfg.OptDone(CStr(v)) Else MatchCond = True
                Case "ABGESCHLOSSEN": If Not blank Then MatchCond = mCfg.OptDone(CStr(v))
            End Select
        Case "multi"
            a = Split(v1, ",")
            Select Case op
                Case "ENTHAELT"
                    For i = 0 To UBound(a)
                        If ListHas(NzS(v), CStr(a(i))) Then MatchCond = True: Exit For
                    Next
                Case "ENTHAELTALLE"
                    MatchCond = True
                    For i = 0 To UBound(a)
                        If Not ListHas(NzS(v), CStr(a(i))) Then MatchCond = False: Exit For
                    Next
                Case "ENTHAELTNICHT"
                    MatchCond = True
                    For i = 0 To UBound(a)
                        If ListHas(NzS(v), CStr(a(i))) Then MatchCond = False: Exit For
                    Next
            End Select
    End Select
End Function

'--- Zeilenauswahl -------------------------------------------------------------------
' Liefert die Zeilenindizes (1..n) der passenden Datensaetze
Public Function FilterRows(ByVal tid As String, ByVal conds As Collection, ByVal search As String, _
                           ByVal searchFields As Collection, ByVal archMode As Long, ByRef n As Long) As Long()
    Dim t As cTableData, res() As Long, r As Long, ok As Boolean, cf() As cField, cc() As Long, k As Long
    Dim nC As Long, c As Variant, grpOk As Boolean, anyOk As Boolean, tokens As Variant, sf() As cField, sc() As Long, nS As Long
    Dim f As cField, txt As String, i As Long, optNames As New cMap, conn() As String, ops() As String, v1s() As String, v2s() As String
    Dim sk() As Long, sv As Variant, tw() As String, tf() As Boolean, nT As Long, hits As Long, srch() As String
    Set t = mData.TData(tid)
    n = 0
    If t.RowCount = 0 Then
        ReDim res(1 To 1)
        FilterRows = res
        Exit Function
    End If
    ReDim res(1 To t.RowCount)
    ' Bedingungen vorbereiten
    nC = conds.Count
    If nC > 0 Then
        ReDim cf(1 To nC): ReDim cc(1 To nC): ReDim conn(1 To nC): ReDim ops(1 To nC): ReDim v1s(1 To nC): ReDim v2s(1 To nC)
        k = 0
        For Each c In conds
            k = k + 1
            Set cf(k) = mCfg.Field(CStr(c(1)))
            If Not cf(k) Is Nothing Then cc(k) = t.ColOf(cf(k).ID)
            conn(k) = c(0): ops(k) = c(2): v1s(k) = c(3): v2s(k) = c(4)
        Next
    End If
    ' Suchbegriffe
    search = Trim$(LCase$(search))
    If Len(search) > 0 Then
        tokens = Split(search, " ")
        ReDim tw(0 To UBound(tokens)): ReDim tf(0 To UBound(tokens))
        For i = 0 To UBound(tokens)
            If Len(tokens(i)) > 0 Then tw(nT) = tokens(i): nT = nT + 1
        Next
        nS = searchFields.Count
        If nS > 0 Then
            ReDim sf(1 To nS): ReDim sc(1 To nS)
            k = 0
            ReDim sk(1 To nS)
            For Each f In searchFields
                k = k + 1
                Set sf(k) = f
                sc(k) = t.ColOf(f.ID)
                ' Art der Textbildung einmal je Feld bestimmen (nicht je Zeile)
                If f.IsLinkType Then
                    sk(k) = 6
                Else
                    Select Case f.ValueType
                        Case FT_TEXT, FT_LTEXT, FT_PFAD, FT_ZAHL, FT_GELD, FT_DATUM: sk(k) = 1
                        Case FT_AUSWAHL, FT_STATUS: sk(k) = 2
                        Case FT_MEHRFACH: sk(k) = 3
                        Case FT_PROZENT: sk(k) = 4
                        Case FT_JANEIN: sk(k) = 5
                        Case Else: sk(k) = 7
                    End Select
                End If
            Next
        End If
    End If
    If Len(search) > 0 And nS > 0 Then srch = SearchIndex(tid, t, sf, sc, sk, nS, optNames)
    If Len(search) > 0 And nS = 0 Then n = 0: FilterRows = res: Exit Function
    For r = 1 To t.RowCount
        ok = True
        Select Case archMode
            Case ARCH_ACTIVE: ok = Not t.IsArchived(r)
            Case ARCH_ONLY: ok = t.IsArchived(r)
        End Select
        If ok And nC > 0 Then
            ' UND bindet staerker als ODER
            anyOk = False: grpOk = True
            For k = 1 To nC
                If k > 1 And conn(k) = "ODER" Then
                    If grpOk Then anyOk = True: Exit For
                    grpOk = True
                End If
                If grpOk Then
                    If cf(k) Is Nothing Then
                        grpOk = False
                    ElseIf cf(k).IsLinkType Then
                        grpOk = MatchCond(cf(k), Empty, t.IdAt(r), ops(k), v1s(k), v2s(k))
                    ElseIf cc(k) > 0 Then
                        grpOk = MatchCond(cf(k), t.Cell(r, cc(k)), t.IdAt(r), ops(k), v1s(k), v2s(k))
                    Else
                        grpOk = False
                    End If
                End If
            Next
            If grpOk Then anyOk = True
            ok = anyOk
        End If
        If ok And Len(search) > 0 Then
            For i = 0 To nT - 1
                If InStr(1, srch(r), tw(i), vbBinaryCompare) = 0 Then ok = False: Exit For
            Next
        End If
        If ok Then n = n + 1: res(n) = r
    Next
    FilterRows = res
End Function

' Suchindex: je Zeile der kleingeschriebene Text aller Suchfelder. Wird wiederverwendet, solange
' sich Daten, Struktur und Spaltenauswahl nicht aendern (z. B. beim Verfeinern eines Suchbegriffs).
Private Function SearchIndex(ByVal tid As String, ByVal t As cTableData, ByRef sf() As cField, ByRef sc() As Long, ByRef sk() As Long, _
                             ByVal nS As Long, ByVal optNames As cMap) As String()
    Dim key As String, k As Long, r As Long, sv As Variant, a() As String, col As Variant
    For k = 1 To nS: key = key & sf(k).ID & ",": Next
    key = tid & "|" & key & "|" & mApp.gDataRev & "|" & mApp.gCfgRev & "|" & t.RowCount
    If key = mSrchKey Then SearchIndex = mSrchRows: Exit Function
    ReDim a(1 To t.RowCount)
    ' spaltenweise aufbauen: Optionen und Verknuepfungsziele ueber Arrays (Nummer der ID) statt Schluesselsuche
    Dim tp0 As Double, prof As New cMap, optArr() As String, maxO As Long, ov As Variant, f As cField, rowOf() As Long, maxR As Long
    Dim parts As Variant, i As Long, tt As String, td As cTableData, lbl() As String, rel As cTableData, rr As Long, pf As String, fromC As Long, toC As Long
    ReDim optArr(0 To 0)
    For k = 1 To nS
        If sk(k) = 2 Or sf(k).ValueType = FT_MEHRFACH Then
            For Each ov In mCfg.OptionIds(sf(k).ID)
                If NumId(CStr(ov)) > maxO Then maxO = NumId(CStr(ov))
            Next
        End If
    Next
    ReDim optArr(0 To maxO)
    For k = 1 To nS
        If sk(k) = 2 Or sf(k).ValueType = FT_MEHRFACH Then
            For Each ov In mCfg.OptionIds(sf(k).ID)
                optArr(NumId(CStr(ov))) = mCfg.OptName(CStr(ov))
            Next
        End If
    Next
    For k = 1 To nS
        tp0 = Timer
        If sk(k) = 6 Then
            ' Verknuepfung: Zuordnungstabelle einmal durchlaufen
            Set f = sf(k)
            If maxR = 0 Then
                For r = 1 To t.RowCount
                    If NumId(t.IdAt(r)) > maxR Then maxR = NumId(t.IdAt(r))
                Next
                ReDim rowOf(0 To maxR)
                For r = 1 To t.RowCount: rowOf(NumId(t.IdAt(r))) = r: Next
            End If
            If f.FType = FT_LINK Then
                tt = f.TargetTable: pf = f.ID: fromC = 2: toC = 3
            Else
                pf = f.PairField: tt = mCfg.Field(pf).TableID: fromC = 3: toC = 2
            End If
            Set td = mData.TData(tt)
            ReDim lbl(0 To 0)
            For rr = 1 To td.RowCount
                If NumId(td.IdAt(rr)) > UBound(lbl) Then ReDim Preserve lbl(0 To NumId(td.IdAt(rr)))
                lbl(NumId(td.IdAt(rr))) = mData.RecordLabel(tt, td.IdAt(rr))
            Next
            Set rel = mCfg.TD(SH_REL)
            For rr = 1 To rel.RowCount
                If CStr(rel.Cell(rr, 1)) = pf Then
                    i = NumId(CStr(rel.Cell(rr, fromC)))
                    If i <= maxR Then
                        r = rowOf(i)
                        If r > 0 Then
                            i = NumId(CStr(rel.Cell(rr, toC)))
                            If i <= UBound(lbl) Then a(r) = a(r) & " " & lbl(i)
                        End If
                    End If
                End If
            Next
        ElseIf sc(k) > 0 Then
            col = t.ColumnArray(sf(k).ID)
            For r = 1 To t.RowCount
                sv = col(r)
                If Not IsEmpty(sv) Then
                    Select Case sk(k)
                        Case 1: If Not IsError(sv) Then a(r) = a(r) & " " & CStr(sv)
                        Case 2
                            i = NumId(CStr(sv))
                            If i <= maxO Then a(r) = a(r) & " " & optArr(i)
                        Case 3
                            parts = Split(CStr(sv), ";")
                            For i = 0 To UBound(parts)
                                If NumId(CStr(parts(i))) <= maxO Then a(r) = a(r) & " " & optArr(NumId(CStr(parts(i))))
                            Next
                        Case 4: If IsNumeric(sv) Then a(r) = a(r) & " " & CStr(Round(CDbl(sv) * 100, 6)) & " %"
                        Case 5: If VarType(sv) = vbBoolean Then a(r) = a(r) & IIf(sv, " ja", " nein")
                        Case Else: a(r) = a(r) & " " & SearchText(sf(k), sv, t.IdAt(r), optNames)
                    End Select
                End If
            Next
        End If
        prof.AddNum sf(k).ValueType & "/" & sk(k), Timer - tp0
    Next
    tp0 = Timer
    For r = 1 To t.RowCount
        a(r) = LCase$(a(r))
    Next
    prof.AddNum "lcase", Timer - tp0
    gSrchProfile = ""
    Dim pk As Variant
    For Each pk In prof.Keys: gSrchProfile = gSrchProfile & pk & "=" & Format$(prof.GetVal(CStr(pk)), "0.00") & " ": Next
    mSrchKey = key
    mSrchRows = a
    SearchIndex = a
End Function

' Nummer einer ID wie "O0012" oder "R0001234" (0 bei unbekanntem Format)
Private Function NumId(ByVal s As String) As Long
    If Len(s) > 1 Then NumId = CLng(Val(Mid$(s, 2)))
    If NumId < 0 Then NumId = 0
End Function

Public Function SearchText(ByVal f As cField, ByVal v As Variant, ByVal rid As String, ByVal optNames As cMap) As String
    Dim a As Variant, i As Long, s As String
    Dim tt As String, ids As Variant
    If f.IsLinkType Then
        ' Bezeichnungen der Zieleintraege je Suchlauf zwischenspeichern (viele Zeilen verweisen auf dieselben Ziele)
        If f.FType = FT_LINK Then tt = f.TargetTable Else tt = mCfg.Field(f.PairField).TableID
        ids = Split(mData.LinkTargets(f.ID, rid), ";")
        For i = 0 To UBound(ids)
            If Not optNames.Exists("L:" & ids(i)) Then optNames.SetVal "L:" & ids(i), mData.RecordLabel(tt, CStr(ids(i)))
            s = s & " " & optNames.GetVal("L:" & ids(i))
        Next
        SearchText = s
        Exit Function
    End If
    If IsBlank(v) Then Exit Function
    Select Case f.ValueType
        Case FT_TEXT, FT_LTEXT, FT_PFAD
            SearchText = CStr(v)
        Case FT_AUSWAHL, FT_STATUS
            If Not optNames.Exists(CStr(v)) Then optNames.SetVal CStr(v), mCfg.OptName(CStr(v))
            SearchText = optNames.GetVal(CStr(v))
        Case FT_MEHRFACH
            a = Split(CStr(v), ";")
            For i = 0 To UBound(a)
                If Not optNames.Exists(CStr(a(i))) Then optNames.SetVal CStr(a(i)), mCfg.OptName(CStr(a(i)))
                s = s & " " & optNames.GetVal(CStr(a(i)))
            Next
            SearchText = s
        Case FT_ZAHL, FT_GELD, FT_DATUM
            ' schnelle Textform (deutsches Format des Systems) statt vollstaendiger Formatierung
            SearchText = CStr(v)
        Case FT_JANEIN
            If VarType(v) = vbBoolean Then SearchText = IIf(v, "ja", "nein") Else SearchText = CStr(v)
        Case Else
            SearchText = mFmt.DisplayText(f, v, rid)
    End Select
End Function

'--- Sortierung ---------------------------------------------------------------------
' sortSpec: "F0001:A;F0003:D"; leere Werte stehen am Ende
Public Sub SortRows(ByVal tid As String, ByRef rows() As Long, ByVal n As Long, ByVal sortSpec As String, Optional ByVal groupFid As String = "", Optional ByVal groupLevel As String = "")
    Dim t As cTableData, parts As Variant, nK As Long, keys() As Variant, dirs() As Long, i As Long, k As Long
    Dim idx() As Long, fid As String, sortDir As String, p As Variant
    If n < 2 Then Exit Sub
    Set t = mData.TData(tid)
    parts = Split(sortSpec, ";")
    nK = 0
    If Len(groupFid) > 0 Then nK = 1
    nK = nK + UBound(parts) + 1
    If nK = 0 Then Exit Sub
    ReDim keys(1 To nK): ReDim dirs(1 To nK)
    k = 0
    If Len(groupFid) > 0 Then
        k = 1
        keys(1) = GroupSortKeys(t, mCfg.Field(groupFid), groupLevel)
        dirs(1) = 1
    End If
    For i = 0 To UBound(parts)
        p = Split(parts(i), ":")
        If UBound(p) >= 0 Then
            fid = p(0)
            If UBound(p) >= 1 Then sortDir = p(1) Else sortDir = "A"
            If mCfg.FieldExists(fid) Then
                k = k + 1
                keys(k) = SortKeys(t, mCfg.Field(fid))
                dirs(k) = IIf(sortDir = "D", -1, 1)
            End If
        End If
    Next
    If k = 0 Then Exit Sub
    ' Sortiert werden Positionen im Zeilenarray; Schluessel nach Tabellenzeile
    ReDim idx(1 To n)
    For i = 1 To n: idx(i) = rows(i): Next
    MergeSortIdx idx, keys, k, dirs
    For i = 1 To n: rows(i) = idx(i): Next
End Sub

Private Function SortKeys(ByVal t As cTableData, ByVal f As cField) As Variant
    Dim a() As Variant, r As Long, c As Long, v As Variant
    ReDim a(1 To t.RowCount)
    c = t.ColOf(f.ID)
    For r = 1 To t.RowCount
        If f.IsLinkType Then
            a(r) = LCase$(mFmt.DisplayText(f, Empty, t.IdAt(r)))
            If Len(a(r)) = 0 Then a(r) = Empty
        ElseIf c > 0 Then
            v = t.Cell(r, c)
            If IsBlank(v) Or IsErrVal(v) Then
                a(r) = Empty
            Else
                Select Case f.ValueType
                    Case FT_AUSWAHL, FT_STATUS: a(r) = CDbl(mCfg.OptOrderIndex(f.ID, CStr(v)))
                    Case FT_MEHRFACH: a(r) = LCase$(mFmt.DisplayText(f, v))
                    Case FT_JANEIN: a(r) = IIf(CBool(v), 0#, 1#)
                    Case FT_TEXT, FT_LTEXT, FT_PFAD: a(r) = LCase$(CStr(v))
                    Case Else
                        If IsNumeric(v) Or IsDate(v) Then a(r) = CDbl(v) Else a(r) = LCase$(CStr(v))
                End Select
            End If
        End If
    Next
    SortKeys = a
End Function

'--- Gruppierung ----------------------------------------------------------------------
Private Function GroupSortKeys(ByVal t As cTableData, ByVal f As cField, ByVal level As String) As Variant
    Dim a() As Variant, r As Long, k As String, lbl As String, ord As Double
    ReDim a(1 To t.RowCount)
    For r = 1 To t.RowCount
        GroupOf f, GroupValue(t, f, r), t.IdAt(r), level, k, lbl, ord
        If Len(k) = 0 Then a(r) = Empty Else a(r) = ord
    Next
    GroupSortKeys = a
End Function

Public Function GroupValue(ByVal t As cTableData, ByVal f As cField, ByVal r As Long) As Variant
    Dim c As Long
    If f.IsLinkType Then GroupValue = Empty: Exit Function
    c = t.ColOf(f.ID)
    If c > 0 Then GroupValue = t.Cell(r, c)
End Function

' Gruppenschluessel, Beschriftung und Sortierwert fuer einen Wert
Public Sub GroupOf(ByVal f As cField, ByVal v As Variant, ByVal rid As String, ByVal level As String, _
                   ByRef key As String, ByRef label As String, ByRef ord As Double)
    Dim d As Date, lst As String
    key = "": label = "Ohne Angabe": ord = 1E+30
    If f.IsLinkType Then
        lst = mData.LinkTargets(f.ID, rid)
        If Len(lst) > 0 Then
            key = Split(lst, ";")(0)
            label = mFmt.LinkLabels(f, key)
            If ListCount(lst) > 1 Then label = label & " (+" & (ListCount(lst) - 1) & ")"
            ord = 0
        End If
        Exit Sub
    End If
    If IsBlank(v) Or IsErrVal(v) Then Exit Sub
    Select Case f.ValueType
        Case FT_AUSWAHL, FT_STATUS
            key = CStr(v)
            ' Beschriftung und Reihenfolge je Option zwischenspeichern (Auswertungen ueber viele Zeilen)
            If mGrpCache Is Nothing Or mGrpRev <> mApp.gCfgRev Then Set mGrpCache = New cMap: mGrpRev = mApp.gCfgRev
            If Not mGrpCache.Exists(f.ID & "|" & key) Then mGrpCache.SetVal f.ID & "|" & key, Array(mCfg.OptName(key), CDbl(mCfg.OptOrderIndex(f.ID, key)))
            label = mGrpCache.GetVal(f.ID & "|" & key)(0): ord = mGrpCache.GetVal(f.ID & "|" & key)(1)
        Case FT_JANEIN
            key = IIf(CBool(v), "J", "N"): label = IIf(CBool(v), "Ja", "Nein"): ord = IIf(CBool(v), 1, 2)
        Case FT_DATUM, FT_DATUMZEIT
            d = CDate(v)
            Select Case level
                Case "Jahr": key = CStr(Year(d)): label = CStr(Year(d)): ord = Year(d)
                Case "Quartal": key = Year(d) & "Q" & ((Month(d) - 1) \ 3 + 1): label = ((Month(d) - 1) \ 3 + 1) & ". Quartal " & Year(d): ord = Year(d) * 10 + (Month(d) - 1) \ 3
                Case "Tag": key = CStr(CLng(Int(CDbl(d)))): label = DateLong_DE(Int(d)): ord = Int(CDbl(d))
                Case Else: key = Year(d) & "-" & Format$(Month(d), "00"): label = MonthName_DE(Month(d)) & " " & Year(d): ord = Year(d) * 100 + Month(d)
            End Select
        Case FT_MEHRFACH
            key = CStr(v)
            If mGrpCache Is Nothing Or mGrpRev <> mApp.gCfgRev Then Set mGrpCache = New cMap: mGrpRev = mApp.gCfgRev
            If Not mGrpCache.Exists(f.ID & "|" & key) Then mGrpCache.SetVal f.ID & "|" & key, Array(mFmt.DisplayText(f, v), 0#)
            label = mGrpCache.GetVal(f.ID & "|" & key)(0): ord = 0
        Case FT_ZAHL, FT_GELD, FT_PROZENT
            key = NumInv(CDbl(v)): label = mFmt.DisplayText(f, v): ord = CDbl(v)
        Case Else
            key = LCase$(Trim$(CStr(v))): label = CStr(v): ord = 0
    End Select
End Sub

Public Function CanGroupBy(ByVal f As cField) As Boolean
    Select Case f.ValueType
        Case FT_AUSWAHL, FT_STATUS, FT_JANEIN, FT_DATUM, FT_DATUMZEIT, FT_TEXT, FT_MEHRFACH, FT_ZAHL
            CanGroupBy = True
    End Select
    If f.FType = FT_LINK Then CanGroupBy = True
End Function
