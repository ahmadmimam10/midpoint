Attribute VB_Name = "mDash"
'==============================================================================
' mDash - Auswertungsassistent und Dashboards. Elemente: Kennzahlkachel,
' Balken, Saeulen, Zeitverlauf, Anteil (Kreis), Ergebnistabelle. Jede
' Auswertung zaehlt eindeutige Datensaetze; Mehrfachzuordnungen werden
' transparent ausgewiesen (Gruppensumme kann die Gesamtzahl uebersteigen).
'==============================================================================
Option Explicit

Private mLastStamp As Date
Private mDrill As cMap           ' Formname -> Array(ElementID, Gruppenschluessel)
Private mDataCol As Long
Public Const DASH_DATA_COL As Long = 120
Public gArrange As Boolean        ' Anordnen: Karten ziehen, sie rasten im 4-Spalten-Raster ein
Private mWatch As cDashWatch      ' rastet nach dem Loslassen einer Karte sofort ein
Public gHighlight As String       ' Element-ID, die hervorgehoben wird (Vorschau)
Private mUnitW As Single, mGap As Single, mX0 As Single
' Raster: jedes Element hat Spalte (0-3), Zeile (0..), Breite (1-4 Spalten) und Hoehe
' (normal/hoch). Gespeichert in _Elemente (Optionen gx, gy, hoch; Groesse).
Private mArr As cMap              ' waehrend "Anordnen": eid -> Array(gx, gy, sz, hoch) (noch nicht gespeichert)
Private mLay As cMap              ' zuletzt gezeichnet: eid -> Array(x, y, w, h, gx, gy)
Private mRowTop() As Single, mRowH() As Single, mNRows As Long

'--- Kataloge -------------------------------------------------------------------------
Public Function ElemTypes() As Variant
    ElemTypes = Array("Kachel", "Balken", "Säulen", "Zeitverlauf", "Fläche", "Anteil", "Kreis", "Netz", "Trichter", "Baumdiagramm", "Pareto", "Rangliste", "Sankey", "Tabelle")
End Function

' Kurze Erklärung je Darstellung (Hilfetext im Dialog)
Public Function ElemTypeHelp(ByVal typ As String) As String
    Select Case typ
        Case "Kachel": ElemTypeHelp = "Eine große Zahl, z. B. „Offene Aufgaben: 50“."
        Case "Balken", "Säulen": ElemTypeHelp = "Vergleich von Gruppen (waagerecht bzw. senkrecht)."
        Case "Zeitverlauf", "Fläche": ElemTypeHelp = "Entwicklung über die Zeit, Gruppierung nach einem Datumsfeld."
        Case "Anteil", "Kreis": ElemTypeHelp = "Anteile am Ganzen (Ring bzw. Kreis), am besten bis 7 Gruppen."
        Case "Netz": ElemTypeHelp = "Netzdiagramm: Profil über mindestens 3 Gruppen, z. B. Bewertungen je Thema."
        Case "Trichter": ElemTypeHelp = "Stufen eines Ablaufs, z. B. Anfragen " & ChrW(&H2192) & " Angebote " & ChrW(&H2192) & " Aufträge (nach Wert sortiert)."
        Case "Baumdiagramm": ElemTypeHelp = "Flächen proportional zum Wert, viele Gruppen auf einen Blick."
        Case "Pareto": ElemTypeHelp = "Säulen absteigend plus Linie mit kumuliertem Anteil (80/20-Regel)."
        Case "Rangliste": ElemTypeHelp = "Beschriftete Fortschrittsbalken je Gruppe, größter Wert zuerst."
        Case "Sankey": ElemTypeHelp = "Flüsse zwischen zwei Gruppierungen, z. B. Status " & ChrW(&H2192) & " Priorität; die Breite der Bänder zeigt Anzahl bzw. Summe."
        Case "Tabelle": ElemTypeHelp = "Gruppen mit Wert als Liste."
    End Select
End Function

Public Function Measures() As Variant
    Measures = Array("ANZAHL", "GEFUELLT", "LEER", "EINDEUTIG", "SUMME", "MITTEL", "MEDIAN", "MIN", "MAX", "SPANNE", "ANTEIL")
End Function

Public Function MeasureLabel(ByVal m As String) As String
    Select Case m
        Case "ANZAHL": MeasureLabel = "Anzahl Datensätze (eindeutig)"
        Case "GEFUELLT": MeasureLabel = "Anzahl ausgefüllter Werte"
        Case "LEER": MeasureLabel = "Anzahl leerer Werte"
        Case "EINDEUTIG": MeasureLabel = "Anzahl verschiedener Werte"
        Case "SUMME": MeasureLabel = "Summe"
        Case "MITTEL": MeasureLabel = "Durchschnitt"
        Case "MEDIAN": MeasureLabel = "Median (mittlerer Wert)"
        Case "MIN": MeasureLabel = "Minimum"
        Case "MAX": MeasureLabel = "Maximum"
        Case "SPANNE": MeasureLabel = "Spannweite (Maximum " & ChrW(&H2212) & " Minimum)"
        Case "ANTEIL": MeasureLabel = "Anteil in % (Grundgesamtheit wählbar)"
    End Select
End Function

Public Function MeasureNeedsField(ByVal m As String) As Boolean
    MeasureNeedsField = (m <> "ANZAHL" And m <> "ANTEIL")
End Function

Public Function MeasureNumeric(ByVal m As String) As Boolean
    Select Case m
        Case "SUMME", "MITTEL", "MEDIAN", "MIN", "MAX", "SPANNE": MeasureNumeric = True
    End Select
End Function

Public Function ElemGet(ByVal eid As String, ByVal col As String) As String
    ElemGet = mCfg.GetS(SH_ELM, eid, col)
End Function

'--- Gestaltung: Stile und Farbpaletten -------------------------------------------------
Public Function StyleNames() As Variant
    StyleNames = Array("Karten", "Flach", "Farbakzent", "Kontur")
End Function

' Farbpaletten = gemeinsame Farbstile (mDocStyle); frühere Namen bleiben lesbar
Public Function PaletteNames() As Variant
    PaletteNames = mDocStyle.StyleNames()
End Function

Public Function DashStyle(ByVal did As String) As String
    DashStyle = mCfg.SysGet("DashStil_" & did, "Karten")
End Function

Public Function DashPaletteName(ByVal did As String) As String
    DashPaletteName = mCfg.SysGet("DashPalette_" & did, mDocStyle.DefaultStyle())
End Function

' Palette eines Elements: eigene Wahl oder die des Dashboards
Public Function ElemPaletteName(ByVal eid As String, ByVal did As String) As String
    ElemPaletteName = KVGet(ElemGet(eid, "Optionen"), "pal", "")
    If Len(ElemPaletteName) = 0 Then ElemPaletteName = DashPaletteName(did)
End Function

' Farben einer Palette (VBA-RGB), mindestens 8
Public Function PaletteColors(ByVal nm As String) As Variant
    Dim a As Long
    If mDocStyle.IsStyleName(nm) Then PaletteColors = mDocStyle.StyleColors(nm): Exit Function
    Select Case nm
        Case "Sonnenuntergang": PaletteColors = Array(RGB(234, 88, 12), RGB(244, 63, 94), RGB(251, 146, 60), RGB(219, 39, 119), RGB(250, 204, 21), RGB(190, 18, 60), RGB(253, 186, 116), RGB(157, 23, 77))
        Case "Pastell": PaletteColors = Array(RGB(147, 197, 253), RGB(252, 165, 165), RGB(134, 239, 172), RGB(253, 224, 71), RGB(196, 181, 253), RGB(249, 168, 212), RGB(103, 232, 249), RGB(253, 186, 116))
        Case "Kontrast": PaletteColors = Array(RGB(0, 114, 178), RGB(230, 159, 0), RGB(0, 158, 115), RGB(204, 121, 167), RGB(86, 180, 233), RGB(213, 94, 0), RGB(240, 228, 66), RGB(0, 0, 0))
        Case "Graustufen": PaletteColors = Array(RGB(31, 41, 55), RGB(75, 85, 99), RGB(107, 114, 128), RGB(156, 163, 175), RGB(55, 65, 81), RGB(209, 213, 219), RGB(17, 24, 39), RGB(229, 231, 235))
        Case Else
            a = CAccent()
            PaletteColors = Array(a, Blend(a, RGB(255, 255, 255), 0.35), Blend(a, RGB(0, 0, 0), 0.3), Blend(a, RGB(255, 255, 255), 0.6), _
                                  Blend(a, RGB(0, 0, 0), 0.5), Blend(a, RGB(255, 255, 255), 0.2), Blend(a, RGB(255, 255, 255), 0.75), Blend(a, RGB(0, 0, 0), 0.15))
    End Select
End Function

Public Function PaletteColor(ByVal nm As String, ByVal i As Long) As Long
    Dim p As Variant
    p = PaletteColors(nm)
    PaletteColor = p(i Mod (UBound(p) + 1))
End Function

'==============================================================================
' Berechnung
'==============================================================================
' Ergebnis: cMap mit Schluesseln
'   labels, keys, values, counts (Arrays 0..n-1), n, total (eindeutig), totalValue, notes, dashApplied, error, unitField
Public Function Compute(ByVal eid As String, ByVal did As String) As cMap
    Dim res As New cMap, tid As String, t As cTableData, rows() As Long, n As Long, k As Long, r As Long, conds As Collection
    Dim measure As String, mf As cField, gf As cField, level As String, dfid As String, dvon As Variant, dbis As Variant
    Dim dashF As String, dashT As String, applied As String, notes As String, groups As New cMap, order As New Collection
    Dim gk As String, gl As String, go As Double, vals As Variant, i As Long, v As Variant, acc As Variant, keysA As Variant
    Dim multiG As Boolean, sumGroups As Double, blanks As Long, baseNo As Double, gKeys() As String, gList As String, j As Long
    Dim ordArr() As Variant, idx() As Long, sk(1 To 1) As Variant, dr(1 To 1) As Long, sortMode As String, topN As Long, ng As Long
    Dim totalV As Variant, totalCnt As Long, hasRows As Boolean

    tid = ElemGet(eid, "Tabelle")
    If Not mCfg.TableExists(tid) Then res.SetVal "error", "Die Datenquelle wurde entfernt.": Set Compute = res: Exit Function
    measure = ElemGet(eid, "Kennzahl")
    If Len(measure) = 0 Then measure = "ANZAHL"
    If mCfg.FieldExists(ElemGet(eid, "Feld")) Then Set mf = mCfg.Field(ElemGet(eid, "Feld"))
    If MeasureNeedsField(measure) And mf Is Nothing Then res.SetVal "error", "Für diese Kennzahl ist ein Feld nötig.": Set Compute = res: Exit Function
    If mCfg.FieldExists(ElemGet(eid, "Gruppierung")) Then Set gf = mCfg.Field(ElemGet(eid, "Gruppierung"))
    level = ElemGet(eid, "Datumsstufe")
    If Len(level) = 0 Then level = "Monat"
    dfid = ElemGet(eid, "Datumsfeld")
    sortMode = KVGet(ElemGet(eid, "Optionen"), "sort", "wert")
    topN = CLng(Val(KVGet(ElemGet(eid, "Optionen"), "top", "12")))
    If topN < 1 Then topN = 12

    ' Zeilen: aktive Datensaetze + Elementfilter
    Set conds = mQuery.ParseConds(ElemGet(eid, "Filter"))
    ' Dashboardfilter
    dashF = mCfg.GetS(SH_DSB, did, "Filter")
    dvon = InvDate(KVGet(dashF, "von")): dbis = InvDate(KVGet(dashF, "bis"))
    dashT = KVGet(dashF, "tab")
    If dashT = tid And Len(KVGet(dashF, "f")) > 0 Then
        Dim dc As Variant, dconds As Collection
        Set dconds = mQuery.ParseConds(KVGet(dashF, "f"))
        For Each dc In dconds
            If conds.Count > 0 And dc(0) = "" Then dc(0) = "UND"
            conds.Add Array(IIf(conds.Count = 0, "", "UND"), dc(1), dc(2), dc(3), dc(4))
        Next
        applied = "Tabellenfilter"
    End If
    If (Not IsEmpty(dvon) Or Not IsEmpty(dbis)) And mCfg.FieldExists(dfid) Then
        If Not IsEmpty(dvon) And Not IsEmpty(dbis) Then
            conds.Add Array(IIf(conds.Count = 0, "", "UND"), dfid, "ZWISCHEN", DateInv(CDate(dvon)), DateInv(CDate(dbis)))
        ElseIf Not IsEmpty(dvon) Then
            conds.Add Array(IIf(conds.Count = 0, "", "UND"), dfid, "NACH", DateInv(CDate(dvon) - 1), "")
        Else
            conds.Add Array(IIf(conds.Count = 0, "", "UND"), dfid, "VOR", DateInv(CDate(dbis) + 1), "")
        End If
        applied = applied & IIf(Len(applied) > 0, " und ", "") & "Zeitraum"
    End If
    If Len(applied) > 0 Then
        res.SetVal "dashApplied", "Dashboardfilter angewendet (" & applied & ")"
    ElseIf Len(dashF) > 0 And (Len(KVGet(dashF, "von")) > 0 Or Len(KVGet(dashF, "bis")) > 0 Or Len(dashT) > 0) Then
        res.SetVal "dashApplied", "Nicht vom Dashboardfilter betroffen (anderes Datum/andere Tabelle)"
        res.SetVal "dashIgnored", True
    End If
    rows = mQuery.FilterRows(tid, conds, "", New Collection, mQuery.ARCH_ACTIVE, n)
    Set t = mData.TData(tid)
    res.SetVal "total", n
    notes = "Basis: " & Plural(n, "Datensatz", "Datensätze") & " aus „" & mCfg.TableName(tid) & "“ (archivierte ausgeschlossen)"
    If conds.Count > 0 Then notes = notes & ", gefiltert"

    ' Gesamtwert (eindeutig)
    acc = NewAcc()
    For k = 1 To n
        AccAdd acc, t, rows(k), mf, measure
    Next
    totalV = AccResult(acc, measure, n, n)
    res.SetVal "totalValue", totalV
    If MeasureNumeric(measure) Then
        If acc(3) > 0 Then notes = notes & "; " & acc(3) & " ohne Wert nicht berücksichtigt"
    End If

    If gf Is Nothing Then
        res.SetVal "n", 0
        res.SetVal "notes", notes
        res.SetVal "measure", measure
        If Not mf Is Nothing Then res.SetVal "unitField", mf.ID
        Set Compute = res
        Exit Function
    End If

    ' Gruppierung (ein Datensatz kann bei Mehrfachfeldern mehreren Gruppen angehoeren)
    multiG = (gf.FType = FT_MEHRFACH Or (gf.IsLinkType And (gf.Multi Or gf.FType = FT_BACKLINK)))
    For k = 1 To n
        r = rows(k)
        gList = GroupKeysOf(gf, t, r, level)
        For Each v In Split(gList, Chr$(30))
            If groups.Exists(CStr(v)) Then
                acc = groups.GetVal(CStr(v))
            Else
                acc = NewAcc()
                order.Add CStr(v)
            End If
            AccAdd acc, t, r, mf, measure
            groups.SetVal CStr(v), acc
        Next
    Next
    ng = order.Count
    ReDim gKeys(0 To Application.Max(0, ng - 1))
    ReDim ordArr(1 To Application.Max(1, ng))
    ReDim idx(1 To Application.Max(1, ng))
    Dim labelsA() As Variant, valuesA() As Variant, countsA() As Variant
    ReDim labelsA(0 To Application.Max(0, ng - 1)): ReDim valuesA(0 To Application.Max(0, ng - 1)): ReDim countsA(0 To Application.Max(0, ng - 1))
    i = 0
    For Each v In order
        acc = groups.GetVal(CStr(v))
        gKeys(i) = CStr(v)
        labelsA(i) = GroupLabel(gf, CStr(v), level)
        countsA(i) = acc(0)
        sumGroups = sumGroups + acc(0)
        If measure = "ANTEIL" Then
            baseNo = n
            If ElemGet(eid, "Basis") = "GEFUELLT" Then baseNo = n - CountNoGroup(groups)
            If baseNo > 0 Then valuesA(i) = acc(0) / baseNo Else valuesA(i) = Empty
            If CStr(v) = "" And ElemGet(eid, "Basis") = "GEFUELLT" Then valuesA(i) = Empty
        Else
            valuesA(i) = AccResult(acc, measure, acc(0), n)
        End If
        i = i + 1
        idx(i) = i
        If sortMode = "name" Or gf.IsDate_ Or gf.IsSingleOption Then
            ordArr(i) = GroupOrder(gf, CStr(v), level)
        Else
            If IsNumeric(valuesA(i - 1)) And Not IsBlank(valuesA(i - 1)) Then ordArr(i) = -CDbl(valuesA(i - 1)) Else ordArr(i) = 1E+30
        End If
    Next
    If ng > 1 Then
        sk(1) = ordArr: dr(1) = 1
        MergeSortIdx idx, sk, 1, dr
    End If
    Dim L2() As Variant, V2() As Variant, C2() As Variant, K2() As Variant, nOut As Long, restCnt As Long, restV As Double, restOk As Boolean
    nOut = Application.Min(ng, topN)
    If gf.IsDate_ Then nOut = ng
    ReDim L2(0 To Application.Max(0, nOut)): ReDim V2(0 To Application.Max(0, nOut)): ReDim C2(0 To Application.Max(0, nOut)): ReDim K2(0 To Application.Max(0, nOut))
    For i = 1 To ng
        j = idx(i) - 1
        If i <= nOut Then
            L2(i - 1) = labelsA(j): V2(i - 1) = valuesA(j): C2(i - 1) = countsA(j): K2(i - 1) = gKeys(j)
        Else
            restCnt = restCnt + 1
        End If
    Next
    If restCnt > 0 Then notes = notes & "; " & restCnt & " weitere Gruppen nicht dargestellt (Top " & nOut & ")"
    If multiG And sumGroups > n Then
        notes = notes & "; Mehrfachzuordnung: Ein Datensatz kann mehreren Gruppen angehören; die Summe der Gruppen (" & FmtInt(sumGroups) & _
                ") ist größer als die eindeutige Gesamtzahl (" & FmtInt(n) & "). Gruppenwerte nicht addieren."
    End If
    If measure = "ANTEIL" Then notes = notes & "; Anteile bezogen auf " & IIf(ElemGet(eid, "Basis") = "GEFUELLT", "Datensätze mit Gruppenangabe", "alle gefilterten Datensätze") & _
                                     IIf(multiG, "; Anteile können sich auf mehr als 100 % summieren", "")
    If ElemGet(eid, "Typ") = "Sankey" Then SankeyFlows eid, res, t, rows, n, gf, mf, measure, level, K2, nOut
    res.SetVal "n", nOut
    res.SetVal "labels", L2
    res.SetVal "values", V2
    res.SetVal "counts", C2
    res.SetVal "keys", K2
    res.SetVal "notes", notes
    res.SetVal "measure", measure
    res.SetVal "multi", multiG
    res.SetVal "groupField", gf.ID
    If Not mf Is Nothing Then res.SetVal "unitField", mf.ID
    Set Compute = res
End Function

' Sankey: Flüsse von den Gruppen der ersten Gruppierung (links, wie im Ergebnis) zu den Gruppen
' einer zweiten Gruppierung (rechts, Option "ziel"). Ergebnis "sankey" = Array(Beschriftungen rechts,
' Flüsse als Array(Index links, Index rechts, Wert)). Rechts höchstens 8 Gruppen, Rest "Weitere".
Private Sub SankeyFlows(ByVal eid As String, ByVal res As cMap, ByVal t As cTableData, rows() As Long, ByVal n As Long, _
                        ByVal gf As cField, ByVal mf As cField, ByVal measure As String, ByVal level As String, K2 As Variant, ByVal nOut As Long)
    Dim zf As String, gf2 As cField, aIdx As New cMap, fl As New cMap, bTot As New cMap, k As Long, r As Long, va As Variant, vb As Variant
    Dim acc As Variant, key As Variant, p As Variant, i As Long, j As Long, nb As Long, bKeys() As String, bVals() As Double, tmpS As String, tmpD As Double
    Dim bPos As New cMap, bLabels() As Variant, flows() As Variant, nf As Long, v As Variant, merged As New cMap, kb As String, bList As Variant
    zf = KVGet(ElemGet(eid, "Optionen"), "ziel", "")
    If Not mCfg.FieldExists(zf) Then res.SetVal "error", "Für das Sankey-Diagramm fehlt die zweite Gruppierung („Nach“).": Exit Sub
    Set gf2 = mCfg.Field(zf)
    For i = 0 To nOut - 1
        aIdx.SetVal CStr(K2(i)), i
    Next
    For k = 1 To n
        r = rows(k)
        For Each va In Split(GroupKeysOf(gf, t, r, level), Chr$(30))
            If aIdx.Exists(CStr(va)) Then
                ' leere Angabe rechts als eigene Gruppe "(ohne Angabe)" statt verloren
                kb = GroupKeysOf(gf2, t, r, level)
                If Len(kb) = 0 Then bList = Array("") Else bList = Split(kb, Chr$(30))
                For Each vb In bList
                    key = CStr(va) & Chr$(31) & CStr(vb)
                    If fl.Exists(CStr(key)) Then acc = fl.GetVal(CStr(key)) Else acc = NewAcc()
                    AccAdd acc, t, r, mf, measure
                    fl.SetVal CStr(key), acc
                Next
            End If
        Next
    Next
    ' Werte je Fluss und Summen rechts
    For Each key In fl.Keys()
        acc = fl.GetVal(CStr(key))
        v = AccResult(acc, measure, acc(0), n)
        If IsNumeric(v) And Not IsBlank(v) Then
            If CDbl(v) > 0 Then
                p = Split(CStr(key), Chr$(31))
                bTot.AddNum CStr(p(1)), CDbl(v)
                merged.SetVal CStr(key), CDbl(v)
            End If
        End If
    Next
    ' rechte Seite nach Summe absteigend; ab der 9. Gruppe zusammengefasst
    nb = bTot.Count
    If nb = 0 Then res.SetVal "sankey", Empty: Exit Sub
    ReDim bKeys(1 To nb): ReDim bVals(1 To nb)
    i = 0
    For Each key In bTot.Keys()
        i = i + 1: bKeys(i) = CStr(key): bVals(i) = bTot.GetVal(CStr(key))
    Next
    For i = 1 To nb - 1
        For j = i + 1 To nb
            If bVals(j) > bVals(i) Then
                tmpS = bKeys(i): bKeys(i) = bKeys(j): bKeys(j) = tmpS
                tmpD = bVals(i): bVals(i) = bVals(j): bVals(j) = tmpD
            End If
        Next
    Next
    ReDim bLabels(0 To Application.Min(nb, 9) - 1)
    For i = 1 To nb
        If i <= 8 Or nb = 9 Then
            bPos.SetVal bKeys(i), i - 1
            bLabels(i - 1) = GroupLabel(gf2, bKeys(i), level)
        Else
            bPos.SetVal bKeys(i), 8
            bLabels(8) = "Weitere (" & (nb - 8) & ")"
        End If
    Next
    ' Flüsse (gleiche Ziele nach dem Zusammenfassen addieren)
    Dim sumF As New cMap
    For Each key In merged.Keys()
        p = Split(CStr(key), Chr$(31))
        sumF.AddNum aIdx.GetVal(CStr(p(0))) & "|" & bPos.GetVal(CStr(p(1))), CDbl(merged.GetVal(CStr(key)))
    Next
    ReDim flows(0 To sumF.Count - 1)
    nf = 0
    For Each key In sumF.Keys()
        p = Split(CStr(key), "|")
        flows(nf) = Array(CLng(p(0)), CLng(p(1)), CDbl(sumF.GetVal(CStr(key))))
        nf = nf + 1
    Next
    res.SetVal "sankey", Array(bLabels, flows)
    res.SetVal "sankeyTo", gf2.Name
End Sub

Private Function CountNoGroup(ByVal groups As cMap) As Long
    Dim a As Variant
    If groups.Exists("") Then
        a = groups.GetVal("")
        CountNoGroup = a(0)
    End If
End Function

' Akkumulator: (0) Anzahl, (1) Summe, (2) gezaehlte Werte, (3) leere, (4) Min, (5) Max, (6) ausgefuellt,
' (7) Zahlenwerte fuer den Median (Collection), (8) verschiedene Werte (cMap), nur bei Bedarf angelegt
Private Function NewAcc() As Variant
    NewAcc = Array(0#, 0#, 0#, 0#, Empty, Empty, 0#, Empty, Empty)
End Function

Private Sub AccAdd(ByRef acc As Variant, ByVal t As cTableData, ByVal r As Long, ByVal mf As cField, ByVal measure As String)
    Dim v As Variant
    acc(0) = acc(0) + 1
    If mf Is Nothing Then Exit Sub
    If mf.IsLinkType Then
        v = mData.LinkTargets(mf.ID, t.IdAt(r))
    Else
        v = t.Val_(r, mf.ID)
    End If
    If IsBlank(v) Or IsErrVal(v) Then
        acc(3) = acc(3) + 1
        Exit Sub
    End If
    acc(6) = acc(6) + 1
    If measure = "EINDEUTIG" Then
        If IsEmpty(acc(8)) Then Set acc(8) = New cMap
        acc(8).SetVal LCase$(CStr(v)), True
    End If
    If IsNumeric(v) Or IsDate(v) Then
        If VarType(v) <> vbBoolean And VarType(v) <> vbString Then
            If measure = "MEDIAN" Then
                If IsEmpty(acc(7)) Then Set acc(7) = New Collection
                acc(7).Add CDbl(v)
            End If
            acc(1) = acc(1) + CDbl(v)
            acc(2) = acc(2) + 1
            If IsEmpty(acc(4)) Then
                acc(4) = CDbl(v)
            ElseIf CDbl(v) < acc(4) Then
                acc(4) = CDbl(v)
            End If
            If IsEmpty(acc(5)) Then
                acc(5) = CDbl(v)
            ElseIf CDbl(v) > acc(5) Then
                acc(5) = CDbl(v)
            End If
        End If
    End If
End Sub

Private Function AccResult(ByVal acc As Variant, ByVal measure As String, ByVal cnt As Double, ByVal total As Double) As Variant
    Select Case measure
        Case "ANZAHL": AccResult = acc(0)
        Case "GEFUELLT": AccResult = acc(6)
        Case "LEER": AccResult = acc(3)
        Case "SUMME": If acc(2) > 0 Then AccResult = acc(1) Else AccResult = Empty
        Case "MITTEL": If acc(2) > 0 Then AccResult = acc(1) / acc(2) Else AccResult = Empty
        Case "MIN": AccResult = acc(4)
        Case "MAX": AccResult = acc(5)
        Case "SPANNE": If acc(2) > 0 Then AccResult = acc(5) - acc(4) Else AccResult = Empty
        Case "MEDIAN": If IsObject(acc(7)) Then AccResult = MedianOf(acc(7)) Else AccResult = Empty
        Case "EINDEUTIG": If IsObject(acc(8)) Then AccResult = CDbl(acc(8).Count) Else AccResult = 0#
        Case "ANTEIL": If total > 0 Then AccResult = acc(6) / total Else AccResult = Empty
    End Select
End Function

Private Function MedianOf(ByVal c As Collection) As Variant
    Dim a() As Double, i As Long, j As Long, n As Long, tmp As Double
    n = c.Count
    If n = 0 Then MedianOf = Empty: Exit Function
    ReDim a(1 To n)
    For i = 1 To n: a(i) = c(i): Next
    ' Einfuegesortierung genuegt fuer Gruppen; grosse Mengen ueber Shell-Abstaende
    Dim gap As Long
    gap = n \ 2
    Do While gap > 0
        For i = gap + 1 To n
            tmp = a(i): j = i
            Do While j > gap
                If a(j - gap) <= tmp Then Exit Do
                a(j) = a(j - gap): j = j - gap
            Loop
            a(j) = tmp
        Next
        gap = gap \ 2
    Loop
    If n Mod 2 = 1 Then MedianOf = a((n + 1) \ 2) Else MedianOf = (a(n \ 2) + a(n \ 2 + 1)) / 2
End Function

' Gruppenschluessel eines Datensatzes (mehrere durch Chr(30) getrennt)
Private Function GroupKeysOf(ByVal gf As cField, ByVal t As cTableData, ByVal r As Long, ByVal level As String) As String
    Dim v As Variant, k As String, l As String, o As Double, a As Variant, i As Long, s As String
    If gf.IsLinkType Then
        s = mData.LinkTargets(gf.ID, t.IdAt(r))
        GroupKeysOf = Replace(s, ";", Chr$(30))
        Exit Function
    End If
    v = t.Val_(r, gf.ID)
    If gf.FType = FT_MEHRFACH Then
        If IsBlank(v) Then GroupKeysOf = "": Exit Function
        GroupKeysOf = Replace(CStr(v), ";", Chr$(30))
        Exit Function
    End If
    mQuery.GroupOf gf, v, t.IdAt(r), level, k, l, o
    GroupKeysOf = k
End Function

Public Function GroupLabel(ByVal gf As cField, ByVal key As String, ByVal level As String) As String
    Dim tt As String, d As Date
    If Len(key) = 0 Then GroupLabel = "Ohne Angabe": Exit Function
    If gf.IsLinkType Then
        If gf.FType = FT_LINK Then tt = gf.TargetTable Else tt = mCfg.Field(gf.PairField).TableID
        GroupLabel = mData.RecordLabel(tt, key)
        Exit Function
    End If
    Select Case gf.ValueType
        Case FT_AUSWAHL, FT_STATUS, FT_MEHRFACH: GroupLabel = mCfg.OptName(key)
        Case FT_JANEIN: GroupLabel = IIf(key = "J", "Ja", "Nein")
        Case FT_DATUM, FT_DATUMZEIT
            Select Case level
                Case "Jahr": GroupLabel = key
                Case "Quartal": GroupLabel = Mid$(key, 6) & ". Quartal " & Left$(key, 4)
                Case "Tag": GroupLabel = FmtDate(CDate(CLng(key)))
                Case Else: GroupLabel = Left$(MonthName_DE(CLng(Mid$(key, 6))), 3) & " " & Left$(key, 4)
            End Select
        Case FT_ZAHL, FT_GELD, FT_PROZENT: GroupLabel = mFmt.DisplayText(gf, InvNum(key))
        Case Else: GroupLabel = key
    End Select
    If Len(GroupLabel) = 0 Then GroupLabel = key
End Function

Private Function GroupOrder(ByVal gf As cField, ByVal key As String, ByVal level As String) As Variant
    If Len(key) = 0 Then GroupOrder = 1E+30: Exit Function
    Select Case gf.ValueType
        Case FT_AUSWAHL, FT_STATUS, FT_MEHRFACH: GroupOrder = CDbl(mCfg.OptOrderIndex(gf.ID, key))
        Case FT_DATUM, FT_DATUMZEIT
            Select Case level
                Case "Jahr": GroupOrder = CDbl(Val(key))
                Case "Quartal": GroupOrder = CDbl(Val(Left$(key, 4))) * 10 + Val(Mid$(key, 6))
                Case "Tag": GroupOrder = CDbl(Val(key))
                Case Else: GroupOrder = CDbl(Val(Left$(key, 4))) * 100 + Val(Mid$(key, 6))
            End Select
        Case FT_ZAHL, FT_GELD, FT_PROZENT: GroupOrder = InvNum(key)
        Case Else: GroupOrder = LCase$(GroupLabel(gf, key, level))
    End Select
End Function

Public Function FormatMeasure(ByVal res As cMap, ByVal v As Variant) As String
    Dim m As String, f As cField
    If IsBlank(v) Then FormatMeasure = "keine Werte": Exit Function
    m = CStr(res.GetVal("measure", "ANZAHL"))
    Select Case m
        Case "ANZAHL", "GEFUELLT", "LEER", "EINDEUTIG": FormatMeasure = FmtInt(CDbl(v))
        Case "ANTEIL": FormatMeasure = FmtNum(CDbl(v) * 100, 1) & " %"
        Case Else
            If mCfg.FieldExists(CStr(res.GetVal("unitField", ""))) Then
                Set f = mCfg.Field(CStr(res.GetVal("unitField")))
                If m = "SPANNE" And f.IsDate_ Then
                    FormatMeasure = FmtNum(CDbl(v), 0) & " Tage"
                ElseIf (m = "MITTEL" Or m = "MEDIAN") And f.ValueType = FT_ZAHL Then
                    FormatMeasure = FmtNum(CDbl(v), 2, True)
                Else
                    FormatMeasure = mFmt.DisplayText(f, v)
                End If
            Else
                FormatMeasure = FmtNum(CDbl(v), 2, True)
            End If
    End Select
End Function

' Datensaetze hinter einer Kennzahl/Gruppe
Public Function DrillRids(ByVal eid As String, ByVal did As String, ByVal groupKey As Variant) As Collection
    Dim res As New Collection, tid As String, t As cTableData, rows() As Long, n As Long, k As Long, conds As Collection, gf As cField
    Dim dashF As String, dvon As Variant, dbis As Variant, dfid As String, keysS As String, level As String, dashT As String, dc As Variant
    tid = ElemGet(eid, "Tabelle")
    Set conds = mQuery.ParseConds(ElemGet(eid, "Filter"))
    dashF = mCfg.GetS(SH_DSB, did, "Filter")
    dashT = KVGet(dashF, "tab")
    If dashT = tid And Len(KVGet(dashF, "f")) > 0 Then
        For Each dc In mQuery.ParseConds(KVGet(dashF, "f"))
            conds.Add Array(IIf(conds.Count = 0, "", "UND"), dc(1), dc(2), dc(3), dc(4))
        Next
    End If
    dvon = InvDate(KVGet(dashF, "von")): dbis = InvDate(KVGet(dashF, "bis"))
    dfid = ElemGet(eid, "Datumsfeld")
    If mCfg.FieldExists(dfid) Then
        If Not IsEmpty(dvon) And Not IsEmpty(dbis) Then
            conds.Add Array(IIf(conds.Count = 0, "", "UND"), dfid, "ZWISCHEN", DateInv(CDate(dvon)), DateInv(CDate(dbis)))
        ElseIf Not IsEmpty(dvon) Then
            conds.Add Array(IIf(conds.Count = 0, "", "UND"), dfid, "NACH", DateInv(CDate(dvon) - 1), "")
        ElseIf Not IsEmpty(dbis) Then
            conds.Add Array(IIf(conds.Count = 0, "", "UND"), dfid, "VOR", DateInv(CDate(dbis) + 1), "")
        End If
    End If
    rows = mQuery.FilterRows(tid, conds, "", New Collection, mQuery.ARCH_ACTIVE, n)
    Set t = mData.TData(tid)
    If mCfg.FieldExists(ElemGet(eid, "Gruppierung")) Then Set gf = mCfg.Field(ElemGet(eid, "Gruppierung"))
    level = ElemGet(eid, "Datumsstufe")
    If Len(level) = 0 Then level = "Monat"
    For k = 1 To n
        If IsEmpty(groupKey) Or gf Is Nothing Then
            res.Add t.IdAt(rows(k))
        Else
            keysS = Chr$(30) & GroupKeysOf(gf, t, rows(k), level) & Chr$(30)
            If CStr(groupKey) = "" Then
                If keysS = Chr$(30) & Chr$(30) Then res.Add t.IdAt(rows(k))
            ElseIf InStr(1, keysS, Chr$(30) & CStr(groupKey) & Chr$(30)) > 0 Then
                res.Add t.IdAt(rows(k))
            End If
        End If
    Next
    Set DrillRids = res
End Function

Public Sub ShowRecordList(ByVal tid As String, ByVal rids As Collection, ByVal title As String)
    Dim items() As Variant, ids() As String, i As Long, v As Variant, f As frmMsg, sel As Long, n As Long
    n = rids.Count
    If n = 0 Then mUI.Alert "Keine Datensätze.", title: Exit Sub
    If n > 500 Then n = 500
    ReDim items(0 To n - 1, 0 To 1): ReDim ids(0 To n - 1)
    For Each v In rids
        If i >= n Then Exit For
        items(i, 0) = mData.RecordLabel(tid, CStr(v))
        items(i, 1) = mData.RecordHint(tid, CStr(v))
        ids(i) = CStr(v)
        i = i + 1
    Next
    Set f = New frmMsg
    f.SetupList title, Plural(rids.Count, "Datensatz", "Datensätze") & IIf(rids.Count > 500, " (die ersten 500 angezeigt)", "") & _
                ". Doppelklick oder „Öffnen“ zeigt den Eintrag.", items, False, "Öffnen", 2, "230;220", 0
    f.Show: mNav.FocusGrid
    sel = -1
    If f.Result = 1 Then sel = f.SelectedIndex
    Unload f
    If sel >= 0 Then mForms.OpenRecord tid, ids(sel)
End Sub

'==============================================================================
' Darstellung
'==============================================================================
Public Sub Render()
    Dim ws As Worksheet, ids As Variant, did As String, i As Long, eids As Variant, x0 As Single, W As Single, unitW As Single, gap As Single
    Dim x As Single, y As Single, rowH As Single, used As Long, sz As Long, eh As Single, ew As Single, e As Long, res As cMap, s As Shape
    Dim dashF As String, txt As String, rowTop As Single, maxBottom As Single, needRows As Long, stamp As String

    Set ws = mRender.Sheet(SH_DASH)
    mRender.BeginPage ws
    Set mDrill = New cMap
    mDataCol = DASH_DATA_COL
    did = mNav.gDash
    ids = mCfg.DashIds()
    mRender.DrawTitle ws, ChrW(&HE9D2), "Dashboards", IIf(Len(did) > 0, mCfg.GetS(SH_DSB, did, "Beschreibung"), "Auswertungen über Ihre Tabellen")
    ' Register
    mRender.FlowBegin ws, R_TABS, mRender.ContentLeft(ws), mRender.PageRight()
    For i = 0 To ArrLen(ids) - 1
        mRender.FlowBtn "dtab_" & ids(i), mCfg.GetS(SH_DSB, CStr(ids(i)), "Name"), "mCmd.Cmd_DashTab", IIf(CStr(ids(i)) = did, "tabActive", "tab"), ChrW(&HE9D2), 2
    Next
    mRender.FlowGap 6
    mRender.FlowBtn "dtab_add", "Dashboard", "mCmd.Cmd_DashNew", "tab", ChrW(&HE710)
    mRender.FlowEnd
    If Len(did) = 0 Then
        mRender.EmptyState ws, R_DATA, "Noch kein Dashboard", "Dashboards bündeln Kennzahlen, Diagramme und Ergebnistabellen über Ihre Tabellen. " & _
                           "Legen Sie ein Dashboard an und fügen Sie Elemente hinzu.", "Dashboard anlegen", "mCmd.Cmd_DashNew", ChrW(&HE9D2)
        mRender.EndPage ws
        Exit Sub
    End If
    ' Werkzeuge
    mRender.FlowBegin ws, R_TOOLS, mRender.ContentLeft(ws), mRender.PageRight()
    If gArrange Then
        mRender.FlowBtn "b_arrdone", "Fertig", "mCmd.Cmd_DashArrangeDone", "primary", ChrW(&HE73E)
        mRender.FlowBtn "b_arrsnap", "Einrasten", "mCmd.Cmd_DashArrangeSnap", "secondary", ChrW(&HE80A)
        mRender.FlowBtn "b_arrcancel", "Abbrechen", "mCmd.Cmd_DashArrangeCancel", "secondary", ChrW(&HE711)
        ws.Cells(R_CHIPS, C_CONTENT).Value = "Karte ziehen und loslassen: sie rastet sofort in einer der 4 Spalten ein. " & _
                                             "Breite am Rand ziehen (¼ bis ganze Breite), Diagramme unten ziehen = hoch/normal."
        ws.Cells(R_CHIPS, C_CONTENT).Font.Color = CAccent()
    Else
        mRender.FlowBtn "b_add", "Element hinzufügen", "mCmd.Cmd_DashAddElement", "primary", ChrW(&HE710)
        mRender.FlowBtn "b_arrange", "Anordnen", "mCmd.Cmd_DashArrange", "secondary", ChrW(&HE8B0)
        mRender.FlowBtn "b_style", "Gestaltung", "mCmd.Cmd_DashStyle", "secondary", ChrW(&HE790)
        mRender.FlowBtn "b_filter", "Filter", "mCmd.Cmd_DashFilter", "secondary", ChrW(&HE71C)
        mRender.FlowBtn "b_refresh", "Aktualisieren", "mCmd.Cmd_DashRefresh", "secondary", ChrW(&HE72C)
        mRender.FlowBtn "b_more", "Mehr", "mCmd.Cmd_DashMore", "secondary", ChrW(&HE712)
    End If
    mRender.FlowEnd
    ' Filterzusammenfassung
    dashF = mCfg.GetS(SH_DSB, did, "Filter")
    mLastStamp = Now
    If Not gArrange Then
        mRender.FlowBegin ws, R_CHIPS, mRender.ContentLeft(ws), mRender.PageRight()
        stamp = "Stand: " & FmtDateTime(mLastStamp) & " · aktuell"
        mRender.FlowText "stamp", stamp, CMuted(), 9, True, 12
        txt = DashFilterText(did)
        If Len(txt) > 0 Then mRender.FlowBtn "chip_D", txt & "  " & ChrW(&HD7), "mCmd.Cmd_DashFilterClear", "chip"
        mRender.FlowEnd
    End If

    ' Raster: 4 Einheiten pro Zeile
    x0 = mRender.ContentLeft(ws)
    W = mRender.PageRight() - x0
    If W > 1400 Then W = 1400
    gap = 14
    unitW = (W - 3 * gap) / 4
    mUnitW = unitW: mGap = gap: mX0 = x0
    rowTop = ws.Rows(R_DATA).Top + 4
    eids = mCfg.ElementIds(did)
    If ArrLen(eids) = 0 Then
        EndArrange
        mRender.EmptyState ws, R_DATA, "Dieses Dashboard ist leer", "Mit „Element hinzufügen“ wählen Sie Datenquelle, Kennzahl, Gruppierung, Zeitraum und Filter, ohne Formeln.", _
                           "Element hinzufügen", "mCmd.Cmd_DashAddElement", ChrW(&HE9D2)
        mRender.EndPage ws
        Exit Sub
    End If
    ' Raster berechnen: Zeilenhoehe = hoechstes Element der Zeile
    Dim gx() As Long, gy() As Long, gsz() As Long, r As Long
    LayoutGrid eids, gx, gy, gsz, Nothing
    mNRows = 0
    For e = 0 To ArrLen(eids) - 1
        If gy(e) + 1 > mNRows Then mNRows = gy(e) + 1
    Next
    ReDim mRowTop(0 To mNRows): ReDim mRowH(0 To mNRows)
    For e = 0 To ArrLen(eids) - 1
        eh = ElemHeight(CStr(eids(e)))
        If eh > mRowH(gy(e)) Then mRowH(gy(e)) = eh
    Next
    y = rowTop
    For r = 0 To mNRows - 1
        mRowTop(r) = y
        y = y + mRowH(r) + gap
    Next
    mRowTop(mNRows) = y
    Set mLay = New cMap
    mApp.BeginWork "Dashboard wird berechnet …"
    If gArrange Then DrawGridGuides ws, x0, rowTop, unitW, gap, y + 120
    For e = 0 To ArrLen(eids) - 1
        sz = gsz(e)
        x = x0 + gx(e) * (unitW + gap)
        y = mRowTop(gy(e))
        ew = unitW * sz + gap * (sz - 1)
        eh = ElemHeight(CStr(eids(e)))
        mLay.SetVal CStr(eids(e)), Array(x, y, ew, eh, gx(e), gy(e))
        If gArrange Then
            DrawArrangeCard ws, CStr(eids(e)), x, y, ew, eh, sz
        Else
            Set res = Compute(CStr(eids(e)), did)
            DrawElement ws, CStr(eids(e)), did, res, x, y, ew, eh
        End If
        If y + eh > maxBottom Then maxBottom = y + eh
    Next
    mApp.EndWork
    If mDataCol > DASH_DATA_COL Then ws.Range(ws.Columns(DASH_DATA_COL + 1), ws.Columns(mDataCol + 2)).Hidden = True
    ' genug Zeilen fuer Druck/PDF
    needRows = R_DATA + CLng((maxBottom - rowTop) / 20) + 3
    ws.Cells(needRows, C_CONTENT).Value = " "
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, C_CONTENT), ws.Cells(needRows, C_CONTENT + 20)).Address
    ' Im Anordnen-Modus sind nur die Karten verschiebbar (nicht gesperrt), Schaltflaechen
    ' und Rasterhilfen bleiben durch den Blattschutz fest
    mRender.EndPage ws
End Sub

Public Function ElemSize(ByVal eid As String) As Long
    If gArrange And Not mArr Is Nothing Then
        If mArr.Exists(eid) Then ElemSize = mArr.GetVal(eid)(2): Exit Function
    End If
    ElemSize = CLng(Val(ElemGet(eid, "Groesse")))
    If ElemSize < 1 Or ElemSize > 4 Then ElemSize = IIf(ElemGet(eid, "Typ") = "Kachel", 1, 2)
End Function

Private Function ElemHoch(ByVal eid As String) As Boolean
    If gArrange And Not mArr Is Nothing Then
        If mArr.Exists(eid) Then ElemHoch = mArr.GetVal(eid)(3): Exit Function
    End If
    ElemHoch = (KVGet(ElemGet(eid, "Optionen"), "hoch", "") = "1")
End Function

Public Function ElemHeight(ByVal eid As String) As Single
    If ElemGet(eid, "Typ") = "Kachel" Then
        ElemHeight = 104
    ElseIf ElemHoch(eid) Then
        ElemHeight = 400
    Else
        ElemHeight = 270
    End If
End Function

' Gewuenschte Rasterposition: Array(gx, gy) als Double (-1 = noch keine, fliesst ans Ende)
Private Function ElemWish(ByVal eid As String) As Variant
    Dim o As String
    If gArrange And Not mArr Is Nothing Then
        If mArr.Exists(eid) Then ElemWish = Array(CDbl(mArr.GetVal(eid)(0)), CDbl(mArr.GetVal(eid)(1))): Exit Function
    End If
    o = ElemGet(eid, "Optionen")
    If Len(KVGet(o, "gx", "")) > 0 And Len(KVGet(o, "gy", "")) > 0 Then
        ElemWish = Array(Val(KVGet(o, "gx", "0")), Val(KVGet(o, "gy", "0")))
    Else
        ElemWish = Array(-1#, -1#)
    End If
End Function

' Ordnet alle Elemente ueberschneidungsfrei im 4-Spalten-Raster an. Elemente aus pri
' (gerade gezogen) bekommen ihren Wunschplatz zuerst; verdraengte rutschen nach rechts
' bzw. in die naechste Zeile. Elemente ohne Position fliessen wie bisher ans Ende.
Private Sub LayoutGrid(ByVal eids As Variant, ByRef gx() As Long, ByRef gy() As Long, ByRef gsz() As Long, ByVal pri As cMap)
    Dim n As Long, i As Long, j As Long, k As Long, wx() As Double, wy() As Double, has() As Boolean, ord() As Long, m As Long
    Dim ys() As Double, nys As Long, r As Long, c As Long, occ() As Boolean, maxR As Long, placed As Boolean, d As Long, cc As Long
    Dim fr As Long, fc As Long, used() As Boolean, newR() As Long, nr As Long, w As Variant, tmp As Long, pass As Long
    n = ArrLen(eids)
    ReDim gx(0 To Application.Max(0, n - 1)): ReDim gy(0 To Application.Max(0, n - 1)): ReDim gsz(0 To Application.Max(0, n - 1))
    If n = 0 Then Exit Sub
    ReDim wx(0 To n - 1): ReDim wy(0 To n - 1): ReDim has(0 To n - 1): ReDim ord(0 To n - 1)
    For i = 0 To n - 1
        gsz(i) = ElemSize(CStr(eids(i)))
        w = ElemWish(CStr(eids(i)))
        wx(i) = w(0): wy(i) = w(1)
        has(i) = (wx(i) >= 0 And wy(i) >= -1)
    Next
    ' Reihenfolge: gezogene zuerst, dann platzierte nach (Zeile, Spalte), dann neue
    For pass = 0 To 2
        For i = 0 To n - 1
            If pass = 2 Then
                If Not has(i) Then ord(m) = i: m = m + 1
            ElseIf has(i) Then
                If GroupOf(i, pri, eids) = pass Then ord(m) = i: m = m + 1
            End If
        Next
    Next
    ' innerhalb der Gruppen 0 und 1 nach (Zeile, Spalte) sortieren
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If has(ord(i)) And has(ord(j)) And GroupOf(ord(i), pri, eids) = GroupOf(ord(j), pri, eids) Then
                If wy(ord(j)) < wy(ord(i)) Or (wy(ord(j)) = wy(ord(i)) And wx(ord(j)) < wx(ord(i))) Then
                    tmp = ord(i): ord(i) = ord(j): ord(j) = tmp
                End If
            End If
        Next
    Next
    ' Wunschzeilen (auch Zwischenwerte wie 1,5 = neue Zeile) auf Rangfolge abbilden
    ReDim ys(0 To n)
    For i = 0 To n - 1
        If has(i) Then
            For k = 0 To nys - 1
                If ys(k) = wy(i) Then Exit For
            Next
            If k = nys Then ys(nys) = wy(i): nys = nys + 1
        End If
    Next
    For i = 0 To nys - 2
        For j = i + 1 To nys - 1
            If ys(j) < ys(i) Then Swap ys(i), ys(j)
        Next
    Next
    ReDim occ(0 To 2 * n + 2, 0 To 3)
    maxR = -1
    For m = 0 To n - 1
        i = ord(m)
        If has(i) Then
            For k = 0 To nys - 1
                If ys(k) = wy(i) Then r = k: Exit For
            Next
            c = CLng(wx(i))
        Else
            ' Fluss: hinter das letzte Element der letzten Zeile
            r = IIf(maxR < 0, 0, maxR)
            c = 0
            For cc = 0 To 3
                If occ(r, cc) Then c = cc + 1
            Next
        End If
        If c > 4 - gsz(i) Then c = 4 - gsz(i)
        If c < 0 Then c = 0
        placed = False
        Do While Not placed
            If Fits(occ, r, c, gsz(i)) Then
                placed = True
            ElseIf has(i) Then
                ' naechste freie Spalte in derselben Zeile (nahe am Wunsch)
                For d = 1 To 3
                    If Fits(occ, r, c + d, gsz(i)) Then
                        c = c + d: placed = True: Exit For
                    ElseIf Fits(occ, r, c - d, gsz(i)) Then
                        c = c - d: placed = True: Exit For
                    End If
                Next
            End If
            If Not placed Then
                r = r + 1
                If Not has(i) Then c = 0
                If r > UBound(occ, 1) Then r = UBound(occ, 1): c = 0: Exit Do
            End If
        Loop
        For cc = c To c + gsz(i) - 1
            occ(r, cc) = True
        Next
        gx(i) = c: gy(i) = r
        If r > maxR Then maxR = r
    Next
    ' leere Zeilen entfernen
    ReDim newR(0 To UBound(occ, 1))
    nr = 0
    For r = 0 To UBound(occ, 1)
        newR(r) = nr
        If occ(r, 0) Or occ(r, 1) Or occ(r, 2) Or occ(r, 3) Then nr = nr + 1
    Next
    For i = 0 To n - 1
        gy(i) = newR(gy(i))
    Next
End Sub

' Rasterplaetze als Text "eid:gx,gy,breite;" nach Element-ID sortiert (Pruefungen)
Public Function LayoutInfo(ByVal did As String) As String
    Dim eids As Variant, gx() As Long, gy() As Long, gsz() As Long, i As Long, j As Long, n As Long, idx() As Long, tmp As Long
    eids = mCfg.ElementIds(did)
    n = ArrLen(eids)
    If n = 0 Then Exit Function
    LayoutGrid eids, gx, gy, gsz, Nothing
    ReDim idx(0 To n - 1)
    For i = 0 To n - 1: idx(i) = i: Next
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If CStr(eids(idx(j))) < CStr(eids(idx(i))) Then tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
        Next
    Next
    For i = 0 To n - 1
        LayoutInfo = LayoutInfo & eids(idx(i)) & ":" & gx(idx(i)) & "," & gy(idx(i)) & "," & gsz(idx(i)) & ";"
    Next
End Function

Private Function GroupOf(ByVal i As Long, ByVal pri As cMap, ByVal eids As Variant) As Long
    If pri Is Nothing Then GroupOf = 1: Exit Function
    If pri.Exists(CStr(eids(i))) Then GroupOf = 0 Else GroupOf = 1
End Function

Private Sub Swap(ByRef a As Double, ByRef b As Double)
    Dim t As Double
    t = a: a = b: b = t
End Sub

Private Function Fits(occ() As Boolean, ByVal r As Long, ByVal c As Long, ByVal sz As Long) As Boolean
    Dim k As Long
    If r > UBound(occ, 1) Then Exit Function
    If c < 0 Or c + sz > 4 Then Exit Function
    For k = c To c + sz - 1
        If occ(r, k) Then Exit Function
    Next
    Fits = True
End Function

' Hellere Spalten als Einrasthilfe (gesperrt, also nicht verschiebbar)
Private Sub DrawGridGuides(ByVal ws As Worksheet, ByVal x0 As Single, ByVal y0 As Single, ByVal unitW As Single, ByVal gap As Single, ByVal y1 As Single)
    Dim c As Long, s As Shape
    For c = 0 To 3
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x0 + c * (unitW + gap), y0, unitW, y1 - y0)
        s.Name = "grid_" & c
        s.Adjustments.Item(1) = 0.01
        s.Fill.ForeColor.RGB = CSurface()
        s.Line.ForeColor.RGB = CBorder(): s.Line.Weight = 0.75: s.Line.DashStyle = msoLineDash
        s.Placement = xlFreeFloating
        s.Locked = True
    Next
End Sub

' Platzhalterkarte im Anordnen-Modus (ohne Makro und nicht gesperrt: ziehbar)
Private Sub DrawArrangeCard(ByVal ws As Worksheet, ByVal eid As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal sz As Long)
    Dim s As Shape, pal As String, wTxt As String
    pal = ElemPaletteName(eid, mNav.gDash)
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    s.Name = "arr_" & eid
    s.Adjustments.Item(1) = 0.04
    s.Fill.ForeColor.RGB = Blend(PaletteColor(pal, 0), RGB(255, 255, 255), 0.85)
    s.Line.ForeColor.RGB = PaletteColor(pal, 0): s.Line.Weight = 1.75
    s.Placement = xlFreeFloating
    s.LockAspectRatio = msoFalse
    s.Locked = False
    wTxt = Choose(sz, "¼ Breite", "½ Breite", "¾ Breite", "ganze Breite")
    If ElemGet(eid, "Typ") <> "Kachel" Then wTxt = wTxt & " · " & IIf(h > 300, "hoch", "normal")
    With s.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoTrue
        With .TextRange
            .Text = ChrW(&HE7C2) & vbCr & ElemGet(eid, "Titel") & vbCr & ElemGet(eid, "Typ") & "  ·  " & wTxt
            .Font.Name = UI_FONT: .Font.Size = 10
            .Font.Fill.ForeColor.RGB = CText()
            .ParagraphFormat.Alignment = msoAlignCenter
            .Paragraphs(1).Font.Name = ICON_FONT
            .Paragraphs(1).Font.Size = 16
            .Paragraphs(1).Font.Fill.ForeColor.RGB = mDocStyle.DocInk(PaletteColor(pal, 0))
            .Paragraphs(2).Font.Name = UI_FONT_SEMI
            .Paragraphs(3).Font.Size = 8.5
            .Paragraphs(3).Font.Fill.ForeColor.RGB = CMuted()
        End With
    End With
End Sub

Public Sub BeginArrange()
    Set mArr = New cMap
    gArrange = True
    ' nach jedem Loslassen einer Karte sofort einrasten
    Set mWatch = New cDashWatch
    mWatch.StartWatch
End Sub

Public Sub EndArrange()
    If Not mWatch Is Nothing Then mWatch.StopWatch
    Set mWatch = Nothing
    Set mArr = Nothing
    gArrange = False
End Sub

' Wurde eine Karte seit dem letzten Aufbau gezogen oder in der Groesse geaendert? (billig)
Public Function CardsMoved() As Boolean
    Dim ws As Worksheet, s As Shape, L As Variant, eid As String
    If mLay Is Nothing Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_DASH)
    For Each s In ws.Shapes
        If Left$(s.Name, 4) = "arr_" Then
            eid = Mid$(s.Name, 5)
            If mLay.Exists(eid) Then
                L = mLay.GetVal(eid)
                If Abs(s.Left - L(0)) > 6 Or Abs(s.Top - L(1)) > 6 Or Abs(s.Width - L(2)) > 6 Or Abs(s.Height - L(3)) > 6 Then CardsMoved = True: Exit Function
            End If
        End If
    Next
End Function

' Liest die gezogenen Karten und rastet sie ein (Spalte aus der linken Kante, Zeile aus der
' Oberkante, Breite 1-4 Spalten, Hoehe normal/hoch). Nur im Speicher; True = etwas geaendert
Public Function SnapArrangement() As Boolean
    Dim ws As Worksheet, eids As Variant, i As Long, s As Shape, L As Variant, eid As String, pri As New cMap
    Dim nsz As Long, ncol As Long, hoch As Boolean, moved As Boolean, k As Long, t As Single, wy As Double
    Dim gx() As Long, gy() As Long, gsz() As Long
    If mLay Is Nothing Or mNRows = 0 Then Exit Function
    If mArr Is Nothing Then Set mArr = New cMap
    Set ws = mRender.Sheet(SH_DASH)
    eids = mCfg.ElementIds(mNav.gDash)
    For i = 0 To ArrLen(eids) - 1
        eid = CStr(eids(i))
        Set s = Nothing
        On Error Resume Next
        Set s = ws.Shapes("arr_" & eid)
        On Error GoTo 0
        If Not s Is Nothing And mLay.Exists(eid) Then
            L = mLay.GetVal(eid)
            moved = (Abs(s.Left - L(0)) > 6 Or Abs(s.Top - L(1)) > 6 Or Abs(s.Width - L(2)) > 6 Or Abs(s.Height - L(3)) > 6)
            nsz = CLng((s.Width + mGap) / (mUnitW + mGap))
            If nsz < 1 Then nsz = 1
            If nsz > 4 Then nsz = 4
            ncol = CLng((s.Left - mX0) / (mUnitW + mGap))
            If ncol > 4 - nsz Then ncol = 4 - nsz
            If ncol < 0 Then ncol = 0
            hoch = ElemHoch(eid)
            If ElemGet(eid, "Typ") <> "Kachel" And Abs(s.Height - L(3)) > 6 Then hoch = (s.Height > 335)
            If moved Then
                SnapArrangement = True
                pri.SetVal eid, True
                t = s.Top
                If t < mRowTop(0) - 20 Then
                    wy = -0.5
                ElseIf t > mRowTop(mNRows - 1) + mRowH(mNRows - 1) * 0.6 Then
                    wy = mNRows - 0.5
                Else
                    For k = mNRows - 1 To 0 Step -1
                        If t >= mRowTop(k) - 20 Then Exit For
                    Next
                    If k < 0 Then k = 0
                    ' Oberkante deutlich unterhalb der Zeilenmitte: neue Zeile darunter
                    If t > mRowTop(k) + mRowH(k) * 0.6 Then wy = k + 0.5 Else wy = k
                End If
                mArr.SetVal eid, Array(ncol, wy, nsz, hoch)
            Else
                mArr.SetVal eid, Array(CLng(L(4)), CDbl(L(5)), ElemSize(eid), hoch)
            End If
        End If
    Next
    ' ueberschneidungsfrei machen und als ganze Zahlen merken
    LayoutGrid eids, gx, gy, gsz, pri
    For i = 0 To ArrLen(eids) - 1
        eid = CStr(eids(i))
        If mArr.Exists(eid) Then mArr.SetVal eid, Array(gx(i), CDbl(gy(i)), gsz(i), mArr.GetVal(eid)(3))
    Next
End Function

' Uebernimmt die Anordnung dauerhaft (Spalte, Zeile, Breite, Hoehe und Reihenfolge)
Public Sub ApplyArrangement()
    Dim eids As Variant, i As Long, j As Long, eid As String, p As Variant, opt As String, idx() As Long, tmp As Long, n As Long
    SnapArrangement
    If mArr Is Nothing Then Exit Sub
    eids = mCfg.ElementIds(mNav.gDash)
    n = ArrLen(eids)
    If n = 0 Then Exit Sub
    ReDim idx(0 To n - 1)
    For i = 0 To n - 1
        idx(i) = i
        eid = CStr(eids(i))
        If mArr.Exists(eid) Then
            p = mArr.GetVal(eid)
            opt = ElemGet(eid, "Optionen")
            opt = KVSet(opt, "gx", CStr(p(0)))
            opt = KVSet(opt, "gy", CStr(CLng(p(1))))
            If ElemGet(eid, "Typ") <> "Kachel" Then opt = KVSet(opt, "hoch", IIf(p(3), "1", ""))
            mCfg.SetV SH_ELM, eid, "Optionen", opt
            mCfg.SetV SH_ELM, eid, "Groesse", CStr(p(2))
        End If
    Next
    ' Reihenfolge = Leserichtung (Zeile, dann Spalte) fuer Export und Tastaturwege
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If PosKey(CStr(eids(idx(j)))) < PosKey(CStr(eids(idx(i)))) Then tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
        Next
    Next
    For i = 0 To n - 1
        mCfg.SetV SH_ELM, CStr(eids(idx(i))), "Reihenfolge", i + 1
    Next
    mLog.Add "Dashboard angeordnet", "", mNav.gDash, "", ""
    mApp.MarkCfgDirty
End Sub

Private Function PosKey(ByVal eid As String) As Double
    Dim p As Variant
    If mArr.Exists(eid) Then p = mArr.GetVal(eid): PosKey = CDbl(p(1)) * 10 + CDbl(p(0)) Else PosKey = 1E+9
End Function

Public Function DashFilterText(ByVal did As String) As String
    Dim f As String, s As String, tid As String
    f = mCfg.GetS(SH_DSB, did, "Filter")
    If Len(KVGet(f, "von")) > 0 Or Len(KVGet(f, "bis")) > 0 Then
        s = "Zeitraum " & IIf(Len(KVGet(f, "von")) > 0, FmtDate(InvDate(KVGet(f, "von"))), "…") & " bis " & IIf(Len(KVGet(f, "bis")) > 0, FmtDate(InvDate(KVGet(f, "bis"))), "…")
    End If
    tid = KVGet(f, "tab")
    If Len(tid) > 0 And Len(KVGet(f, "f")) > 0 And mCfg.TableExists(tid) Then
        s = s & IIf(Len(s) > 0, " · ", "") & mCfg.TableName(tid) & ": " & mQuery.CondsText(mQuery.ParseConds(KVGet(f, "f")))
    End If
    DashFilterText = s
End Function

Private Sub DrawElement(ByVal ws As Worksheet, ByVal eid As String, ByVal did As String, ByVal res As cMap, ByVal x As Single, ByVal y As Single, _
                        ByVal w As Single, ByVal h As Single)
    Dim card As Shape, s As Shape, typ As String, n As Long, labels As Variant, values As Variant, i As Long, ch As ChartObject
    Dim vals() As Double, cats() As String, maxRows As Long, rowY As Single, nm As String, note As String, ttl As String, keys As Variant
    Dim sty As String, pal As String, c0 As Long
    typ = ElemGet(eid, "Typ")
    ttl = ElemGet(eid, "Titel")
    sty = DashStyle(did)
    pal = ElemPaletteName(eid, did)
    c0 = PaletteColor(pal, 0)
    Set card = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    card.Name = "card_" & eid
    card.Adjustments.Item(1) = 0.04
    card.Placement = xlFreeFloating
    Select Case sty
        Case "Flach"
            card.Fill.ForeColor.RGB = CSurface()
            card.Line.Visible = msoFalse
        Case "Farbakzent"
            card.Fill.ForeColor.RGB = Blend(c0, RGB(255, 255, 255), 0.93)
            card.Line.ForeColor.RGB = Blend(c0, RGB(255, 255, 255), 0.6): card.Line.Weight = 0.75
        Case "Kontur"
            card.Fill.ForeColor.RGB = CWhite()
            card.Line.ForeColor.RGB = CBorderStrong(): card.Line.Weight = 1.5
        Case Else
            card.Fill.ForeColor.RGB = CWhite()
            card.Line.ForeColor.RGB = CBorder(): card.Line.Weight = 0.75
            With card.Shadow
                .Visible = msoTrue: .ForeColor.RGB = RGB(0, 0, 0): .Transparency = 0.93: .OffsetX = 0: .OffsetY = 1: .Blur = 4
            End With
    End Select
    If eid = gHighlight Then
        ' Vorschau eines neuen/geaenderten Elements deutlich markieren
        card.Line.Visible = msoTrue
        card.Line.ForeColor.RGB = CAccent(): card.Line.Weight = 2.5
    End If
    card.OnAction = mRender.MacroRef("mCmd.Cmd_DashElementMenu")
    If sty = "Farbakzent" Then
        Set s = ws.Shapes.AddShape(msoShapeRectangle, x + 14, y + 8, 4, 18)
        s.Name = "ex_" & eid
        s.Fill.ForeColor.RGB = c0: s.Line.Visible = msoFalse
        s.Placement = xlFreeFloating
        Set s = mRender.TextShape(ws, "et_" & eid, ttl, x + 24, y + 8, w - 70, 20, 10.5, CText(), True, 1, True)
        mRender.FitOneLine s, w - 70, 10.5
    Else
        Set s = mRender.TextShape(ws, "et_" & eid, ttl, x + 14, y + 8, w - 60, 20, 10.5, CText(), True, 1, True)
        mRender.FitOneLine s, w - 60, 10.5
    End If
    s.OnAction = mRender.MacroRef("mCmd.Cmd_DashElementMenu")
    Set s = mRender.Btn(ws, "em_" & eid, "", x + w - 34, y + 6, "mCmd.Cmd_DashElementMenu", "ghost", ChrW(&HE712), 22, 0, 26)
    note = CStr(res.GetVal("notes", ""))
    If res.Exists("dashApplied") Then note = CStr(res.GetVal("dashApplied")) & "; " & note
    If res.Exists("error") Then
        mRender.TextShape ws, "ee_" & eid, ChrW(&H26A0) & " " & res.GetVal("error"), x + 14, y + 36, w - 28, 40, 9.5, CDanger(), False, 1
        Exit Sub
    End If
    n = CLng(res.GetVal("n", 0))
    Select Case typ
        Case "Kachel"
            Set s = mRender.TextShape(ws, "ev_" & eid, FormatMeasure(res, res.GetVal("totalValue")), x + 14, y + 30, w - 28, 36, 22, mDocStyle.DocInk(c0), True, 1, True)
            mRender.FitOneLine s, w - 28, 14
            s.OnAction = mRender.MacroRef("mCmd.Cmd_DashDrill")
            mDrill.SetVal s.Name, Array(eid, Empty)
            mRender.TextShape ws, "en_" & eid, MeasureLabel(CStr(res.GetVal("measure"))) & IIf(res.Exists("dashIgnored"), "  ·  ungefiltert", ""), _
                              x + 14, y + 66, w - 28, 14, 8, CMuted(), False, 1
            mRender.TextShape ws, "eb_" & eid, Truncate(note, 140), x + 14, y + 80, w - 28, 20, 7.5, CMuted(), False, 1
        Case "Tabelle"
            labels = res.GetVal("labels"): values = res.GetVal("values"): keys = res.GetVal("keys")
            maxRows = Int((h - 74) / 19)
            rowY = y + 36
            If n = 0 Then
                Set s = mRender.TextShape(ws, "ev_" & eid, "Gesamt: " & FormatMeasure(res, res.GetVal("totalValue")), x + 14, rowY, w - 28, 20, 11, CText(), True, 1)
                s.OnAction = mRender.MacroRef("mCmd.Cmd_DashDrill"): mDrill.SetVal s.Name, Array(eid, Empty)
            End If
            For i = 0 To Application.Min(n, maxRows) - 1
                nm = "er_" & eid & "_" & i
                Set s = mRender.TextShape(ws, nm & "a", Truncate(CStr(labels(i)), CLng((w - 140) / 5.6)), x + 14, rowY, w - 130, 18, 9.5, CText(), False, 1)
                s.OnAction = mRender.MacroRef("mCmd.Cmd_DashDrill"): mDrill.SetVal s.Name, Array(eid, keys(i))
                Set s = mRender.TextShape(ws, nm & "b", FormatMeasure(res, values(i)), x + w - 120, rowY, 106, 18, 9.5, CText(), True, 3)
                s.OnAction = mRender.MacroRef("mCmd.Cmd_DashDrill"): mDrill.SetVal s.Name, Array(eid, keys(i))
                Set s = ws.Shapes.AddLine(x + 14, rowY + 18.5, x + w - 14, rowY + 18.5)
                s.Line.ForeColor.RGB = CBorder(): s.Line.Weight = 0.5: s.Name = nm & "l"
                rowY = rowY + 19
            Next
            If n > 0 Then
                Set s = mRender.TextShape(ws, "et2_" & eid, "Gesamt (eindeutig): " & FormatMeasure(res, res.GetVal("totalValue")), x + 14, rowY + 2, w - 28, 18, 9, CMuted(), True, 3)
                s.OnAction = mRender.MacroRef("mCmd.Cmd_DashDrill"): mDrill.SetVal s.Name, Array(eid, Empty)
            End If
            mRender.TextShape ws, "eb_" & eid, Truncate(note, 260), x + 14, y + h - 34, w - 28, 30, 7.5, CMuted(), False, 1
        Case Else
            labels = res.GetVal("labels"): values = res.GetVal("values")
            If n = 0 Then
                mRender.TextShape ws, "ev_" & eid, IIf(res.Exists("groupField"), "Keine Daten für diese Auswahl.", "Für Diagramme bitte eine Gruppierung wählen."), _
                                  x + 14, y + 40, w - 28, 20, 9.5, CMuted(), False, 1
            ElseIf typ = "Netz" And n < 3 Then
                mRender.TextShape ws, "ev_" & eid, "Ein Netzdiagramm braucht mindestens 3 Gruppen (zurzeit " & n & ").", _
                                  x + 14, y + 40, w - 28, 20, 9.5, CMuted(), False, 1
            ElseIf typ = "Rangliste" Then
                DrawRanking ws, eid, res, x, y, w, h, pal
            ElseIf typ = "Sankey" Then
                DrawSankey ws, eid, res, x, y, w, h, pal
            Else
                ' Diagrammdaten in ausgeblendete Hilfsspalten desselben Blattes (keine externen Bezuege)
                mDataCol = mDataCol + 3
                For i = 0 To n - 1
                    ws.Cells(R_DATA + i, mDataCol).NumberFormat = "@"
                    ws.Cells(R_DATA + i, mDataCol).Value = Truncate(CStr(labels(i)), 28)
                    If IsNumeric(values(i)) And Not IsBlank(values(i)) Then ws.Cells(R_DATA + i, mDataCol + 1).Value = CDbl(values(i)) Else ws.Cells(R_DATA + i, mDataCol + 1).Value = 0
                Next
                Set ch = NewChart(ws, typ, x + 10, y + 32, w - 20, h - 74, ws.Range(ws.Cells(R_DATA, mDataCol), ws.Cells(R_DATA + n - 1, mDataCol)), _
                                  ws.Range(ws.Cells(R_DATA, mDataCol + 1), ws.Cells(R_DATA + n - 1, mDataCol + 1)), res, pal)
                ch.Name = "ch_" & eid
                ch.Placement = xlFreeFloating
                ch.OnAction = mRender.MacroRef("mCmd.Cmd_DashElementMenu")
            End If
            Set s = mRender.TextShape(ws, "eb_" & eid, Truncate(note, 260), x + 14, y + h - 40, w - 120, 36, 7.5, CMuted(), False, 1)
            If n > 0 Then
                Set s = mRender.Btn(ws, "ed_" & eid, "Details", x + w - 96, y + h - 34, "mCmd.Cmd_DashDetails", "link", ChrW(&HE8A7))
            End If
    End Select
End Sub

' Diagramm anlegen. Trichter und Baumdiagramm sind neuere Excel-Diagrammarten (ab Excel 2016/2019):
' sie entstehen über AddChart2; fehlen sie in der Excel-Version, wird ein Balkendiagramm gezeichnet.
Public Function NewChart(ByVal ws As Worksheet, ByVal typ As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                         ByVal catRng As Range, ByVal valRng As Range, ByVal res As cMap, Optional ByVal palName As String = "Akzentfarbe") As ChartObject
    Dim shp As Shape, code As Long, co As ChartObject, i As Long
    If typ = "Trichter" Then code = 123
    If typ = "Baumdiagramm" Then code = 117
    If code <> 0 Then
        On Error Resume Next
        ' diese Diagrammarten übernehmen ihre Daten nur aus der Markierung beim Anlegen
        ' (nachträgliches SetSourceData ist für sie nicht umgesetzt)
        Dim ev As Boolean
        ev = Application.EnableEvents
        Application.EnableEvents = False
        If Not ws.Parent Is ActiveWorkbook Then ws.Parent.Activate
        If Not ws Is ActiveSheet Then ws.Activate
        ws.Range(catRng.Cells(1, 1), valRng.Cells(valRng.Rows.Count, 1)).Select
        Set shp = ws.Shapes.AddChart2(-1, code, l, t, w, h)
        Application.EnableEvents = ev
        If Err.Number = 0 And Not shp Is Nothing Then
            If Err.Number = 0 Then
                Set co = ws.ChartObjects(shp.Name)
                With co.Chart
                    ' Daten stehen in später ausgeblendeten Hilfsspalten
                    .PlotVisibleOnly = False
                    .HasTitle = False
                    .HasLegend = False
                    .ChartArea.Format.Line.Visible = msoFalse
                    .ChartArea.Format.Fill.Visible = msoFalse
                    .ChartArea.Font.Name = UI_FONT
                    .ChartArea.Font.Size = 8
                    For i = 1 To .SeriesCollection(1).Points.Count
                        .SeriesCollection(1).Points(i).Format.Fill.ForeColor.RGB = IIf(typ = "Trichter" And palName = "Akzentfarbe", _
                            Blend(PaletteColor(palName, 0), RGB(255, 255, 255), Application.Min(0.6, (i - 1) * 0.12)), PaletteColor(palName, i - 1))
                    Next
                    .SeriesCollection(1).HasDataLabels = True
                    .SeriesCollection(1).DataLabels.Font.Size = 8
                    If typ = "Baumdiagramm" Then
                        .SeriesCollection(1).DataLabels.ShowCategoryName = True
                        .SeriesCollection(1).DataLabels.ShowValue = True
                    End If
                    .ChartArea.Border.LineStyle = xlNone
                End With
                Err.Clear
                On Error GoTo 0
                Set NewChart = co
                Exit Function
            End If
            shp.Delete
        End If
        Err.Clear
        On Error GoTo 0
    End If
    Set co = ws.ChartObjects.Add(l, t, w, h)
    BuildChart co.Chart, typ, catRng, valRng, res, palName
    Set NewChart = co
End Function

' Sankey aus Formen (Excel kennt keine Sankey-Diagrammart): links die Gruppen der ersten Gruppierung,
' rechts die der zweiten, dazwischen gebogene Bänder in der Breite des Wertes
Public Sub DrawSankey(ByVal ws As Worksheet, ByVal eid As String, ByVal res As cMap, ByVal x As Single, ByVal y As Single, _
                      ByVal w As Single, ByVal h As Single, ByVal pal As String, Optional ByVal showValues As Boolean = True)
    Dim sk As Variant, bLabels As Variant, flows As Variant, labels As Variant, keys As Variant, nA As Long, nB As Long, nf As Long
    Dim aTot() As Double, bTot() As Double, i As Long, j As Long, tot As Double, top_ As Single, hh As Single, gap As Single, sc As Single
    Dim aY() As Single, bY() As Single, aOff() As Single, bOff() As Single, ya() As Single, yb() As Single, th() As Single, ord() As Long, tmp As Long
    Dim lw As Single, xa As Single, xb As Single, nw As Single, s As Shape, ff As FreeformBuilder, xm As Single, nm As String, col As Long
    sk = res.GetVal("sankey", Empty)
    If IsEmpty(sk) Then
        mRender.TextShape ws, "ev_" & eid, "Keine Flüsse für diese Auswahl.", x + 14, y + 40, w - 28, 20, 9.5, CMuted(), False, 1
        Exit Sub
    End If
    bLabels = sk(0): flows = sk(1)
    labels = res.GetVal("labels"): keys = res.GetVal("keys")
    nA = CLng(res.GetVal("n", 0)): nB = UBound(bLabels) + 1: nf = UBound(flows) + 1
    ReDim aTot(0 To nA - 1): ReDim bTot(0 To nB - 1)
    For i = 0 To nf - 1
        aTot(flows(i)(0)) = aTot(flows(i)(0)) + flows(i)(2)
        bTot(flows(i)(1)) = bTot(flows(i)(1)) + flows(i)(2)
        tot = tot + flows(i)(2)
    Next
    If tot <= 0 Then Exit Sub
    ' Fläche: Beschriftungen links/rechts, Knoten 10 pt breit
    lw = Application.Min(w * 0.26, 150)
    nw = 10
    xa = x + 14 + lw
    xb = x + w - 14 - lw - nw
    top_ = y + 38
    hh = h - 38 - 46
    gap = Application.Min(8, hh * 0.04)
    sc = Application.Min((hh - gap * (nA - 1)) / tot, (hh - gap * (nB - 1)) / tot)
    If sc <= 0 Then sc = 0.1
    ReDim aY(0 To nA - 1): ReDim bY(0 To nB - 1): ReDim aOff(0 To nA - 1): ReDim bOff(0 To nB - 1)
    aY(0) = top_ + (hh - (tot * sc + gap * (nA - 1))) / 2
    For i = 1 To nA - 1: aY(i) = aY(i - 1) + aTot(i - 1) * sc + gap: Next
    bY(0) = top_ + (hh - (tot * sc + gap * (nB - 1))) / 2
    For i = 1 To nB - 1: bY(i) = bY(i - 1) + bTot(i - 1) * sc + gap: Next
    ' Reihenfolge: links nach (links, rechts), rechts nach (rechts, links); Bänder kreuzen sich wenig
    ReDim ya(0 To nf - 1): ReDim yb(0 To nf - 1): ReDim th(0 To nf - 1): ReDim ord(0 To nf - 1)
    For i = 0 To nf - 1: ord(i) = i: th(i) = flows(i)(2) * sc: Next
    For i = 0 To nf - 2
        For j = i + 1 To nf - 1
            If flows(ord(j))(0) * 100 + flows(ord(j))(1) < flows(ord(i))(0) * 100 + flows(ord(i))(1) Then tmp = ord(i): ord(i) = ord(j): ord(j) = tmp
        Next
    Next
    For i = 0 To nf - 1
        j = ord(i)
        ya(j) = aY(flows(j)(0)) + aOff(flows(j)(0)): aOff(flows(j)(0)) = aOff(flows(j)(0)) + th(j)
    Next
    For i = 0 To nf - 2
        For j = i + 1 To nf - 1
            If flows(ord(j))(1) * 100 + flows(ord(j))(0) < flows(ord(i))(1) * 100 + flows(ord(i))(0) Then tmp = ord(i): ord(i) = ord(j): ord(j) = tmp
        Next
    Next
    For i = 0 To nf - 1
        j = ord(i)
        yb(j) = bY(flows(j)(1)) + bOff(flows(j)(1)): bOff(flows(j)(1)) = bOff(flows(j)(1)) + th(j)
    Next
    ' Bänder
    xm = (xa + nw + xb) / 2
    For i = 0 To nf - 1
        col = PaletteColor(pal, flows(i)(0))
        Set ff = ws.Shapes.BuildFreeform(msoEditingAuto, xa + nw, ya(i))
        ff.AddNodes msoSegmentCurve, msoEditingCorner, xm, ya(i), xm, yb(i), xb, yb(i)
        ff.AddNodes msoSegmentLine, msoEditingAuto, xb, yb(i) + Application.Max(0.8, th(i))
        ff.AddNodes msoSegmentCurve, msoEditingCorner, xm, yb(i) + Application.Max(0.8, th(i)), xm, ya(i) + Application.Max(0.8, th(i)), xa + nw, ya(i) + Application.Max(0.8, th(i))
        ff.AddNodes msoSegmentLine, msoEditingAuto, xa + nw, ya(i)
        Set s = ff.ConvertToShape
        s.Name = "es_" & eid & "_f" & i
        s.Line.Visible = msoFalse
        s.Fill.ForeColor.RGB = col
        s.Fill.Transparency = IIf(IsLightColor(col), 0.25, 0.55)
        s.Placement = xlFreeFloating
        s.OnAction = mRender.MacroRef("mCmd.Cmd_DashDrill")
        If Not mDrill Is Nothing Then mDrill.SetVal s.Name, Array(eid, keys(flows(i)(0)))
    Next
    ' Knoten und Beschriftungen
    For i = 0 To nA - 1
        nm = "es_" & eid & "_a" & i
        Set s = ws.Shapes.AddShape(msoShapeRectangle, xa, aY(i), nw, Application.Max(1.5, aTot(i) * sc))
        s.Name = nm: s.Line.Visible = msoFalse: s.Fill.ForeColor.RGB = PaletteColor(pal, i): s.Placement = xlFreeFloating
        Set s = mRender.TextShape(ws, nm & "t", Truncate(CStr(labels(i)), CLng(lw / 5.2)) & IIf(showValues, "  " & FormatMeasure(res, aTot(i)), ""), _
                                  x + 10, aY(i) + aTot(i) * sc / 2 - 8, lw - 8, 16, 8.5, CText(), False, 3)
        s.OnAction = mRender.MacroRef("mCmd.Cmd_DashDrill")
        If Not mDrill Is Nothing Then mDrill.SetVal s.Name, Array(eid, keys(i))
    Next
    For i = 0 To nB - 1
        nm = "es_" & eid & "_b" & i
        Set s = ws.Shapes.AddShape(msoShapeRectangle, xb, bY(i), nw, Application.Max(1.5, bTot(i) * sc))
        s.Name = nm: s.Line.Visible = msoFalse: s.Fill.ForeColor.RGB = RGB(100, 116, 139): s.Placement = xlFreeFloating
        mRender.TextShape ws, nm & "t", Truncate(CStr(bLabels(i)), CLng(lw / 5.2)) & IIf(showValues, "  " & FormatMeasure(res, bTot(i)), ""), _
                          xb + nw + 6, bY(i) + bTot(i) * sc / 2 - 8, lw - 4, 16, 8.5, CText(), False, 1
    Next
End Sub

' Rangliste: je Gruppe Beschriftung, Wert und ein Balken anteilig zum größten Wert
Private Sub DrawRanking(ByVal ws As Worksheet, ByVal eid As String, ByVal res As cMap, ByVal x As Single, ByVal y As Single, _
                        ByVal w As Single, ByVal h As Single, ByVal pal As String)
    Dim labels As Variant, values As Variant, keys As Variant, n As Long, i As Long, mx As Double, rowY As Single, s As Shape, nm As String
    Dim maxRows As Long, v As Double, bw As Single
    labels = res.GetVal("labels"): values = res.GetVal("values"): keys = res.GetVal("keys")
    n = CLng(res.GetVal("n", 0))
    For i = 0 To n - 1
        If IsNumeric(values(i)) And Not IsBlank(values(i)) Then If CDbl(values(i)) > mx Then mx = CDbl(values(i))
    Next
    maxRows = Int((h - 80) / 30)
    rowY = y + 36
    For i = 0 To Application.Min(n, maxRows) - 1
        nm = "er_" & eid & "_" & i
        v = 0
        If IsNumeric(values(i)) And Not IsBlank(values(i)) Then v = CDbl(values(i))
        Set s = mRender.TextShape(ws, nm & "a", Truncate(CStr(labels(i)), CLng((w - 130) / 5.6)), x + 14, rowY, w - 130, 15, 9, CText(), False, 1)
        s.OnAction = mRender.MacroRef("mCmd.Cmd_DashDrill"): mDrill.SetVal s.Name, Array(eid, keys(i))
        Set s = mRender.TextShape(ws, nm & "b", FormatMeasure(res, values(i)), x + w - 116, rowY, 102, 15, 9, CText(), True, 3)
        s.OnAction = mRender.MacroRef("mCmd.Cmd_DashDrill"): mDrill.SetVal s.Name, Array(eid, keys(i))
        ' Spur und Balken
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x + 14, rowY + 17, w - 28, 7)
        s.Name = nm & "t": s.Line.Visible = msoFalse: s.Fill.ForeColor.RGB = CSurface2()
        On Error Resume Next: s.Adjustments.Item(1) = 0.5: On Error GoTo 0
        If mx > 0 And v > 0 Then
            bw = (w - 28) * v / mx
            If bw < 7 Then bw = 7
            Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x + 14, rowY + 17, bw, 7)
            s.Name = nm & "f": s.Line.Visible = msoFalse
            s.Fill.ForeColor.RGB = IIf(pal = "Akzentfarbe", PaletteColor(pal, 0), PaletteColor(pal, i))
            On Error Resume Next: s.Adjustments.Item(1) = 0.5: On Error GoTo 0
            s.OnAction = mRender.MacroRef("mCmd.Cmd_DashDrill"): mDrill.SetVal s.Name, Array(eid, keys(i))
        End If
        rowY = rowY + 30
    Next
    If n > maxRows Then mRender.TextShape ws, "em_" & eid, "+ " & (n - maxRows) & " weitere; „Details“ zeigt alle", x + 14, rowY, w - 28, 14, 8, CMuted(), False, 1
End Sub

Public Sub BuildChart(ByVal c As Chart, ByVal typ As String, ByVal catRng As Range, ByVal valRng As Range, ByVal res As cMap, _
                      Optional ByVal palName As String = "Akzentfarbe")
    Dim se As Series, i As Long, se2 As Series, tot As Double, run As Double, pie As Boolean
    Do While c.SeriesCollection.Count > 0
        c.SeriesCollection(1).Delete
    Loop
    pie = (typ = "Anteil" Or typ = "Kreis")
    Select Case typ
        Case "Balken", "Trichter", "Baumdiagramm": c.ChartType = xlBarClustered
        Case "Zeitverlauf": c.ChartType = xlLineMarkers
        Case "Fläche": c.ChartType = xlArea
        Case "Anteil": c.ChartType = xlDoughnut
        Case "Kreis": c.ChartType = xlPie
        Case "Netz": c.ChartType = xlRadarFilled
        Case Else: c.ChartType = xlColumnClustered
    End Select
    Set se = c.SeriesCollection.NewSeries
    se.Values = valRng
    se.XValues = catRng
    c.PlotVisibleOnly = False
    se.Name = MeasureLabel(CStr(res.GetVal("measure")))
    ' Pareto: zweite Reihe mit kumuliertem Anteil (rechte Achse)
    If typ = "Pareto" Then
        For i = 1 To valRng.Rows.Count
            If IsNumeric(valRng.Cells(i, 1).Value) Then tot = tot + CDbl(valRng.Cells(i, 1).Value)
        Next
        For i = 1 To valRng.Rows.Count
            If IsNumeric(valRng.Cells(i, 1).Value) Then run = run + CDbl(valRng.Cells(i, 1).Value)
            valRng.Cells(i, 1).Offset(0, 1).Value = IIf(tot > 0, run / tot, 0)
        Next
        Set se2 = c.SeriesCollection.NewSeries
        se2.Values = valRng.Offset(0, 1)
        se2.XValues = catRng
        se2.Name = "kumuliert"
        se2.ChartType = xlLineMarkers
        se2.AxisGroup = xlSecondary
    End If
    c.HasTitle = True
    c.HasTitle = False
    On Error Resume Next
    c.SetElement msoElementChartTitleNone
    On Error GoTo 0
    c.HasLegend = pie
    If c.HasLegend Then
        c.Legend.Position = xlLegendPositionRight
        c.Legend.Font.Size = 8
        c.Legend.Font.Color = CMuted()
    End If
    c.ChartArea.Format.Line.Visible = msoFalse
    c.ChartArea.Format.Fill.Visible = msoFalse
    c.PlotArea.Format.Fill.Visible = msoFalse
    c.ChartArea.Font.Name = UI_FONT
    c.ChartArea.Font.Size = 8
    If pie Then
        For i = 1 To se.Points.Count
            se.Points(i).Format.Fill.ForeColor.RGB = PaletteColor(palName, i - 1)
            se.Points(i).Format.Line.ForeColor.RGB = RGB(255, 255, 255)
        Next
        On Error Resume Next
        If typ = "Anteil" Then c.ChartGroups(1).DoughnutHoleSize = 55
        se.HasDataLabels = True
        se.DataLabels.ShowPercentage = True
        se.DataLabels.ShowValue = False
        se.DataLabels.Font.Size = 7.5
        On Error GoTo 0
    Else
        se.Format.Fill.ForeColor.RGB = PaletteColor(palName, 0)
        se.Format.Line.ForeColor.RGB = PaletteColor(palName, 0)
        On Error Resume Next
        If typ = "Zeitverlauf" Then
            se.Format.Line.Weight = 2.25
            se.MarkerBackgroundColor = PaletteColor(palName, 0)
            se.MarkerForegroundColor = PaletteColor(palName, 0)
        ElseIf typ = "Fläche" Then
            se.Format.Fill.ForeColor.RGB = Blend(PaletteColor(palName, 0), RGB(255, 255, 255), 0.45)
            se.Format.Line.Visible = msoTrue
            se.Format.Line.Weight = 2
        ElseIf typ = "Netz" Then
            ' halbtransparente Fläche mit kräftiger Kontur
            se.Format.Fill.ForeColor.RGB = PaletteColor(palName, 0)
            se.Format.Fill.Transparency = 0.65
            se.Format.Line.Visible = msoTrue
            se.Format.Line.ForeColor.RGB = PaletteColor(palName, 0)
            se.Format.Line.Weight = 2
            ' Werte stehen an den Punkten; Skalenbeschriftung würde sich damit überlagern
            c.Axes(xlValue).TickLabelPosition = xlTickLabelPositionNone
        ElseIf typ = "Pareto" Then
            se2.Format.Line.ForeColor.RGB = PaletteColor(IIf(palName = "Akzentfarbe", "Kontrast (barrierearm)", palName), 1)
            se2.Format.Line.Weight = 2
            se2.MarkerBackgroundColor = se2.Format.Line.ForeColor.RGB
            se2.MarkerForegroundColor = se2.Format.Line.ForeColor.RGB
            c.Axes(xlValue, xlSecondary).MinimumScale = 0
            c.Axes(xlValue, xlSecondary).MaximumScale = 1
            c.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0%"
            c.Axes(xlValue, xlSecondary).TickLabels.Font.Color = CMuted()
            c.Axes(xlValue, xlSecondary).Format.Line.Visible = msoFalse
        ElseIf palName <> "Akzentfarbe" Then
            ' bunte Paletten: jede Gruppe in eigener Farbe
            For i = 1 To se.Points.Count
                se.Points(i).Format.Fill.ForeColor.RGB = PaletteColor(palName, i - 1)
            Next
        End If
        c.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = CBorder()
        c.Axes(xlValue).Format.Line.Visible = msoFalse
        c.Axes(xlCategory).Format.Line.ForeColor.RGB = CBorderStrong()
        c.Axes(xlValue).TickLabels.Font.Color = CMuted()
        c.Axes(xlCategory).TickLabels.Font.Color = CMuted()
        If res.GetVal("measure") = "ANTEIL" Then c.Axes(xlValue).TickLabels.NumberFormat = "0%"
        c.ChartGroups(1).GapWidth = 60
        se.HasDataLabels = (valRng.Rows.Count < 12)
        If se.HasDataLabels Then
            se.DataLabels.Font.Size = 7.5
            se.DataLabels.Font.Color = CMuted()
            If res.GetVal("measure") = "ANTEIL" Then se.DataLabels.NumberFormat = "0.0%"
        End If
        On Error GoTo 0
    End If
End Sub

Public Function DrillInfo(ByVal shapeName As String) As Variant
    If mDrill Is Nothing Then DrillInfo = Empty: Exit Function
    If mDrill.Exists(shapeName) Then DrillInfo = mDrill.GetVal(shapeName) Else DrillInfo = Empty
End Function

Public Function ElementFromShape(ByVal shapeName As String) As String
    Dim p As Long
    p = InStr(shapeName, "_")
    If p = 0 Then Exit Function
    ElementFromShape = Mid$(shapeName, p + 1)
    If InStr(ElementFromShape, "_") > 0 Then ElementFromShape = Left$(ElementFromShape, InStr(ElementFromShape, "_") - 1)
End Function

' Element als Bild in die Zwischenablage (fuer PowerPoint und Word): Karte mit Titel, Kennzahl, Tabelle oder
' Diagramm, ohne Menueknopf, "Details" und Quellenhinweis. Vektorgrafik (EMF): sieht aus wie im Dashboard und
' bleibt beim Vergroessern scharf.
Public Function CopyElement(ByVal eid As String) As Boolean
    Dim ws As Worksheet, s As Shape, names() As Variant, n As Long, pre As String, p As Long
    Set ws = ActiveSheet
    For Each s In ws.Shapes
        If ElementFromShape(s.Name) = eid And s.Visible Then
            p = InStr(s.Name, "_")
            pre = Left$(s.Name, p - 1)
            If pre <> "em" And pre <> "ed" And pre <> "eb" And pre <> "arr" Then
                n = n + 1
                ReDim Preserve names(1 To n)
                names(n) = s.Name
            End If
        End If
    Next
    If n = 0 Then Exit Function
    mProtect.Unprot ws
    On Error GoTo done
    ws.Shapes.Range(names).Select
    On Error Resume Next
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    If Err.Number <> 0 Then
        ' ohne sichtbares Fenster (z. B. minimiert): Darstellung wie beim Drucken
        Err.Clear
        Selection.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
    End If
    CopyElement = (Err.Number = 0)
done:
    On Error Resume Next
    ws.Cells(R_TOP, 1).Select
    mProtect.ProtectUi ws
End Function

' Detail: erst Gruppe waehlen, dann Datensaetze
Public Sub Details(ByVal eid As String)
    Dim res As cMap, n As Long, labels As Variant, values As Variant, keys As Variant, items() As Variant, i As Long, sel As Long
    Set res = Compute(eid, mNav.gDash)
    n = CLng(res.GetVal("n", 0))
    If n = 0 Then
        ShowRecordList ElemGet(eid, "Tabelle"), DrillRids(eid, mNav.gDash, Empty), ElemGet(eid, "Titel")
        Exit Sub
    End If
    labels = res.GetVal("labels"): values = res.GetVal("values"): keys = res.GetVal("keys")
    ReDim items(0 To n - 1, 0 To 1)
    For i = 0 To n - 1
        items(i, 0) = labels(i)
        items(i, 1) = FormatMeasure(res, values(i))
    Next
    sel = mUI.ChooseOne("Gruppe wählen, um die zugrunde liegenden Datensätze zu sehen.", ElemGet(eid, "Titel"), items, "Datensätze zeigen")
    If sel < 0 Then Exit Sub
    ShowRecordList ElemGet(eid, "Tabelle"), DrillRids(eid, mNav.gDash, keys(sel)), ElemGet(eid, "Titel") & ": " & labels(sel)
End Sub

'--- Verwaltung ---------------------------------------------------------------------
Public Function CreateDashboard(ByVal nm As String, ByVal desc As String) As String
    Dim m As New cMap, did As String
    did = mCfg.NextId("D", 4)
    m.SetVal "ID", did
    m.SetVal "Name", nm
    m.SetVal "Beschreibung", desc
    m.SetVal "Reihenfolge", mCfg.MaxOrder(SH_DSB, "", "") + 1
    m.SetVal "Filter", ""
    mCfg.AddRow SH_DSB, m
    mLog.Add "Dashboard angelegt", "", did, "", nm
    CreateDashboard = did
End Function

Public Function CreateElement(ByVal did As String, ByVal props As cMap) As String
    Dim m As New cMap, eid As String, k As Variant
    eid = mCfg.NextId("E", 4)
    m.SetVal "ID", eid
    m.SetVal "Dashboard", did
    m.SetVal "Reihenfolge", mCfg.MaxOrder(SH_ELM, "Dashboard", did) + 1
    For Each k In props.Keys
        m.SetVal CStr(k), props.GetVal(CStr(k))
    Next
    mCfg.AddRow SH_ELM, m
    CreateElement = eid
End Function

Public Sub DeleteDashboard(ByVal did As String)
    Dim e As Variant
    For Each e In mCfg.ElementIds(did)
        mCfg.DeleteRow SH_ELM, CStr(e)
    Next
    mCfg.DeleteRow SH_DSB, did
    mLog.Add "Dashboard gelöscht", "", did, ""
End Sub
