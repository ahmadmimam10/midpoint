Attribute VB_Name = "mStruct"
'==============================================================================
' mStruct - Strukturaenderungen: Tabellen, Felder, Auswahloptionen, Ansichten.
' Interne IDs bleiben bei Umbenennen/Verschieben stabil. Vor dem Entfernen
' werden Abhaengigkeiten ermittelt; Typwechsel erhalten eine Konfliktvorschau.
' Vorlagen nutzen dieselbe Engine (konfigurationsbasierte Starter).
'==============================================================================
Option Explicit

'==============================================================================
' Tabellen
'==============================================================================
Public Function CreateTable(ByVal nm As String, ByVal desc As String, Optional ByVal icon As Long = &HE8FD&) As String
    Dim tid As String, ws As Worksheet, hdr As Variant, i As Long, lo_ As ListObject, m As New cMap
    nm = Trim$(nm)
    If Len(nm) = 0 Then RaiseUser "Bitte einen Tabellennamen angeben."
    If Len(mCfg.TableIdByName(nm)) > 0 Then RaiseUser "Eine Tabelle mit dem Namen „" & nm & "“ gibt es bereits."
    tid = mCfg.NextId("T", 4)
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = mCfg.DataSheetName(tid)
    hdr = Array(C_ID, C_CREATED, C_CHANGED, C_USER, C_ARCH)
    For i = 0 To 4
        ws.Cells(1, i + 1).Value = hdr(i)
    Next
    ws.Columns(1).NumberFormat = "@"
    ws.Columns(4).NumberFormat = "@"
    Set lo_ = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E2"), , xlYes)
    lo_.Name = mCfg.DataSheetName(tid)
    lo_.TableStyle = "TableStyleLight1"
    ws.Visible = xlSheetVeryHidden
    m.SetVal "ID", tid
    m.SetVal "Name", nm
    m.SetVal "Beschreibung", desc
    m.SetVal "Symbol", CStr(icon)
    m.SetVal "Reihenfolge", mCfg.MaxOrder(SH_TAB, "", "") + 1
    m.SetVal "Status", "aktiv"
    m.SetVal "Erstellt", Now
    m.SetVal "Blatt", ws.Name
    mCfg.AddRow SH_TAB, m
    mLog.Add "Tabelle angelegt", tid, tid, "", nm
    mApp.MarkCfgDirty
    CreateTable = tid
End Function

Public Sub UpdateTable(ByVal tid As String, ByVal nm As String, ByVal desc As String, ByVal icon As String)
    Dim other As String
    nm = Trim$(nm)
    If Len(nm) = 0 Then RaiseUser "Bitte einen Tabellennamen angeben."
    other = mCfg.TableIdByName(nm)
    If Len(other) > 0 And other <> tid Then RaiseUser "Eine andere Tabelle heißt bereits „" & nm & "“."
    mCfg.SetV SH_TAB, tid, "Name", nm
    mCfg.SetV SH_TAB, tid, "Beschreibung", desc
    If Len(icon) > 0 Then mCfg.SetV SH_TAB, tid, "Symbol", icon
    mLog.Add "Tabelle geändert", tid, tid, ""
End Sub

Public Sub SetLabelField(ByVal tid As String, ByVal fid As String)
    mCfg.SetV SH_TAB, tid, "Bezeichnungsfeld", fid
End Sub

Public Sub SetTableHidden(ByVal tid As String, ByVal hidden As Boolean)
    mCfg.SetV SH_TAB, tid, "Status", IIf(hidden, "ausgeblendet", "aktiv")
    mLog.Add IIf(hidden, "Tabelle ausgeblendet", "Tabelle eingeblendet"), tid, tid, ""
End Sub

Public Sub MoveTable(ByVal tid As String, ByVal delta As Long)
    MoveInOrder SH_TAB, "", "", tid, delta
End Sub

' Verschiebt einen Eintrag in einer nach "Reihenfolge" sortierten Liste
Public Sub MoveInOrder(ByVal sh As String, ByVal filterCol As String, ByVal filterVal As String, ByVal id As String, ByVal delta As Long)
    Dim ids As Variant, i As Long, pos As Long, other As String, o1 As Variant, o2 As Variant
    If Len(filterCol) > 0 Then ids = mCfg.IdsWhere(sh, filterCol, filterVal, "Reihenfolge") Else ids = mCfg.AllIds(sh, "Reihenfolge")
    pos = -1
    For i = 0 To ArrLen(ids) - 1
        If CStr(ids(i)) = id Then pos = i
    Next
    If pos < 0 Then Exit Sub
    If pos + delta < 0 Or pos + delta > UBound(ids) Then Exit Sub
    ' Reihenfolge neu durchnummerieren, dann tauschen
    For i = 0 To UBound(ids)
        mCfg.SetV sh, CStr(ids(i)), "Reihenfolge", i + 1
    Next
    other = CStr(ids(pos + delta))
    mCfg.SetV sh, id, "Reihenfolge", pos + delta + 1
    mCfg.SetV sh, other, "Reihenfolge", pos + 1
End Sub

'==============================================================================
' Felder
'==============================================================================
' props: Hilfetext, Pflicht, Standard, Gruppe, Eindeutig, MaxLaenge, Dezimalstellen, Minimum, Maximum,
'        LeerErlaubt, Zieltabelle, Mehrfach, Formel, Beschreibung, Ergebnistyp, NichtVor, Sichtbar, Gegenname
Public Function AddField(ByVal tid As String, ByVal nm As String, ByVal ftype As String, Optional ByVal props As cMap = Nothing) As String
    Dim kind As String, fid As String, m As New cMap, k As Variant, lo_ As ListObject, lc As ListColumn, pairId As String, tgt As String
    nm = Trim$(nm)
    If Len(nm) = 0 Then RaiseUser "Bitte einen Feldnamen angeben."
    If Len(mCfg.FieldIdByName(tid, nm)) > 0 Then RaiseUser "Diese Tabelle hat bereits ein Feld „" & nm & "“."
    If props Is Nothing Then Set props = New cMap
    If ftype = FT_ROLLUP Then ftype = FT_FORMEL
    kind = ftype
    ftype = BaseTypeOf(kind)
    fid = mCfg.NextId("F", 4)
    m.SetVal "ID", fid
    m.SetVal "Tabelle", tid
    m.SetVal "Name", nm
    m.SetVal "Typ", kind
    m.SetVal "Reihenfolge", mCfg.MaxOrder(SH_FLD, "Tabelle", tid) + 1
    m.SetVal "Sichtbar", "Ja"
    m.SetVal "Pflicht", "Nein"
    m.SetVal "Status", "aktiv"
    m.SetVal "Erstellt", Now
    For Each k In props.Keys
        If CStr(k) <> "Gegenname" And CStr(k) <> "Optionen" Then m.SetVal CStr(k), props.GetVal(CStr(k))
    Next
    If ftype = FT_DATUM Or ftype = FT_DATUMZEIT Or ftype = FT_ZAHL Then
        If Not m.Exists("Dezimalstellen") Then m.SetVal "Dezimalstellen", IIf(ftype = FT_ZAHL, 0, Empty)
    End If
    If ftype = FT_GELD Then m.SetVal "Dezimalstellen", 2
    mCfg.AddRow SH_FLD, m
    ' Datenspalte
    Set lo_ = mData.DataLO(tid)
    mProtect.Unprot lo_.Parent
    Set lc = lo_.ListColumns.Add
    lc.Name = fid
    ApplyColumnFormat tid, fid, ftype
    mData.InvalidateTable tid
    mCfg.Invalidate SH_FLD
    ' Bezeichnungsfeld setzen, falls noch keins
    If Len(mCfg.LabelField(tid)) = 0 And (ftype = FT_TEXT Or ftype = FT_FORMEL) Then SetLabelField tid, fid
    ' Gegenseite einer Verknuepfung
    If ftype = FT_LINK Then
        tgt = CStr(props.GetVal("Zieltabelle", ""))
        If Not mCfg.TableExists(tgt) Then RaiseUser "Bitte eine Zieltabelle für die Verknüpfung wählen."
        pairId = AddBacklink(tgt, fid, CStr(props.GetVal("Gegenname", mCfg.TableName(tid))))
        mCfg.SetV SH_FLD, fid, "Gegenfeld", pairId
    End If
    ' Optionen
    If props.Exists("Optionen") Then AddOptionsFromSpec fid, CStr(props.GetVal("Optionen"))
    If ftype = FT_AUSWAHL Or ftype = FT_STATUS Or ftype = FT_MEHRFACH Then RebuildLists
    If IsAutoKind(kind) Then BackfillAuto tid, fid
    mLog.Add "Feld angelegt", tid, fid, fid, nm & " (" & FieldTypeName(kind) & ")"
    mApp.MarkCfgDirty
    AddField = fid
End Function

' Fuellt automatisch gepflegte Felder (Erstellt/Geaendert um/von, ID) fuer vorhandene Datensaetze
Public Sub BackfillAuto(ByVal tid As String, Optional ByVal onlyFid As String = "")
    Dim f As cField, t As cTableData, r As Long, c As Long, vals() As Variant, changed As Boolean, n As Long, pre As String
    Set t = mData.TData(tid)
    If t.RowCount = 0 Then Exit Sub
    For Each f In mCfg.Fields(tid, False)
        If IsAutoKind(f.Kind) And f.Kind <> FT_BUTTON And (Len(onlyFid) = 0 Or f.ID = onlyFid) Then
            c = t.ColOf(f.ID)
            If c > 0 Then
                ReDim vals(1 To t.RowCount)
                changed = False
                For r = 1 To t.RowCount
                    vals(r) = t.Cell(r, c)
                    If IsBlank(vals(r)) Then
                        Select Case f.Kind
                            Case FT_ERST_UM: vals(r) = t.Cell(r, 2)
                            Case FT_AEND_UM: vals(r) = t.Cell(r, 3)
                            Case FT_ERST_VON, FT_AEND_VON: vals(r) = t.Cell(r, 4)
                            Case FT_AUTOID: vals(r) = NextAutoId(f)
                        End Select
                        changed = True
                    End If
                Next
                If changed Then mData.WriteColumnRaw tid, f.ID, vals
            End If
        End If
    Next
End Sub

' Naechste fortlaufende ID eines ID-Feldes (Praefix aus der Standardwert-Einstellung)
Public Function NextAutoId(ByVal f As cField) As String
    Dim key As String, n As Long
    key = "IDZaehler_" & f.ID
    n = mCfg.SysLong(key, 0) + 1
    mCfg.SysSet key, CStr(n)
    NextAutoId = f.DefaultVal & IIf(Len(f.DefaultVal) > 0, "-", "") & CStr(n)
End Function

Private Function AddBacklink(ByVal tid As String, ByVal srcFid As String, ByVal nm As String) As String
    Dim fid As String, m As New cMap, lo_ As ListObject, lc As ListColumn, baseNo As String, i As Long
    baseNo = Trim$(nm)
    If Len(baseNo) = 0 Then baseNo = "Verknüpft"
    nm = baseNo
    Do While Len(mCfg.FieldIdByName(tid, nm)) > 0
        i = i + 1
        nm = baseNo & " (" & i + 1 & ")"
    Loop
    fid = mCfg.NextId("F", 4)
    m.SetVal "ID", fid
    m.SetVal "Tabelle", tid
    m.SetVal "Name", nm
    m.SetVal "Typ", FT_BACKLINK
    m.SetVal "Reihenfolge", mCfg.MaxOrder(SH_FLD, "Tabelle", tid) + 1
    m.SetVal "Sichtbar", "Ja"
    m.SetVal "Pflicht", "Nein"
    m.SetVal "Gegenfeld", srcFid
    m.SetVal "Mehrfach", "Ja"
    m.SetVal "Status", "aktiv"
    m.SetVal "Erstellt", Now
    mCfg.AddRow SH_FLD, m
    Set lo_ = mData.DataLO(tid)
    mProtect.Unprot lo_.Parent
    Set lc = lo_.ListColumns.Add
    lc.Name = fid
    mData.InvalidateTable tid
    AddBacklink = fid
End Function

Public Sub ApplyColumnFormat(ByVal tid As String, ByVal fid As String, ByVal ftype As String)
    Dim lo_ As ListObject, ws As Worksheet, c As Long
    Set lo_ = mData.DataLO(tid)
    Set ws = lo_.Parent
    c = lo_.ListColumns(fid).Range.Column
    Select Case ftype
        Case FT_TEXT, FT_LTEXT, FT_PFAD, FT_AUSWAHL, FT_STATUS, FT_MEHRFACH, FT_LINK, FT_BACKLINK
            ws.Columns(c).NumberFormat = "@"
        Case FT_DATUM
            ws.Columns(c).NumberFormat = "yyyy-mm-dd"
        Case FT_DATUMZEIT
            ws.Columns(c).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        Case Else
            ws.Columns(c).NumberFormat = "General"
    End Select
    ws.Cells(lo_.HeaderRowRange.Row, c).NumberFormat = "@"
End Sub

' Aendert Feldeinstellungen (keine Typaenderung); props: Spaltenname -> Wert
Public Sub UpdateField(ByVal fid As String, ByVal props As cMap)
    Dim k As Variant, f As cField, other As String
    Set f = mCfg.Field(fid)
    If props.Exists("Name") Then
        other = mCfg.FieldIdByName(f.TableID, CStr(props.GetVal("Name")))
        If Len(other) > 0 And other <> fid Then RaiseUser "Diese Tabelle hat bereits ein Feld „" & props.GetVal("Name") & "“."
        If Len(Trim$(CStr(props.GetVal("Name")))) = 0 Then RaiseUser "Der Feldname darf nicht leer sein."
    End If
    For Each k In props.Keys
        If CStr(k) <> "Typ" And CStr(k) <> "ID" And CStr(k) <> "Tabelle" And CStr(k) <> "Optionen" And CStr(k) <> "Gegenname" Then
            If NzS(mCfg.GetV(SH_FLD, fid, CStr(k))) <> NzS(props.GetVal(CStr(k))) Then mCfg.SetV SH_FLD, fid, CStr(k), props.GetVal(CStr(k))
        End If
    Next
    ' Beschreibungen von Formeln mit neuem Namen aktualisieren
    RefreshFormulaDescriptions
    mLog.Add "Feld geändert", f.TableID, fid, fid, Join(props.Keys, ",")
End Sub

Public Sub RefreshFormulaDescriptions()
    Dim tids As Variant, i As Long, f As cField
    tids = mCfg.TableIds(True)
    For i = 0 To ArrLen(tids) - 1
        For Each f In mCfg.Fields(CStr(tids(i)), False)
            If f.FType = FT_FORMEL Then
                If mCalc.Describe(f.Formula) <> f.FormulaDesc Then mCfg.SetV SH_FLD, f.ID, "Beschreibung", mCalc.Describe(f.Formula)
            End If
        Next
    Next
End Sub

Public Sub SetFieldHidden(ByVal fid As String, ByVal hidden As Boolean)
    Dim f As cField
    Set f = mCfg.Field(fid)
    If hidden And mCfg.LabelField(f.TableID) = fid Then RaiseUser "Das Bezeichnungsfeld kann nicht ausgeblendet werden. Wählen Sie vorher ein anderes Bezeichnungsfeld."
    mCfg.SetV SH_FLD, fid, "Status", IIf(hidden, "ausgeblendet", "aktiv")
    mLog.Add IIf(hidden, "Feld ausgeblendet", "Feld eingeblendet"), f.TableID, fid, fid
End Sub

Public Sub MoveField(ByVal fid As String, ByVal delta As Long)
    MoveInOrder SH_FLD, "Tabelle", mCfg.Field(fid).TableID, fid, delta
End Sub

'--- Optionen ---------------------------------------------------------------------------
Public Function AddOption(ByVal fid As String, ByVal nm As String, ByVal color As String, ByVal done As Boolean) As String
    Dim m As New cMap, oid As String, n As Long, keys As Variant
    nm = Trim$(nm)
    If Len(nm) = 0 Then RaiseUser "Bitte einen Namen für die Option angeben."
    If Len(mCfg.OptIdByName(fid, nm)) > 0 Then RaiseUser "Die Option „" & nm & "“ gibt es bei diesem Feld bereits."
    If InStr(nm, ";") > 0 Or InStr(nm, ",") > 0 Then RaiseUser "Optionsnamen dürfen kein Semikolon und kein Komma enthalten (diese trennen Mehrfachangaben)."
    If Len(color) = 0 Then
        keys = PaletteKeys()
        n = mCfg.OptionIds(fid).Count
        color = CStr(keys((n + 5) Mod (UBound(keys) + 1)))
    End If
    oid = mCfg.NextId("O", 4)
    m.SetVal "ID", oid
    m.SetVal "Feld", fid
    m.SetVal "Name", nm
    m.SetVal "Farbe", color
    m.SetVal "Reihenfolge", mCfg.MaxOrder(SH_OPT, "Feld", fid) + 1
    m.SetVal "Aktiv", "Ja"
    m.SetVal "Abgeschlossen", IIf(done, "Ja", "Nein")
    mCfg.AddRow SH_OPT, m
    AddOption = oid
End Function

' Viele Optionen auf einmal (Import, Liste übernehmen): vorhandene Namen werden übersprungen,
' Farben und Reihenfolge wie bei AddOption, aber ein Schreibvorgang statt einer Zeile je Option
Public Function AddOptionsBulk(ByVal fid As String, ByVal names As Collection) As Long
    Dim have As New cMap, v As Variant, nm As String, rows As New Collection, m As cMap, keys As Variant
    Dim n As Long, ord As Double, firstNo As Long, i As Long
    For Each v In mCfg.OptionIds(fid)
        have.SetVal LCase$(Trim$(mCfg.OptName(CStr(v)))), True
        n = n + 1
    Next
    keys = PaletteKeys()
    ord = mCfg.MaxOrder(SH_OPT, "Feld", fid)
    For Each v In names
        nm = Trim$(Replace(Replace(CStr(v), ";", " "), ",", " "))
        Do While InStr(nm, "  ") > 0: nm = Replace(nm, "  ", " "): Loop
        If Len(nm) > 0 Then
            If Not have.Exists(LCase$(nm)) Then
                have.SetVal LCase$(nm), True
                Set m = New cMap
                m.SetVal "Feld", fid
                m.SetVal "Name", nm
                m.SetVal "Farbe", CStr(keys((n + 5) Mod (UBound(keys) + 1)))
                ord = ord + 1
                m.SetVal "Reihenfolge", ord
                m.SetVal "Aktiv", "Ja"
                m.SetVal "Abgeschlossen", "Nein"
                rows.Add m
                n = n + 1
            End If
        End If
    Next
    If rows.Count = 0 Then Exit Function
    firstNo = mCfg.ReserveIds("O", rows.Count)
    i = 0
    For Each m In rows
        m.SetVal "ID", "O" & Format$(firstNo + i, "0000")
        i = i + 1
    Next
    mCfg.AddRows SH_OPT, rows
    AddOptionsBulk = rows.Count
End Function

' Spezifikation "Name:farbe:x,Name2:farbe" (x = abgeschlossen)
Public Sub AddOptionsFromSpec(ByVal fid As String, ByVal spec As String)
    Dim a As Variant, i As Long, p As Variant
    If Len(spec) = 0 Then Exit Sub
    a = Split(spec, ",")
    For i = 0 To UBound(a)
        p = Split(a(i), ":")
        If Len(Trim$(p(0))) > 0 Then
            AddOption fid, Trim$(p(0)), ColorOrAuto(ArgOr(p, 1)), (Len(ArgOr(p, 2)) > 0)
        End If
    Next
End Sub

Private Function ColorOrAuto(ByVal s As String) As String
    If Len(Trim$(s)) > 0 Then ColorOrAuto = PaletteKeyFromName(s)
End Function

Private Function ArgOr(ByVal p As Variant, ByVal i As Long) As String
    If i <= UBound(p) Then ArgOr = CStr(p(i))
End Function

Public Sub UpdateOption(ByVal oid As String, ByVal nm As String, ByVal color As String, ByVal active As Boolean, ByVal done As Boolean)
    Dim fid As String, other As String
    fid = mCfg.GetS(SH_OPT, oid, "Feld")
    nm = Trim$(nm)
    If Len(nm) = 0 Then RaiseUser "Der Optionsname darf nicht leer sein."
    If InStr(nm, ";") > 0 Or InStr(nm, ",") > 0 Then RaiseUser "Optionsnamen dürfen kein Semikolon und kein Komma enthalten."
    other = mCfg.OptIdByName(fid, nm)
    If Len(other) > 0 And other <> oid Then RaiseUser "Die Option „" & nm & "“ gibt es bei diesem Feld bereits."
    mCfg.SetV SH_OPT, oid, "Name", nm
    mCfg.SetV SH_OPT, oid, "Farbe", color
    mCfg.SetV SH_OPT, oid, "Aktiv", IIf(active, "Ja", "Nein")
    mCfg.SetV SH_OPT, oid, "Abgeschlossen", IIf(done, "Ja", "Nein")
End Sub

' Anzahl Datensaetze, die eine Option verwenden
Public Function OptionUsage(ByVal oid As String) As Long
    Dim fid As String, f As cField, t As cTableData, c As Long, r As Long, n As Long
    fid = mCfg.GetS(SH_OPT, oid, "Feld")
    Set f = mCfg.Field(fid)
    If f Is Nothing Then Exit Function
    Set t = mData.TData(f.TableID)
    c = t.ColOf(fid)
    If c = 0 Then Exit Function
    For r = 1 To t.RowCount
        If f.FType = FT_MEHRFACH Then
            If ListHas(NzS(t.Cell(r, c)), oid) Then n = n + 1
        ElseIf NzS(t.Cell(r, c)) = oid Then
            n = n + 1
        End If
    Next
    OptionUsage = n
End Function

Public Sub DeleteOption(ByVal oid As String)
    Dim n As Long
    n = OptionUsage(oid)
    If n > 0 Then RaiseUser "Die Option wird noch von " & Plural(n, "Eintrag", "Einträgen") & " verwendet. " & _
                           "Deaktivieren Sie sie für neue Eingaben oder ändern Sie zuerst die betroffenen Einträge (z. B. per Sammeländerung)."
    mCfg.DeleteRow SH_OPT, oid
End Sub

' Hilfslisten fuer Zell-Auswahllisten (benannte Bereiche, regionsunabhaengig)
Public Sub RebuildLists()
    Dim ws As Worksheet, nm As Name, tids As Variant, i As Long, f As cField, col As Long, oid As Variant, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_LST)
    mProtect.Unprot ws
    ws.Cells.Clear
    On Error Resume Next
    For Each nm In ThisWorkbook.Names
        If Left$(nm.Name, 2) = "L_" Then nm.Delete
    Next
    On Error GoTo 0
    ws.Cells(1, 1).Value = "Ja": ws.Cells(2, 1).Value = "Nein"
    ThisWorkbook.Names.Add Name:="L_JANEIN", RefersTo:="='" & SH_LST & "'!$A$1:$A$2"
    col = 1
    tids = mCfg.TableIds(True)
    For i = 0 To ArrLen(tids) - 1
        For Each f In mCfg.Fields(CStr(tids(i)), False)
            If f.IsSingleOption Then
                col = col + 1
                r = 0
                ws.Columns(col).NumberFormat = "@"
                For Each oid In mCfg.OptionIds(f.ID, True)
                    r = r + 1
                    ws.Cells(r, col).Value = mCfg.OptName(CStr(oid))
                Next
                If r = 0 Then r = 1
                ThisWorkbook.Names.Add Name:="L_" & f.ID, RefersTo:="='" & SH_LST & "'!" & ws.Range(ws.Cells(1, col), ws.Cells(r, col)).Address
            End If
        Next
    Next
End Sub

'==============================================================================
' Abhaengigkeiten und Entfernen
'==============================================================================
' Liefert Collection von Array(Art, Beschreibung, Objekt-ID, blockierend 1/0)
Public Function FieldDependencies(ByVal fid As String) As Collection
    Dim res As New Collection, f As cField, tids As Variant, i As Long, g As cField, t As cTableData, r As Long, s As String
    Dim did As String, pf As cField
    Set f = mCfg.Field(fid)
    If mCfg.LabelField(f.TableID) = fid Then res.Add Array("Bezeichnungsfeld", "Ist das Bezeichnungsfeld der Tabelle „" & mCfg.TableName(f.TableID) & "“; vorher ein anderes festlegen.", f.TableID, 1)
    tids = mCfg.TableIds(True)
    For i = 0 To ArrLen(tids) - 1
        For Each g In mCfg.Fields(CStr(tids(i)), False)
            If g.ID <> fid Then
                If g.FType = FT_FORMEL Then
                    If InStr(1, g.Formula & "|", "F:" & fid & "|") > 0 Or InStr(1, g.Formula, "{" & fid & "}") > 0 Then
                        res.Add Array("Berechnung", "Berechnetes Feld „" & g.Name & "“ (" & mCfg.TableName(g.TableID) & ") verwendet dieses Feld.", g.ID, 1)
                    ElseIf (f.IsLinkType) And Len(f.PairField) > 0 Then
                        If InStr(1, g.Formula & "|", "F:" & f.PairField & "|") > 0 Then
                            res.Add Array("Berechnung", "Berechnetes Feld „" & g.Name & "“ wertet die Gegenseite dieser Verknüpfung aus.", g.ID, 1)
                        End If
                    End If
                End If
                If g.NotBefore = fid Then res.Add Array("Datumsregel", "„" & g.Name & "“ darf nicht vor diesem Feld liegen (Regel wird entfernt).", g.ID, 0)
            End If
        Next
    Next
    ' Ansichten
    Set t = mCfg.TD(SH_VIEW)
    For r = 1 To t.RowCount
        s = NzS(t.Val_(r, "Spalten")) & "|" & NzS(t.Val_(r, "Filter")) & "|" & NzS(t.Val_(r, "Sortierung")) & "|" & NzS(t.Val_(r, "Gruppierung")) & "|" & NzS(t.Val_(r, "Optionen"))
        If InStr(1, s, fid) > 0 Then
            If InStr(1, NzS(t.Val_(r, "Optionen")), fid) > 0 And NzS(t.Val_(r, "Typ")) <> VT_TABLE Then
                res.Add Array("Ansicht", "Ansicht „" & NzS(t.Val_(r, "Name")) & "“ (" & NzS(t.Val_(r, "Typ")) & ") nutzt das Feld als Datums-/Gruppenfeld; die Ansicht wird mit entfernt.", CStr(t.Cell(r, 1)), 0)
            Else
                res.Add Array("Ansicht", "Ansicht „" & NzS(t.Val_(r, "Name")) & "“ zeigt, filtert oder sortiert nach dem Feld und wird automatisch bereinigt.", CStr(t.Cell(r, 1)), 0)
            End If
        End If
    Next
    ' Dashboards
    Set t = mCfg.TD(SH_ELM)
    For r = 1 To t.RowCount
        s = NzS(t.Val_(r, "Feld")) & "|" & NzS(t.Val_(r, "Gruppierung")) & "|" & NzS(t.Val_(r, "Datumsfeld")) & "|" & NzS(t.Val_(r, "Filter"))
        If InStr(1, s, fid) > 0 Then
            did = NzS(t.Val_(r, "Dashboard"))
            res.Add Array("Dashboard", "Element „" & NzS(t.Val_(r, "Titel")) & "“ im Dashboard „" & mCfg.GetS(SH_DSB, did, "Name") & "“ verwendet das Feld.", CStr(t.Cell(r, 1)), 1)
        End If
    Next
    If InStr(1, ";" & mCfg.GetS(SH_TAB, f.TableID, "Dublettenfelder") & ";", ";" & fid & ";") > 0 Then
        res.Add Array("Dublettenprüfung", "Das Feld gehört zur Dublettenprüfung (wird daraus entfernt).", f.TableID, 0)
    End If
    If f.IsLinkType And Len(f.PairField) > 0 Then
        Set pf = mCfg.Field(f.PairField)
        If Not pf Is Nothing Then res.Add Array("Gegenseite", "Das Gegenfeld „" & pf.Name & "“ in „" & mCfg.TableName(pf.TableID) & "“ und alle Zuordnungen werden mit entfernt.", pf.ID, 0)
    End If
    Set FieldDependencies = res
End Function

Public Function HasBlocking(ByVal deps As Collection) As Boolean
    Dim d As Variant
    For Each d In deps
        If d(3) = 1 Then HasBlocking = True: Exit Function
    Next
End Function

' Entfernt ein Feld nach geprueften Abhaengigkeiten (nicht blockierende werden bereinigt)
Public Sub DeleteField(ByVal fid As String)
    Dim f As cField, pf As cField, deps As Collection
    Set f = mCfg.Field(fid)
    If f Is Nothing Then Exit Sub
    Set deps = FieldDependencies(fid)
    If HasBlocking(deps) Then RaiseUser "Das Feld wird noch verwendet und kann erst entfernt werden, wenn die blockierenden Abhängigkeiten aufgelöst sind."
    If f.IsLinkType And Len(f.PairField) > 0 Then
        Set pf = mCfg.Field(f.PairField)
        If f.FType = FT_LINK Then mData.RelPurgeField fid Else mData.RelPurgeField f.PairField
        If Not pf Is Nothing Then RemoveFieldRaw pf.ID
    End If
    RemoveFieldRaw fid
    mData.ClearUndo
    RebuildLists
    mLog.Add "Feld entfernt", f.TableID, fid, fid, f.Name
End Sub

Private Sub RemoveFieldRaw(ByVal fid As String)
    Dim f As cField, lo_ As ListObject, t As cTableData, r As Long, oids As Collection, v As Variant, ids As Variant, i As Long
    Set f = mCfg.Field(fid)
    If f Is Nothing Then Exit Sub
    ' Ansichten bereinigen
    ids = mCfg.ViewIds(f.TableID)
    For i = 0 To ArrLen(ids) - 1
        CleanViewOfField CStr(ids(i)), fid
    Next
    ' Datumsregeln anderer Felder
    Set t = mCfg.TD(SH_FLD)
    For r = 1 To t.RowCount
        If NzS(t.Val_(r, "NichtVor")) = fid Then mCfg.SetV SH_FLD, CStr(t.Cell(r, 1)), "NichtVor", ""
    Next
    mCfg.SetV SH_TAB, f.TableID, "Dublettenfelder", ListRemove(mCfg.GetS(SH_TAB, f.TableID, "Dublettenfelder"), fid)
    ' Optionen
    Set oids = mCfg.OptionIds(fid)
    For Each v In oids
        mCfg.DeleteRow SH_OPT, CStr(v)
    Next
    ' Datenspalte
    Set lo_ = mData.DataLO(f.TableID)
    mProtect.Unprot lo_.Parent
    lo_.ListColumns(fid).Delete
    mData.InvalidateTable f.TableID
    mCfg.DeleteRow SH_FLD, fid
    mCfg.InvalidateAll
End Sub

Public Sub CleanViewOfField(ByVal vid As String, ByVal fid As String)
    Dim s As String, parts As Variant, i As Long, res As String, conds As Collection, c As Variant, nc As New Collection, vt As String, opt As String
    vt = mCfg.ViewGet(vid, "Typ")
    opt = mCfg.ViewGet(vid, "Optionen")
    If vt <> VT_TABLE Then
        If KVGet(opt, "datum") = fid Or KVGet(opt, "start") = fid Or KVGet(opt, "gruppe") = fid Or KVGet(opt, "ende") = fid Then
            If KVGet(opt, "ende") = fid And KVGet(opt, "datum") <> fid And KVGet(opt, "start") <> fid Then
                mCfg.ViewSetOpt vid, "ende", ""
            Else
                mCfg.DeleteRow SH_VIEW, vid
                Exit Sub
            End If
        End If
        If KVGet(opt, "fortschritt") = fid Then mCfg.ViewSetOpt vid, "fortschritt", ""
        If ListHas(KVGet(opt, "karte"), fid, ",") Then mCfg.ViewSetOpt vid, "karte", ListRemove(KVGet(opt, "karte"), fid, ",")
    End If
    parts = Split(mCfg.ViewGet(vid, "Spalten"), ";")
    For i = 0 To UBound(parts)
        If Split(parts(i), ":")(0) <> fid Then res = res & IIf(Len(res) > 0, ";", "") & parts(i)
    Next
    mCfg.SetV SH_VIEW, vid, "Spalten", res
    Set conds = mQuery.ParseConds(mCfg.ViewGet(vid, "Filter"))
    For Each c In conds
        If c(1) <> fid Then nc.Add c
    Next
    mCfg.SetV SH_VIEW, vid, "Filter", mQuery.SerializeConds(nc)
    res = ""
    parts = Split(mCfg.ViewGet(vid, "Sortierung"), ";")
    For i = 0 To UBound(parts)
        If Split(parts(i), ":")(0) <> fid Then res = res & IIf(Len(res) > 0, ";", "") & parts(i)
    Next
    mCfg.SetV SH_VIEW, vid, "Sortierung", res
    If KVGet(mCfg.ViewGet(vid, "Gruppierung"), "f") = fid Then mCfg.SetV SH_VIEW, vid, "Gruppierung", ""
    opt = mCfg.ViewGet(vid, "Optionen")
    If InStr(1, KVGet(opt, "summen"), fid) > 0 Then
        parts = Split(KVGet(opt, "summen"), ",")
        res = ""
        For i = 0 To UBound(parts)
            If Split(parts(i), ":")(0) <> fid Then res = res & IIf(Len(res) > 0, ",", "") & parts(i)
        Next
        mCfg.ViewSetOpt vid, "summen", res
    End If
    mNav.ResetState vid
End Sub

'--- Tabelle entfernen -------------------------------------------------------------------
Public Function TableDependencies(ByVal tid As String) As Collection
    Dim res As New Collection, tids As Variant, i As Long, g As cField, t As cTableData, r As Long, f As cField
    tids = mCfg.TableIds(True)
    For i = 0 To ArrLen(tids) - 1
        If CStr(tids(i)) <> tid Then
            For Each g In mCfg.Fields(CStr(tids(i)), False)
                If g.FType = FT_LINK And g.TargetTable = tid Then
                    res.Add Array("Verknüpfung", "Feld „" & g.Name & "“ in „" & mCfg.TableName(g.TableID) & "“ verweist auf diese Tabelle (wird mit entfernt).", g.ID, 0)
                End If
                If g.FType = FT_FORMEL And InStr(g.Formula, "ROLLUP") = 1 Then
                    For Each f In mCfg.Fields(tid, False)
                        If InStr(1, g.Formula & "|", "F:" & f.ID & "|") > 0 Or (Len(f.PairField) > 0 And InStr(1, g.Formula & "|", "F:" & f.PairField & "|") > 0) Then
                            res.Add Array("Berechnung", "Berechnetes Feld „" & g.Name & "“ in „" & mCfg.TableName(g.TableID) & "“ wertet diese Tabelle aus; bitte vorher anpassen oder entfernen.", g.ID, 1)
                            Exit For
                        End If
                    Next
                End If
            Next
        End If
    Next
    Set t = mCfg.TD(SH_ELM)
    For r = 1 To t.RowCount
        If NzS(t.Val_(r, "Tabelle")) = tid Then
            res.Add Array("Dashboard", "Element „" & NzS(t.Val_(r, "Titel")) & "“ im Dashboard „" & mCfg.GetS(SH_DSB, NzS(t.Val_(r, "Dashboard")), "Name") & "“ nutzt diese Tabelle (wird mit entfernt).", CStr(t.Cell(r, 1)), 0)
        End If
    Next
    Set TableDependencies = res
End Function

Public Sub DeleteTable(ByVal tid As String)
    Dim deps As Collection, d As Variant, f As cField, ids As Variant, i As Long, ws As Worksheet, nm As String, t As cTableData, r As Long
    Set deps = TableDependencies(tid)
    If HasBlocking(deps) Then RaiseUser "Die Tabelle wird noch von Berechnungen anderer Tabellen ausgewertet. Bitte diese zuerst anpassen."
    nm = mCfg.TableName(tid)
    ' eingehende Verknuepfungsfelder anderer Tabellen
    For Each d In deps
        If d(0) = "Verknüpfung" Then
            mData.RelPurgeField CStr(d(2))
            RemoveFieldRaw CStr(d(2))
        ElseIf d(0) = "Dashboard" Then
            mCfg.DeleteRow SH_ELM, CStr(d(2))
        End If
    Next
    ' eigene Verknuepfungsfelder samt Gegenseiten
    For Each f In mCfg.Fields(tid, False)
        If f.FType = FT_LINK Then
            mData.RelPurgeField f.ID
            If mCfg.FieldExists(f.PairField) Then
                If mCfg.Field(f.PairField).TableID <> tid Then RemoveFieldRaw f.PairField
            End If
        ElseIf f.FType = FT_BACKLINK Then
            mData.RelPurgeField f.PairField
        End If
    Next
    ' Konfiguration
    ids = mCfg.ViewIds(tid)
    For i = 0 To ArrLen(ids) - 1
        mCfg.DeleteRow SH_VIEW, CStr(ids(i))
    Next
    For Each f In mCfg.Fields(tid, False)
        For Each d In mCfg.OptionIds(f.ID)
            mCfg.DeleteRow SH_OPT, CStr(d)
        Next
        mCfg.DeleteRow SH_FLD, f.ID
    Next
    mCfg.DeleteRow SH_TAB, tid
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Worksheets(mCfg.DataSheetName(tid))
    ws.Visible = xlSheetVisible
    ws.Delete
    Application.DisplayAlerts = True
    mCfg.InvalidateAll
    mData.InvalidateAllData
    mData.ClearUndo
    RebuildLists
    mLog.Add "Tabelle entfernt", tid, tid, "", nm
End Sub

'==============================================================================
' Feldtyp wechseln
'==============================================================================
Public Function CanConvert(ByVal fromT As String, ByVal toT As String) As Boolean
    If fromT = toT Then Exit Function
    Select Case fromT
        Case FT_LINK, FT_BACKLINK, FT_FORMEL: Exit Function
    End Select
    Select Case toT
        ' Verknuepfung/Berechnung brauchen eine Einrichtung; Systemfelder werden automatisch gefuellt
        Case FT_LINK, FT_BACKLINK, FT_FORMEL, FT_ROLLUP, FT_BUTTON, FT_ERST_UM, FT_ERST_VON, FT_AEND_UM, FT_AEND_VON: Exit Function
    End Select
    CanConvert = True
End Function

' Prueft alle Werte; liefert Konflikte als Collection von Array(Datensatz, alter Wert, Grund)
' und die umgewandelten Werte in newVals (1..RowCount)
Public Function ConversionPreview(ByVal fid As String, ByVal toT As String, ByRef newVals As Variant, ByRef createOpts As Collection) As Collection
    Dim f As cField, t As cTableData, r As Long, c As Long, v As Variant, txt As String, res As New Collection, tmp As New cField
    Dim outV As Variant, msg As String, a As Variant, i As Long, oid As String, seen As New cMap, nv() As Variant, dn As Double
    Set f = mCfg.Field(fid)
    Set t = mData.TData(f.TableID)
    c = t.ColOf(fid)
    Set createOpts = New Collection
    tmp.ID = fid: tmp.TableID = f.TableID: tmp.FType = toT: tmp.Decimals = IIf(toT = FT_ZAHL, 6, 2): tmp.AllowEmpty = True
    If t.RowCount = 0 Then
        newVals = Empty
        Set ConversionPreview = res
        Exit Function
    End If
    ReDim nv(1 To t.RowCount)
    For r = 1 To t.RowCount
        v = t.Cell(r, c)
        If IsBlank(v) Then
            nv(r) = Empty
        Else
            ' Ausgangswert als Text
            Select Case f.FType
                Case FT_AUSWAHL, FT_STATUS, FT_MEHRFACH: txt = mFmt.DisplayText(f, v)
                Case FT_ZAHL, FT_GELD: txt = FmtNum(CDbl(v), 6, True)
                Case FT_PROZENT: txt = FmtNum(CDbl(v) * 100, 6, True)
                Case FT_JANEIN: txt = IIf(CBool(v), "Ja", "Nein")
                Case FT_DATUM: txt = FmtDate(v)
                Case FT_DATUMZEIT: txt = FmtDateTime(v)
                Case Else: txt = CStr(v)
            End Select
            Select Case toT
                Case FT_AUSWAHL, FT_STATUS
                    If f.FType = FT_AUSWAHL Or f.FType = FT_STATUS Then
                        nv(r) = v          ' gleiche Options-IDs
                    ElseIf f.FType = FT_MEHRFACH Then
                        If ListCount(CStr(v)) > 1 Then
                            res.Add Array(mData.RecordLabel(f.TableID, t.IdAt(r)), txt, "Mehrere Werte; Einfachauswahl erlaubt nur einen.")
                        Else
                            nv(r) = CStr(v)
                        End If
                    Else
                        ' Text -> neue Optionen
                        If InStr(txt, ";") > 0 Or InStr(txt, ",") > 0 Then
                            res.Add Array(mData.RecordLabel(f.TableID, t.IdAt(r)), txt, "Enthält ; oder , und ist als Option nicht zulässig.")
                        Else
                            If Not seen.Exists(LCase$(Trim$(txt))) Then
                                seen.SetVal LCase$(Trim$(txt)), Trim$(txt)
                                createOpts.Add Trim$(txt)
                            End If
                            nv(r) = "#NEU:" & Trim$(txt)
                        End If
                    End If
                Case FT_MEHRFACH
                    If f.IsOptionType Then
                        nv(r) = CStr(v)
                    Else
                        a = Split(Replace(txt, ",", ";"), ";")
                        txt = ""
                        For i = 0 To UBound(a)
                            If Len(Trim$(a(i))) > 0 Then
                                If Not seen.Exists(LCase$(Trim$(a(i)))) Then
                                    seen.SetVal LCase$(Trim$(a(i))), Trim$(a(i))
                                    createOpts.Add Trim$(a(i))
                                End If
                                txt = txt & IIf(Len(txt) > 0, ";", "") & Trim$(a(i))
                            End If
                        Next
                        nv(r) = "#NEU:" & txt
                    End If
                Case FT_ZAHL, FT_GELD
                    If f.FType = FT_ZAHL Or f.FType = FT_GELD Or f.FType = FT_PROZENT Then
                        nv(r) = CDbl(v)
                        If toT = FT_GELD And Abs(Round(CDbl(v), 2) - CDbl(v)) > 0.0000001 Then
                            res.Add Array(mData.RecordLabel(f.TableID, t.IdAt(r)), txt, "Mehr als 2 Nachkommastellen; Geldbeträge sind centgenau.")
                        End If
                    ElseIf ParseNumberDE(txt, dn) Then
                        nv(r) = dn
                    Else
                        res.Add Array(mData.RecordLabel(f.TableID, t.IdAt(r)), txt, "Keine Zahl.")
                    End If
                Case FT_PROZENT
                    If f.FType = FT_ZAHL Or f.FType = FT_GELD Then
                        nv(r) = CDbl(v) / 100
                    ElseIf ParseNumberDE(txt, dn) Then
                        nv(r) = dn / 100
                    Else
                        res.Add Array(mData.RecordLabel(f.TableID, t.IdAt(r)), txt, "Kein Prozentwert.")
                    End If
                Case Else
                    If mValid.ParseInput(tmp, txt, outV, msg) Then
                        nv(r) = outV
                    Else
                        res.Add Array(mData.RecordLabel(f.TableID, t.IdAt(r)), txt, msg)
                    End If
            End Select
        End If
    Next
    newVals = nv
    Set ConversionPreview = res
End Function

' Fuehrt den Typwechsel aus (nach Vorschau ohne Konflikte oder mit ausdruecklich geleerten Konfliktwerten)
Public Sub ApplyTypeChange(ByVal fid As String, ByVal toT As String, ByVal newVals As Variant, ByVal createOpts As Collection)
    Dim f As cField, t As cTableData, r As Long, v As Variant, a As Variant, i As Long, s As String, oldOpts As Collection
    Set f = mCfg.Field(fid)
    If Not (f.IsOptionType And (toT = FT_AUSWAHL Or toT = FT_STATUS Or toT = FT_MEHRFACH)) And f.IsOptionType Then
        Set oldOpts = mCfg.OptionIds(fid)
    End If
    Dim nm As Variant
    For Each nm In createOpts
        If Len(mCfg.OptIdByName(fid, CStr(nm))) = 0 Then AddOption fid, CStr(nm), "", False
    Next
    Set t = mData.TData(f.TableID)
    If Not IsEmpty(newVals) Then
        For r = 1 To t.RowCount
            v = newVals(r)
            If VarType(v) = vbString Then
                If Left$(v, 5) = "#NEU:" Then
                    a = Split(Mid$(v, 6), ";")
                    s = ""
                    For i = 0 To UBound(a)
                        s = ListAdd(s, mCfg.OptIdByName(fid, CStr(a(i))))
                    Next
                    newVals(r) = s
                End If
            End If
        Next
    End If
    mCfg.SetV SH_FLD, fid, "Typ", toT
    If toT = FT_GELD Then mCfg.SetV SH_FLD, fid, "Dezimalstellen", 2
    If toT = FT_ZAHL And f.FType <> FT_ZAHL Then mCfg.SetV SH_FLD, fid, "Dezimalstellen", 2
    ApplyColumnFormat f.TableID, fid, toT
    If Not IsEmpty(newVals) Then mData.WriteColumnRaw f.TableID, fid, newVals
    If Not oldOpts Is Nothing Then
        For Each v In oldOpts
            mCfg.DeleteRow SH_OPT, CStr(v)
        Next
    End If
    mData.InvalidateTable f.TableID
    mData.ClearUndo
    mCfg.InvalidateAll
    RebuildLists
    mLog.Add "Feldtyp geändert", f.TableID, fid, fid, FieldTypeName(f.FType) & " " & ChrW(&H2192) & " " & FieldTypeName(toT)
End Sub

'==============================================================================
' Ansichten
'==============================================================================
Public Function CreateView(ByVal tid As String, ByVal nm As String, ByVal vt As String, Optional ByVal opts As String = "", _
                           Optional ByVal cols As String = "", Optional ByVal filt As String = "", Optional ByVal sortS As String = "", _
                           Optional ByVal grp As String = "") As String
    Dim vid As String, m As New cMap
    vid = mCfg.NextId("V", 4)
    m.SetVal "ID", vid
    m.SetVal "Tabelle", tid
    m.SetVal "Name", nm
    m.SetVal "Typ", vt
    m.SetVal "Reihenfolge", mCfg.MaxOrder(SH_VIEW, "Tabelle", tid) + 1
    m.SetVal "Spalten", cols
    m.SetVal "Filter", filt
    m.SetVal "Sortierung", sortS
    m.SetVal "Gruppierung", grp
    m.SetVal "Optionen", opts
    mCfg.AddRow SH_VIEW, m
    mLog.Add "Ansicht angelegt", tid, vid, "", nm & " (" & vt & ")"
    CreateView = vid
End Function

Public Function CreateDefaultView(ByVal tid As String) As String
    CreateDefaultView = CreateView(tid, "Alle Einträge", VT_TABLE)
End Function

Public Sub DeleteView(ByVal vid As String)
    Dim tid As String
    tid = mCfg.ViewGet(vid, "Tabelle")
    If ArrLen(mCfg.ViewIds(tid)) <= 1 Then RaiseUser "Die letzte Ansicht einer Tabelle kann nicht gelöscht werden."
    mCfg.DeleteRow SH_VIEW, vid
    mNav.ResetState vid
    mLog.Add "Ansicht gelöscht", tid, vid, ""
End Sub

Public Function DuplicateView(ByVal vid As String, ByVal nm As String) As String
    DuplicateView = CreateView(mCfg.ViewGet(vid, "Tabelle"), nm, mCfg.ViewGet(vid, "Typ"), mCfg.ViewGet(vid, "Optionen"), _
                               mCfg.ViewGet(vid, "Spalten"), mCfg.ViewGet(vid, "Filter"), mCfg.ViewGet(vid, "Sortierung"), mCfg.ViewGet(vid, "Gruppierung"))
End Function

'==============================================================================
' Duplizieren (nur Struktur oder Struktur mit Daten): neue IDs, Beziehungen umgesetzt
'==============================================================================
Public Function DuplicateTable(ByVal tid As String, ByVal newName As String, ByVal withData As Boolean) As String
    Dim nt As String, fmap As New cMap, omap As New cMap, rmap As New cMap, f As cField, nf As String, props As cMap, oid As Variant, noid As String
    Dim t As cTableData, nt_ As cTableData, r As Long, c As Long, a() As Variant, rows() As Variant, n As Long, rid As String, v As Variant
    Dim relT As cTableData, rr As Long, rels() As Variant, nr As Long, src As String, dst As String, ids As Variant, i As Long, k As Variant
    Dim def As String, p As Variant, j As Long, s As String
    nt = CreateTable(newName, mCfg.GetS(SH_TAB, tid, "Beschreibung"), CLng(Val(mCfg.GetS(SH_TAB, tid, "Symbol"))))
    ' Felder (ohne Rueckverknuepfungen eingehender Beziehungen)
    For Each f In mCfg.Fields(tid, False)
        If f.FType <> FT_BACKLINK And f.FType <> FT_FORMEL Then
            Set props = FieldProps(f)
            If f.FType = FT_LINK Then
                If f.TargetTable = tid Then props.SetVal "Zieltabelle", nt
                props.SetVal "Gegenname", newName
            End If
            nf = AddField(nt, f.Name, f.FType, props)
            fmap.SetVal f.ID, nf
            For Each oid In mCfg.OptionIds(f.ID)
                noid = AddOption(nf, mCfg.OptName(CStr(oid)), mCfg.OptColor(CStr(oid)), mCfg.OptDone(CStr(oid)))
                If Not mCfg.OptActive(CStr(oid)) Then mCfg.SetV SH_OPT, noid, "Aktiv", "Nein"
                omap.SetVal CStr(oid), noid
            Next
            If f.Status = "ausgeblendet" Then mCfg.SetV SH_FLD, nf, "Status", "ausgeblendet"
        End If
    Next
    ' Formeln (Feldbezuege umsetzen)
    For Each f In mCfg.Fields(tid, False)
        If f.FType = FT_FORMEL Then
            def = RemapDef(f.Formula, fmap, omap)
            Set props = FieldProps(f)
            props.SetVal "Formel", def
            props.SetVal "Beschreibung", ""
            nf = AddField(nt, f.Name, FT_FORMEL, props)
            fmap.SetVal f.ID, nf
        End If
    Next
    RefreshFormulaDescriptions
    ' Bezeichnungsfeld, Dubletten, Datumsregeln
    If fmap.Exists(mCfg.LabelField(tid)) Then SetLabelField nt, fmap.GetVal(mCfg.LabelField(tid))
    s = ""
    For Each v In Split(mCfg.GetS(SH_TAB, tid, "Dublettenfelder"), ";")
        If fmap.Exists(CStr(v)) Then s = ListAdd(s, fmap.GetVal(CStr(v)))
    Next
    mCfg.SetV SH_TAB, nt, "Dublettenfelder", s
    For Each f In mCfg.Fields(tid, False)
        If Len(f.NotBefore) > 0 And fmap.Exists(f.ID) And fmap.Exists(f.NotBefore) Then mCfg.SetV SH_FLD, fmap.GetVal(f.ID), "NichtVor", fmap.GetVal(f.NotBefore)
    Next
    ' Ansichten
    ids = mCfg.ViewIds(tid)
    For i = 0 To ArrLen(ids) - 1
        CreateView nt, mCfg.ViewGet(CStr(ids(i)), "Name"), mCfg.ViewGet(CStr(ids(i)), "Typ"), RemapText(mCfg.ViewGet(CStr(ids(i)), "Optionen"), fmap, omap), _
                   RemapText(mCfg.ViewGet(CStr(ids(i)), "Spalten"), fmap, omap), RemapText(mCfg.ViewGet(CStr(ids(i)), "Filter"), fmap, omap), _
                   RemapText(mCfg.ViewGet(CStr(ids(i)), "Sortierung"), fmap, omap), RemapText(mCfg.ViewGet(CStr(ids(i)), "Gruppierung"), fmap, omap)
    Next
    RebuildLists
    ' Daten
    If withData Then
        Set t = mData.TData(tid)
        Set nt_ = mData.TData(nt)
        n = mCfg.ReserveIds("R", Application.Max(1, t.RowCount))
        For r = 1 To t.RowCount
            rmap.SetVal t.IdAt(r), "R" & Format$(n + r - 1, "0000000")
        Next
        If t.RowCount > 0 Then
            ReDim rows(1 To t.RowCount, 1 To nt_.ColCount)
            For r = 1 To t.RowCount
                rows(r, 1) = rmap.GetVal(t.IdAt(r))
                rows(r, 2) = Now: rows(r, 3) = Now: rows(r, 4) = UserId(): rows(r, 5) = t.Cell(r, 5)
                For Each k In fmap.Keys
                    c = nt_.ColOf(CStr(fmap.GetVal(CStr(k))))
                    v = t.Val_(r, CStr(k))
                    Set f = mCfg.Field(CStr(k))
                    If f.IsOptionType And Not IsBlank(v) Then
                        s = ""
                        For Each oid In Split(CStr(v), ";")
                            If omap.Exists(CStr(oid)) Then s = ListAdd(s, omap.GetVal(CStr(oid)))
                        Next
                        v = s
                    End If
                    If c > 0 And Not f.IsLinkType Then rows(r, c) = v
                Next
            Next
            mData.TxBegin "Tabelle duplizieren"
            mData.InsertRecordsBulk nt, rows
            ' Zuordnungen der eigenen Verknuepfungsfelder
            Set relT = mCfg.TD(SH_REL)
            ReDim rels(1 To Application.Max(1, relT.RowCount), 1 To 4)
            nr = 0
            For rr = 1 To relT.RowCount
                If fmap.Exists(CStr(relT.Cell(rr, 1))) Then
                    If mCfg.Field(CStr(relT.Cell(rr, 1))).FType = FT_LINK Then
                        src = CStr(relT.Cell(rr, 2)): dst = CStr(relT.Cell(rr, 3))
                        If rmap.Exists(src) Then
                            nr = nr + 1
                            rels(nr, 1) = fmap.GetVal(CStr(relT.Cell(rr, 1)))
                            rels(nr, 2) = rmap.GetVal(src)
                            If rmap.Exists(dst) And mCfg.Field(CStr(relT.Cell(rr, 1))).TargetTable = tid Then rels(nr, 3) = rmap.GetVal(dst) Else rels(nr, 3) = dst
                            rels(nr, 4) = Now
                        End If
                    End If
                End If
            Next
            If nr > 0 Then mData.RelAddBulk mData.SliceRows(rels, nr, 4)
            mData.TxCommit
            mData.ClearUndo
            mCalc.RecalcAll
        End If
    End If
    mLog.Add "Tabelle dupliziert", nt, nt, "", mCfg.TableName(tid) & IIf(withData, " mit Daten", " (nur Struktur)")
    DuplicateTable = nt
End Function

Private Function FieldProps(ByVal f As cField) As cMap
    Dim m As New cMap
    m.SetVal "Hilfetext", f.Help
    m.SetVal "Pflicht", IIf(f.Required, "Ja", "Nein")
    m.SetVal "Standard", f.DefaultVal
    m.SetVal "Gruppe", f.Group
    m.SetVal "Sichtbar", IIf(f.Visible, "Ja", "Nein")
    m.SetVal "Eindeutig", IIf(f.Unique, "Ja", "Nein")
    If f.MaxLen > 0 Then m.SetVal "MaxLaenge", f.MaxLen
    m.SetVal "Dezimalstellen", f.Decimals
    If Not IsEmpty(f.MinVal) Then m.SetVal "Minimum", NumInv(CDbl(f.MinVal))
    If Not IsEmpty(f.MaxVal) Then m.SetVal "Maximum", NumInv(CDbl(f.MaxVal))
    m.SetVal "LeerErlaubt", IIf(f.AllowEmpty, "Ja", "Nein")
    If f.FType = FT_LINK Then m.SetVal "Zieltabelle", f.TargetTable: m.SetVal "Mehrfach", IIf(f.Multi, "Ja", "Nein")
    If f.FType = FT_FORMEL Then m.SetVal "Formel", f.Formula: m.SetVal "Ergebnistyp", f.ResultType
    Set FieldProps = m
End Function

Private Function RemapDef(ByVal def As String, ByVal fmap As cMap, ByVal omap As cMap) As String
    Dim p As Variant, i As Long, res As String, k As Variant
    If mCalc.IsXl(def) Then
        For Each k In fmap.Keys
            def = Replace(def, "{" & CStr(k) & "}", "{" & CStr(fmap.GetVal(CStr(k))) & "}")
        Next
        RemapDef = def
        Exit Function
    End If
    p = Split(def, "|")
    For i = 0 To UBound(p)
        If Left$(p(i), 2) = "F:" Then
            If fmap.Exists(Mid$(p(i), 3)) Then p(i) = "F:" & fmap.GetVal(Mid$(p(i), 3))
        ElseIf Left$(p(i), 2) = "O:" Then
            If omap.Exists(Mid$(p(i), 3)) Then p(i) = "O:" & omap.GetVal(Mid$(p(i), 3))
        End If
        res = res & IIf(i > 0, "|", "") & p(i)
    Next
    RemapDef = res
End Function

Private Function RemapText(ByVal s As String, ByVal fmap As cMap, ByVal omap As cMap) As String
    Dim k As Variant
    For Each k In fmap.Keys
        s = Replace(s, CStr(k), CStr(fmap.GetVal(CStr(k))))
    Next
    For Each k In omap.Keys
        s = Replace(s, CStr(k), CStr(omap.GetVal(CStr(k))))
    Next
    RemapText = s
End Function

'==============================================================================
' Vorlagen (konfigurationsbasiert, Blatt _Vorlagen)
'==============================================================================
Public Function TemplateList() As Variant
    Dim ws As Worksheet, r As Long, res() As Variant, n As Long, p As Variant
    Set ws = ThisWorkbook.Worksheets(SH_TPL)
    ReDim res(0 To 49, 0 To 2)
    r = 1
    Do While Len(NzS(ws.Cells(r, 1).Value)) > 0
        p = Split(CStr(ws.Cells(r, 1).Value), "|")
        If p(0) = "VORLAGE" Then
            res(n, 0) = p(1): res(n, 1) = p(2): res(n, 2) = p(3)
            n = n + 1
        End If
        r = r + 1
    Loop
    If n = 0 Then TemplateList = Empty: Exit Function
    Dim out() As Variant, i As Long
    ReDim out(0 To n - 1, 0 To 2)
    For i = 0 To n - 1
        out(i, 0) = res(i, 0): out(i, 1) = res(i, 1): out(i, 2) = res(i, 2)
    Next
    TemplateList = out
End Function

' Wendet eine Vorlage an; liefert die erste angelegte Tabellen-ID
Public Function ApplyTemplate(ByVal key As String) As String
    Dim ws As Worksheet, r As Long, lnTxt As String, p As Variant, inTpl As Boolean, aliases As New cMap, tid As String, props As cMap
    Dim fid As String, first As String, k As Variant, kv As Variant, i As Long, nm As String, def As String, rt As String
    Set ws = ThisWorkbook.Worksheets(SH_TPL)
    r = 1
    Do While Len(NzS(ws.Cells(r, 1).Value)) > 0
        lnTxt = CStr(ws.Cells(r, 1).Value)
        p = Split(lnTxt, "|")
        If p(0) = "VORLAGE" Then
            inTpl = (p(1) = key)
        ElseIf inTpl Then
            Select Case p(0)
                Case "TABELLE"
                    nm = UniqueTableName(CStr(p(2)))
                    tid = CreateTable(nm, CStr(p(3)), CLng(Val(p(4))))
                    aliases.SetVal CStr(p(1)), tid
                    If Len(first) = 0 Then first = tid
                Case "FELD"
                    tid = aliases.GetVal(CStr(p(1)))
                    Set props = ParseProps(ArgOr(p, 4), tid, aliases)
                    fid = AddField(tid, CStr(p(2)), CStr(p(3)), props)
                    If props.Exists("bez") Then SetLabelField tid, fid
                    If props.Exists("Standard") And (p(3) = FT_AUSWAHL Or p(3) = FT_STATUS) Then
                        mCfg.SetV SH_FLD, fid, "Standard", mCfg.OptIdByName(fid, CStr(props.GetVal("Standard")))
                    End If
                Case "BUTTON"
                    ' BUTTON|Alias|Name|{Feld}=Wert,{Feld}=Wert
                    tid = aliases.GetVal(CStr(p(1)))
                    def = "AKTION"
                    kv = Split(ArgOr(p, 3), ",")
                    For i = 0 To UBound(kv)
                        If InStr(kv(i), "=") > 0 Then
                            def = def & "|F:" & ResolveNames(Left$(kv(i), InStr(kv(i), "=") - 1), tid) & "|" & Esc(Mid$(kv(i), InStr(kv(i), "=") + 1))
                        End If
                    Next
                    Set props = New cMap
                    props.SetVal "Formel", def
                    AddField tid, CStr(p(2)), FT_BUTTON, props
                Case "LINK"
                    tid = aliases.GetVal(CStr(p(1)))
                    Set props = New cMap
                    props.SetVal "Zieltabelle", aliases.GetVal(CStr(p(3)))
                    props.SetVal "Mehrfach", IIf(p(4) = "1", "Ja", "Nein")
                    props.SetVal "Gegenname", CStr(p(5))
                    AddField tid, CStr(p(2)), FT_LINK, props
                Case "FORMEL"
                    tid = aliases.GetVal(CStr(p(1)))
                    def = ResolveTemplateDef(Replace(CStr(p(3)), "~", "|"), tid, aliases)
                    Set props = New cMap
                    props.SetVal "Formel", def
                    rt = CStr(p(4))
                    If Len(rt) = 0 Then rt = mCalc.DefaultResultType(def, tid)
                    props.SetVal "Ergebnistyp", rt
                    props.SetVal "Dezimalstellen", IIf(rt = FT_GELD, 2, IIf(UBound(p) >= 5, Val(ArgOr(p, 5)), 0))
                    props.SetVal "Beschreibung", mCalc.Describe(def)
                    If UBound(p) >= 6 Then props.SetVal "Gruppe", CStr(p(6))
                    AddField tid, CStr(p(2)), FT_FORMEL, props
                Case "ANSICHT"
                    tid = aliases.GetVal(CStr(p(1)))
                    CreateView tid, CStr(p(2)), CStr(p(3)), ResolveNames(ArgOr(p, 4), tid), ResolveNames(ArgOr(p, 5), tid), _
                               ResolveNames(Replace(ArgOr(p, 6), "~", "|"), tid), ResolveNames(ArgOr(p, 7), tid), ResolveNames(ArgOr(p, 8), tid)
                Case "DUBLETTE"
                    tid = aliases.GetVal(CStr(p(1)))
                    mCfg.SetV SH_TAB, tid, "Dublettenfelder", Replace(ResolveNames(CStr(p(2)), tid), ",", ";")
                Case "NICHTVOR"
                    tid = aliases.GetVal(CStr(p(1)))
                    mCfg.SetV SH_FLD, mCfg.FieldIdByName(tid, CStr(p(2))), "NichtVor", mCfg.FieldIdByName(tid, CStr(p(3)))
            End Select
        End If
        r = r + 1
    Loop
    ' Standardansicht fuer Tabellen ohne Ansicht
    For Each k In aliases.Keys
        If ArrLen(mCfg.ViewIds(CStr(aliases.GetVal(CStr(k))))) = 0 Then CreateDefaultView CStr(aliases.GetVal(CStr(k)))
    Next
    RebuildLists
    mCfg.InvalidateAll
    ApplyTemplate = first
End Function

Private Function UniqueTableName(ByVal nm As String) As String
    Dim i As Long, baseNo As String
    baseNo = nm
    Do While Len(mCfg.TableIdByName(nm)) > 0
        i = i + 1
        nm = baseNo & " " & (i + 1)
    Loop
    UniqueTableName = nm
End Function

' "pflicht=1;bez=1;optionen=Offen:blau,Erledigt:gruen:x;dez=2;min=0;std=heute;hilfe=..."
Private Function ParseProps(ByVal s As String, ByVal tid As String, ByVal aliases As cMap) As cMap
    Dim m As New cMap, a As Variant, i As Long, p As Long, k As String, v As String
    If Len(s) > 0 Then
        a = Split(s, ";")
        For i = 0 To UBound(a)
            p = InStr(a(i), "=")
            If p > 0 Then
                k = LCase$(Left$(a(i), p - 1)): v = Mid$(a(i), p + 1)
                Select Case k
                    Case "pflicht": m.SetVal "Pflicht", IIf(v = "1", "Ja", "Nein")
                    Case "bez": m.SetVal "bez", "1"
                    Case "optionen": m.SetVal "Optionen", v
                    Case "dez": m.SetVal "Dezimalstellen", Val(v)
                    Case "min": m.SetVal "Minimum", v
                    Case "max": m.SetVal "Maximum", v
                    Case "std": m.SetVal "Standard", v
                    Case "hilfe": m.SetVal "Hilfetext", v
                    Case "gruppe": m.SetVal "Gruppe", v
                    Case "eindeutig": m.SetVal "Eindeutig", IIf(v = "1", "Ja", "Nein")
                    Case "leer": m.SetVal "LeerErlaubt", IIf(v = "1", "Ja", "Nein")
                    Case "max_len": m.SetVal "MaxLaenge", Val(v)
                End Select
            End If
        Next
    End If
    Set ParseProps = m
End Function

' Ersetzt {Feldname} durch Feld-IDs in Ansichtsoptionen/Spalten etc.
Public Function ResolveNames(ByVal s As String, ByVal tid As String) As String
    Dim p1 As Long, p2 As Long, nm As String, fid As String, opt As String, eq As Long
    Do
        p1 = InStr(s, "{")
        If p1 = 0 Then Exit Do
        p2 = InStr(p1, s, "}")
        If p2 = 0 Then Exit Do
        nm = Mid$(s, p1 + 1, p2 - p1 - 1)
        eq = InStr(nm, "=")
        If eq > 0 Then
            fid = mCfg.FieldIdByName(tid, Left$(nm, eq - 1))
            opt = mCfg.OptIdByName(fid, Mid$(nm, eq + 1))
            s = Left$(s, p1 - 1) & opt & Mid$(s, p2 + 1)
        Else
            fid = mCfg.FieldIdByName(tid, nm)
            s = Left$(s, p1 - 1) & fid & Mid$(s, p2 + 1)
        End If
    Loop
    ResolveNames = s
End Function

' Formeldefinition mit {Feld}, {Alias.Feld}, {Feld=Option}
Private Function ResolveTemplateDef(ByVal def As String, ByVal tid As String, ByVal aliases As cMap) As String
    Dim p1 As Long, p2 As Long, nm As String, dot As Long, t2 As String, eq As Long, fid As String
    Do
        p1 = InStr(def, "{")
        If p1 = 0 Then Exit Do
        p2 = InStr(p1, def, "}")
        nm = Mid$(def, p1 + 1, p2 - p1 - 1)
        t2 = tid
        dot = InStr(nm, ".")
        If dot > 0 Then
            t2 = aliases.GetVal(Left$(nm, dot - 1), tid)
            nm = Mid$(nm, dot + 1)
        End If
        eq = InStr(nm, "=")
        If eq > 0 Then
            fid = mCfg.FieldIdByName(t2, Left$(nm, eq - 1))
            def = Left$(def, p1 - 1) & mCfg.OptIdByName(fid, Mid$(nm, eq + 1)) & Mid$(def, p2 + 1)
        Else
            def = Left$(def, p1 - 1) & mCfg.FieldIdByName(t2, nm) & Mid$(def, p2 + 1)
        End If
    Loop
    ResolveTemplateDef = def
End Function
