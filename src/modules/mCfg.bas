Attribute VB_Name = "mCfg"
'==============================================================================
' mCfg - Konfiguration: Tabellen, Felder, Optionen, Ansichten, Dashboards,
' Systemwerte. Getrennte ListObjects je Konfigurationsart, im Speicher gepuffert.
' Technische IDs sind stabil; sichtbare Namen frei aenderbar.
'==============================================================================
Option Explicit

Private mCache As cMap          ' Blattname -> cTableData
Private mFieldRev As Long
Private mFields As cMap         ' FeldID -> cField
Private mTableFields As cMap    ' TabellenID -> Collection(cField) sortiert
Private mOptRev As Long
Private mOptByField As cMap     ' FeldID -> Collection(OptionID) sortiert

'--- Spaltendefinitionen (Formatversion 1) ------------------------------------
Public Function CfgHeaders(ByVal sh As String) As Variant
    Select Case sh
        Case SH_SYS: CfgHeaders = Array("Schluessel", "Wert")
        Case SH_CAP: CfgHeaders = Array("ID", "Name", "Tabelle", "Felder", "Modus", "Titel", "Einleitung", "Zeilen", "Absender", "AbsenderFeld", "Schutz", "Erstellt", "Version", "Optionen")
        Case SH_RET: CfgHeaders = Array("DateiID", "Datei", "Erfassung", "Absender", "Zeilen", "Importiert", "Benutzer", "Empfaenger")
        Case SH_TAB: CfgHeaders = Array("ID", "Name", "Beschreibung", "Symbol", "Bezeichnungsfeld", "Reihenfolge", "Status", "Dublettenfelder", "Erstellt", "Blatt")
        Case SH_FLD: CfgHeaders = Array("ID", "Tabelle", "Name", "Typ", "Hilfetext", "Pflicht", "Standard", "Reihenfolge", "Sichtbar", "Gruppe", _
                                        "Eindeutig", "MaxLaenge", "Dezimalstellen", "Minimum", "Maximum", "LeerErlaubt", "Zieltabelle", "Mehrfach", _
                                        "Gegenfeld", "Formel", "Beschreibung", "Ergebnistyp", "NichtVor", "Status", "Erstellt")
        Case SH_OPT: CfgHeaders = Array("ID", "Feld", "Name", "Farbe", "Reihenfolge", "Aktiv", "Abgeschlossen")
        Case SH_REL: CfgHeaders = Array("Feld", "Von", "Nach", "Erstellt")
        Case SH_VIEW: CfgHeaders = Array("ID", "Tabelle", "Name", "Typ", "Reihenfolge", "Spalten", "Filter", "Sortierung", "Gruppierung", "Optionen")
        Case SH_DSB: CfgHeaders = Array("ID", "Name", "Beschreibung", "Reihenfolge", "Filter")
        Case SH_ELM: CfgHeaders = Array("ID", "Dashboard", "Titel", "Typ", "Tabelle", "Kennzahl", "Feld", "Gruppierung", "Datumsstufe", _
                                        "Datumsfeld", "Filter", "Reihenfolge", "Groesse", "Basis", "Optionen")
        Case SH_LOG: CfgHeaders = Array("Zeit", "Benutzer", "Art", "Tabelle", "Objekt", "Felder", "Hinweis")
        Case Else: CfgHeaders = Array()
    End Select
End Function

Public Function CfgSheets() As Variant
    CfgSheets = Array(SH_SYS, SH_TAB, SH_FLD, SH_OPT, SH_REL, SH_VIEW, SH_DSB, SH_ELM, SH_LOG, SH_CAP, SH_RET)
End Function

'--- Puffer ---------------------------------------------------------------------
Public Sub InvalidateAll()
    Set mCache = Nothing
    Set mFields = Nothing
    Set mOptByField = Nothing
    mDocStyle.ResetFontCache
    mApp.gCfgRev = mApp.gCfgRev + 1
End Sub

Public Sub Invalidate(ByVal sh As String)
    If Not mCache Is Nothing Then mCache.Remove sh
    If sh = SH_FLD Or sh = SH_TAB Then Set mFields = Nothing
    If sh = SH_OPT Then Set mOptByField = Nothing
End Sub

Public Function LO(ByVal sh As String) As ListObject
    Set LO = ThisWorkbook.Worksheets(sh).ListObjects(1)
End Function

' Fehlende Konfigurationsspalte (neuere Programmversion) am Ende ergaenzen
Public Sub EnsureColumn(ByVal sh As String, ByVal nm As String)
    Dim lo_ As ListObject, lc As ListColumn
    Set lo_ = LO(sh)
    For Each lc In lo_.ListColumns
        If lc.Name = nm Then Exit Sub
    Next
    mProtect.Unprot lo_.Parent
    Set lc = lo_.ListColumns.Add
    lc.Name = nm
    lc.Range.NumberFormat = "@"
    Invalidate sh
End Sub

Public Function TD(ByVal sh As String) As cTableData
    Dim t As cTableData
    If mCache Is Nothing Then Set mCache = New cMap
    If mCache.Exists(sh) Then
        Set TD = mCache.GetVal(sh)
        Exit Function
    End If
    Set t = New cTableData
    t.LoadFrom LO(sh)
    t.TableID = sh
    mCache.SetVal sh, t
    Set TD = t
End Function

'--- Generischer Zugriff -------------------------------------------------------
Public Function GetV(ByVal sh As String, ByVal id As String, ByVal col As String) As Variant
    Dim t As cTableData, r As Long
    Set t = TD(sh)
    r = t.RowOf(id)
    If r = 0 Then GetV = Empty Else GetV = t.Val_(r, col)
End Function

Public Function GetS(ByVal sh As String, ByVal id As String, ByVal col As String) As String
    GetS = NzS(GetV(sh, id, col))
End Function

Public Sub SetV(ByVal sh As String, ByVal id As String, ByVal col As String, ByVal v As Variant)
    Dim t As cTableData, r As Long, c As Long, lo_ As ListObject
    Set t = TD(sh)
    r = t.RowOf(id)
    c = t.ColOf(col)
    If r = 0 Or c = 0 Then RaiseUser "Konfigurationseintrag " & id & "/" & col & " nicht gefunden."
    Set lo_ = LO(sh)
    mProtect.Unprot lo_.Parent
    If VarType(v) = vbString Then lo_.DataBodyRange.Cells(r, c).NumberFormat = "@"
    lo_.DataBodyRange.Cells(r, c).Value = CellText(v)
    t.SetCell r, c, v
    If sh = SH_FLD Or sh = SH_TAB Then Set mFields = Nothing
    If sh = SH_OPT Then Set mOptByField = Nothing
    If sh <> SH_LOG Then mApp.MarkCfgDirty
End Sub

' Neue Zeile; vals: Spaltenname -> Wert
Public Sub AddRow(ByVal sh As String, ByVal vals As cMap)
    Dim lo_ As ListObject, hdr As Variant, a() As Variant, i As Long, n As Long, rng As Range, r As Long
    Set lo_ = LO(sh)
    mProtect.Unprot lo_.Parent
    hdr = lo_.HeaderRowRange.Value
    n = UBound(hdr, 2)
    ReDim a(1 To 1, 1 To n)
    For i = 1 To n
        a(1, i) = vals.GetVal(CStr(hdr(1, i)), Empty)
    Next
    r = AppendRowsRaw(lo_, a)
    Set rng = lo_.DataBodyRange.Rows(r)
    For i = 1 To n
        If VarType(a(1, i)) = vbString Then rng.Cells(1, i).NumberFormat = "@"
    Next
    rng.Value = SafeArray(a)
    Invalidate sh
    If sh <> SH_LOG And sh <> SH_REL Then mApp.MarkCfgDirty
End Sub

' Mehrere neue Zeilen in einem Schritt (z. B. viele Optionen beim Import); je Zeile Spaltenname -> Wert.
' Ein Schreibvorgang und ein Zuruecksetzen des Puffers statt je Zeile (sonst quadratischer Aufwand).
Public Sub AddRows(ByVal sh As String, ByVal rowsVals As Collection)
    Dim lo_ As ListObject, hdr As Variant, a() As Variant, i As Long, n As Long, j As Long, r As Long, m As cMap, rng As Range
    If rowsVals.Count = 0 Then Exit Sub
    Set lo_ = LO(sh)
    mProtect.Unprot lo_.Parent
    hdr = lo_.HeaderRowRange.Value
    n = UBound(hdr, 2)
    ReDim a(1 To rowsVals.Count, 1 To n)
    For Each m In rowsVals
        j = j + 1
        For i = 1 To n
            a(j, i) = m.GetVal(CStr(hdr(1, i)), Empty)
        Next
    Next
    r = AppendRowsRaw(lo_, a)
    Set rng = lo_.DataBodyRange.Rows(r).Resize(rowsVals.Count)
    For i = 1 To n
        If VarType(a(1, i)) = vbString Then rng.Columns(i).NumberFormat = "@"
    Next
    rng.Value = SafeArray(a)
    Invalidate sh
    If sh <> SH_LOG And sh <> SH_REL Then mApp.MarkCfgDirty
End Sub

Private Function SafeArray(ByVal a As Variant) As Variant
    SafeArray = CellTexts(a)
End Function

' Haengt n Zeilen an ein ListObject an und liefert die Startzeile (1-basiert im Datenbereich)
Public Function AppendRowsRaw(ByVal lo_ As ListObject, ByRef a As Variant) As Long
    Dim cur As Long, add As Long, hdrRow As Long, emptyFirst As Boolean
    add = UBound(a, 1) - LBound(a, 1) + 1
    If lo_.DataBodyRange Is Nothing Then
        ' Liste ohne Datenzeile (z. B. nachdem die letzte Zeile geloescht wurde): Excel legt
        ' per Groessenaenderung keinen Datenbereich an - erst eine leere Zeile einfuegen
        lo_.ListRows.Add
        cur = 0
    Else
        cur = lo_.DataBodyRange.Rows.Count
        If cur = 1 Then
            If Application.WorksheetFunction.CountA(lo_.DataBodyRange) = 0 Then emptyFirst = True
        End If
    End If
    If emptyFirst Then cur = 0
    hdrRow = lo_.HeaderRowRange.Row
    lo_.Resize lo_.Parent.Range(lo_.HeaderRowRange.Cells(1, 1), lo_.Parent.Cells(hdrRow + cur + add, lo_.HeaderRowRange.Column + lo_.ListColumns.Count - 1))
    lo_.DataBodyRange.Rows(cur + 1).Resize(add).Value = SafeArray(a)
    AppendRowsRaw = cur + 1
End Function

Public Sub DeleteRow(ByVal sh As String, ByVal id As String)
    Dim t As cTableData, r As Long, lo_ As ListObject
    Set t = TD(sh)
    r = t.RowOf(id)
    If r = 0 Then Exit Sub
    Set lo_ = LO(sh)
    mProtect.Unprot lo_.Parent
    If lo_.ListRows.Count = 1 Then
        ' letzte Zeile: als leere Platzhalterzeile stehen lassen (wie in einer neuen Datei)
        lo_.DataBodyRange.ClearContents
    Else
        lo_.ListRows(r).Delete
    End If
    Invalidate sh
    mApp.MarkCfgDirty
End Sub

Public Function IdsWhere(ByVal sh As String, ByVal col As String, ByVal value As String, Optional ByVal orderCol As String = "") As Variant
    Dim t As cTableData, r As Long, c As Long, n As Long, ids() As Variant, ord() As Variant, oc As Long
    Set t = TD(sh)
    c = t.ColOf(col)
    If t.RowCount = 0 Or c = 0 Then IdsWhere = Array(): Exit Function
    ReDim ids(1 To t.RowCount): ReDim ord(1 To t.RowCount)
    If Len(orderCol) > 0 Then oc = t.ColOf(orderCol)
    For r = 1 To t.RowCount
        If StrComp(NzS(t.Cell(r, c)), value, vbTextCompare) = 0 Then
            n = n + 1
            ids(n) = t.Cell(r, 1)
            If oc > 0 Then ord(n) = t.Cell(r, oc)
        End If
    Next
    IdsWhere = SortIds(ids, ord, n, oc > 0)
End Function

Public Function AllIds(ByVal sh As String, Optional ByVal orderCol As String = "") As Variant
    Dim t As cTableData, r As Long, ids() As Variant, ord() As Variant, oc As Long
    Set t = TD(sh)
    If t.RowCount = 0 Then AllIds = Array(): Exit Function
    ReDim ids(1 To t.RowCount): ReDim ord(1 To t.RowCount)
    If Len(orderCol) > 0 Then oc = t.ColOf(orderCol)
    For r = 1 To t.RowCount
        ids(r) = t.Cell(r, 1)
        If oc > 0 Then ord(r) = t.Cell(r, oc)
    Next
    AllIds = SortIds(ids, ord, t.RowCount, oc > 0)
End Function

Private Function SortIds(ids() As Variant, ord() As Variant, ByVal n As Long, ByVal doSort As Boolean) As Variant
    Dim res() As Variant, idx() As Long, i As Long, keys(1 To 1) As Variant, dirs(1 To 1) As Long, k() As Variant
    If n = 0 Then SortIds = Array(): Exit Function
    ReDim idx(1 To n)
    For i = 1 To n: idx(i) = i: Next
    If doSort Then
        ReDim k(1 To n)
        For i = 1 To n
            If IsNumeric(ord(i)) And Not IsBlank(ord(i)) Then k(i) = CDbl(ord(i)) Else k(i) = 1E+30
        Next
        keys(1) = k: dirs(1) = 1
        MergeSortIdx idx, keys, 1, dirs
    End If
    ReDim res(0 To n - 1)
    For i = 1 To n: res(i - 1) = CStr(ids(idx(i))): Next
    SortIds = res
End Function

Public Function MaxOrder(ByVal sh As String, ByVal col As String, ByVal value As String) As Double
    Dim t As cTableData, r As Long, c As Long, oc As Long, m As Double
    Set t = TD(sh)
    c = t.ColOf(col): oc = t.ColOf("Reihenfolge")
    For r = 1 To t.RowCount
        Dim hit As Boolean
        If c = 0 Then
            hit = True
        Else
            hit = (StrComp(NzS(t.Cell(r, c)), value, vbTextCompare) = 0)
        End If
        If hit And oc > 0 Then
            If IsNumeric(t.Cell(r, oc)) And Not IsBlank(t.Cell(r, oc)) Then
                If CDbl(t.Cell(r, oc)) > m Then m = CDbl(t.Cell(r, oc))
            End If
        End If
    Next
    MaxOrder = m
End Function

'--- Systemwerte -------------------------------------------------------------------
Public Function SysGet(ByVal key As String, Optional ByVal dflt As String = "") As String
    Dim t As cTableData, r As Long
    On Error GoTo nope
    Set t = TD(SH_SYS)
    r = t.RowOf(key)
    If r = 0 Then SysGet = dflt Else SysGet = NzS(t.Cell(r, 2))
    If r > 0 And Len(SysGet) = 0 Then SysGet = dflt
    Exit Function
nope:
    SysGet = dflt
End Function

Public Function SysLong(ByVal key As String, ByVal dflt As Long) As Long
    Dim s As String
    s = SysGet(key, "")
    If Len(s) = 0 Or Not IsNumeric(s) Then SysLong = dflt Else SysLong = CLng(Val(s))
End Function

Public Sub SysSet(ByVal key As String, ByVal value As String)
    Dim t As cTableData, r As Long, m As cMap, lo_ As ListObject
    Set t = TD(SH_SYS)
    r = t.RowOf(key)
    If r = 0 Then
        Set m = New cMap
        m.SetVal "Schluessel", key
        m.SetVal "Wert", value
        AddRow SH_SYS, m
    Else
        Set lo_ = LO(SH_SYS)
        mProtect.Unprot lo_.Parent
        lo_.DataBodyRange.Cells(r, 2).NumberFormat = "@"
        lo_.DataBodyRange.Cells(r, 2).Value = value
        t.SetCell r, 2, value
    End If
End Sub

' Fortlaufende IDs (z. B. T0001, F0001, R0000001)
Public Function NextId(ByVal prefix As String, ByVal digits As Long) As String
    Dim key As String, n As Long
    key = "Zaehler_" & prefix
    n = SysLong(key, 0) + 1
    SysSet key, CStr(n)
    NextId = prefix & Format$(n, String$(digits, "0"))
End Function

' Reserviert einen Block von IDs; liefert erste Nummer
Public Function ReserveIds(ByVal prefix As String, ByVal count As Long) As Long
    Dim key As String, n As Long
    key = "Zaehler_" & prefix
    n = SysLong(key, 0)
    SysSet key, CStr(n + count)
    ReserveIds = n + 1
End Function

'--- Tabellen -------------------------------------------------------------------------
Public Function TableIds(Optional ByVal includeHidden As Boolean = False) As Variant
    Dim all As Variant, i As Long, res() As Variant, n As Long
    all = AllIds(SH_TAB, "Reihenfolge")
    If ArrLen(all) = 0 Then TableIds = Array(): Exit Function
    ReDim res(0 To UBound(all))
    For i = 0 To UBound(all)
        If includeHidden Or GetS(SH_TAB, CStr(all(i)), "Status") <> "ausgeblendet" Then
            res(n) = all(i): n = n + 1
        End If
    Next
    If n = 0 Then TableIds = Array(): Exit Function
    ReDim Preserve res(0 To n - 1)
    TableIds = res
End Function

Public Function TableExists(ByVal tid As String) As Boolean
    If Len(tid) = 0 Then Exit Function
    TableExists = (TD(SH_TAB).RowOf(tid) > 0)
End Function

Public Function TableName(ByVal tid As String) As String
    TableName = GetS(SH_TAB, tid, "Name")
    If Len(TableName) = 0 Then TableName = "(gelöschte Tabelle)"
End Function

Public Function TableIcon(ByVal tid As String) As String
    TableIcon = IconChar(GetS(SH_TAB, tid, "Symbol"))
End Function

Public Function LabelField(ByVal tid As String) As String
    LabelField = GetS(SH_TAB, tid, "Bezeichnungsfeld")
End Function

Public Function DataSheetName(ByVal tid As String) As String
    DataSheetName = DATA_PREFIX & tid
End Function

Public Function TableIdByName(ByVal nm As String) As String
    Dim ids As Variant, i As Long
    ids = TableIds(True)
    For i = 0 To ArrLen(ids) - 1
        If StrComp(TableName(CStr(ids(i))), Trim$(nm), vbTextCompare) = 0 Then TableIdByName = CStr(ids(i)): Exit Function
    Next
End Function

'--- Felder -----------------------------------------------------------------------------
Private Sub EnsureFields()
    Dim t As cTableData, r As Long, f As cField, col As Collection, tids As Variant, i As Long
    Dim ordered As Variant, j As Long
    If Not mFields Is Nothing Then Exit Sub
    Set mFields = New cMap
    Set mTableFields = New cMap
    Set t = TD(SH_FLD)
    For r = 1 To t.RowCount
        Set f = New cField
        f.ID = NzS(t.Cell(r, t.ColOf("ID")))
        f.TableID = NzS(t.Cell(r, t.ColOf("Tabelle")))
        f.Name = NzS(t.Cell(r, t.ColOf("Name")))
        f.Kind = NzS(t.Cell(r, t.ColOf("Typ")))
        f.FType = BaseTypeOf(f.Kind)
        f.Help = NzS(t.Cell(r, t.ColOf("Hilfetext")))
        f.Required = (NzS(t.Cell(r, t.ColOf("Pflicht"))) = "Ja")
        f.DefaultVal = NzS(t.Cell(r, t.ColOf("Standard")))
        f.Order = Val(NzS(t.Cell(r, t.ColOf("Reihenfolge"))))
        f.Visible = (NzS(t.Cell(r, t.ColOf("Sichtbar"))) <> "Nein")
        f.Group = NzS(t.Cell(r, t.ColOf("Gruppe")))
        f.Unique = (NzS(t.Cell(r, t.ColOf("Eindeutig"))) = "Ja")
        f.MaxLen = CLng(Val(NzS(t.Cell(r, t.ColOf("MaxLaenge")))))
        f.Decimals = CLng(Val(NzS(t.Cell(r, t.ColOf("Dezimalstellen")))))
        If IsBlank(t.Cell(r, t.ColOf("Minimum"))) Then f.MinVal = Empty Else f.MinVal = InvNum(NzS(t.Cell(r, t.ColOf("Minimum"))))
        If IsBlank(t.Cell(r, t.ColOf("Maximum"))) Then f.MaxVal = Empty Else f.MaxVal = InvNum(NzS(t.Cell(r, t.ColOf("Maximum"))))
        f.AllowEmpty = (NzS(t.Cell(r, t.ColOf("LeerErlaubt"))) = "Ja") And f.Kind <> FT_KONTROLL
        f.TargetTable = NzS(t.Cell(r, t.ColOf("Zieltabelle")))
        f.Multi = (NzS(t.Cell(r, t.ColOf("Mehrfach"))) = "Ja")
        f.PairField = NzS(t.Cell(r, t.ColOf("Gegenfeld")))
        f.Formula = NzS(t.Cell(r, t.ColOf("Formel")))
        f.FormulaDesc = NzS(t.Cell(r, t.ColOf("Beschreibung")))
        f.ResultType = NzS(t.Cell(r, t.ColOf("Ergebnistyp")))
        f.NotBefore = NzS(t.Cell(r, t.ColOf("NichtVor")))
        f.Status = NzS(t.Cell(r, t.ColOf("Status")))
        f.CfgRow = r
        If Len(f.ID) > 0 Then mFields.SetVal f.ID, f
    Next
    ' je Tabelle sortiert
    tids = TableIds(True)
    For i = 0 To ArrLen(tids) - 1
        Set col = New Collection
        ordered = IdsWhere(SH_FLD, "Tabelle", CStr(tids(i)), "Reihenfolge")
        For j = 0 To ArrLen(ordered) - 1
            col.Add mFields.GetVal(CStr(ordered(j)))
        Next
        mTableFields.SetVal CStr(tids(i)), col
    Next
End Sub

Public Function Field(ByVal fid As String) As cField
    EnsureFields
    If mFields.Exists(fid) Then Set Field = mFields.GetVal(fid)
End Function

Public Function FieldExists(ByVal fid As String) As Boolean
    EnsureFields
    FieldExists = mFields.Exists(fid)
End Function

' Felder einer Tabelle in Reihenfolge; activeOnly blendet ausgeblendete aus
Public Function Fields(ByVal tid As String, Optional ByVal activeOnly As Boolean = True) As Collection
    Dim res As New Collection, f As cField, src As Collection
    EnsureFields
    If mTableFields.Exists(tid) Then
        Set src = mTableFields.GetVal(tid)
        For Each f In src
            If Not activeOnly Or f.IsActive Then res.Add f
        Next
    End If
    Set Fields = res
End Function

Public Function FieldName(ByVal fid As String) As String
    Dim f As cField
    Set f = Field(fid)
    If f Is Nothing Then FieldName = "(entfernt)" Else FieldName = f.Name
End Function

Public Function FieldIdByName(ByVal tid As String, ByVal nm As String) As String
    Dim f As cField
    For Each f In Fields(tid, False)
        If StrComp(f.Name, Trim$(nm), vbTextCompare) = 0 Then FieldIdByName = f.ID: Exit Function
    Next
End Function

'--- Optionen ----------------------------------------------------------------------------
Private Sub EnsureOpts()
    Dim t As cTableData, r As Long, fc As Long, col As Collection, ids As Variant, i As Long, k As Variant
    Dim fieldIds As New cMap
    If Not mOptByField Is Nothing Then Exit Sub
    Set mOptByField = New cMap
    Set t = TD(SH_OPT)
    fc = t.ColOf("Feld")
    For r = 1 To t.RowCount
        fieldIds.SetVal NzS(t.Cell(r, fc)), True
    Next
    For Each k In fieldIds.Keys
        ids = IdsWhere(SH_OPT, "Feld", CStr(k), "Reihenfolge")
        Set col = New Collection
        For i = 0 To ArrLen(ids) - 1
            col.Add CStr(ids(i))
        Next
        mOptByField.SetVal CStr(k), col
    Next
End Sub

Public Function OptionIds(ByVal fid As String, Optional ByVal activeOnly As Boolean = False) As Collection
    Dim res As New Collection, v As Variant
    EnsureOpts
    If mOptByField.Exists(fid) Then
        For Each v In mOptByField.GetVal(fid)
            If Not activeOnly Or OptActive(CStr(v)) Then res.Add CStr(v)
        Next
    End If
    Set OptionIds = res
End Function

Public Function OptName(ByVal oid As String) As String
    OptName = GetS(SH_OPT, oid, "Name")
End Function

Public Function OptColor(ByVal oid As String) As String
    OptColor = GetS(SH_OPT, oid, "Farbe")
    If Len(OptColor) = 0 Then OptColor = "grau"
End Function

Public Function OptActive(ByVal oid As String) As Boolean
    OptActive = (GetS(SH_OPT, oid, "Aktiv") <> "Nein")
End Function

Public Function OptDone(ByVal oid As String) As Boolean
    OptDone = (GetS(SH_OPT, oid, "Abgeschlossen") = "Ja")
End Function

Public Function OptIdByName(ByVal fid As String, ByVal nm As String) As String
    Dim v As Variant
    nm = Trim$(nm)
    If Len(nm) = 0 Then Exit Function
    For Each v In OptionIds(fid)
        If StrComp(OptName(CStr(v)), nm, vbTextCompare) = 0 Then OptIdByName = CStr(v): Exit Function
    Next
End Function

Public Function OptOrderIndex(ByVal fid As String, ByVal oid As String) As Long
    Dim v As Variant, i As Long
    For Each v In OptionIds(fid)
        i = i + 1
        If CStr(v) = oid Then OptOrderIndex = i: Exit Function
    Next
    OptOrderIndex = 9999
End Function

'--- Ansichten -----------------------------------------------------------------------------
Public Function ViewIds(ByVal tid As String) As Variant
    ViewIds = IdsWhere(SH_VIEW, "Tabelle", tid, "Reihenfolge")
End Function

Public Function ViewGet(ByVal vid As String, ByVal col As String) As String
    ViewGet = GetS(SH_VIEW, vid, col)
End Function

Public Function ViewOpt(ByVal vid As String, ByVal key As String, Optional ByVal dflt As String = "") As String
    ViewOpt = KVGet(ViewGet(vid, "Optionen"), key, dflt)
End Function

Public Sub ViewSetOpt(ByVal vid As String, ByVal key As String, ByVal value As String)
    SetV SH_VIEW, vid, "Optionen", KVSet(ViewGet(vid, "Optionen"), key, value)
End Sub

Public Function ViewExists(ByVal vid As String) As Boolean
    If Len(vid) = 0 Then Exit Function
    ViewExists = (TD(SH_VIEW).RowOf(vid) > 0)
End Function

'--- Dashboards ----------------------------------------------------------------------------
Public Function DashIds() As Variant
    DashIds = AllIds(SH_DSB, "Reihenfolge")
End Function

Public Function ElementIds(ByVal did As String) As Variant
    ElementIds = IdsWhere(SH_ELM, "Dashboard", did, "Reihenfolge")
End Function
