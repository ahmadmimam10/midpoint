Attribute VB_Name = "mData"
'==============================================================================
' mData - Datenzugriff auf Fachtabellen und Zuordnungen mit gemeinsamen
' Schreibwegen, Transaktionen (vollstaendige Ruecknahme) und eigener
' Rueckgaengig-Funktion. Datensaetze werden ausschliesslich ueber IDs adressiert.
'==============================================================================
Option Explicit

Private mTables As cMap                 ' TabellenID -> cTableData
Private mRelFwd As cMap                 ' Feld|Von -> "Nach;Nach"
Private mRelBwd As cMap                 ' Feld|Nach -> "Von;Von"
Private mRelBuilt As cMap               ' Feld -> True

Private mTxDepth As Long
Private mTxOps As Collection
Private mTxLabel As String
Private mTxTouched As cMap              ' Tabelle|ID -> Array(tid, rid)
Private mUndo As Collection             ' Array(label, ops, zeit)
Private mUndoing As Boolean
Private Const UNDO_MAX As Long = 20

'==============================================================================
' Tabellenabbild
'==============================================================================
Public Function DataLO(ByVal tid As String) As ListObject
    Set DataLO = ThisWorkbook.Worksheets(mCfg.DataSheetName(tid)).ListObjects(1)
End Function

Public Function TData(ByVal tid As String) As cTableData
    Dim t As cTableData
    If mTables Is Nothing Then Set mTables = New cMap
    If mTables.Exists(tid) Then
        Set TData = mTables.GetVal(tid)
        Exit Function
    End If
    Set t = New cTableData
    t.LoadFrom DataLO(tid)
    t.TableID = tid
    mTables.SetVal tid, t
    Set TData = t
End Function

Public Sub InvalidateTable(ByVal tid As String)
    If Not mTables Is Nothing Then mTables.Remove tid
End Sub

Public Sub InvalidateAllData()
    Set mTables = Nothing
    InvalidateRel
End Sub

Public Sub InvalidateRel()
    Set mRelFwd = Nothing
    Set mRelBwd = Nothing
    Set mRelBuilt = Nothing
    mCfg.Invalidate SH_REL
End Sub

Public Function RecordExists(ByVal tid As String, ByVal rid As String) As Boolean
    If Len(rid) = 0 Or Not mCfg.TableExists(tid) Then Exit Function
    RecordExists = (TData(tid).RowOf(rid) > 0)
End Function

Public Function GetValue(ByVal tid As String, ByVal rid As String, ByVal colName As String) As Variant
    Dim t As cTableData, r As Long
    Set t = TData(tid)
    r = t.RowOf(rid)
    If r = 0 Then GetValue = Empty Else GetValue = t.Val_(r, colName)
End Function

Public Function IsArchivedRec(ByVal tid As String, ByVal rid As String) As Boolean
    Dim t As cTableData, r As Long
    Set t = TData(tid)
    r = t.RowOf(rid)
    If r > 0 Then IsArchivedRec = t.IsArchived(r)
End Function

' Tabelle zu einer Datensatz-ID (IDs sind dateiweit eindeutig)
Public Function TableOfRecord(ByVal rid As String) As String
    Dim ids As Variant, i As Long
    ids = mCfg.TableIds(True)
    For i = 0 To ArrLen(ids) - 1
        If TData(CStr(ids(i))).RowOf(rid) > 0 Then TableOfRecord = CStr(ids(i)): Exit Function
    Next
End Function

' Lesbare Bezeichnung eines Datensatzes
Public Function RecordLabel(ByVal tid As String, ByVal rid As String) As String
    Dim lf As String, v As Variant, f As cField
    lf = mCfg.LabelField(tid)
    If Len(lf) > 0 And mCfg.FieldExists(lf) Then
        Set f = mCfg.Field(lf)
        v = GetValue(tid, rid, lf)
        RecordLabel = mFmt.DisplayText(f, v, rid)
    End If
    If Len(RecordLabel) = 0 Then RecordLabel = "(ohne Bezeichnung)"
End Function

' Zusatzmerkmale zur Unterscheidung gleichnamiger Eintraege
Public Function RecordHint(ByVal tid As String, ByVal rid As String, Optional ByVal maxParts As Long = 2) As String
    Dim f As cField, lf As String, n As Long, s As String, v As Variant, t As String
    lf = mCfg.LabelField(tid)
    For Each f In mCfg.Fields(tid)
        If f.ID <> lf And f.Visible And Not f.IsLinkType And f.FType <> FT_LTEXT Then
            v = GetValue(tid, rid, f.ID)
            If Not IsBlank(v) Then
                t = Truncate(mFmt.DisplayText(f, v, rid), 28)
                If Len(t) > 0 Then
                    s = s & IIf(Len(s) > 0, " · ", "") & t
                    n = n + 1
                    If n >= maxParts Then Exit For
                End If
            End If
        End If
    Next
    RecordHint = s
End Function

'==============================================================================
' Zuordnungen (Verknuepfungen)
'==============================================================================
Private Sub EnsureRelMaps()
    If mRelFwd Is Nothing Then
        Set mRelFwd = New cMap
        Set mRelBwd = New cMap
        Set mRelBuilt = New cMap
    End If
End Sub

' Index einer Verknuepfung aufbauen: Paare sortieren und je Schluessel genau einmal ablegen
' (wiederholtes Ersetzen in der Collection ist bei vielen Zuordnungen sehr langsam)
Private Sub BuildRel(ByVal fid As String)
    Dim t As cTableData, r As Long, m As Long, a() As String, b() As String
    EnsureRelMaps
    If mRelBuilt.Exists(fid) Then Exit Sub
    Set t = mCfg.TD(SH_REL)
    If t.RowCount > 0 Then
        ReDim a(1 To t.RowCount): ReDim b(1 To t.RowCount)
        For r = 1 To t.RowCount
            If CStr(t.Cell(r, 1)) = fid Then
                m = m + 1
                a(m) = CStr(t.Cell(r, 2)): b(m) = CStr(t.Cell(r, 3))
            End If
        Next
    End If
    If m > 0 Then
        ReDim Preserve a(1 To m): ReDim Preserve b(1 To m)
        StoreGrouped mRelFwd, fid, a, b, m
        StoreGrouped mRelBwd, fid, b, a, m
    End If
    mRelBuilt.SetVal fid, True
End Sub

' Sortiert die Paare (Schluessel+Wert zusammengesetzt) und legt je Schluessel "V1;V2" ohne Doppelte ab
Private Sub StoreGrouped(ByVal target As cMap, ByVal fid As String, ByRef keys() As String, ByRef vals() As String, ByVal m As Long)
    Dim c() As String, i As Long, p As Long, k As String, v As String, cur As String, lst As String, prev As String
    ReDim c(1 To m)
    For i = 1 To m
        c(i) = keys(i) & vbTab & vals(i)
    Next
    SortStrings c, 1, m
    For i = 1 To m
        If c(i) <> prev Then
            p = InStr(c(i), vbTab)
            k = Left$(c(i), p - 1): v = Mid$(c(i), p + 1)
            If k <> cur Then
                If Len(cur) > 0 Then target.SetVal fid & "|" & cur, lst
                cur = k: lst = v
            Else
                lst = lst & ";" & v
            End If
            prev = c(i)
        End If
    Next
    If Len(cur) > 0 Then target.SetVal fid & "|" & cur, lst
End Sub

' Quicksort (Textvergleich binaer); kleinerer Teil rekursiv, groesserer iterativ
Private Sub SortStrings(ByRef a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As String, tmp As String
    Do While lo < hi
        i = lo: j = hi
        p = a((lo + hi) \ 2)
        Do While i <= j
            Do While a(i) < p: i = i + 1: Loop
            Do While a(j) > p: j = j - 1: Loop
            If i <= j Then
                tmp = a(i): a(i) = a(j): a(j) = tmp
                i = i + 1: j = j - 1
            End If
        Loop
        If j - lo < hi - i Then
            If lo < j Then SortStrings a, lo, j
            lo = i
        Else
            If i < hi Then SortStrings a, i, hi
            hi = j
        End If
    Loop
End Sub
' Zieldatensaetze einer Verknuepfung (Hin- oder Rueckrichtung) als "R1;R2"
Public Function LinkTargets(ByVal fid As String, ByVal rid As String) As String
    Dim f As cField
    Set f = mCfg.Field(fid)
    If f Is Nothing Then Exit Function
    If f.FType = FT_LINK Then
        BuildRel fid
        LinkTargets = CStr(mRelFwd.GetVal(fid & "|" & rid, ""))
    ElseIf f.FType = FT_BACKLINK Then
        BuildRel f.PairField
        LinkTargets = CStr(mRelBwd.GetVal(f.PairField & "|" & rid, ""))
    End If
End Function

Private Sub RelMapAdd(ByVal fid As String, ByVal fromId As String, ByVal toId As String)
    Dim k1 As String, k2 As String
    EnsureRelMaps
    If Not mRelBuilt.Exists(fid) Then Exit Sub
    k1 = fid & "|" & fromId: k2 = fid & "|" & toId
    mRelFwd.SetVal k1, ListAdd(CStr(mRelFwd.GetVal(k1, "")), toId)
    mRelBwd.SetVal k2, ListAdd(CStr(mRelBwd.GetVal(k2, "")), fromId)
End Sub

Private Sub RelMapRemove(ByVal fid As String, ByVal fromId As String, ByVal toId As String)
    Dim k1 As String, k2 As String
    EnsureRelMaps
    If Not mRelBuilt.Exists(fid) Then Exit Sub
    k1 = fid & "|" & fromId: k2 = fid & "|" & toId
    mRelFwd.SetVal k1, ListRemove(CStr(mRelFwd.GetVal(k1, "")), toId)
    mRelBwd.SetVal k2, ListRemove(CStr(mRelBwd.GetVal(k2, "")), fromId)
End Sub

Private Sub RelWriteAdd(ByVal fid As String, ByVal fromId As String, ByVal toId As String)
    Dim a(1 To 1, 1 To 4) As Variant, lo_ As ListObject
    a(1, 1) = fid: a(1, 2) = fromId: a(1, 3) = toId: a(1, 4) = Now
    Set lo_ = mCfg.LO(SH_REL)
    mProtect.Unprot lo_.Parent
    mCfg.AppendRowsRaw lo_, a
    mCfg.Invalidate SH_REL
    RelMapAdd fid, fromId, toId
    TxRecord Array("RADD", fid, fromId, toId)
End Sub

Private Sub RelWriteRemove(ByVal fid As String, ByVal fromId As String, ByVal toId As String)
    Dim t As cTableData, r As Long, lo_ As ListObject
    Set t = mCfg.TD(SH_REL)
    For r = t.RowCount To 1 Step -1
        If CStr(t.Cell(r, 1)) = fid And CStr(t.Cell(r, 2)) = fromId And CStr(t.Cell(r, 3)) = toId Then
            Set lo_ = mCfg.LO(SH_REL)
            mProtect.Unprot lo_.Parent
            lo_.ListRows(r).Delete
            mCfg.Invalidate SH_REL
            RelMapRemove fid, fromId, toId
            TxRecord Array("RDEL", fid, fromId, toId)
            Exit Sub
        End If
    Next
End Sub

' Setzt die Verknuepfungen eines Datensatzes fuer ein Verknuepfungs- oder
' Rueckverknuepfungsfeld. targets: "R1;R2". Liefert Hinweise zu Umhaengungen.
Public Function SetLinks(ByVal fid As String, ByVal rid As String, ByVal targets As String) As String
    Dim f As cField, cur As String, a As Variant, i As Long, srcF As cField, other As String, note As String
    Dim pairTid As String
    Set f = mCfg.Field(fid)
    cur = LinkTargets(fid, rid)
    If f.FType = FT_LINK Then
        ' entfernen
        a = Split(cur, ";")
        For i = 0 To UBound(a)
            If Len(a(i)) > 0 And Not ListHas(targets, CStr(a(i))) Then
                RelWriteRemove fid, rid, CStr(a(i))
                TouchRecord f.TargetTable, CStr(a(i))
            End If
        Next
        a = Split(targets, ";")
        For i = 0 To UBound(a)
            If Len(a(i)) > 0 And Not ListHas(cur, CStr(a(i))) Then
                RelWriteAdd fid, rid, CStr(a(i))
                TouchRecord f.TargetTable, CStr(a(i))
            End If
        Next
    ElseIf f.FType = FT_BACKLINK Then
        Set srcF = mCfg.Field(f.PairField)
        pairTid = srcF.TableID
        a = Split(cur, ";")
        For i = 0 To UBound(a)
            If Len(a(i)) > 0 And Not ListHas(targets, CStr(a(i))) Then
                RelWriteRemove srcF.ID, CStr(a(i)), rid
                TouchRecord pairTid, CStr(a(i))
            End If
        Next
        a = Split(targets, ";")
        For i = 0 To UBound(a)
            If Len(a(i)) > 0 And Not ListHas(cur, CStr(a(i))) Then
                If Not srcF.Multi Then
                    ' Einfachzuordnung der Gegenseite: bisherige Zuordnung loesen
                    other = LinkTargets(srcF.ID, CStr(a(i)))
                    If Len(other) > 0 Then
                        RelWriteRemove srcF.ID, CStr(a(i)), other
                        TouchRecord f.TableID, other
                        note = note & RecordLabel(pairTid, CStr(a(i))) & " war " & RecordLabel(f.TableID, other) & " zugeordnet. "
                    End If
                End If
                RelWriteAdd srcF.ID, CStr(a(i)), rid
                TouchRecord pairTid, CStr(a(i))
            End If
        Next
    End If
    SetLinks = note
End Function

' Alle Zuordnungen, an denen ein Datensatz beteiligt ist (fuer Loeschpruefung)
Public Function RelCountFor(ByVal rid As String) As Long
    Dim t As cTableData, r As Long, n As Long
    Set t = mCfg.TD(SH_REL)
    For r = 1 To t.RowCount
        If CStr(t.Cell(r, 2)) = rid Or CStr(t.Cell(r, 3)) = rid Then n = n + 1
    Next
    RelCountFor = n
End Function

Public Sub RemoveAllRelsOf(ByVal rid As String)
    Dim t As cTableData, r As Long, fid As String, fr As String, tt As String, f As cField
    Set t = mCfg.TD(SH_REL)
    For r = t.RowCount To 1 Step -1
        Set t = mCfg.TD(SH_REL)
        If r <= t.RowCount Then
            If CStr(t.Cell(r, 2)) = rid Or CStr(t.Cell(r, 3)) = rid Then
                fid = CStr(t.Cell(r, 1)): fr = CStr(t.Cell(r, 2)): tt = CStr(t.Cell(r, 3))
                Set f = mCfg.Field(fid)
                RelWriteRemove fid, fr, tt
                If Not f Is Nothing Then
                    TouchRecord f.TableID, fr
                    TouchRecord f.TargetTable, tt
                End If
            End If
        End If
    Next
End Sub

' Massenanlage von Zuordnungen (Import/Duplizieren); a: (1..n, 1..4)
Public Sub RelAddBulk(ByRef a As Variant)
    Dim lo_ As ListObject, i As Long
    If IsEmpty(a) Then Exit Sub
    Set lo_ = mCfg.LO(SH_REL)
    mProtect.Unprot lo_.Parent
    mCfg.AppendRowsRaw lo_, a
    For i = LBound(a, 1) To UBound(a, 1)
        TxRecord Array("RADD", a(i, 1), a(i, 2), a(i, 3))
    Next
    InvalidateRel
End Sub

' Entfernt alle Zuordnungen eines Feldes (Feld loeschen); ohne Ruecknahme
Public Sub RelPurgeField(ByVal fid As String)
    Dim t As cTableData, r As Long, keep() As Variant, n As Long, c As Long, lo_ As ListObject, removed As Long
    Set t = mCfg.TD(SH_REL)
    If t.RowCount = 0 Then Exit Sub
    ReDim keep(1 To t.RowCount, 1 To 4)
    For r = 1 To t.RowCount
        If CStr(t.Cell(r, 1)) <> fid Then
            n = n + 1
            For c = 1 To 4: keep(n, c) = t.Cell(r, c): Next
        Else
            removed = removed + 1
        End If
    Next
    If removed = 0 Then Exit Sub
    RelRewrite keep, n
End Sub

Public Sub RelRewrite(ByRef keep As Variant, ByVal n As Long)
    Dim lo_ As ListObject, hdr As Range, ws As Worksheet, total As Long
    Set lo_ = mCfg.LO(SH_REL)
    Set ws = lo_.Parent
    mProtect.Unprot ws
    If Not lo_.DataBodyRange Is Nothing Then
        total = lo_.DataBodyRange.Rows.Count
        lo_.DataBodyRange.ClearContents
    End If
    Set hdr = lo_.HeaderRowRange
    If n > 0 Then
        lo_.Resize ws.Range(hdr.Cells(1, 1), ws.Cells(hdr.Row + n, hdr.Column + 3))
        lo_.DataBodyRange.Value = SliceRows(keep, n, 4)
    Else
        lo_.Resize ws.Range(hdr.Cells(1, 1), ws.Cells(hdr.Row + 1, hdr.Column + 3))
    End If
    If total > n Then ws.Range(ws.Cells(hdr.Row + n + 2, hdr.Column), ws.Cells(hdr.Row + total + 1, hdr.Column + 3)).Clear
    InvalidateRel
End Sub

Public Function SliceRows(ByRef a As Variant, ByVal n As Long, ByVal cols As Long) As Variant
    Dim b() As Variant, r As Long, c As Long
    ReDim b(1 To n, 1 To cols)
    For r = 1 To n
        For c = 1 To cols
            b(r, c) = a(r, c)
        Next
    Next
    SliceRows = b
End Function

'==============================================================================
' Datensaetze schreiben
'==============================================================================
Private Function To2D(ByRef a As Variant) As Variant
    Dim b() As Variant, c As Long
    ReDim b(1 To 1, 1 To UBound(a))
    For c = 1 To UBound(a)
        b(1, c) = CellText(a(c))
    Next
    To2D = b
End Function

Private Sub WriteRowToSheet(ByVal tid As String, ByVal r As Long, ByRef a As Variant)
    Dim lo_ As ListObject
    Set lo_ = DataLO(tid)
    mProtect.Unprot lo_.Parent
    lo_.DataBodyRange.Rows(r).Value = To2D(a)
End Sub

' Neue leere Datensatzzeile mit Systemwerten und Standardwerten
Public Function NewRowArray(ByVal tid As String, ByVal rid As String) As Variant
    Dim t As cTableData, a() As Variant, f As cField, c As Long
    Set t = TData(tid)
    ReDim a(1 To t.ColCount)
    a(1) = rid: a(2) = Now: a(3) = Now: a(4) = UserId(): a(5) = Empty
    For Each f In mCfg.Fields(tid, False)
        c = t.ColOf(f.ID)
        If c > 0 Then
            Select Case f.Kind
                Case FT_ERST_UM, FT_AEND_UM: a(c) = a(2)
                Case FT_ERST_VON, FT_AEND_VON: a(c) = a(4)
                Case FT_AUTOID: a(c) = mStruct.NextAutoId(f)
                Case Else: a(c) = Empty
            End Select
        End If
    Next
    NewRowArray = a
End Function

Public Function NewRecordId() As String
    NewRecordId = mCfg.NextId("R", 7)
End Function

Public Sub InsertRecord(ByVal tid As String, ByRef a As Variant)
    Dim lo_ As ListObject, t As cTableData, r As Long
    Set t = TData(tid)
    Set lo_ = DataLO(tid)
    mProtect.Unprot lo_.Parent
    r = mCfg.AppendRowsRaw(lo_, To2D(a))
    If r <> t.RowCount + 1 Then
        InvalidateTable tid
    Else
        t.AppendRow a
    End If
    TxRecord Array("ADD", tid, CStr(a(1)), a)
    TouchRecord tid, CStr(a(1))
End Sub

' Aendert Werte eines Datensatzes; vals: Spaltenname -> Wert
Public Sub UpdateRecord(ByVal tid As String, ByVal rid As String, ByVal vals As cMap, Optional ByVal stamp As Boolean = True)
    Dim t As cTableData, r As Long, oldA As Variant, newA As Variant, k As Variant, c As Long, changed As Boolean
    Set t = TData(tid)
    r = t.RowOf(rid)
    If r = 0 Then RaiseUser "Der Datensatz " & rid & " wurde nicht gefunden. Möglicherweise wurde er inzwischen gelöscht."
    oldA = t.RowArray(r)
    newA = oldA
    For Each k In vals.Keys
        c = t.ColOf(CStr(k))
        If c > SYS_COLS Or CStr(k) = C_ARCH Then
            If Not SameValue(newA(c), vals.GetVal(CStr(k))) Then
                newA(c) = vals.GetVal(CStr(k))
                changed = True
            End If
        End If
    Next
    If Not changed Then Exit Sub
    If stamp Then
        newA(3) = Now
        newA(4) = UserId()
        Dim af As cField
        For Each af In mCfg.Fields(tid, False)
            If af.Kind = FT_AEND_UM Then newA(t.ColOf(af.ID)) = newA(3)
            If af.Kind = FT_AEND_VON Then newA(t.ColOf(af.ID)) = newA(4)
        Next
    End If
    WriteRowToSheet tid, r, newA
    t.SetRow r, newA
    TxRecord Array("ROW", tid, rid, oldA, newA)
    TouchRecord tid, rid
End Sub

Public Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsBlank(a) And IsBlank(b) Then SameValue = True: Exit Function
    If IsBlank(a) Or IsBlank(b) Then Exit Function
    If IsError(a) Or IsError(b) Then Exit Function
    If VarType(a) = vbString Or VarType(b) = vbString Then
        SameValue = (CStr(a) = CStr(b))
    ElseIf VarType(a) = vbBoolean Or VarType(b) = vbBoolean Then
        SameValue = (CBool(a) = CBool(b))
    Else
        SameValue = (Abs(CDbl(a) - CDbl(b)) < 0.0000001)
    End If
End Function

Public Sub SetArchived(ByVal tid As String, ByVal rid As String, ByVal archived As Boolean)
    Dim m As New cMap
    If archived Then m.SetVal C_ARCH, Now Else m.SetVal C_ARCH, Empty
    UpdateRecord tid, rid, m
    TouchLinkedOf tid, rid
End Sub

' Verknuepfte Datensaetze als betroffen markieren (fuer Rollups)
Public Sub TouchLinkedOf(ByVal tid As String, ByVal rid As String)
    Dim f As cField, a As Variant, i As Long, tt As String
    For Each f In mCfg.Fields(tid)
        If f.IsLinkType Then
            a = Split(LinkTargets(f.ID, rid), ";")
            If f.FType = FT_LINK Then tt = f.TargetTable Else tt = mCfg.Field(f.PairField).TableID
            For i = 0 To UBound(a)
                If Len(a(i)) > 0 Then TouchRecord tt, CStr(a(i))
            Next
        End If
    Next
End Sub

' Endgueltiges Loeschen (Zuordnungen werden vorher entfernt; keine Kaskade auf Datensaetze)
Public Sub DeleteRecordHard(ByVal tid As String, ByVal rid As String)
    Dim t As cTableData, r As Long, oldA As Variant, lo_ As ListObject
    TouchLinkedOf tid, rid
    RemoveAllRelsOf rid
    Set t = TData(tid)
    r = t.RowOf(rid)
    If r = 0 Then Exit Sub
    oldA = t.RowArray(r)
    Set lo_ = DataLO(tid)
    mProtect.Unprot lo_.Parent
    lo_.ListRows(r).Delete
    InvalidateTable tid
    TxRecord Array("DEL", tid, rid, oldA)
End Sub

' Massenanlage (Import, Duplizieren). a: (1..n, 1..ColCount)
Public Sub InsertRecordsBulk(ByVal tid As String, ByRef a As Variant)
    Dim lo_ As ListObject, i As Long, c As Long, row() As Variant, n As Long, cols As Long
    Set lo_ = DataLO(tid)
    mProtect.Unprot lo_.Parent
    mCfg.AppendRowsRaw lo_, a
    InvalidateTable tid
    n = UBound(a, 1): cols = UBound(a, 2)
    For i = 1 To n
        ReDim row(1 To cols)
        For c = 1 To cols: row(c) = a(i, c): Next
        TxRecord Array("ADD", tid, CStr(a(i, 1)), row)
        TouchRecord tid, CStr(a(i, 1))
    Next
    mStruct.BackfillAuto tid
End Sub

' Schreibt eine ganze Spalte (Neuberechnung) ohne Rueckgaengig-Eintrag
Public Sub WriteColumnRaw(ByVal tid As String, ByVal colName As String, ByRef values As Variant)
    Dim t As cTableData, c As Long, r As Long, lo_ As ListObject, b() As Variant
    Set t = TData(tid)
    c = t.ColOf(colName)
    If c = 0 Or t.RowCount = 0 Then Exit Sub
    ReDim b(1 To t.RowCount, 1 To 1)
    For r = 1 To t.RowCount
        b(r, 1) = CellText(values(r))
        t.SetCell r, c, values(r)
    Next
    Set lo_ = DataLO(tid)
    mProtect.Unprot lo_.Parent
    lo_.DataBodyRange.Columns(c).Value = b
End Sub

'==============================================================================
' Transaktionen und Rueckgaengig
'==============================================================================
Public Function TxDepth() As Long
    TxDepth = mTxDepth
End Function

Public Sub TxBegin(ByVal label As String)
    If mTxDepth = 0 Then
        Set mTxOps = New Collection
        Set mTxTouched = New cMap
        mTxLabel = label
    End If
    mTxDepth = mTxDepth + 1
End Sub

Public Sub TxRecord(ByVal op As Variant)
    If mTxDepth > 0 Then mTxOps.Add op
End Sub

Public Sub TouchRecord(ByVal tid As String, ByVal rid As String)
    If mTxDepth = 0 Or Len(rid) = 0 Or Len(tid) = 0 Then Exit Sub
    mTxTouched.SetVal tid & "|" & rid, Array(tid, rid)
End Sub

Public Function TxTouched() As cMap
    Set TxTouched = mTxTouched
End Function

Public Sub TxCommit()
    Dim entry As Variant
    If mTxDepth = 0 Then Exit Sub
    If mTxDepth > 1 Then mTxDepth = mTxDepth - 1: Exit Sub
    ' abhaengige Berechnungen innerhalb derselben Transaktion
    mCalc.RecalcTouched mTxTouched
    mTxDepth = 0
    If mTxOps.Count > 0 Then
        If Not mUndoing Then
            If mUndo Is Nothing Then Set mUndo = New Collection
            mUndo.Add Array(mTxLabel, mTxOps, Now)
            Do While mUndo.Count > UNDO_MAX
                mUndo.Remove 1
            Loop
        End If
        mApp.MarkDirty
    End If
    mLog.Flush
    Set mTxOps = Nothing
End Sub

Public Sub TxRollback()
    Dim i As Long
    If mTxDepth = 0 Then Exit Sub
    mTxDepth = 0
    On Error Resume Next
    For i = mTxOps.Count To 1 Step -1
        ApplyInverse mTxOps(i)
    Next
    On Error GoTo 0
    mLog.Discard
    Set mTxOps = Nothing
    InvalidateAllData
End Sub

Public Sub TxAbortIfOpen()
    If mTxDepth > 0 Then TxRollback
End Sub

Private Sub ApplyInverse(ByVal op As Variant)
    Dim t As cTableData, r As Long, lo_ As ListObject
    Select Case op(0)
        Case "ROW"
            InvalidateTable CStr(op(1))
            Set t = TData(CStr(op(1)))
            r = t.RowOf(CStr(op(2)))
            If r > 0 Then
                WriteRowToSheet CStr(op(1)), r, op(3)
                t.SetRow r, op(3)
            End If
        Case "ADD"
            InvalidateTable CStr(op(1))
            Set t = TData(CStr(op(1)))
            r = t.RowOf(CStr(op(2)))
            If r > 0 Then
                Set lo_ = DataLO(CStr(op(1)))
                mProtect.Unprot lo_.Parent
                lo_.ListRows(r).Delete
                InvalidateTable CStr(op(1))
            End If
        Case "DEL"
            Set lo_ = DataLO(CStr(op(1)))
            mProtect.Unprot lo_.Parent
            mCfg.AppendRowsRaw lo_, To2D(op(3))
            InvalidateTable CStr(op(1))
        Case "RADD"
            RelRemoveRaw CStr(op(1)), CStr(op(2)), CStr(op(3))
        Case "RDEL"
            RelAddRaw CStr(op(1)), CStr(op(2)), CStr(op(3))
    End Select
End Sub

Private Sub RelAddRaw(ByVal fid As String, ByVal fromId As String, ByVal toId As String)
    Dim a(1 To 1, 1 To 4) As Variant, lo_ As ListObject
    a(1, 1) = fid: a(1, 2) = fromId: a(1, 3) = toId: a(1, 4) = Now
    Set lo_ = mCfg.LO(SH_REL)
    mProtect.Unprot lo_.Parent
    mCfg.AppendRowsRaw lo_, a
    InvalidateRel
End Sub

Private Sub RelRemoveRaw(ByVal fid As String, ByVal fromId As String, ByVal toId As String)
    Dim t As cTableData, r As Long, lo_ As ListObject
    Set t = mCfg.TD(SH_REL)
    For r = t.RowCount To 1 Step -1
        If CStr(t.Cell(r, 1)) = fid And CStr(t.Cell(r, 2)) = fromId And CStr(t.Cell(r, 3)) = toId Then
            Set lo_ = mCfg.LO(SH_REL)
            mProtect.Unprot lo_.Parent
            lo_.ListRows(r).Delete
            Exit For
        End If
    Next
    InvalidateRel
End Sub

Public Function UndoAvailable() As Boolean
    If mUndo Is Nothing Then Exit Function
    UndoAvailable = (mUndo.Count > 0)
End Function

Public Function UndoLabel() As String
    Dim e As Variant
    If UndoAvailable() Then
        e = mUndo.Item(mUndo.Count)
        UndoLabel = CStr(e(0))
    End If
End Function

Public Function UndoCount() As Long
    If Not mUndo Is Nothing Then UndoCount = mUndo.Count
End Function

' Nimmt die letzte zusammenhaengende Datenaenderung der Sitzung zurueck
Public Function UndoLast() As String
    Dim entry As Variant, ops As Collection, i As Long, op As Variant, touched As New cMap
    If Not UndoAvailable() Then Exit Function
    entry = mUndo(mUndo.Count)
    Set ops = entry(1)
    mUndoing = True
    TxBegin "Rückgängig: " & entry(0)
    On Error GoTo failed
    For i = ops.Count To 1 Step -1
        op = ops(i)
        ApplyInverse op
        Select Case op(0)
            Case "ROW", "ADD", "DEL": TouchRecord CStr(op(1)), CStr(op(2))
            Case "RADD", "RDEL": TouchRel CStr(op(1)), CStr(op(2)), CStr(op(3))
        End Select
    Next
    ' Die inversen Operationen selbst nicht als neue Ruecknahme stapeln
    Set mTxOps = New Collection
    mLog.Add "Rückgängig", "", "", "", CStr(entry(0))
    mUndo.Remove mUndo.Count
    TxCommit
    mUndoing = False
    mApp.MarkDirty
    UndoLast = CStr(entry(0))
    Exit Function
failed:
    mUndoing = False
    mTxDepth = 0
    Set mTxOps = Nothing
    InvalidateAllData
    Err.Raise Err.Number, , "Rücknahme nicht vollständig möglich: " & Err.Description & _
              " Bitte den Datenbestand prüfen oder eine Sicherung wiederherstellen."
End Function

Private Sub TouchRel(ByVal fid As String, ByVal fromId As String, ByVal toId As String)
    Dim f As cField
    Set f = mCfg.Field(fid)
    If f Is Nothing Then Exit Sub
    TouchRecord f.TableID, fromId
    TouchRecord f.TargetTable, toId
End Sub

Public Sub ClearUndo()
    Set mUndo = Nothing
End Sub

