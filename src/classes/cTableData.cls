VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTableData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cTableData - Arbeitsspeicherabbild einer Fachtabelle (Array-basiert).
' Zeilen werden immer ueber die Datensatz-ID gefunden, nie ueber die
' sichtbare Zeilennummer einer Ansicht.
'==============================================================================
Option Explicit

Public TableID As String
Public Rev As Long
Public RowCount As Long
Public ColCount As Long
Public Data As Variant          ' (1..RowCount, 1..ColCount)
Private mCol As cMap            ' Spaltenname -> Index
Private mIdx As cMap            ' Datensatz-ID -> Zeile
Private mHeaders As Variant

Public Sub LoadFrom(ByVal lo As ListObject)
    Dim i As Long, hdr As Variant
    Set mCol = New cMap
    Set mIdx = New cMap
    hdr = lo.HeaderRowRange.Value
    ColCount = UBound(hdr, 2)
    mHeaders = hdr
    For i = 1 To ColCount
        mCol.SetVal CStr(hdr(1, i)), i
    Next
    If lo.DataBodyRange Is Nothing Then
        RowCount = 0
        Data = Empty
    Else
        RowCount = lo.DataBodyRange.Rows.Count
        If RowCount = 1 And ColCount = 1 Then
            Dim one(1 To 1, 1 To 1) As Variant
            one(1, 1) = lo.DataBodyRange.Value
            Data = one
        Else
            Data = lo.DataBodyRange.Value
        End If
        ' leere Zeile einer frisch angelegten Tabelle nicht als Datensatz zaehlen
        If RowCount = 1 Then
            If IsBlank(Data(1, 1)) Then RowCount = 0
        End If
    End If
    For i = 1 To RowCount
        ' erstes Vorkommen genuegt (bei Zuordnungen/Protokoll ist Spalte 1 keine eindeutige ID;
        ' wiederholtes Ersetzen desselben Schluessels waere bei vielen Zeilen sehr langsam)
        If Not IsBlank(Data(i, 1)) Then
            If Not mIdx.Exists(CStr(Data(i, 1))) Then mIdx.SetVal CStr(Data(i, 1)), i
        End If
    Next
End Sub

Public Function ColOf(ByVal colName As String) As Long
    ColOf = CLng(mCol.GetVal(colName, 0))
End Function

Public Function HasCol(ByVal colName As String) As Boolean
    HasCol = mCol.Exists(colName)
End Function

Public Function RowOf(ByVal rid As String) As Long
    RowOf = CLng(mIdx.GetVal(rid, 0))
End Function

' Einzelwert lesen. Von aussen IMMER Cell statt Data(r, c) verwenden: Der Zugriff auf das
' oeffentliche Variant-Array kopiert sonst bei jedem Aufruf das gesamte Array (sehr langsam).
Public Function Cell(ByVal r As Long, ByVal c As Long) As Variant
    Cell = Data(r, c)
End Function

Public Function IdAt(ByVal r As Long) As String
    IdAt = CStr(Data(r, 1))
End Function

Public Function Val_(ByVal r As Long, ByVal colName As String) As Variant
    Dim c As Long
    c = ColOf(colName)
    If c = 0 Or r < 1 Or r > RowCount Then Val_ = Empty Else Val_ = Data(r, c)
End Function

Public Function IsArchived(ByVal r As Long) As Boolean
    IsArchived = Not IsBlank(Data(r, 5))
End Function

Public Sub SetCell(ByVal r As Long, ByVal c As Long, ByVal v As Variant)
    Data(r, c) = v
End Sub

' Zeile als 1-dimensionales Array (1..ColCount)
Public Function RowArray(ByVal r As Long) As Variant
    Dim a() As Variant, c As Long
    ReDim a(1 To ColCount)
    For c = 1 To ColCount
        a(c) = Data(r, c)
    Next
    RowArray = a
End Function

Public Sub SetRow(ByVal r As Long, ByRef a As Variant)
    Dim c As Long
    For c = 1 To ColCount
        Data(r, c) = a(c)
    Next
End Sub

' Neue Zeile am Ende anhaengen (nur Abbild; Blatt schreibt mData)
Public Sub AppendRow(ByRef a As Variant)
    Dim nd() As Variant, r As Long, c As Long
    ReDim nd(1 To RowCount + 1, 1 To ColCount)
    For r = 1 To RowCount
        For c = 1 To ColCount
            nd(r, c) = Data(r, c)
        Next
    Next
    For c = 1 To ColCount
        nd(RowCount + 1, c) = a(c)
    Next
    RowCount = RowCount + 1
    Data = nd
    mIdx.SetVal CStr(a(1)), RowCount
End Sub

Public Function HeaderName(ByVal c As Long) As String
    HeaderName = CStr(mHeaders(1, c))
End Function

' Spaltenwerte als 1-basiertes Array
Public Function ColumnArray(ByVal colName As String) As Variant
    Dim a() As Variant, r As Long, c As Long
    c = ColOf(colName)
    If RowCount = 0 Then
        ColumnArray = Array()
        Exit Function
    End If
    ReDim a(1 To RowCount)
    If c > 0 Then
        For r = 1 To RowCount
            a(r) = Data(r, c)
        Next
    End If
    ColumnArray = a
End Function
