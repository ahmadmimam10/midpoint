Attribute VB_Name = "mLog"
'==============================================================================
' mLog - Aenderungsprotokoll: Zeitpunkt, Windows-Kennung (kein Identitaets-
' nachweis), Art, Tabelle, Objekt-ID und betroffene Feld-IDs. Feldinhalte werden
' bewusst NICHT kopiert (Datensparsamkeit).
'==============================================================================
Option Explicit

Private mBuffer As Collection

Public Sub Add(ByVal art As String, ByVal tid As String, ByVal objId As String, ByVal fieldIds As String, Optional ByVal note As String = "")
    Dim a(1 To 7) As Variant
    a(1) = Now: a(2) = UserId(): a(3) = art: a(4) = tid: a(5) = objId: a(6) = fieldIds: a(7) = note
    If mBuffer Is Nothing Then Set mBuffer = New Collection
    mBuffer.Add a
    If mBuffer.Count >= 1 And mData.TxDepth = 0 Then Flush
End Sub

' Schreibt gepufferte Eintraege in einem Schritt
Public Sub Flush()
    Dim a() As Variant, i As Long, j As Long, item As Variant, lo_ As ListObject, r As Long
    If mBuffer Is Nothing Then Exit Sub
    If mBuffer.Count = 0 Then Exit Sub
    ReDim a(1 To mBuffer.Count, 1 To 7)
    i = 0
    For Each item In mBuffer
        i = i + 1
        For j = 1 To 7
            a(i, j) = item(j)
        Next
    Next
    Set lo_ = mCfg.LO(SH_LOG)
    mProtect.Unprot lo_.Parent
    r = mCfg.AppendRowsRaw(lo_, a)
    Set mBuffer = Nothing
    mCfg.Invalidate SH_LOG
End Sub

Public Sub Discard()
    Set mBuffer = Nothing
End Sub

Public Function RecentRows(ByVal maxRows As Long) As Variant
    ' liefert die letzten Eintraege (neueste zuerst) als 2D-Array (1..n, 1..7)
    Dim t As cTableData, n As Long, r As Long, i As Long, j As Long, a() As Variant
    Set t = mCfg.TD(SH_LOG)
    n = t.RowCount
    If n = 0 Then RecentRows = Empty: Exit Function
    If maxRows > n Then maxRows = n
    ReDim a(1 To maxRows, 1 To 7)
    For i = 1 To maxRows
        r = n - i + 1
        For j = 1 To 7
            a(i, j) = t.Cell(r, j)
        Next
    Next
    RecentRows = a
End Function
