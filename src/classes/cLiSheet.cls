VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLiSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cLiSheet - gelesenes Quellblatt fuer "Liste uebernehmen" (nur Werte, keine
' Verbindung zur Quelldatei). Zeilen/Spalten 1-basiert ab Zelle A1 des Blatts.
' Die Werte liegen privat im Objekt; Cell() liest einzelne Zellen, ohne das
' ganze Feld zu kopieren.
'==============================================================================
Option Explicit

Public Name As String
Public R As Long
Public C As Long
Public Formulas As cMap            ' "r|c" -> Formeltext
Public MergeTop As cMap            ' "r|c" -> "r0|c0" (Zelle liegt in einem verbundenen Bereich)
Public HiddenCols As cMap          ' Spalte -> True
Public ColFmt As cMap              ' Spalte -> Zahlenformat (einheitlich oder der letzten Zahl)
Public CellFmt As cMap             ' "r|c" -> Zahlenformat (nur in Spalten mit gemischtem Format)
Public Tables As Collection        ' Array(r1, c1, r2, c2, Name) der Excel-Tabellen (ListObjects)
Public ErrCells As Long             ' Anzahl Fehlerwerte (#WERT! usw.), als leer gelesen
Public IsCsv As Boolean            ' Quelle ist CSV (alle Werte Text)
Public Truncated As Boolean        ' Blatt groesser als die Lesegrenze
Public MergeMaxRow As Long         ' unterste Zeile mit verbundenen Zellen (0 = keine)
Private mV As Variant
' Spalten mit Formeln bzw. eigenen Zellformaten: Nachschlagen nur dort (ein fehlender Schluessel
' kostet in cMap eine Fehlerbehandlung, bei 5.000 Zeilen je Spalte spuerbar)
Private mFormCol() As Boolean
Private mFmtCol() As Boolean

Private Sub Class_Initialize()
    Set Formulas = New cMap
    Set MergeTop = New cMap
    Set HiddenCols = New cMap
    Set ColFmt = New cMap
    Set CellFmt = New cMap
    Set Tables = New Collection
End Sub

Public Sub SetValues(ByRef v As Variant)
    mV = v
    If IsArray(mV) Then
        R = UBound(mV, 1): C = UBound(mV, 2)
    Else
        R = 0: C = 0
    End If
    ReDim mFormCol(0 To C)
    ReDim mFmtCol(0 To C)
End Sub

' Formel, Zellformat und verbundene Zelle merken (nach SetValues aufrufen)
Public Sub AddFormula(ByVal rr As Long, ByVal cc As Long, ByVal txt As String)
    Formulas.SetVal rr & "|" & cc, txt
    If cc >= 1 And cc <= C Then mFormCol(cc) = True
End Sub

Public Sub AddCellFmt(ByVal rr As Long, ByVal cc As Long, ByVal fmt As String)
    CellFmt.SetVal rr & "|" & cc, fmt
    If cc >= 1 And cc <= C Then mFmtCol(cc) = True
End Sub

Public Sub AddMerge(ByVal rr As Long, ByVal cc As Long, ByVal origin As String)
    MergeTop.SetVal rr & "|" & cc, origin
    If rr > MergeMaxRow Then MergeMaxRow = rr
End Sub

Public Function Cell(ByVal rr As Long, ByVal cc As Long) As Variant
    If rr < 1 Or cc < 1 Or rr > R Or cc > C Then Exit Function
    Cell = mV(rr, cc)
End Function

' Anzahl gefuellter Zellen einer Zeile (c2 = 0: bis zur letzten Spalte)
Public Function FilledInRow(ByVal rr As Long, Optional ByVal c1 As Long = 1, Optional ByVal c2 As Long = 0) As Long
    Dim cc As Long, n As Long
    If rr < 1 Or rr > R Then Exit Function
    If c1 < 1 Then c1 = 1
    If c2 < 1 Or c2 > C Then c2 = C
    For cc = c1 To c2
        If Not IsBlank(mV(rr, cc)) Then n = n + 1
    Next
    FilledInRow = n
End Function

Public Function HasFormulaCol(ByVal cc As Long) As Boolean
    If cc < 1 Or cc > C Then Exit Function
    HasFormulaCol = mFormCol(cc)
End Function

Public Function IsFormula(ByVal rr As Long, ByVal cc As Long) As Boolean
    If cc < 1 Or cc > C Then Exit Function
    If Not mFormCol(cc) Then Exit Function
    IsFormula = Formulas.Exists(rr & "|" & cc)
End Function

Public Function FormulaText(ByVal rr As Long, ByVal cc As Long) As String
    If cc < 1 Or cc > C Then Exit Function
    If Not mFormCol(cc) Then Exit Function
    FormulaText = CStr(Formulas.GetVal(rr & "|" & cc, ""))
End Function

' Oberste linke Zelle eines verbundenen Bereichs ("" wenn nicht verbunden)
Public Function MergeOrigin(ByVal rr As Long, ByVal cc As Long) As String
    If rr < 1 Or rr > MergeMaxRow Then Exit Function
    MergeOrigin = CStr(MergeTop.GetVal(rr & "|" & cc, ""))
End Function

Public Function CellFormat(ByVal rr As Long, ByVal cc As Long) As String
    Dim own As Boolean
    If cc >= 1 And cc <= C Then own = mFmtCol(cc)
    If own Then own = CellFmt.Exists(rr & "|" & cc)
    If own Then
        CellFormat = CStr(CellFmt.GetVal(rr & "|" & cc))
    Else
        CellFormat = CStr(ColFmt.GetVal(CStr(cc), ""))
    End If
End Function
