VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLiTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cLiTab - gefundene Tabelle in einer Quelldatei ("Liste uebernehmen").
' Art: "liste" (Kopfzeile + Datenzeilen), "formular" (Bezeichnung: Wert),
'      "zusammen" (gleich aufgebaute Tabellen mehrerer Blaetter).
'==============================================================================
Option Explicit

Public Kind As String
Public SheetIdx As Long            ' Index in gSheets (bei "zusammen": erstes Blatt)
Public R1 As Long                  ' erste Zeile des Bereichs (inkl. Titel)
Public R2 As Long
Public C1 As Long
Public C2 As Long
Public HdrRow As Long              ' (erste) Kopfzeile
Public HdrRows As Long             ' 1 oder 2
Public AnnRow As Long              ' Hinweiszeile unter dem Kopf (Stichtag, Uebersetzung, Schluessel/Bezeichnung), 0 = keine
Public RowClass As cMap            ' Zeile -> "daten" | "leer" | "summe" | "gruppe" | "titel" | "notiz"
Public ColName As cMap             ' Spalte -> Ueberschrift
Public ColTop As cMap              ' Spalte -> obere Ueberschrift (zweizeiliger Kopf)
Public ColHelp As cMap             ' Spalte -> Hilfetext (vollstaendige Ueberschrift, Hinweiszeile)
Public Title As String             ' Titelzeile oberhalb (falls vorhanden)
Public Members As Collection       ' bei "zusammen": cLiTab je Blatt
Public FromListObject As Boolean
Public Note As String              ' kurze Beschreibung fuer die Auswahlliste
Public Score As Double             ' Wahrscheinlichkeit "echte Tabelle"

Private Sub Class_Initialize()
    Set RowClass = New cMap
    Set ColName = New cMap
    Set ColTop = New cMap
    Set ColHelp = New cMap
    Set Members = New Collection
    HdrRows = 1
End Sub

' erste Zeile nach Kopf (und Hinweiszeile)
Public Function FirstDataRow() As Long
    FirstDataRow = HdrRow + HdrRows + IIf(AnnRow > 0, 1, 0)
End Function

Public Function DataRowCount() As Long
    Dim k As Variant, n As Long, m As cLiTab
    If Kind = "zusammen" Then
        For Each m In Members: n = n + m.DataRowCount(): Next
        DataRowCount = n
        Exit Function
    End If
    If Kind = "formular" Then DataRowCount = 1: Exit Function
    For Each k In RowClass.Keys
        If RowClass.GetVal(CStr(k)) = "daten" Then n = n + 1
    Next
    DataRowCount = n
End Function

Public Function CountClass(ByVal cls As String) As Long
    Dim k As Variant, n As Long
    For Each k In RowClass.Keys
        If RowClass.GetVal(CStr(k)) = cls Then n = n + 1
    Next
    CountClass = n
End Function

Public Function ColCount() As Long
    ColCount = C2 - C1 + 1
End Function
