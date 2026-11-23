VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cViewState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cViewState - Sitzungszustand einer gespeicherten Ansicht (Suche, Filter,
' Sortierung, Gruppierung, Seite, Archivmodus, Zeitraum). Wird beim Oeffnen der
' Datei neu aus der gespeicherten Ansicht abgeleitet; Aenderungen veraendern die
' gespeicherte Ansicht erst bei "Ansicht speichern".
'==============================================================================
Option Explicit

Public ViewID As String
Public Search As String
Public Conds As String          ' serialisierte Bedingungen
Public SortSpec As String
Public GroupFid As String
Public GroupLevel As String
Public Page As Long
Public ArchMode As Long
Public CalMonth As Date
Public CalMode As String         ' Monat | Liste
Public GanttStart As Date
Public GanttScale As String      ' Tag | Woche | Monat
Public BoardPage As Long
Public SelRid As String

Public Sub InitFrom(ByVal vid As String)
    ViewID = vid
    Search = ""
    Conds = mCfg.ViewGet(vid, "Filter")
    SortSpec = mCfg.ViewGet(vid, "Sortierung")
    GroupFid = KVGet(mCfg.ViewGet(vid, "Gruppierung"), "f")
    GroupLevel = KVGet(mCfg.ViewGet(vid, "Gruppierung"), "stufe", "Monat")
    Page = 1
    ArchMode = 0
    CalMonth = DateSerial(Year(Date), Month(Date), 1)
    CalMode = "Monat"
    GanttScale = mCfg.ViewOpt(vid, "massstab", "Woche")
    GanttStart = DefaultGanttStart(GanttScale)
    BoardPage = 1
End Sub

Public Function DefaultGanttStart(ByVal scale_ As String) As Date
    Select Case scale_
        Case "Tag": DefaultGanttStart = Date - Weekday(Date, vbMonday) + 1 - 7
        Case "Monat": DefaultGanttStart = DateSerial(Year(Date), Month(Date) - 2, 1)
        Case Else: DefaultGanttStart = Date - Weekday(Date, vbMonday) + 1 - 14
    End Select
End Function

' Weicht der Sitzungszustand von der gespeicherten Ansicht ab?
Public Function IsModified() As Boolean
    If Len(Search) > 0 Then IsModified = True: Exit Function
    If ArchMode <> 0 Then IsModified = True: Exit Function
    If Conds <> mCfg.ViewGet(ViewID, "Filter") Then IsModified = True: Exit Function
    If SortSpec <> mCfg.ViewGet(ViewID, "Sortierung") Then IsModified = True: Exit Function
    If GroupFid <> KVGet(mCfg.ViewGet(ViewID, "Gruppierung"), "f") Then IsModified = True: Exit Function
End Function

' Unterscheidet sich die speicherbare Einstellung (ohne Suche) von der Ansicht?
Public Function HasUnsavedSettings() As Boolean
    If Conds <> mCfg.ViewGet(ViewID, "Filter") Then HasUnsavedSettings = True: Exit Function
    If SortSpec <> mCfg.ViewGet(ViewID, "Sortierung") Then HasUnsavedSettings = True: Exit Function
    If GroupFid <> KVGet(mCfg.ViewGet(ViewID, "Gruppierung"), "f") Then HasUnsavedSettings = True: Exit Function
End Function

' Alles, was die Darstellung bestimmt (ohne Auswahl) - fuer die Wiederverwendung einer Seite
Public Function Signature() As String
    Signature = Search & Chr$(1) & Conds & Chr$(1) & SortSpec & Chr$(1) & GroupFid & Chr$(1) & GroupLevel & Chr$(1) & Page & Chr$(1) & _
                ArchMode & Chr$(1) & CDbl(CalMonth) & Chr$(1) & CalMode & Chr$(1) & CDbl(GanttStart) & Chr$(1) & GanttScale & Chr$(1) & BoardPage
End Function

Public Function GroupSpec() As String
    If Len(GroupFid) = 0 Then Exit Function
    GroupSpec = "f=" & GroupFid & ";stufe=" & GroupLevel
End Function
