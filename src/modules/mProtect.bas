Attribute VB_Name = "mProtect"
'==============================================================================
' mProtect - Blattschutz gegen Fehlbedienung (ohne Passwort; keine Zugriffs-
' kontrolle). Oberflaechenblaetter: UserInterfaceOnly, damit Code zeichnen kann.
' Daten-/Konfigurationsblaetter sind nur bei ausdruecklichem Datenzugang
' sichtbar und dann schreibgeschuetzt.
'==============================================================================
Option Explicit

Private mPending As Collection

Public Function IsUiSheet(ByVal nm As String) As Boolean
    Select Case nm
        Case SH_START, SH_TABLE, SH_CAL, SH_BOARD, SH_GANTT, SH_DASH, SH_SETUP, SH_HELP, SH_TECH, SH_SEARCH, SH_CAPTURE
            IsUiSheet = True
    End Select
End Function

Public Sub ProtectUi(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFormattingColumns:=False, AllowFormattingRows:=False
    ws.EnableSelection = xlNoRestrictions
End Sub

' Vor Schreibzugriff auf ein Daten-/Konfigurationsblatt
Public Sub Unprot(ByVal ws As Worksheet)
    If ws.ProtectContents Then
        ws.Unprotect Password:=""
        If mPending Is Nothing Then Set mPending = New Collection
        On Error Resume Next
        mPending.Add ws.Name, ws.Name
        On Error GoTo 0
    End If
End Sub

Public Sub ReprotectAll()
    Dim v As Variant
    On Error Resume Next
    If mPending Is Nothing Then Exit Sub
    For Each v In mPending
        ThisWorkbook.Worksheets(CStr(v)).Protect Password:="", DrawingObjects:=True, Contents:=True
    Next
    Set mPending = Nothing
End Sub

Public Sub ProtectAllUi()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsUiSheet(ws.Name) Then ProtectUi ws
    Next
End Sub

Public Function IsInternalSheet(ByVal nm As String) As Boolean
    IsInternalSheet = (Left$(nm, 1) = "_" Or Left$(nm, Len(DATA_PREFIX)) = DATA_PREFIX)
End Function

' Datenzugang: interne Blaetter sichtbar (schreibgeschuetzt) oder wieder ausblenden
Public Sub ShowDataSheets(ByVal show As Boolean)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsInternalSheet(ws.Name) And ws.Name <> SH_CODE And ws.Name <> SH_CALC And ws.Name <> SH_RES Then
            If show Then
                ws.Visible = xlSheetVisible
                ws.Protect Password:="", DrawingObjects:=True, Contents:=True
            Else
                If ws.ProtectContents Then ws.Unprotect Password:=""
                ws.Visible = xlSheetVeryHidden
            End If
        End If
    Next
End Sub
