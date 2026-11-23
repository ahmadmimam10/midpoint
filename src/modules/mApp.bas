Attribute VB_Name = "mApp"
'==============================================================================
' mApp - Anwendungszustand: Start/Ende, ruhige Excel-Zustaende mit sicherer
' Wiederherstellung, zentrale Fehlergrenze, Modus, Lese-/Schreibstatus.
'==============================================================================
Option Explicit

Private mDepth As Long
Private mPrevScreen As Boolean
Private mPrevEvents As Boolean
Private mPrevCalc As Long
Private mPrevStatusBar As Variant
Private mSetupUnlocked As Boolean

Public gStarted As Boolean
Public gDataDirty As Boolean      ' Fachdaten/Konfiguration seit letzter Speicherung geaendert
Public gReadOnly As Boolean
Public gSaving As Boolean
Public gClosing As Boolean
Public gDataRev As Long           ' steigt bei jeder Datenaenderung
Public gCfgRev As Long            ' steigt bei jeder Strukturaenderung
Public gLastError As String

'--- Ruhige Ausfuehrung (verschachtelbar) --------------------------------------
Public Sub BeginWork(Optional ByVal statusText As String = "")
    If mDepth = 0 Then
        mPrevScreen = Application.ScreenUpdating
        mPrevEvents = Application.EnableEvents
        mPrevCalc = Application.Calculation
        mPrevStatusBar = Application.StatusBar
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        If Application.Calculation <> xlCalculationManual Then Application.Calculation = xlCalculationManual
    End If
    mDepth = mDepth + 1
    If Len(statusText) > 0 Then Application.StatusBar = statusText
End Sub

Public Sub EndWork()
    If mDepth <= 0 Then mDepth = 0: Exit Sub
    mDepth = mDepth - 1
    If mDepth = 0 Then RestoreState
End Sub

Private Sub RestoreState()
    On Error Resume Next
    If Application.Calculation <> mPrevCalc Then Application.Calculation = mPrevCalc
    Application.EnableEvents = mPrevEvents
    Application.ScreenUpdating = mPrevScreen
    If VarType(mPrevStatusBar) = vbString And UCase$(CStr(mPrevStatusBar)) <> "FALSE" And UCase$(CStr(mPrevStatusBar)) <> "FALSCH" Then
        Application.StatusBar = mPrevStatusBar
    Else
        Application.StatusBar = False
    End If
End Sub

' Nach Fehlern: alle offenen Ebenen schliessen und Ursprungszustand herstellen
Public Sub ResetWork()
    If mDepth > 0 Then
        mDepth = 0
        RestoreState
    End If
    If Not Application.EnableEvents Then Application.EnableEvents = True
    If Not Application.ScreenUpdating Then Application.ScreenUpdating = True
End Sub

Public Function WorkDepth() As Long
    WorkDepth = mDepth
End Function

'--- Zentrale Fehlergrenze ----------------------------------------------------
' Aufruf ausschliesslich in Fehlerbehandlern an Bediengrenzen.
Public Sub Boundary(ByVal context As String)
    Dim num As Long, desc As String, src As String
    num = Err.Number: desc = Err.Description: src = Err.Source
    On Error Resume Next
    mData.TxAbortIfOpen
    ResetWork
    mProtect.ReprotectAll
    gLastError = Format$(Now, "dd.mm.yyyy hh:nn:ss") & " | " & context & " | " & num & " | " & desc
    If num = vbObjectError + 513 Then
        ' fachliche Meldung (von RaiseUser ausgeloest)
        mUI.Alert desc, context
    Else
        mUI.Alert "Bei """ & context & """ ist ein Fehler aufgetreten. Die Aktion wurde abgebrochen; " & _
                  "bereits begonnene Datenänderungen wurden zurückgenommen." & vbLf & vbLf & _
                  "Fehler " & num & ": " & desc, "Fehler", True
    End If
    Err.Clear
End Sub

' Fachliche Abbruchmeldung (wird in Boundary verstaendlich angezeigt)
Public Sub RaiseUser(ByVal msg As String)
    Err.Raise vbObjectError + 513, APP_TITLE, msg
End Sub

'--- Modus ------------------------------------------------------------------------
' Seit 1.2.0 gibt es keinen getrennten Einrichtungsmodus mehr (alles in einer Oberflaeche
' wie in Notion). Struktur aendern darf, wer schreiben darf.
Public Function IsSetupMode() As Boolean
    IsSetupMode = Not gReadOnly
End Function

Public Function CanEditData() As Boolean
    CanEditData = Not gReadOnly
End Function

Public Function RequireWritable(Optional ByVal action As String = "Diese Aktion") As Boolean
    If gReadOnly Then
        mUI.Alert action & " ist nicht möglich: Die Datei ist schreibgeschützt geöffnet (Lesemodus). " & _
                  "Änderungen könnten nicht gespeichert werden. Öffnen Sie die Datei mit Schreibrechten, " & _
                  "sobald sie nicht mehr von einer anderen Person geöffnet ist.", "Lesemodus"
        Exit Function
    End If
    RequireWritable = True
End Function

' Strukturaenderungen: schreibbar und - falls ein Einrichtungsschutz gesetzt ist -
' einmal je Sitzung das Schutzwort
Public Function RequireSetup(Optional ByVal action As String = "Diese Strukturänderung") As Boolean
    Dim pw As String, entered As String
    If Not RequireWritable(action) Then Exit Function
    pw = mCfg.SysGet("Einrichtungsschutz", "")
    If Len(pw) > 0 And Not mSetupUnlocked Then
        entered = mUI.AskText(action & " ist durch ein Schutzwort gesichert. Bitte eingeben." & vbLf & _
                              "(Schutz gegen Fehlbedienung, keine Zugriffskontrolle.)", "Einrichtungsschutz", "", True)
        If entered <> pw Then
            If Len(entered) > 0 Then mUI.Alert "Das Schutzwort stimmt nicht.", "Einrichtungsschutz"
            Exit Function
        End If
        mSetupUnlocked = True
    End If
    RequireSetup = True
End Function

Public Function SwitchMode(ByVal target As String) As Boolean
    Dim pw As String, entered As String
    If target = "Einrichten" Then
        If Not RequireWritable("Der Einrichtungsmodus") Then Exit Function
        pw = mCfg.SysGet("Einrichtungsschutz", "")
        If Len(pw) > 0 Then
            entered = mUI.AskText("Bitte das Schutzwort für den Einrichtungsmodus eingeben." & vbLf & _
                                  "(Schutz gegen Fehlbedienung, keine Zugriffskontrolle.)", "Einrichtungsschutz", "", True)
            If entered <> pw Then
                If Len(entered) > 0 Then mUI.Alert "Das Schutzwort stimmt nicht. Der Arbeitsmodus bleibt aktiv.", "Einrichtungsschutz"
                Exit Function
            End If
        End If
    End If
    mCfg.SysSet "Modus", target
    mLog.Add "Modus", "", "", "", "Wechsel in " & target
    SwitchMode = True
    mNav.RefreshNav
    mNav.RefreshCurrent
End Function

'--- Datenstand -----------------------------------------------------------------
Public Sub MarkDirty()
    gDataDirty = True
    gDataRev = gDataRev + 1
    mBackup.ScheduleAutoSave
End Sub

Public Sub MarkCfgDirty()
    gDataDirty = True
    gCfgRev = gCfgRev + 1
    gDataRev = gDataRev + 1
    mBackup.ScheduleAutoSave
End Sub

' Renderaenderungen an Oberflaechenblaettern zaehlen nicht als Datenaenderung
Public Sub SyncSavedFlag()
    On Error Resume Next
    If Not gDataDirty And Not gReadOnly Then ThisWorkbook.Saved = True
End Sub

Public Function Wb() As Workbook
    Set Wb = ThisWorkbook
End Function

Public Function Ws(ByVal nm As String) As Worksheet
    Set Ws = ThisWorkbook.Worksheets(nm)
End Function

Public Function SheetExists(ByVal nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = ThisWorkbook.Sheets(nm)
    SheetExists = Not s Is Nothing
End Function
