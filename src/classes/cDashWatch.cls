VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDashWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cDashWatch - Einrasten im Anordnen-Modus direkt nach dem Loslassen einer Karte.
' Excel meldet das Verschieben von Formen nicht; Office meldet aber jede
' Aktualisierung der Oberflaeche (CommandBars.OnUpdate), auch nach dem Ziehen.
' Die Pruefung ist billig (Lage der Karten vergleichen) und laeuft nur, solange
' der Anordnen-Modus aktiv ist. Kein Zeitgeber, kein Application.OnTime.
'==============================================================================
Option Explicit

Private WithEvents cb As Office.CommandBars
Private mBusy As Boolean

Public Sub StartWatch()
    Set cb = Application.CommandBars
End Sub

Public Sub StopWatch()
    Set cb = Nothing
End Sub

Private Sub cb_OnUpdate()
    Dim evOff As Boolean
    If mBusy Then Exit Sub
    If Not mDash.gArrange Then Exit Sub
    ' noch am Ziehen: erst nach dem Loslassen einrasten
    If mWin.MouseDown() Then Exit Sub
    mBusy = True
    On Error GoTo done
    If ActiveWorkbook Is ThisWorkbook Then
        If ActiveSheet.Name = SH_DASH Then
            If mDash.CardsMoved() Then
                If mDash.SnapArrangement() Then
                    mNav.RenderCurrent
                    ' die gezogene Form gibt es nicht mehr: Auswahl auf eine Zelle setzen,
                    ' sonst geht der naechste Klick (z. B. auf "Fertig") ins Leere
                    Application.EnableEvents = False: evOff = True
                    ActiveSheet.Cells(1, 1).Select
                    Application.EnableEvents = True
                End If
            End If
        End If
    End If
done:
    If evOff Then Application.EnableEvents = True
    mBusy = False
End Sub

Private Sub Class_Terminate()
    Set cb = Nothing
End Sub
