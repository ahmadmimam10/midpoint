Attribute VB_Name = "mEvents"
'==============================================================================
' mEvents - Weiterleitung der Arbeitsmappenereignisse an die Seitenmodule mit
' Wiedereintrittsschutz und Fehlergrenze.
'==============================================================================
Option Explicit

Private mBusy As Boolean

Private Function Ready() As Boolean
    Ready = mApp.gStarted And Not mBusy And mApp.WorkDepth = 0 And Not mApp.gSaving
End Function

Public Sub SheetChange(ByVal sh As Object, ByVal target As Range)
    Dim st As cViewState
    If Not Ready() Then Exit Sub
    mBusy = True
    On Error GoTo f
    Select Case sh.Name
        Case SH_TABLE
            mBusy = False
            mViewTable.OnChange target
        Case SH_CAL, SH_BOARD, SH_GANTT
            If target.Row = R_CHIPS Then
                Set st = mNav.CurState
                If Not st Is Nothing Then
                    st.Search = Trim$(NzS(sh.Cells(R_CHIPS, C_CONTENT).Value))
                    mBusy = False
                    mNav.RenderCurrent
                End If
            End If
        Case SH_START, SH_SEARCH
            If target.Row = R_TABS Then
                mNav.gGlobalSearch = Trim$(NzS(sh.Cells(R_TABS, C_CONTENT).Value))
                mBusy = False
                If Len(mNav.gGlobalSearch) > 0 Or sh.Name = SH_SEARCH Then mNav.GoPage "Suche"
            End If
    End Select
    mBusy = False
    Exit Sub
f: mBusy = False
    mApp.Boundary "Eingabe"
End Sub

Public Sub SelectionChange(ByVal sh As Object, ByVal target As Range)
    If Not Ready() Then Exit Sub
    On Error GoTo f
    mNav.CheckNavSize
    Select Case sh.Name
        Case SH_TABLE
            ' Formelleiste offen: Klick auf eine Spalte fuegt sie in die Formel ein
            If mForms.XlBarOpen() Then
                If target.Row >= R_HEAD And target.Cells.Count = 1 Then
                    If mForms.XlBarTakesClick(target) Then Exit Sub
                End If
            End If
            mViewTable.OnSelect target
        Case SH_BOARD: mViewBoard.OnSelect target
        Case SH_DASH
            ' Anordnen: nach dem Ziehen einer Karte genuegt ein Klick ins Blatt zum Einrasten
            If mDash.gArrange Then
                If mDash.SnapArrangement() Then mNav.RenderCurrent
            End If
        Case SH_SETUP: mPages.OnSetupSelect target
        Case SH_CAPTURE: mCapture.OnSelect target
    End Select
    Exit Sub
f: mApp.Boundary "Auswahl"
End Sub

Public Sub BeforeDoubleClick(ByVal sh As Object, ByVal target As Range, ByRef Cancel As Boolean)
    Dim tg As String, p As Long, f As frmTableEditor
    If Not Ready() Then Exit Sub
    On Error GoTo fl
    Select Case sh.Name
        Case SH_TABLE: mViewTable.OnDoubleClick target, Cancel
        Case SH_CAL: mViewCal.OnDoubleClick target, Cancel
        Case SH_BOARD: mViewBoard.OnDoubleClick target, Cancel
        Case SH_GANTT: mViewGantt.OnDoubleClick target, Cancel
        Case SH_START, SH_SEARCH
            tg = mPages.RowTarget(sh.Name, target.Row)
            If Len(tg) > 0 Then
                Cancel = True
                p = InStr(tg, "|")
                mForms.OpenRecord Left$(tg, p - 1), Mid$(tg, p + 1)
            End If
        Case SH_CAPTURE: mCapture.OnDoubleClick target, Cancel
        Case SH_SETUP
            tg = mPages.RowTarget(SH_SETUP, target.Row)
            If Left$(tg, 2) = "T|" Then
                Cancel = True
                mPages.OnSetupSelect target
                mCmd.Cmd_SetupEditTable
            ElseIf Left$(tg, 2) = "D|" Then
                Cancel = True
                mNav.GoDash Mid$(tg, 3)
            End If
    End Select
    Exit Sub
fl: mApp.Boundary "Doppelklick"
End Sub

Public Sub WindowResize()
    If Not mApp.gStarted Then Exit Sub
    On Error Resume Next
    mNav.PositionNav
    ' Seite an neue Fensterbreite anpassen (Umbruch der Schaltflaechen, Statusband)
    If Abs(ActiveWindow.UsableWidth - mRender.LastRenderWidth()) > 30 And mApp.WorkDepth = 0 Then mNav.RenderCurrent
End Sub

Public Sub WindowActivate()
    If Not mApp.gStarted Or mApp.gClosing Then Exit Sub
    On Error Resume Next
    mNav.ShowNav
    mNav.CheckNavSize
    mRender.RefreshStatus
End Sub

' Direkte Blattaktivierung (z. B. bei eingeblendeten Datenblaettern) fuehrt zu frischem Seitenaufbau
Public Sub SheetActivate(ByVal sh As Object)
    If Not Ready() Then Exit Sub
    On Error GoTo f
    If mProtect.IsUiSheet(sh.Name) Then
        If sh.Name <> mNav.CurrentSheetName() Then
            Select Case sh.Name
                Case SH_DASH: mNav.GoDash ""
                Case SH_SETUP: mNav.GoPage "Einrichten"
                Case SH_CAPTURE: mNav.GoPage "Erfassung"
                Case SH_HELP: mNav.GoPage "Hilfe"
                Case SH_TECH: mNav.GoPage "Technik"
                Case Else: mNav.GoStart
            End Select
        End If
    End If
    Exit Sub
f: mApp.Boundary "Seitenwechsel"
End Sub
