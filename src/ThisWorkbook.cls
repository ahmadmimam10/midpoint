'==============================================================================
' ThisWorkbook - Ereignisse der Arbeitsmappe; Weiterleitung an mStart/mEvents.
'==============================================================================
Option Explicit

Private Sub Workbook_Open()
    mStart.OnOpen
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    mStart.OnBeforeClose Cancel
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If mApp.gSaving Then Exit Sub
    Cancel = True
    mStart.OnBeforeSave SaveAsUI
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    mEvents.SheetChange Sh, Target
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    mEvents.SelectionChange Sh, Target
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    mEvents.BeforeDoubleClick Sh, Target, Cancel
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    mEvents.SheetActivate Sh
End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)
    mEvents.WindowResize
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    mEvents.WindowActivate
    mNav.StartWatch
End Sub

Private Sub Workbook_Deactivate()
    ' Zoom-/Größenprüfung nur, solange diese Datei aktiv ist
    mNav.StopWatch
End Sub
