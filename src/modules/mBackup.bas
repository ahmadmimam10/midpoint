Attribute VB_Name = "mBackup"
'==============================================================================
' mBackup - Kontrolliertes Speichern (Statusband erst nach Erfolg), Sicherungen
' im Unterordner neben der Datei (Zeitstempel, Version, Aufbewahrung), Wieder-
' herstellung aus der Sicherungsliste, neue leere Kopie, Gesamtreset.
' Keine eigene Verschluesselung; Datei-/Zugriffsschutz ueber Office/Windows.
'==============================================================================
Option Explicit

Private mRollbackOk As Boolean

Private mLastBackupFail As String
Private mAutoPending As Boolean
Private mAutoAt As Date
Private mSaveSheet As String
Private mSaveSel As String
Private mSaveScroll As Long

'--- Status -------------------------------------------------------------------------
Public Function StatusText() As String
    Dim s As String, lb As String
    If mApp.gReadOnly Then
        s = "Lesemodus (schreibgeschützt)"
    ElseIf AutoSaveOn() Then
        s = "Auto-Speichern ein"
    Else
        s = "Bearbeiten"
    End If
    If Len(mCfg.SysGet("LetzteSpeicherung", "")) > 0 And Len(ThisWorkbook.path) > 0 Then
        s = s & "  ·  Gespeichert " & FmtStamp(InvDate(mCfg.SysGet("LetzteSpeicherung")))
    Else
        s = s & "  ·  Noch nicht gespeichert"
    End If
    If Len(mLastBackupFail) > 0 Then
        s = s & "  ·  Letzte Sicherung FEHLGESCHLAGEN"
    Else
        lb = LastBackupStamp()
        If Len(lb) > 0 Then s = s & "  ·  Sicherung " & lb Else s = s & "  ·  Noch keine Sicherung"
    End If
    StatusText = s
End Function

Public Function BackupFolder() As String
    Dim baseNo As String
    If Len(ThisWorkbook.path) = 0 Then Exit Function
    baseNo = Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    BackupFolder = ThisWorkbook.path & "\" & baseNo & "_Sicherungen"
End Function

Private Function BaseName() As String
    BaseName = Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
End Function

' Liste eigener Sicherungen: Array von Array(Dateiname, Zeitpunkt, Groesse) neueste zuerst
Public Function ListBackups() As Collection
    Dim res As New Collection, folder As String, f As String, items As New Collection, d As Date, names() As String, n As Long, i As Long, j As Long
    Dim stamps() As Double, tmpS As String, tmpD As Double
    folder = BackupFolder()
    Set ListBackups = res
    If Not DirExists(folder) Then Exit Function
    f = Dir$(folder & "\" & BaseName() & "_????-??-??_??????*.xlsm")
    Do While Len(f) > 0
        n = n + 1
        ReDim Preserve names(1 To n): ReDim Preserve stamps(1 To n)
        names(n) = f
        stamps(n) = CDbl(FileDateTime(folder & "\" & f))
        f = Dir$
    Loop
    For i = 1 To n - 1
        For j = i + 1 To n
            If names(j) > names(i) Then
                tmpS = names(i): names(i) = names(j): names(j) = tmpS
                tmpD = stamps(i): stamps(i) = stamps(j): stamps(j) = tmpD
            End If
        Next
    Next
    For i = 1 To n
        res.Add Array(names(i), CDate(stamps(i)), FileLen(folder & "\" & names(i)))
    Next
End Function

Public Function LastBackupStamp() As String
    Dim l As Collection
    On Error Resume Next
    Set l = ListBackups()
    If l.Count > 0 Then LastBackupStamp = FmtStamp(l(1)(1))
End Function

'--- Automatisch speichern ---------------------------------------------------------------
' Ein: nach jeder Daten- oder Strukturaenderung wird die Datei gespeichert (kurz verzoegert,
' damit mehrere Schritte einer Aktion zusammen gespeichert werden). Aus: wie bisher.
Public Function AutoSaveOn() As Boolean
    AutoSaveOn = (mCfg.SysGet("AutoSpeichern", "Nein") = "Ja")
End Function

Public Sub ToggleAutoSave()
    If Not mApp.RequireWritable("Automatisch speichern") Then Exit Sub
    If AutoSaveOn() Then
        mCfg.SysSet "AutoSpeichern", "Nein"
        mLog.Add "Einstellungen", "", "", "", "Automatisch speichern aus"
    Else
        If Len(ThisWorkbook.path) = 0 Then
            mUI.Alert "Bitte die Datei zuerst einmal speichern. Danach kann jede Änderung automatisch gespeichert werden.", "Automatisch speichern"
            Exit Sub
        End If
        mCfg.SysSet "AutoSpeichern", "Ja"
        mLog.Add "Einstellungen", "", "", "", "Automatisch speichern ein"
    End If
    mApp.gDataDirty = True
    mNav.RefreshNav
    If AutoSaveOn() Then ControlledSave False Else mRender.RefreshStatus
End Sub

Public Sub ScheduleAutoSave()
    If mAutoPending Or mApp.gSaving Or mApp.gReadOnly Then Exit Sub
    If Not mApp.gStarted Then Exit Sub
    If Not AutoSaveOn() Then Exit Sub
    If Len(ThisWorkbook.path) = 0 Then Exit Sub
    mAutoPending = True
    On Error Resume Next
    mAutoAt = Now + TimeSerial(0, 0, 1)
    Application.OnTime mAutoAt, mRender.MacroRef("mBackup.AutoSaveTick")
    If Err.Number <> 0 Then mAutoPending = False
End Sub

Public Sub CancelAutoSave()
    On Error Resume Next
    If mAutoPending Then Application.OnTime mAutoAt, mRender.MacroRef("mBackup.AutoSaveTick"), , False
    mAutoPending = False
End Sub

Public Sub AutoSaveTick()
    Dim uf As Object
    mAutoPending = False
    On Error GoTo f
    If Not mApp.gDataDirty Or mApp.gSaving Or Not AutoSaveOn() Then Exit Sub
    ' waehrend ein Dialog offen ist oder gearbeitet wird: spaeter erneut
    For Each uf In VBA.UserForms
        If TypeName(uf) <> "frmNav" Then
            If uf.Visible Then ScheduleAutoSave: Exit Sub
        End If
    Next
    If mApp.WorkDepth > 0 Or mData.TxDepth > 0 Then ScheduleAutoSave: Exit Sub
    ControlledSave False
    Exit Sub
f:  mApp.Boundary "Automatisch speichern"
End Sub

'--- Speichern -----------------------------------------------------------------------
' Wird aus Workbook_BeforeSave (Cancel=True) aufgerufen
Public Function ControlledSave(ByVal saveAsUI As Boolean) As Boolean
    Dim prevStamp As String, target As Variant, prevEv As Boolean, ok As Boolean, msg As String, wasNew As Boolean
    If mApp.gReadOnly Then
        mUI.Alert "Die Datei ist schreibgeschützt geöffnet (Lesemodus) und kann nicht gespeichert werden. " & _
                  "Nutzen Sie bei Bedarf „Datei > Speichern unter“ für eine eigene Kopie.", "Speichern", True
        Exit Function
    End If
    If mApp.gSaving Then Exit Function
    mApp.gSaving = True
    prevEv = Application.EnableEvents
    wasNew = (Len(ThisWorkbook.path) = 0)
    If saveAsUI Or wasNew Then
        target = Application.GetSaveAsFilename(InitialFileName:=ThisWorkbook.FullName, FileFilter:="Excel-Arbeitsmappe mit Makros (*.xlsm),*.xlsm", Title:="Speichern unter")
        If VarType(target) <> vbString Then mApp.gSaving = False: Exit Function
    End If
    ' automatische Sicherung bei regulaerem Speichern (vor dem Speichern = gleicher Datenstand)
    If Not wasNew And Not saveAsUI Then MaybeAutoBackupOnSave
    prevStamp = mCfg.SysGet("LetzteSpeicherung", "")
    mCfg.SysSet "LetzteSpeicherung", DateInv(Now)
    mLog.Flush
    PrepareForSave
    Application.EnableEvents = False
    On Error GoTo failed
    If saveAsUI Or wasNew Then
        Application.DisplayAlerts = False
        ThisWorkbook.SaveAs Filename:=CStr(target), FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.DisplayAlerts = True
    Else
        ThisWorkbook.Save
    End If
    ok = True
cleanup:
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.EnableEvents = prevEv
    RestoreAfterSave
    If ok Then
        mApp.gDataDirty = False
        ThisWorkbook.Saved = True
    Else
        mCfg.SysSet "LetzteSpeicherung", prevStamp
        mUI.Alert "Die Datei konnte NICHT gespeichert werden. Ihre Änderungen sind weiterhin nur in Excel geöffnet." & vbLf & vbLf & _
                  "Ursache: " & msg & vbLf & vbLf & "Mögliche Gründe: fehlende Schreibrechte, Netzlaufwerk nicht erreichbar, Datei von anderer Person geöffnet. " & _
                  "Bitte das Problem beheben und erneut speichern oder „Datei > Speichern unter“ an einem anderen Ort verwenden.", "Speichern fehlgeschlagen", True
    End If
    mApp.gSaving = False
    mRender.RefreshStatus
    ControlledSave = ok
    Exit Function
failed:
    msg = Err.Description
    Resume cleanup
End Function

' Hinweisblatt sichtbar machen (fuer Oeffnen ohne Makros)
Private Sub PrepareForSave()
    Dim ws As Worksheet
    On Error Resume Next
    Application.ScreenUpdating = False
    ' aktuelle Seite merken: nach dem Speichern ohne Neuaufbau wieder zeigen
    mSaveSheet = "": mSaveSel = "": mSaveScroll = 1
    If ActiveWorkbook Is ThisWorkbook Then
        mSaveSheet = ActiveSheet.Name
        mSaveSel = Selection.Address
        mSaveScroll = ActiveWindow.ScrollRow
    End If
    With ThisWorkbook.Worksheets(SH_HINWEIS)
        .Visible = xlSheetVisible
        .Activate
    End With
    ActiveWindow.DisplayWorkbookTabs = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SH_HINWEIS And ws.Visible = xlSheetVisible Then ws.Visible = xlSheetHidden
    Next
End Sub

Private Sub RestoreAfterSave()
    Dim ws As Worksheet
    On Error Resume Next
    If Len(mSaveSheet) > 0 And mProtect.IsUiSheet(mSaveSheet) Then
        Set ws = ThisWorkbook.Worksheets(mSaveSheet)
        mApp.BeginWork
        mRender.ShowSheet ws
        ActiveWindow.ScrollRow = mSaveScroll
        If Len(mSaveSel) > 0 Then ws.Range(mSaveSel).Select
        mApp.EndWork
        mRender.RefreshStatus
    Else
        mNav.RenderCurrent
    End If
    ThisWorkbook.Worksheets(SH_HINWEIS).Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
    ThisWorkbook.Saved = True
End Sub

'--- Sicherungen ---------------------------------------------------------------------
' Erstellt eine Sicherung; liefert True nur nach geprueftem Erfolg
Public Function BackupNow(ByVal reason As String, Optional ByVal silent As Boolean = False) As Boolean
    Dim folder As String, path As String, b() As Byte, ok As Boolean, msg As String
    mLastBackupFail = ""
    If Len(ThisWorkbook.path) = 0 Then
        mLastBackupFail = "Datei noch nicht gespeichert"
        If Not silent Then mUI.Alert "Sicherungen werden neben der Datei abgelegt. Bitte die Datei zuerst einmal speichern.", "Sicherung", True
        Exit Function
    End If
    folder = BackupFolder()
    If Not EnsureDir(folder) Then
        mLastBackupFail = "Ordner nicht anlegbar"
        msg = "Der Sicherungsordner konnte nicht angelegt werden:" & vbLf & folder
        GoTo fail
    End If
    path = folder & "\" & BaseName() & "_" & StampForFile() & "_v" & PROG_VERSION & ".xlsm"
    ' zwei Sicherungen in derselben Sekunde duerfen sich nicht ueberschreiben
    Dim sfx As Long
    Do While FileExists(path)
        sfx = sfx + 1
        path = folder & "\" & BaseName() & "_" & StampForFile() & Chr$(96 + sfx) & "_v" & PROG_VERSION & ".xlsm"
    Loop
    On Error GoTo writeFail
    mLog.Flush
    ThisWorkbook.SaveCopyAs path
    On Error GoTo 0
    ' Pruefung: Datei vorhanden, nicht leer, ZIP-Signatur (PK)
    If FileExists(path) Then
        If FileLen(path) > 1000 Then
            b = ReadHead(path)
            ok = (b(0) = &H50 And b(1) = &H4B)
        End If
    End If
    If Not ok Then
        mLastBackupFail = "Prüfung fehlgeschlagen"
        msg = "Die Sicherungsdatei wurde nicht vollständig geschrieben:" & vbLf & path
        GoTo fail
    End If
    mCfg.SysSet "LetzteSicherung", DateInv(Now)
    mLog.Add "Sicherung", "", "", "", Mid$(path, InStrRev(path, "\") + 1) & " (" & reason & ")"
    Retention
    BackupNow = True
    If Not silent Then mUI.Alert "Sicherung erstellt und geprüft:" & vbLf & path, "Sicherung"
    mRender.RefreshStatus
    Exit Function
writeFail:
    msg = "Die Sicherung konnte nicht geschrieben werden (" & Err.Description & ")." & vbLf & "Ziel: " & path
    mLastBackupFail = Err.Description
fail:
    mRender.RefreshStatus
    mUI.Alert msg & vbLf & vbLf & "Es wurde KEINE gültige Sicherung erstellt.", "Sicherung fehlgeschlagen", True
End Function

Private Function ReadHead(ByVal p As String) As Byte()
    Dim f As Integer, b(0 To 3) As Byte
    f = FreeFile
    Open p For Binary Access Read Shared As #f
    Get #f, 1, b
    Close #f
    ReadHead = b
End Function

' Behaelt die konfigurierte Anzahl eigener Sicherungen (Standard 10)
Private Sub Retention()
    Dim l As Collection, keep As Long, i As Long, folder As String
    keep = mCfg.SysLong("SicherungenBehalten", 10)
    If keep < 1 Then keep = 10
    Set l = ListBackups()
    folder = BackupFolder()
    On Error Resume Next
    For i = keep + 1 To l.Count
        Kill folder & "\" & l(i)(0)
    Next
End Sub

Private Sub MaybeAutoBackupOnSave()
    Dim mins As Long, last As Variant
    If mCfg.SysGet("SicherungBeimSpeichern", "Ja") <> "Ja" Then Exit Sub
    mins = mCfg.SysLong("SicherungAbstandMinuten", 60)
    last = InvDate(mCfg.SysGet("LetzteSicherung", ""))
    If Not IsEmpty(last) Then
        If (Now - CDate(last)) * 1440 < mins Then Exit Sub
    End If
    BackupNow "automatisch beim Speichern", True
End Sub

' Automatische Sicherung vor riskanten Strukturaenderungen; False = Vorgang abbrechen
Public Function AutoBackup(ByVal reason As String) As Boolean
    If Len(ThisWorkbook.path) = 0 Then
        AutoBackup = mUI.Confirm("Vor dieser Änderung wird normalerweise automatisch gesichert. Die Datei wurde aber noch nie gespeichert, " & _
                                 "daher ist keine Sicherung möglich. Trotzdem fortfahren?", "Sicherung", "Fortfahren", "Abbrechen")
        Exit Function
    End If
    If BackupNow(reason, True) Then
        AutoBackup = True
    Else
        AutoBackup = mUI.Confirm("Die automatische Sicherung " & reason & " ist fehlgeschlagen (" & mLastBackupFail & "). " & _
                                 "Ohne Sicherung fortfahren?", "Sicherung fehlgeschlagen", "Ohne Sicherung fortfahren", "Abbrechen", True)
    End If
End Function

' Bietet eine Sicherung an; False = Vorgang abbrechen
Public Function OfferBackup(ByVal reason As String) As Boolean
    Dim ans As Long
    If Len(ThisWorkbook.path) = 0 Then OfferBackup = True: Exit Function
    ans = mUI.Ask3("Eine Sicherung " & reason & " wird empfohlen. So lässt sich der vorherige Stand jederzeit wiederherstellen.", "Sicherung", "Sichern und fortfahren", "Ohne Sicherung", "Abbrechen")
    Select Case ans
        Case 1: OfferBackup = BackupNow(reason, True)
            If Not OfferBackup Then OfferBackup = mUI.Confirm("Die Sicherung ist fehlgeschlagen. Trotzdem fortfahren?", "Sicherung", "Fortfahren", "Abbrechen", True)
        Case 2: OfferBackup = True
        Case Else: OfferBackup = False
    End Select
End Function

'--- Wiederherstellung --------------------------------------------------------------------
Public Sub RestoreFlow()
    Dim l As Collection, items() As Variant, i As Long, sel As Long, path As String, v As Variant
    Set l = ListBackups()
    If l.Count = 0 Then mUI.Alert "Im Sicherungsordner wurden keine Sicherungen dieser Datei gefunden:" & vbLf & BackupFolder(), "Wiederherstellen": Exit Sub
    ReDim items(0 To l.Count - 1, 0 To 2)
    For Each v In l
        items(i, 0) = FmtDateTime(v(1))
        items(i, 1) = v(0)
        items(i, 2) = FmtNum(v(2) / 1024, 0) & " KB"
        i = i + 1
    Next
    Dim f As frmMsg
    Set f = New frmMsg
    f.SetupList "Sicherung wiederherstellen", "Welche Sicherung soll wiederhergestellt werden? Der aktuelle Stand wird vorher gesichert. " & _
                "Übernommen werden Tabellen, Felder, Optionen, Daten, Verknüpfungen, Ansichten, Dashboards, Einstellungen und Protokoll.", items, False, "Wiederherstellen", 3, "120;300;80", 0
    f.Show: mNav.FocusGrid
    sel = -1
    If f.Result = 1 Then sel = f.SelectedIndex
    Unload f
    If sel < 0 Then Exit Sub
    path = BackupFolder() & "\" & l(sel + 1)(0)
    If Not mUI.Confirm("Den Datenstand aus" & vbLf & l(sel + 1)(0) & vbLf & "wiederherstellen? Der aktuelle Stand wird zuvor gesichert. Nicht gesicherte Änderungen seit der Sicherung gehen im Arbeitsstand verloren.", _
                       "Wiederherstellen", "Wiederherstellen", "Abbrechen", True) Then Exit Sub
    RestoreFrom path
End Sub

Public Sub RestoreFrom(ByVal path As String, Optional ByVal isRollback As Boolean = False)
    Dim safety As String, l As Collection
    Dim src As Workbook, prevSec As MsoAutomationSecurity, prevEv As Boolean, sh As Variant, ws As Worksheet, fmtV As Long, ok As Boolean
    Dim srcSys As Worksheet, r As Long, msg As String, dataSheets As New Collection, s As Worksheet, cnt1 As Long, cnt2 As Long
    If Not mApp.RequireWritable("Die Wiederherstellung") Then Exit Sub
    mRollbackOk = False
    If isRollback Then
        ' Ruecksicherung nach fehlgeschlagener Wiederherstellung: keine weitere Sicherung
    ElseIf BackupNow("vor der Wiederherstellung", True) Then
        Set l = ListBackups()
        If l.Count > 0 Then safety = BackupFolder() & "\" & l(1)(0)
    Else
        If Not mUI.Confirm("Der aktuelle Stand konnte nicht gesichert werden. Trotzdem wiederherstellen?", "Wiederherstellen", "Trotzdem", "Abbrechen", True) Then Exit Sub
    End If
    prevSec = Application.AutomationSecurity
    prevEv = Application.EnableEvents
    mApp.BeginWork "Wiederherstellung läuft …"
    On Error GoTo failed
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set src = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, Notify:=False)
    Application.AutomationSecurity = prevSec
    src.Windows(1).Visible = False
    ' Formatversion pruefen
    fmtV = CLng(Val(SysValueIn(src, "Formatversion")))
    If fmtV > FORMAT_VERSION Then RaiseUser "Die Sicherung stammt aus einer neueren Programmversion (Format " & fmtV & ") und kann hier nicht wiederhergestellt werden."
    ' Konfiguration und Daten uebernehmen
    For Each sh In Array(SH_SYS, SH_TAB, SH_FLD, SH_OPT, SH_REL, SH_VIEW, SH_DSB, SH_ELM, SH_LOG)
        CopySheetData src.Worksheets(CStr(sh)), ThisWorkbook.Worksheets(CStr(sh))
    Next
    ' Erfassungsdateien (ab Programmversion 1.1; aeltere Sicherungen enthalten sie nicht)
    For Each sh In Array(SH_CAP, SH_RET)
        If SheetIn(src, CStr(sh)) And mApp.SheetExists(CStr(sh)) Then CopySheetData src.Worksheets(CStr(sh)), ThisWorkbook.Worksheets(CStr(sh))
    Next
    ' Datenblaetter: vorhandene entfernen, aus Sicherung kopieren
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, Len(DATA_PREFIX)) = DATA_PREFIX Then dataSheets.Add ws.Name
    Next
    For Each sh In dataSheets
        ThisWorkbook.Worksheets(CStr(sh)).Visible = xlSheetVisible
        ThisWorkbook.Worksheets(CStr(sh)).Delete
    Next
    For Each s In src.Worksheets
        If Left$(s.Name, Len(DATA_PREFIX)) = DATA_PREFIX Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = s.Name
            CopyDataSheet s, ws
            ws.Visible = xlSheetVeryHidden
        End If
    Next
    Application.DisplayAlerts = True
    src.Close SaveChanges:=False
    Set src = Nothing
    ' Pruefen
    mCfg.InvalidateAll
    mData.InvalidateAllData
    mData.ClearUndo
    ' aeltere Sicherungen: fehlende Spalten/Blaetter ergaenzen
    mCapture.EnsureSheets
    msg = mStart.IntegrityReport()
    mCfg.SysSet "Formatversion", CStr(FORMAT_VERSION)
    mCfg.SysSet "Programmversion", PROG_VERSION
    mLog.Add "Wiederherstellung", "", "", "", Mid$(path, InStrRev(path, "\") + 1)
    mStruct.RebuildLists
    mCalc.RecalcAll
    mApp.MarkCfgDirty
    mApp.EndWork
    mNav.ResetAllStates
    mNav.GoStart
    If isRollback Then mRollbackOk = True: Exit Sub
    mUI.Alert "Wiederherstellung abgeschlossen. Bitte die Datei speichern, um den wiederhergestellten Stand dauerhaft zu übernehmen." & vbLf & vbLf & msg, "Wiederherstellen"
    Exit Sub
failed:
    msg = Err.Description
    Resume restoreFail
restoreFail:
    On Error Resume Next
    Application.AutomationSecurity = prevSec
    Application.DisplayAlerts = True
    If Not src Is Nothing Then src.Close SaveChanges:=False
    mApp.ResetWork
    mCfg.InvalidateAll
    mData.InvalidateAllData
    If isRollback Then Exit Sub
    ' Teilweise uebernommene Daten zuruecknehmen: Sicherung von unmittelbar vor dem Versuch einspielen
    If Len(safety) > 0 Then
        On Error Resume Next
        RestoreFrom safety, True
        On Error GoTo 0
    End If
    If mRollbackOk Then
        mUI.Alert "Die Wiederherstellung ist fehlgeschlagen: " & msg & vbLf & vbLf & "Der Stand von unmittelbar vor dem Versuch wurde automatisch wieder eingespielt; es ist nichts verloren gegangen.", "Wiederherstellen", True
    Else
        mUI.Alert "Die Wiederherstellung ist fehlgeschlagen: " & msg & vbLf & vbLf & "Bitte die Datei NICHT speichern, sondern ohne Speichern schließen und neu öffnen. Der vorherige Stand liegt zusätzlich als Sicherung im Sicherungsordner.", "Wiederherstellen", True
    End If
End Sub

' Zahlenformat des Datenbereichs einer Spalte (ganze Spalte liefert Null bei gemischten Formaten)
Private Function DataFormat(ByVal ws As Worksheet, ByVal col As Long, ByVal firstDataRow As Long) As String
    Dim v As Variant
    v = ws.Columns(col).NumberFormat
    If IsNull(v) Then v = ws.Cells(firstDataRow, col).NumberFormat
    If IsNull(v) Then v = "General"
    DataFormat = CStr(v)
End Function

Private Function SheetIn(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = nm Then SheetIn = True: Exit Function
    Next
End Function

Private Function SysValueIn(ByVal wb As Workbook, ByVal key As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = wb.Worksheets(SH_SYS)
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 1).Value) = key Then SysValueIn = CStr(ws.Cells(r, 2).Value): Exit Function
    Next
End Function

Private Sub CopySheetData(ByVal src As Worksheet, ByVal dst As Worksheet)
    Dim loS As ListObject, loD As ListObject, n As Long, v As Variant
    Dim c As Long
    Set loS = src.ListObjects(1)
    Set loD = dst.ListObjects(1)
    mProtect.Unprot dst
    If Not loD.DataBodyRange Is Nothing Then loD.DataBodyRange.Rows.Delete
    For c = 1 To loS.ListColumns.Count
        dst.Columns(loD.Range.Column + c - 1).NumberFormat = DataFormat(src, loS.Range.Column + c - 1, loS.Range.Row + 1)
    Next
    If loS.DataBodyRange Is Nothing Then Exit Sub
    v = loS.DataBodyRange.Value
    If Not IsArray(v) Then
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = v
        v = one
    End If
    mCfg.AppendRowsRaw loD, v
End Sub

Private Sub CopyDataSheet(ByVal src As Worksheet, ByVal dst As Worksheet)
    Dim loS As ListObject, loD As ListObject, rng As Range, c As Long
    Set loS = src.ListObjects(1)
    For c = 1 To loS.ListColumns.Count
        dst.Columns(c).NumberFormat = DataFormat(src, c, loS.Range.Row + 1)
        dst.Cells(1, c).NumberFormat = "@"
    Next
    Set rng = dst.Range(dst.Cells(1, 1), dst.Cells(loS.Range.Rows.Count, loS.ListColumns.Count))
    rng.Value = CellTexts(loS.Range.Value)
    Set loD = dst.ListObjects.Add(xlSrcRange, rng, , xlYes)
    loD.Name = loS.Name
    loD.TableStyle = "TableStyleLight1"
End Sub

'--- Neue leere Kopie ---------------------------------------------------------------------
' Erstellt eine Kopie mit Struktur, Ansichten und Dashboards, aber ohne Datensaetze.
' Die geoeffnete Originaldatei bleibt unveraendert.
Public Sub NewEmptyCopy()
    Dim target As Variant, prevSec As MsoAutomationSecurity, wb As Workbook, ws As Worksheet, lo_ As ListObject, msg As String
    target = Application.GetSaveAsFilename(InitialFileName:=ThisWorkbook.path & "\" & BaseName() & "_leer.xlsm", _
                                           FileFilter:="Excel-Arbeitsmappe mit Makros (*.xlsm),*.xlsm", Title:="Neue leere Kopie speichern unter")
    If VarType(target) <> vbString Then Exit Sub
    If StrComp(CStr(target), ThisWorkbook.FullName, vbTextCompare) = 0 Then mUI.Alert "Bitte einen anderen Dateinamen als die geöffnete Datei wählen.", "Leere Kopie", True: Exit Sub
    If FileExists(CStr(target)) Then
        If Not mUI.Confirm("Die Datei existiert bereits. Überschreiben?", "Leere Kopie", "Überschreiben", "Abbrechen", True) Then Exit Sub
        Kill CStr(target)
    End If
    mApp.BeginWork "Leere Kopie wird erstellt …"
    On Error GoTo failed
    mLog.Flush
    ThisWorkbook.SaveCopyAs CStr(target)
    prevSec = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set wb = Application.Workbooks.Open(Filename:=CStr(target), UpdateLinks:=0, AddToMru:=False)
    Application.AutomationSecurity = prevSec
    wb.Windows(1).Visible = False
    For Each ws In wb.Worksheets
        If Left$(ws.Name, Len(DATA_PREFIX)) = DATA_PREFIX Or ws.Name = SH_REL Or ws.Name = SH_LOG Or ws.Name = SH_IMP Then
            If ws.ProtectContents Then ws.Unprotect ""
            If ws.ListObjects.Count > 0 Then
                Set lo_ = ws.ListObjects(1)
                If Not lo_.DataBodyRange Is Nothing Then lo_.DataBodyRange.Rows.Delete
            Else
                ws.Cells.Clear
            End If
        End If
    Next
    SetSysIn wb, "LetzteSpeicherung", ""
    SetSysIn wb, "LetzteSicherung", ""
    wb.Windows(1).Visible = True
    wb.Save
    wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    mApp.EndWork
    mLog.Add "Leere Kopie", "", "", "", Mid$(CStr(target), InStrRev(CStr(target), "\") + 1)
    mUI.Alert "Die leere Kopie wurde erstellt (Struktur, Ansichten, Dashboards und Einstellungen ohne Einträge):" & vbLf & target & vbLf & vbLf & _
              "Die geöffnete Datei wurde nicht verändert.", "Leere Kopie"
    Exit Sub
failed:
    msg = Err.Description
    Resume copyFail
copyFail:
    On Error Resume Next
    Application.AutomationSecurity = prevSec
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    mApp.ResetWork
    ThisWorkbook.Activate
    mUI.Alert "Die leere Kopie konnte nicht erstellt werden: " & msg, "Leere Kopie", True
End Sub

Private Sub SetSysIn(ByVal wb As Workbook, ByVal key As String, ByVal value As String)
    Dim ws As Worksheet, r As Long
    Set ws = wb.Worksheets(SH_SYS)
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 1).Value) = key Then ws.Cells(r, 2).Value = value
    Next
End Sub

'--- Gesamtreset --------------------------------------------------------------------------
' Leert alle Eintraege und Zuordnungen; Struktur, Ansichten, Dashboards, Einstellungen und
' Protokoll bleiben. Nur nach erfolgreicher Sicherung und Bestaetigungswort.
Public Sub ResetData()
    Dim word As String, ok As Boolean, ids As Variant, i As Long, lo_ As ListObject, n As Long
    ids = mCfg.TableIds(True)
    For i = 0 To ArrLen(ids) - 1
        n = n + mData.TData(CStr(ids(i))).RowCount
    Next
    If Not mUI.Confirm("Datei zurücksetzen löscht ALLE " & FmtInt(n) & " Einträge und alle Verknüpfungen in allen Tabellen." & vbLf & vbLf & _
                       "Erhalten bleiben: Tabellen, Felder, Auswahloptionen, Berechnungen, Ansichten, Dashboards, Einstellungen und das Änderungsprotokoll." & vbLf & _
                       "Vorher wird zwingend eine Sicherung erstellt; ohne erfolgreiche Sicherung wird nichts gelöscht.", "Datei zurücksetzen", "Weiter", "Abbrechen", True) Then Exit Sub
    word = mUI.AskText("Zur Bestätigung bitte ZURÜCKSETZEN eingeben.", "Datei zurücksetzen", "", False, ok)
    If Not ok Then Exit Sub
    If UCase$(Trim$(word)) <> "ZURÜCKSETZEN" Then mUI.Alert "Die Bestätigung stimmt nicht. Es wurde nichts gelöscht.", "Datei zurücksetzen": Exit Sub
    If Not BackupNow("vor dem Zurücksetzen", True) Then
        mUI.Alert "Die Sicherung ist fehlgeschlagen. Aus Sicherheitsgründen wurde nichts gelöscht.", "Datei zurücksetzen", True
        Exit Sub
    End If
    mApp.BeginWork "Datei wird zurückgesetzt …"
    For i = 0 To ArrLen(ids) - 1
        Set lo_ = mData.DataLO(CStr(ids(i)))
        mProtect.Unprot lo_.Parent
        If Not lo_.DataBodyRange Is Nothing Then lo_.DataBodyRange.Rows.Delete
    Next
    Set lo_ = mCfg.LO(SH_REL)
    If Not lo_.DataBodyRange Is Nothing Then lo_.DataBodyRange.Rows.Delete
    mData.InvalidateAllData
    mData.ClearUndo
    mCfg.InvalidateAll
    mLog.Add "Zurücksetzen", "", "", "", FmtInt(n) & " Einträge gelöscht (Sicherung vorhanden)"
    mApp.MarkCfgDirty
    mApp.EndWork
    mNav.ResetAllStates
    mNav.GoStart
    mUI.Alert "Alle Einträge wurden gelöscht. Die vorherige Fassung liegt als Sicherung vor:" & vbLf & BackupFolder(), "Datei zurücksetzen"
End Sub
