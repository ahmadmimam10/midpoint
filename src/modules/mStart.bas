Attribute VB_Name = "mStart"
'==============================================================================
' mStart - Start und Schliessen der Anwendung, Formatpruefung/Migration,
' Lesemodus, Integritaetspruefung.
'==============================================================================
Option Explicit

Public Sub OnOpen()
    ' Start erst nach vollstaendigem Laden des Fensters
    Application.OnTime Now, mRender.MacroRef("mStart.DeferredStart")
End Sub

Public Sub DeferredStart()
    On Error GoTo f
    StartApp
    Exit Sub
f: mApp.Boundary "Start der Anwendung"
End Sub

Public Sub StartApp()
    Dim missing As String, sh As Variant
    If Not ThisWorkbook Is ActiveWorkbook Then ThisWorkbook.Activate
    For Each sh In Array(SH_START, SH_TABLE, SH_CAL, SH_BOARD, SH_GANTT, SH_DASH, SH_SETUP, SH_HELP, SH_TECH, SH_SEARCH, SH_SYS, SH_TAB, SH_FLD, SH_OPT, SH_REL, SH_VIEW, SH_DSB, SH_ELM, SH_LOG, SH_LST)
        If Not mApp.SheetExists(CStr(sh)) Then missing = missing & " " & sh
    Next
    If Len(missing) > 0 Then
        mUI.Alert "Die Datei ist unvollständig (fehlende Blätter:" & missing & "). Bitte eine Sicherung verwenden.", "Start", True
        Exit Sub
    End If
    mApp.gReadOnly = ThisWorkbook.ReadOnly
    mCfg.InvalidateAll
    mData.InvalidateAllData
    mData.ClearUndo
    mNav.ResetAllStates
    MigrateIfNeeded
    mProtect.ProtectAllUi
    mApp.gDataDirty = False
    If Not mApp.gReadOnly Then
        mCalc.RecalcIfNewDay
        If mApp.gDataRev > 0 Then mApp.gDataDirty = False
    End If
    mApp.gStarted = True
    mNav.GoStart
    ThisWorkbook.Worksheets(SH_HINWEIS).Visible = xlSheetVeryHidden
    mNav.ShowNav
    ThisWorkbook.Saved = True
    If mApp.gReadOnly Then
        mUI.Alert "Die Datei ist schreibgeschützt geöffnet (Lesemodus), zum Beispiel, weil sie gerade von einer anderen Person bearbeitet wird " & _
                  "oder Schreibrechte fehlen. Sie können lesen, suchen, filtern und exportieren; Änderungen sind gesperrt.", "Lesemodus"
    End If
End Sub

Private Sub MigrateIfNeeded()
    Dim v As Long
    v = mCfg.SysLong("Formatversion", FORMAT_VERSION)
    If v > FORMAT_VERSION Then
        mUI.Alert "Diese Datei wurde mit einer neueren Programmversion (Format " & v & ") erstellt. Bitte die passende Programmversion verwenden; " & _
                  "die Datei wird im Lesemodus betrieben, um Datenverlust zu vermeiden.", "Version", True
        mApp.gReadOnly = True
        Exit Sub
    End If
    If v < FORMAT_VERSION And Not mApp.gReadOnly Then
        ' Migrationsschritte je Formatversion, ohne Neuerfassung
        ' Format 2: Blaetter fuer Erfassungsdateien (Seite, _Erfassungen, _Ruecklauf)
        ' Format 3: Zusatzspalten fuer Erfassungsdateien (Optionen, Empfaenger)
        If v < 3 Then
            mCapture.EnsureSheets
            ' bereits benannte Dateien gelten als "Titel festgelegt"
            If mCfg.SysGet("AppName", APP_TITLE) <> APP_TITLE Then mCfg.SysSet "TitelFestgelegt", "Ja"
        End If
        mCfg.SysSet "Formatversion", CStr(FORMAT_VERSION)
        mLog.Add "Migration", "", "", "", "Format " & v & " " & ChrW(&H2192) & " " & FORMAT_VERSION
    End If
    If mCfg.SysGet("Programmversion", "") <> PROG_VERSION And Not mApp.gReadOnly Then
        mCfg.SysSet "Programmversion", PROG_VERSION
    End If
End Sub

Public Sub OnBeforeSave(ByVal saveAsUI As Boolean)
    On Error GoTo f
    If Not mApp.gStarted Then
        ' ohne gestartete Anwendung normal speichern lassen
        mApp.gSaving = True
        Application.EnableEvents = False
        ThisWorkbook.Save
        Application.EnableEvents = True
        mApp.gSaving = False
        Exit Sub
    End If
    mBackup.ControlledSave saveAsUI
    Exit Sub
f: mApp.gSaving = False
    Application.EnableEvents = True
    mApp.Boundary "Speichern"
End Sub

Public Sub OnBeforeClose(ByRef Cancel As Boolean)
    Dim ans As Long
    On Error GoTo f
    If Not mApp.gStarted Then Exit Sub
    If mApp.gDataDirty And Not mApp.gReadOnly Then
        ans = mUI.Ask3("Es gibt ungespeicherte Änderungen. Vor dem Schließen speichern?", "Schließen", "Speichern", "Nicht speichern", "Abbrechen")
        Select Case ans
            Case 1
                If Not mBackup.ControlledSave(False) Then Cancel = True: Exit Sub
            Case 2
                ThisWorkbook.Saved = True
            Case Else
                Cancel = True
                Exit Sub
        End Select
    Else
        ThisWorkbook.Saved = True
    End If
    mApp.gClosing = True
    ' geplantes automatisches Speichern abmelden (sonst oeffnet Excel die Datei erneut)
    mBackup.CancelAutoSave
    mNav.StopWatch
    mForms.CloseXlBar
    If mDash.gArrange Then mDash.EndArrange
    mNav.HideNav
    On Error Resume Next
    Application.CommandBars("DBK_Menu").Delete
    Exit Sub
f: mApp.Boundary "Schließen"
End Sub

' Prueft Struktur und Bezuege; liefert lesbaren Bericht
Public Function IntegrityReport() As String
    Dim ids As Variant, i As Long, tid As String, f As cField, t As cTableData, s As String, problems As Long, r As Long, rel As cTableData
    Dim seen As cMap, c As Long, orphan As Long, fid As String, v As Variant, a As Variant, k As Long, nRec As Long
    ids = mCfg.TableIds(True)
    For i = 0 To ArrLen(ids) - 1
        tid = CStr(ids(i))
        If Not mApp.SheetExists(mCfg.DataSheetName(tid)) Then
            s = s & "• Datenblatt fehlt für „" & mCfg.TableName(tid) & "“." & vbLf: problems = problems + 1
        Else
            Set t = mData.TData(tid)
            nRec = nRec + t.RowCount
            For Each f In mCfg.Fields(tid, False)
                If Not t.HasCol(f.ID) Then s = s & "• Spalte für Feld „" & f.Name & "“ fehlt (" & mCfg.TableName(tid) & ")." & vbLf: problems = problems + 1
            Next
            Set seen = New cMap
            For r = 1 To t.RowCount
                If seen.Exists(t.IdAt(r)) Then s = s & "• Doppelte Datensatz-ID " & t.IdAt(r) & "." & vbLf: problems = problems + 1
                seen.SetVal t.IdAt(r), True
            Next
            ' Optionsbezuege
            For Each f In mCfg.Fields(tid, False)
                If f.IsOptionType Then
                    c = t.ColOf(f.ID)
                    For r = 1 To t.RowCount
                        If Not IsBlank(t.Cell(r, c)) Then
                            a = Split(CStr(t.Cell(r, c)), ";")
                            For k = 0 To UBound(a)
                                If Len(mCfg.GetS(SH_OPT, CStr(a(k)), "Name")) = 0 Then orphan = orphan + 1
                            Next
                        End If
                    Next
                End If
            Next
            If Len(mCfg.LabelField(tid)) = 0 Or Not mCfg.FieldExists(mCfg.LabelField(tid)) Then
                s = s & "• „" & mCfg.TableName(tid) & "“ hat kein gültiges Bezeichnungsfeld." & vbLf: problems = problems + 1
            End If
        End If
    Next
    If orphan > 0 Then s = s & "• " & orphan & " Werte verweisen auf nicht vorhandene Auswahloptionen." & vbLf: problems = problems + 1
    ' Zuordnungen
    Set rel = mCfg.TD(SH_REL)
    orphan = 0
    For r = 1 To rel.RowCount
        fid = CStr(rel.Cell(r, 1))
        If Not mCfg.FieldExists(fid) Then
            orphan = orphan + 1
        Else
            Set f = mCfg.Field(fid)
            If Not mData.RecordExists(f.TableID, CStr(rel.Cell(r, 2))) Or Not mData.RecordExists(f.TargetTable, CStr(rel.Cell(r, 3))) Then orphan = orphan + 1
        End If
    Next
    If orphan > 0 Then s = s & "• " & orphan & " Zuordnungen verweisen auf nicht vorhandene Einträge oder Felder." & vbLf: problems = problems + 1
    If problems = 0 Then
        IntegrityReport = "Integritätsprüfung ohne Befund: " & ArrLen(ids) & " Tabellen, " & FmtInt(nRec) & " Einträge, " & FmtInt(rel.RowCount) & " Zuordnungen geprüft."
    Else
        IntegrityReport = "Integritätsprüfung: " & problems & " Auffälligkeit(en)." & vbLf & s & vbLf & "Bitte gegebenenfalls eine Sicherung wiederherstellen."
    End If
End Function
