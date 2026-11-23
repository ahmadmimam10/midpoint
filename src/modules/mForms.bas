Attribute VB_Name = "mForms"
'==============================================================================
' mForms - Oeffnen der Formulare (Eintrag, Auswahl, Datum) und gemeinsame
' Datensatzaktionen (Archivieren, Wiederherstellen, endgueltig Loeschen).
'==============================================================================
Option Explicit

Private mDepth As Long
Private mXlBar As frmXlBar

'--- Schwebende Formelleiste (Formelspalten) ------------------------------------------
Public Sub OpenXlBar(ByVal tid As String, ByVal fid As String, ByVal afterFid As String)
    CloseXlBar
    Set mXlBar = New frmXlBar
    mXlBar.Setup tid, fid, afterFid
    mXlBar.Show vbModeless
    On Error Resume Next
    ' Fenster richtig aktivieren, sonst wirkt der erste Klick in ein Feld nicht
    mWin.ActivateWindow mWin.FormHwnd(mXlBar.caption)
    mXlBar.Controls("xl_txt").SetFocus
End Sub

Public Sub CloseXlBar()
    On Error Resume Next
    If mXlBar Is Nothing Then Exit Sub
    mXlBar.Teardown
    Unload mXlBar
    Set mXlBar = Nothing
End Sub

' Klick in eine Tabellenspalte, waehrend die Formelleiste offen ist: True = eingefuegt
Public Function XlBarTakesClick(ByVal target As Range) As Boolean
    Dim fid As String, ev As Boolean
    If mXlBar Is Nothing Then Exit Function
    fid = mViewTable.FidAtCol(target.Column)
    If Len(fid) = 0 Then Exit Function
    If mXlBar.TableId() <> mNav.gTid Then Exit Function
    ' Auswahl in die (unsichtbare) ID-Spalte derselben Zeile setzen: so loest auch ein
    ' erneuter Klick auf dieselbe Zelle wieder ein Einfuegen aus
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    target.Worksheet.Cells(target.Row, mViewTable.COL_ID).Select
    On Error GoTo 0
    Application.EnableEvents = ev
    mXlBar.InsertFromSheet fid
    XlBarTakesClick = True
End Function

Public Function XlBarOpen() As Boolean
    XlBarOpen = Not mXlBar Is Nothing
End Function

Public Sub OpenRecord(ByVal tid As String, ByVal rid As String)
    Dim f As frmRecord, saved As String
    mApp.ResetWork
    Set f = New frmRecord
    mDepth = mDepth + 1
    saved = f.OpenFor(tid, rid)
    Unload f
    mDepth = mDepth - 1
    mNav.FocusGrid
    If Len(saved) > 0 Then NotifySaved tid
End Sub

Public Sub OpenRecordStacked(ByVal tid As String, ByVal rid As String)
    OpenRecord tid, rid
End Sub

Public Function NewRecord(ByVal tid As String, Optional ByVal presets As cMap = Nothing) As String
    Dim f As frmRecord
    If Not mApp.RequireWritable("Das Anlegen eines Eintrags") Then Exit Function
    mApp.ResetWork
    Set f = New frmRecord
    mDepth = mDepth + 1
    NewRecord = f.OpenFor(tid, "", presets)
    Unload f
    mDepth = mDepth - 1
    mNav.FocusGrid
    If Len(NewRecord) > 0 Then NotifySaved tid
End Function

Public Function NewRecordStacked(ByVal tid As String, ByVal presets As cMap) As String
    NewRecordStacked = NewRecord(tid, presets)
End Function

' Nach Speichern: sichtbare Seite aktualisieren (auch waehrend weitere Formulare offen sind)
Public Sub NotifySaved(ByVal tid As String)
    On Error Resume Next
    mNav.RenderCurrent
    mNav.RefreshNav
End Sub

Public Function PickRecords(ByVal tid As String, ByVal cur As String, ByVal multi As Boolean, ByVal title As String, ByRef ok As Boolean) As String
    Dim f As frmPicker
    Set f = New frmPicker
    f.Setup tid, cur, multi, title
    f.Show: mNav.FocusGrid
    ok = f.OK
    If ok Then PickRecords = f.Result
    Unload f
End Function

' "Welches Diagramm passt?": Übersicht mit Vorschaubildern; liefert die gewählte Darstellung oder ""
Public Function ChooseChartType(ByVal cur As String) As String
    Dim f As frmChartGallery
    On Error GoTo fl
    Application.StatusBar = "Vorschaubilder werden vorbereitet …"
    Set f = New frmChartGallery
    f.Setup cur
    Application.StatusBar = False
    f.Show
    If f.OK Then ChooseChartType = f.Result
    Unload f
    Exit Function
fl:
    Application.StatusBar = False
    mApp.Boundary "Diagrammübersicht"
End Function

Public Function PickDate(ByVal cur As String, ByVal withTime As Boolean, ByVal fieldName As String, ByRef ok As Boolean) As String
    Dim f As frmDate
    Set f = New frmDate
    f.Setup cur, withTime, fieldName
    f.Show: mNav.FocusGrid
    ok = f.OK
    If ok Then PickDate = f.Result
    Unload f
End Function

' Wert einer Button-Aktion fuer ein Zielfeld: heute, jetzt, ich, leer, umschalten oder Eingabetext
Public Function ButtonValue(ByVal target As cField, ByVal spec As String, ByVal cur As Variant, ByRef v As Variant, ByRef msg As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(spec))
    Select Case s
        Case "heute": v = Date: ButtonValue = True
        Case "jetzt": v = Now: ButtonValue = True
        Case "leer": v = Empty: ButtonValue = True
        Case "umschalten"
            If IsBlank(cur) Then v = True Else v = Not CBool(cur)
            ButtonValue = True
        Case "ich"
            If target.IsOptionType Then
                ButtonValue = mValid.ParseInput(target, UserId(), v, msg)
                If Not ButtonValue Then msg = "Die Person „" & UserId() & "“ ist in „" & target.Name & "“ nicht angelegt."
            Else
                v = UserId(): ButtonValue = True
            End If
        Case Else
            ButtonValue = mValid.ParseInput(target, spec, v, msg)
    End Select
End Function

' Fuehrt die Aktionen eines Button-Feldes fuer einen Datensatz aus (Pruefung, Transaktion, Protokoll)
Public Sub RunButton(ByVal tid As String, ByVal rid As String, ByVal fid As String)
    Dim f As cField, p As Variant, i As Long, tf As cField, vals As New cMap, v As Variant, msg As String, errs As cMap
    If Not mApp.RequireWritable("Die Button-Aktion") Then Exit Sub
    Set f = mCfg.Field(fid)
    p = Split(f.Formula, "|")
    If UBound(p) < 2 Then mUI.Alert "Für den Button „" & f.Name & "“ sind noch keine Aktionen festgelegt (Eigenschaft bearbeiten).", "Button": Exit Sub
    For i = 1 To UBound(p) - 1 Step 2
        If Left$(p(i), 2) = "F:" Then
            Set tf = mCfg.Field(Mid$(p(i), 3))
            If Not tf Is Nothing Then
                If Not ButtonValue(tf, Unesc(CStr(p(i + 1))), mData.GetValue(tid, rid, tf.ID), v, msg) Then
                    mUI.Alert "Button „" & f.Name & "“, Feld „" & tf.Name & "“: " & msg, "Button", True
                    Exit Sub
                End If
                vals.SetVal tf.ID, v
            End If
        End If
    Next
    If Not mValid.ValidateRecord(tid, rid, vals, New cMap, errs) Then
        mUI.Alert "Der Button kann nicht ausgeführt werden: " & Join(errs.Items, " "), "Button", True
        Exit Sub
    End If
    mApp.BeginWork
    On Error GoTo failed
    mData.TxBegin "Button „" & f.Name & "“ bei „" & mData.RecordLabel(tid, rid) & "“"
    mData.UpdateRecord tid, rid, vals
    mLog.Add "Button", tid, rid, Join(vals.Keys, ";"), f.Name
    mData.TxCommit
    mApp.EndWork
    NotifySaved tid
    Exit Sub
failed:
    mData.TxAbortIfOpen
    mApp.ResetWork
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ButtonDescription(ByVal f As cField) As String
    Dim p As Variant, i As Long, s As String
    p = Split(f.Formula, "|")
    For i = 1 To UBound(p) - 1 Step 2
        If Left$(p(i), 2) = "F:" Then s = s & IIf(Len(s) > 0, "; ", "") & mCfg.FieldName(Mid$(p(i), 3)) & " = " & Unesc(CStr(p(i + 1)))
    Next
    If Len(s) = 0 Then s = "keine Aktion festgelegt"
    ButtonDescription = s
End Function

Public Sub SetArchived(ByVal tid As String, ByVal rid As String, ByVal arch As Boolean)
    mApp.BeginWork
    On Error GoTo failed
    mData.TxBegin IIf(arch, "Archivieren", "Wiederherstellen") & " von „" & mData.RecordLabel(tid, rid) & "“"
    mData.SetArchived tid, rid, arch
    mLog.Add IIf(arch, "Archiviert", "Wiederhergestellt"), tid, rid, ""
    mData.TxCommit
    mApp.EndWork
    NotifySaved tid
    Exit Sub
failed:
    mData.TxAbortIfOpen
    mApp.ResetWork
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ArchiveMany(ByVal tid As String, ByVal rids As Collection, ByVal arch As Boolean)
    Dim v As Variant
    mApp.BeginWork
    On Error GoTo failed
    mData.TxBegin IIf(arch, "Archivieren", "Wiederherstellen") & " (" & Plural(rids.Count, "Eintrag", "Einträge") & ")"
    For Each v In rids
        mData.SetArchived tid, CStr(v), arch
        mLog.Add IIf(arch, "Archiviert", "Wiederhergestellt"), tid, CStr(v), ""
    Next
    mData.TxCommit
    mApp.EndWork
    NotifySaved tid
    Exit Sub
failed:
    mData.TxAbortIfOpen
    mApp.ResetWork
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Endgueltiges Loeschen mit Referenzpruefung und klarer Bestaetigung
Public Sub DeleteForever(ByVal tid As String, ByVal rids As Collection)
    Dim v As Variant, refs As Long, total As Long, items() As Variant, n As Long, msg As String, word As String, ok As Boolean
    For Each v In rids
        If Not mData.IsArchivedRec(tid, CStr(v)) Then
            mUI.Alert "Endgültig gelöscht werden können nur archivierte Einträge. Bitte zuerst archivieren und dann im Archiv löschen.", "Endgültig löschen"
            Exit Sub
        End If
    Next
    ReDim items(0 To rids.Count - 1, 0 To 2)
    For Each v In rids
        refs = mData.RelCountFor(CStr(v))
        total = total + refs
        items(n, 0) = mData.RecordLabel(tid, CStr(v))
        items(n, 1) = CStr(v)
        items(n, 2) = IIf(refs = 0, "keine Verknüpfungen", refs & " Verknüpfung(en) werden mit gelöscht")
        n = n + 1
    Next
    msg = "Diese Einträge werden dauerhaft entfernt. Verknüpfte andere Einträge bleiben bestehen; nur die Zuordnungen zu den gelöschten Einträgen " & _
          "werden entfernt (keine verwaisten Verknüpfungen, kein kaskadierendes Löschen)." & vbLf & _
          "Die Aktion lässt sich in dieser Sitzung mit „Rückgängig“ zurücknehmen; vorher wird eine Sicherung angeboten."
    If mUI.ShowReport("Endgültig löschen", msg, items, 3, "220;80;220", Array("Weiter", "Abbrechen"), True) <> 1 Then Exit Sub
    word = mUI.AskText("Zur Bestätigung bitte LÖSCHEN eingeben.", "Endgültig löschen", "", False, ok)
    If Not ok Or UCase$(Trim$(word)) <> "LÖSCHEN" Then
        If ok Then mUI.Alert "Die Bestätigung stimmt nicht. Es wurde nichts gelöscht.", "Endgültig löschen"
        Exit Sub
    End If
    If Not mBackup.OfferBackup("vor dem endgültigen Löschen") Then Exit Sub
    mApp.BeginWork
    On Error GoTo failed
    mData.TxBegin "Endgültiges Löschen (" & Plural(rids.Count, "Eintrag", "Einträge") & ")"
    For Each v In rids
        mData.DeleteRecordHard tid, CStr(v)
        mLog.Add "Gelöscht", tid, CStr(v), ""
    Next
    mData.TxCommit
    mApp.EndWork
    NotifySaved tid
    Exit Sub
failed:
    mData.TxAbortIfOpen
    mApp.ResetWork
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
