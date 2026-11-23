Attribute VB_Name = "mNav"
'==============================================================================
' mNav - Navigation: aktuelle Seite/Tabelle/Ansicht, Sitzungszustaende der
' Ansichten, angedocktes Navigationsfenster (frmNav) und Seitenaufbau.
'==============================================================================
Option Explicit

Public gPage As String          ' Start | Tabelle | Dashboard | Einrichten | Hilfe | Technik | Suche
Public gTid As String
Public gVid As String
Public gDash As String
Public gGlobalSearch As String

Private mStates As cMap         ' ViewID -> cViewState
Private mNavForm As frmNav
Private mNavHwnd As LongPtr
Private mGridHwnd As LongPtr
Private mLastH As Long
Private mLastW As Long           ' zuletzt gesetzte Breite der Navigation (Pixel)
Private mRendering As Boolean
Private mLastCheck As Single
Private mPageKey As cMap        ' Blattname -> Stand beim letzten Aufbau
Public gZoom As Long            ' Excel-Zoom der Nutzerin/des Nutzers (gilt für alle Seiten)
Private mLastZoom As Long
Private mWatchOn As Boolean, mWatchAt As Date, mWatchProc As String

'--- Sitzungszustaende ---------------------------------------------------------------
Public Function State(ByVal vid As String) As cViewState
    Dim s As cViewState
    If mStates Is Nothing Then Set mStates = New cMap
    If mStates.Exists(vid) Then
        Set State = mStates.GetVal(vid)
    Else
        Set s = New cViewState
        s.InitFrom vid
        mStates.SetVal vid, s
        Set State = s
    End If
End Function

Public Sub ResetState(ByVal vid As String)
    If Not mStates Is Nothing Then mStates.Remove vid
End Sub

Public Sub ResetAllStates()
    Set mStates = Nothing
    gGlobalSearch = ""
End Sub

Public Function CurState() As cViewState
    If Len(gVid) > 0 Then Set CurState = State(gVid)
End Function

'--- Navigationsmodell ------------------------------------------------------------------
' Liefert Collection von Array(Schluessel, Symbol, Text, Art)
Public Function NavItems() As Collection
    Dim c As New Collection, ids As Variant, i As Long
    c.Add Array("START", ChrW(&HE80F), "Start", "item")
    c.Add Array("", "", "TABELLEN", "section")
    ids = mCfg.TableIds()
    For i = 0 To ArrLen(ids) - 1
        c.Add Array("T:" & ids(i), mCfg.TableIcon(CStr(ids(i))), mCfg.TableName(CStr(ids(i))), "item")
    Next
    ' eigener Schluessel: oeffnet direkt den Dialog (nicht doppelt "SETUP" markieren)
    If ArrLen(ids) = 0 Then c.Add Array("NEWTAB", ChrW(&HE710), "Erste Tabelle anlegen", "item")
    c.Add Array("", "", "AUSWERTEN", "section")
    c.Add Array("DASH", ChrW(&HE9D2), "Dashboards", "item")
    c.Add Array("", "", "ERFASSUNG", "section")
    c.Add Array("CAPT", ChrW(&HE8A5), "Erfassungsdateien", "item")
    c.Add Array("", "", "", "bottom")
    c.Add Array("SETUP", ChrW(&HE713), "Einrichten", "item")
    c.Add Array("HELP", ChrW(&HE897), "Hilfe & Sicherung", "item")
    c.Add Array("TECH", ChrW(&HE943), "Technische Dokumentation", "item")
    Set NavItems = c
End Function

Public Function ActiveNavKey() As String
    Select Case gPage
        Case "Start": ActiveNavKey = "START"
        Case "Tabelle": ActiveNavKey = "T:" & gTid
        Case "Dashboard": ActiveNavKey = "DASH"
        Case "Erfassung": ActiveNavKey = "CAPT"
        Case "Einrichten": ActiveNavKey = "SETUP"
        Case "Hilfe": ActiveNavKey = "HELP"
        Case "Technik": ActiveNavKey = "TECH"
        Case Else: ActiveNavKey = ""
    End Select
End Function

Public Sub NavClick(ByVal key As String)
    On Error GoTo failed
    Select Case True
        Case key = "START": GoStart
        Case Left$(key, 2) = "T:": GoTable Mid$(key, 3)
        Case key = "DASH": GoDash ""
        Case key = "CAPT": GoPage "Erfassung"
        Case key = "SETUP": GoPage "Einrichten"
        Case key = "NEWTAB": mCmd.Cmd_NewTable
        Case key = "HELP": GoPage "Hilfe"
        Case key = "TECH": GoPage "Technik"
        Case key = "MODE": mBackup.ToggleAutoSave
        Case key = "TITLE": mCmd.EditTitle
    End Select
    Exit Sub
failed:
    mApp.Boundary "Navigation"
End Sub

'--- Seitenwechsel ---------------------------------------------------------------------
Public Sub GoStart()
    GoPage "Start"
End Sub

Public Function FirstViewOf(ByVal tid As String) As String
    Dim ids As Variant
    ids = mCfg.ViewIds(tid)
    If ArrLen(ids) = 0 Then FirstViewOf = mStruct.CreateDefaultView(tid) Else FirstViewOf = CStr(ids(0))
End Function

Public Sub GoTable(ByVal tid As String, Optional ByVal vid As String = "")
    If Not mCfg.TableExists(tid) Then GoStart: Exit Sub
    If Len(vid) = 0 Or Not mCfg.ViewExists(vid) Then
        If gTid = tid And Len(gVid) > 0 And mCfg.ViewExists(gVid) Then
            If mCfg.ViewGet(gVid, "Tabelle") = tid Then vid = gVid
        End If
        If Len(vid) = 0 Then vid = FirstViewOf(tid)
    End If
    gTid = tid
    gVid = vid
    GoPage "Tabelle"
End Sub

Public Sub GoDash(ByVal did As String)
    Dim ids As Variant
    If Len(did) = 0 Then
        If Len(gDash) > 0 And mCfg.TD(SH_DSB).RowOf(gDash) > 0 Then
            did = gDash
        Else
            ids = mCfg.DashIds()
            If ArrLen(ids) > 0 Then did = CStr(ids(0))
        End If
    End If
    gDash = did
    GoPage "Dashboard"
End Sub

Public Sub GoPage(ByVal page As String)
    If page <> "Tabelle" Then mForms.CloseXlBar
    If page <> "Dashboard" And mDash.gArrange Then mDash.EndArrange
    gPage = page
    ' Seite unveraendert seit dem letzten Aufbau? Dann nur zeigen (schneller Seitenwechsel)
    If Not TryReuse() Then RenderCurrent
    RefreshNav
End Sub

'--- Wiederverwendung aufgebauter Seiten -------------------------------------------------
' Jede Seite ist ein eigenes Blatt. Hat sich seit ihrem letzten Aufbau nichts geaendert
' (Daten, Struktur, Ansichtszustand, Fensterbreite, Datum), wird sie nur eingeblendet.
Private Function PageKey() As String
    Dim k As String
    On Error Resume Next
    ' Fensterbreite in Pixeln des Rasterfensters: UsableWidth meldet Excel je Blatt leicht
    ' unterschiedlich; der Schlüssel passte dann nie und jede Seite wurde neu aufgebaut
    k = gPage & "|" & gTid & "|" & gVid & "|" & gDash & "|" & mApp.gDataRev & "|" & mApp.gCfgRev & "|" & _
        CLng(mWin.ClientWidthPx(mGridHwnd) / 8) & "|" & CLng(Date) & "|" & mApp.gReadOnly & "|" & mData.UndoAvailable & "|" & _
        ThisWorkbook.Name & "|" & mDash.gArrange & "|" & mDash.gHighlight & "|" & mPages.DataSheetsVisible() & "|" & gZoom & "|" & mWin.WindowDpi(mGridHwnd)
    If gPage = "Tabelle" And Len(gVid) > 0 Then k = k & "|" & State(gVid).Signature()
    PageKey = k
End Function

Private Function TryReuse() As Boolean
    Dim sh As String, ws As Worksheet
    On Error GoTo nope
    If mPageKey Is Nothing Then Exit Function
    Select Case gPage
        Case "Start", "Dashboard", "Einrichten", "Hilfe", "Technik"
        Case "Tabelle"
            If Not mCfg.TableExists(gTid) Then Exit Function
            If Not mCfg.ViewExists(gVid) Then Exit Function
        Case Else: Exit Function
    End Select
    sh = CurrentSheetName()
    If CStr(mPageKey.GetVal(sh, "")) <> PageKey() Then Exit Function
    Set ws = ThisWorkbook.Worksheets(sh)
    If ws.Shapes.Count = 0 Then Exit Function
    mApp.BeginWork
    mRender.ShowSheet ws
    mApp.EndWork
    mRender.RefreshStatus
    mApp.SyncSavedFlag
    TryReuse = True
    Exit Function
nope:
    mApp.ResetWork
End Function

Private Sub RememberPage()
    On Error Resume Next
    If mPageKey Is Nothing Then Set mPageKey = New cMap
    mPageKey.SetVal CurrentSheetName(), PageKey()
End Sub

' Nach Aenderungen ohne Daten-/Strukturstand (z. B. Wiederherstellung): alles neu aufbauen
Public Sub ForgetPages()
    Set mPageKey = Nothing
End Sub

' Baut die aktuelle Seite neu auf
Public Sub RenderCurrent()
    Dim vt As String
    If mRendering Then Exit Sub
    mRendering = True
    On Error GoTo failed
    Select Case gPage
        Case "Tabelle"
            If Not mCfg.TableExists(gTid) Then
                gPage = "Start"
                mPages.RenderStart
                GoTo done
            End If
            If Not mCfg.ViewExists(gVid) Then gVid = FirstViewOf(gTid)
            vt = mCfg.ViewGet(gVid, "Typ")
            Select Case vt
                Case VT_CAL: mViewCal.Render
                Case VT_BOARD: mViewBoard.Render
                Case VT_GANTT: mViewGantt.Render
                Case Else: mViewTable.Render
            End Select
        Case "Dashboard": mDash.Render
        Case "Erfassung": mCapture.Render
        Case "Einrichten": mPages.RenderSetup
        Case "Hilfe": mPages.RenderHelp
        Case "Technik": mPages.RenderTech
        Case "Suche": mPages.RenderSearch
        Case Else
            gPage = "Start"
            mPages.RenderStart
    End Select
done:
    mRendering = False
    RememberPage
    Exit Sub
failed:
    mRendering = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub RefreshCurrent()
    RenderCurrent
End Sub

' Nach Datenaenderungen: aktuelle Seite und Statusband aktualisieren
Public Sub AfterDataChange()
    RenderCurrent
End Sub

Public Function CurrentSheetName() As String
    Select Case gPage
        Case "Tabelle"
            Select Case mCfg.ViewGet(gVid, "Typ")
                Case VT_CAL: CurrentSheetName = SH_CAL
                Case VT_BOARD: CurrentSheetName = SH_BOARD
                Case VT_GANTT: CurrentSheetName = SH_GANTT
                Case Else: CurrentSheetName = SH_TABLE
            End Select
        Case "Dashboard": CurrentSheetName = SH_DASH
        Case "Erfassung": CurrentSheetName = SH_CAPTURE
        Case "Einrichten": CurrentSheetName = SH_SETUP
        Case "Hilfe": CurrentSheetName = SH_HELP
        Case "Technik": CurrentSheetName = SH_TECH
        Case "Suche": CurrentSheetName = SH_SEARCH
        Case Else: CurrentSheetName = SH_START
    End Select
End Function

'==============================================================================
' Angedocktes Navigationsfenster
'==============================================================================
Public Sub ShowNav()
    Dim hp As LongPtr, cap As String
    On Error GoTo failed
    If Not ThisWorkbook Is ActiveWorkbook Then Exit Sub
    hp = mWin.WorkbookGridHwnd()
    If hp = 0 Then Exit Sub
    If Not mNavForm Is Nothing Then
        If mWin.HwndAlive(mNavHwnd) And mWin.ParentOf(mNavHwnd) = hp Then
            PositionNav
            Exit Sub
        End If
        HideNav
    End If
    Set mNavForm = New frmNav
    cap = "DBK_NAV_" & Format$(Timer * 100, "0")
    mNavForm.caption = cap
    mNavForm.Show vbModeless
    mNavHwnd = mWin.FormHwnd(cap)
    If mNavHwnd = 0 Then GoTo failed
    mWin.DockAsChild mNavHwnd, hp
    mGridHwnd = hp
    mLastH = 0
    PositionNav
    Exit Sub
failed:
    ' Navigation ist Komfort; Fehler hier duerfen die Anwendung nicht blockieren
    On Error Resume Next
    HideNav
End Sub

Public Sub PositionNav()
    Dim h As Long, w As Long, hFull As Long
    If mNavForm Is Nothing Then Exit Sub
    If Not mWin.HwndAlive(mGridHwnd) Then Exit Sub
    h = NavHeightPx()
    w = NavWidthPx()
    ' erst die Fenstergroesse setzen, dann misst das Formular seine Punkte selbst. Das Fenster
    ' reicht bis ganz unten (die Bildlaufleiste liegt darueber, kein Spalt), die Eintraege
    ' nutzen nur die sichtbare Hoehe h.
    hFull = h
    mWin.PlaceWindow mNavHwnd, 0, 0, w, hFull
    If h <> mLastH Or w <> mLastW Then
        mLastH = h
        mLastW = w
        mNavForm.Relayout w, h, hFull
    End If
End Sub

' Breite der Navigation: fest 184 pt in Bildschirmpixeln (Skalierung und Monitorwechsel
' eingeschlossen), unabhaengig vom Excel-Zoom; Spalte A wird beim Seitenaufbau passend zum
' Zoom gesetzt (mRender.NavColPt). Plus 5 Pixel (skaliert): Excel rundet die Spaltenbreite je
' Zoom auf, die Kante der festgehaltenen Spalte A und der Auswahlrahmen in A1 liegen so sicher
' unter der Navigation.
Private Function NavWidthPx() As Long
    Dim px As Long
    On Error Resume Next
    ' DPI des Fensters selbst (zweiter Bildschirm mit anderer Skalierung), nicht nur des Hauptbildschirms
    px = CLng(NAV_WIDTH_PT * mWin.WindowDpi(mGridHwnd) / 72#)
    On Error GoTo 0
    If px < 40 Or px > 2000 Then px = mWin.PtToPx(NAV_WIDTH_PT)
    NavWidthPx = px + CLng(5# * mWin.WindowDpi(mGridHwnd) / 96#)
End Function

' Prueft guenstig, ob sich die Fenstergroesse geaendert hat (z. B. Menueband ein/aus, anderer Bildschirm)
Public Sub CheckNavSize()
    StartWatch
    If mNavForm Is Nothing Then Exit Sub
    If Not mWin.HwndAlive(mGridHwnd) Then Exit Sub
    If NavHeightPx() <> mLastH Or NavWidthPx() <> mLastW Then PositionNav
End Sub

' Sichtbare Hoehe: Das Rasterfenster meldet seine Hoehe einschliesslich der unteren
' Leiste (Blattregister, waagerechte Bildlaufleiste), die ueber der Navigation liegt.
Private Function NavHeightPx() As Long
    Dim h As Long, uh As Long
    h = mWin.ClientHeightPx(mGridHwnd)
    On Error Resume Next
    ' sichtbare Rasterhoehe laut Excel (endet genau an der Bildlaufleiste), in Pixeln dieses Fensters
    uh = CLng(ThisWorkbook.Windows(1).UsableHeight * mWin.WindowDpi(mGridHwnd) / 72#)
    ' Rundung ausgleichen: lieber 3 Pixel unter die Bildlaufleiste (sie liegt darüber) als ein heller Spalt
    If uh > 50 Then uh = uh + CLng(3# * mWin.WindowDpi(mGridHwnd) / 96#)
    On Error GoTo 0
    If uh > 50 And uh < h Then h = uh
    NavHeightPx = h
End Function

'--- Fenstergroesse ohne Zeitgeber ------------------------------------------------------
' Das Ein-/Ausklappen des Menuebands meldet Excel nicht als Ereignis. Die Navigation
' prueft ihre Hoehe deshalb bei Mausbewegung ueber der Navigation, bei Zellauswahl,
' Fensteraktivierung und jedem Seitenaufbau (nur GetClientRect, praktisch kostenlos).
' Bewusst kein Application.OnTime: ein verbliebener geplanter Aufruf wuerde die Datei
' nach dem Schliessen erneut oeffnen.
Public Sub CheckNavSizeThrottled(Optional ByVal shiftOnly As Boolean = False)
    Dim h As Long
    If Abs(Timer - mLastCheck) < 0.25 Then Exit Sub
    mLastCheck = Timer
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook And mApp.WorkDepth = 0 And Not mRendering) Then Exit Sub
    If mNavForm Is Nothing Or Not mWin.HwndAlive(mGridHwnd) Then Exit Sub
    h = NavHeightPx()
    If h = mLastH And NavWidthPx() = mLastW Then Exit Sub
    If shiftOnly And NavWidthPx() = mLastW Then
        ' aus einem Ereignis eines Eintrags: Steuerelemente nicht neu anlegen
        mWin.PlaceWindow mNavHwnd, 0, 0, mLastW, h
        If mNavForm.TryShift(h, h) Then mLastH = h
    Else
        PositionNav
    End If
End Sub
'--- Excel-Zoom (Statusleiste, Strg+Mausrad) -------------------------------------------
' Excel meldet Zoomänderungen nicht. Ein leichter Takt (Application.OnTime, 1 s) prüft Zoom und
' Fenstergröße; er läuft nur, solange diese Datei aktiv ist, und wird beim Schließen und beim
' Wechsel in eine andere Mappe abgemeldet (sonst würde Excel die Datei erneut öffnen).
' Bei geändertem Zoom wird die Seite neu aufgebaut: Spalte A passt dann wieder genau unter die
' Navigation, die ihre Größe behält. Der gewählte Zoom gilt für alle Seiten.
Public Sub ApplyZoom()
    On Error Resume Next
    If gZoom < 10 Or gZoom > 400 Then gZoom = 100
    If ActiveWindow.Zoom <> gZoom Then ActiveWindow.Zoom = gZoom
    mLastZoom = gZoom
End Sub

Public Sub StartWatch()
    If mWatchOn Or mApp.gClosing Then Exit Sub
    On Error Resume Next
    mWatchProc = mRender.MacroRef("mNav.WatchTick")
    mWatchAt = Now + TimeSerial(0, 0, 1)
    Application.OnTime mWatchAt, mWatchProc
    If Err.Number = 0 Then mWatchOn = True
End Sub

Public Sub StopWatch()
    If Not mWatchOn Then Exit Sub
    On Error Resume Next
    Application.OnTime mWatchAt, mWatchProc, , False
    mWatchOn = False
End Sub

Public Sub WatchTick()
    Dim z As Long
    mWatchOn = False
    On Error GoTo done
    If mApp.gClosing Then Exit Sub
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    If mApp.WorkDepth = 0 And Not mRendering Then
        z = ActiveWindow.Zoom
        If mLastZoom > 0 And z <> mLastZoom Then
            gZoom = z
            mLastZoom = z
            RenderCurrent
        Else
            mLastZoom = z
            CheckNavSize
            mRender.ClampScroll
        End If
    End If
done:
    If Not mApp.gClosing Then
        On Error Resume Next
        If ActiveWorkbook Is ThisWorkbook Then StartWatch
    End If
End Sub

Public Sub HideNav()
    On Error Resume Next
    If Not mNavForm Is Nothing Then
        mNavForm.Teardown
        Unload mNavForm
    End If
    Set mNavForm = Nothing
    mNavHwnd = 0
    mLastH = 0
    mLastW = 0
End Sub

' force = Eintraege sicher neu aufbauen (z. B. nach Farbwechsel); sonst nur, wenn sich
' Eintraege, Titel oder Farben geaendert haben - ein Seitenwechsel setzt nur die Markierung um
Public Sub RefreshNav(Optional ByVal force As Boolean = False)
    On Error Resume Next
    If mNavForm Is Nothing Then
        ShowNav
    Else
        If Not mWin.HwndAlive(mNavHwnd) Then
            Set mNavForm = Nothing
            ShowNav
        Else
            mNavForm.Update force
        End If
    End If
End Sub

' Alles, was die Eintraege der Navigation bestimmt (ohne die aktive Seite)
Public Function NavSignature() As String
    Dim it As Variant, s As String
    For Each it In NavItems()
        s = s & it(0) & Chr$(1) & it(1) & Chr$(1) & it(2) & Chr$(2)
    Next
    NavSignature = s & mCfg.SysGet("AppName", APP_TITLE) & Chr$(1) & mCfg.SysGet("Untertitel", "") & Chr$(1) & _
                   CNavBg() & Chr$(1) & CAccentRaw() & Chr$(1) & mBackup.AutoSaveOn() & Chr$(1) & mApp.gReadOnly
End Function

Public Function NavHwnd() As LongPtr
    NavHwnd = mNavHwnd
End Function

' Nach dem Schliessen eines Dialogs: Fokus zurueck auf das Blatt, damit der naechste
' Klick auf eine Schaltflaeche sofort wirkt. Nicht, solange noch ein Dialog offen ist.
Public Sub FocusGrid()
    Dim uf As Object, vis As Boolean
    On Error Resume Next
    For Each uf In VBA.UserForms
        vis = False
        If TypeName(uf) <> "frmNav" Then vis = uf.Visible
        If vis Then Exit Sub
    Next
    mWin.ForegroundExcel
    ' wie beim Klick in die Navigation: erst Navigation, dann Blatt fokussieren
    If mNavHwnd <> 0 Then mWin.FocusWindow mNavHwnd
    If mGridHwnd <> 0 Then mWin.FocusWindow mGridHwnd
End Sub

