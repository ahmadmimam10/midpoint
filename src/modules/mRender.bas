Attribute VB_Name = "mRender"
'==============================================================================
' mRender - Gemeinsamer Seitenaufbau der Oberflaechenblaetter: Seitenrahmen,
' Titel, Statusband, Schaltflaechen (Shapes), Fliesslayout mit Umbruch,
' Filterkacheln, Suchfeld, Leerzustaende und Kontextmenues.
'==============================================================================
Option Explicit

Public Const R_TOP As Long = 1
Public Const R_TITLE As Long = 2
Public Const R_TABS As Long = 3
Public Const R_TOOLS As Long = 4
Public Const R_CHIPS As Long = 5
Public Const R_GAP As Long = 6
Public Const R_HEAD As Long = 7
Public Const R_DATA As Long = 8
Public Const C_CONTENT As Long = 3
Public Const BTN_HT As Single = 24

Private fWs As Worksheet
Private fX As Single, fY As Single, fLeft As Single, fRight As Single, fLines As Long, fRow As Long
Private mBtnSeq As Long
Private mLastW As Single
Private mContentRight As cMap  ' Blattname -> rechte Kante des Inhalts (Punkte)
Private mWCache As Collection   ' gemessene Schaltflaechenbreiten (Art|Text -> Breite)
Public gProf As String          ' Zwischenzeiten (nur Entwicklung, mSelfTest.ProfilePages)
Private mLapT As Double

' Zwischenzeit fuer die Leistungsmessung; kostet im Normalbetrieb nichts Messbares
Public Sub Lap(ByVal label As String)
    Dim t As Double
    t = Timer
    If Len(label) > 0 Then gProf = gProf & label & "=" & Format$((t - mLapT) * 1000, "0") & " "
    mLapT = t
End Sub

'--- Seitenrahmen -----------------------------------------------------------------------
Public Function Sheet(ByVal nm As String) As Worksheet
    Set Sheet = ThisWorkbook.Worksheets(nm)
End Function

' Macht ein Oberflaechenblatt sichtbar und aktiv; andere Oberflaechenblaetter ausblenden
Public Sub ShowSheet(ByVal ws As Worksheet)
    Dim s As Worksheet
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    If Not ActiveSheet Is ws Then ws.Activate
    For Each s In ThisWorkbook.Worksheets
        If s.Name <> ws.Name And s.Visible = xlSheetVisible Then
            If mProtect.IsUiSheet(s.Name) Or s.Name = SH_HINWEIS Then s.Visible = xlSheetHidden
        End If
    Next
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayWorkbookTabs = mPages.DataSheetsVisible()
    End With
    ' gewählter Zoom gilt auf allen Seiten (Excel merkt ihn sonst je Blatt)
    mNav.ApplyZoom
End Sub

Public Sub BeginPage(ByVal ws As Worksheet, Optional ByVal keepContent As Boolean = False)
    Dim i As Long
    Lap ""
    mApp.BeginWork
    ShowSheet ws
    Lap "show"
    ws.Unprotect Password:=""
    ' keine gestrichelten Seitenumbruchlinien (erscheinen sonst nach PDF-Export/Drucken)
    On Error Resume Next
    ws.DisplayPageBreaks = False
    ' Fixierung lösen: Excel lehnt das selten ab (Fehler 1004, z. B. direkt nach dem Schließen
    ' eines Dialogs); LockHorizontal setzt sie am Seitenende ohnehin neu
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitRow = 0: .SplitColumn = 0
        .ScrollRow = 1: .ScrollColumn = 1
    End With
    On Error GoTo 0
    ' am Ende ausgeblendete Spalten rechts vom Inhalt wieder freigeben (EndPage blendet neu aus)
    On Error Resume Next
    ws.ScrollArea = ""
    If ws.Columns(ws.Columns.Count).Hidden Then ws.Cells.EntireColumn.Hidden = False
    On Error GoTo 0
    If keepContent Then
        For i = ws.Shapes.Count To 1 Step -1
            If Left$(ws.Shapes(i).Name, 4) <> "img_" Then ws.Shapes(i).Delete
        Next
    ElseIf ws.Shapes.Count > 0 Then
        ' alle Formen in einem Aufruf entfernen (deutlich schneller als einzeln)
        On Error Resume Next
        ws.DrawingObjects.Delete
        On Error GoTo 0
        For i = ws.Shapes.Count To 1 Step -1
            ws.Shapes(i).Delete
        Next
    End If
    mBtnSeq = 0
    ' Navigationsbreite immer neu setzen (auch bei vorbereiteten Seiten wie Hilfe/Technik)
    SetColWidthPt ws, 1, NavColPt()
    SetColWidthPt ws, 2, GUTTER_PT
    If keepContent Then
        ' Kopfbereich auch bei vorbereiteten Seiten einheitlich (Titel, Register, Werkzeuge)
        ws.Rows(R_TOP).RowHeight = 10
        ws.Rows(R_TITLE).RowHeight = 34
        ws.Rows(R_TABS).RowHeight = 30
        ws.Rows(R_TOOLS).RowHeight = 30
        ws.Rows(R_CHIPS).RowHeight = 26
        ws.Rows(R_GAP).RowHeight = 6
        Exit Sub
    End If
    On Error Resume Next
    ws.Cells.Validation.Delete
    ws.Cells.UnMerge
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells.Font.Name = UI_FONT
    ws.Cells.Font.Size = 10
    ws.Cells.Font.Color = CText()
    ws.Cells.VerticalAlignment = xlCenter
    ws.Cells.RowHeight = 20
    ws.Cells.ColumnWidth = 10
    ws.Cells.Locked = True
    SetColWidthPt ws, 1, NavColPt()
    SetColWidthPt ws, 2, GUTTER_PT
    ws.Rows(R_TOP).RowHeight = 10
    ws.Rows(R_TITLE).RowHeight = 34
    ws.Rows(R_TABS).RowHeight = 30
    ws.Rows(R_TOOLS).RowHeight = 30
    ws.Rows(R_CHIPS).RowHeight = 26
    ws.Rows(R_GAP).RowHeight = 6
    mBtnSeq = 0
    Lap "begin"
End Sub

Public Sub EndPage(ByVal ws As Worksheet, Optional ByVal frozenRows As Long = 0, Optional ByVal frozenCols As Long = 0, Optional ByVal homeCell As String = "")
    On Error Resume Next
    Lap "inhalt"
    ' Bewusst keine Fensterfixierung (keine Fixierungslinien); Parameter bleiben fuer Kompatibilitaet
    If False Then
        With ActiveWindow
            .ScrollRow = 1: .ScrollColumn = 1
            .SplitColumn = frozenCols
            .SplitRow = frozenRows
            .FreezePanes = True
        End With
    End If
    LockHorizontal ws
    ' ohne Ziel: Auswahl unter der Navigation (sonst steht oben links ein kleiner grüner Rahmen)
    If Len(homeCell) > 0 Then ws.Range(homeCell).Select Else ws.Cells(R_TOP, 1).Select
    ws.DisplayPageBreaks = False
    On Error GoTo 0
    mProtect.ProtectUi ws
    mApp.EndWork
    mApp.SyncSavedFlag
    mNav.CheckNavSize
    Lap "end"
End Sub

' Waagerechtes Scrollen bei breiten Seiten (viele Spalten, Zoom): Spalte A unter der Navigation
' bleibt stehen, ihre Fixierungskante liegt unter der Navigation. Bewusst ohne Bildlaufbereich und
' ohne ausgeblendete Spalten: beides ließ die Bildlaufleiste springen (ganz nach links) und stottern.
Public Sub LockHorizontal(ByVal ws As Worksheet)
    Dim rightPt As Double, s As Shape, f As Range, c As Long
    On Error Resume Next
    ' rechte Kante des Inhalts merken (Formen und beschriebene Zellen): weiter als bis dorthin
    ' zieht der Zoom-/Größen-Takt die Ansicht beim Scrollen sanft zurück (mNav.WatchTick)
    rightPt = 0
    For Each s In ws.Shapes
        If s.Visible Then
            If s.Left + s.Width > rightPt And s.Left + s.Width < 30000 Then rightPt = s.Left + s.Width
        End If
    Next
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then
        For c = f.Column To 3 Step -1
            If Not ws.Columns(c).Hidden Then
                If Application.WorksheetFunction.CountA(ws.Columns(c)) > 0 Then
                    If ws.Columns(c).Left + ws.Columns(c).Width > rightPt Then rightPt = ws.Columns(c).Left + ws.Columns(c).Width
                    Exit For
                End If
            End If
        Next
    End If
    If mContentRight Is Nothing Then Set mContentRight = New cMap
    mContentRight.SetVal ws.Name, rightPt + 24
    With ActiveWindow
        .ScrollColumn = 1: .ScrollRow = 1
        .SplitRow = 0
        .SplitColumn = 1
        .FreezePanes = True
    End With
    On Error GoTo 0
End Sub
' Nach zu weitem Scrollen nach rechts: Ansicht so zurücksetzen, dass der Inhalt rechts bündig endet
Public Sub ClampScroll()
    Dim ws As Worksheet, rightPt As Double, viewW As Double, z As Double, c As Long, sc As Long
    On Error GoTo done
    If mContentRight Is Nothing Then Exit Sub
    Set ws = ActiveSheet
    If Not mContentRight.Exists(ws.Name) Then Exit Sub
    rightPt = mContentRight.GetVal(ws.Name)
    z = ActiveWindow.Zoom / 100#
    If z <= 0 Then z = 1
    ' sichtbarer Bereich rechts der festen Spalte A (in Blatt-Punkten)
    viewW = (ActiveWindow.UsableWidth - mWin.VScrollWidthPt()) / z - ws.Columns(1).Width
    sc = ActiveWindow.ScrollColumn
    If sc <= 2 Then Exit Sub
    If ws.Columns(sc).Left - ws.Columns(2).Left <= rightPt - ws.Columns(2).Left - viewW + 30 Then Exit Sub
    ' kleinste Startspalte, bei der der Inhalt noch bis zum rechten Rand reicht
    c = 2
    Do While c < sc And ws.Columns(c + 1).Left - ws.Columns(2).Left < rightPt - ws.Columns(2).Left - viewW
        c = c + 1
    Loop
    If c < sc Then ActiveWindow.ScrollColumn = c
done:
End Sub

Public Sub SetColWidthPt(ByVal ws As Worksheet, ByVal col As Long, ByVal pts As Double)
    Dim i As Long, cur As Double
    If pts <= 0 Then ws.Columns(col).ColumnWidth = 0: Exit Sub
    ws.Columns(col).ColumnWidth = pts / 5.25
    For i = 1 To 3
        cur = ws.Columns(col).Width
        If cur <= 0 Then Exit For
        If Abs(cur - pts) < 0.8 Then Exit For
        ws.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth * pts / cur
    Next
End Sub

Public Function ContentLeft(ByVal ws As Worksheet) As Single
    ContentLeft = ws.Columns(C_CONTENT).Left
End Function

' Rechte Kante der sichtbaren Inhaltsflaeche (Punkte, ab Spalte A)
Public Function LastRenderWidth() As Single
    LastRenderWidth = mLastW
End Function

Public Function PageRight() As Single
    Dim w As Single, z As Double
    On Error Resume Next
    z = ActiveWindow.Zoom / 100#
    If z <= 0 Then z = 1
    mLastW = ActiveWindow.UsableWidth
    ' sichtbare Breite in Blatt-Punkten (ohne die senkrechte Bildlaufleiste, die UsableWidth
    ' einschließt): bei Vergrößerung entsprechend weniger Platz
    w = (ActiveWindow.UsableWidth - mWin.VScrollWidthPt()) / z - 18
    ' kleine Fenster: Schaltflaechen umbrechen statt waagerecht scrollen
    If w < 480 Then w = 480
    PageRight = w
End Function

' Breite von Spalte A (unter der Navigation): so, dass sie bei jedem Excel-Zoom genau so breit
' ist wie die Navigation (die ihre Größe behält, damit sie immer lesbar bleibt)
Public Function NavColPt() As Double
    Dim z As Double
    On Error Resume Next
    z = ActiveWindow.Zoom / 100#
    If z <= 0 Then z = 1
    NavColPt = NAV_WIDTH_PT / z
End Function

Public Function RowTop(ByVal ws As Worksheet, ByVal r As Long) As Single
    RowTop = ws.Rows(r).Top
End Function

Public Function MacroRef(ByVal macro As String) As String
    MacroRef = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macro
End Function
'--- Formen ---------------------------------------------------------------------------------
Public Function TextShape(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, ByVal x As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal h As Single, Optional ByVal size As Single = 10, Optional ByVal color As Long = -1, _
                          Optional ByVal bold As Boolean = False, Optional ByVal align As Long = 1, Optional ByVal autoW As Boolean = False) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddTextbox(1, x, y, w, h)
    s.Name = nm
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoFalse
    s.Placement = xlFreeFloating
    With s.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = IIf(autoW, msoFalse, msoTrue)
        With .TextRange
            .Text = txt
            .Font.Name = IIf(bold, UI_FONT_SEMI, UI_FONT)
            .Font.Size = size
            .Font.Fill.ForeColor.RGB = IIf(color = -1, CText(), color)
            .ParagraphFormat.Alignment = align
        End With
        If autoW Then
            .AutoSize = msoAutoSizeShapeToFitText
            .AutoSize = msoAutoSizeNone
            s.Height = h
            s.Top = y
        End If
    End With
    s.Locked = True
    Set TextShape = s
End Function

' Einzeiliger Text (TextShape mit autoW) passt in maxW: erst Schrift bis minSize verkleinern,
' dann mit "…" kürzen. Nie umbrechen und unten abschneiden (z. B. Kennzahl bei starkem Zoom).
Public Sub FitOneLine(ByVal s As Shape, ByVal maxW As Single, ByVal minSize As Single)
    Dim t As String, sz As Single, h As Single, tp As Single
    h = s.Height: tp = s.Top
    With s.TextFrame2
        sz = .TextRange.Font.Size
        Do While TextW(s) > maxW And sz > minSize
            sz = sz - 1
            If sz < minSize Then sz = minSize
            .TextRange.Font.Size = sz
        Loop
        t = .TextRange.Text
        Do While TextW(s) > maxW And Len(t) > 2
            t = RTrim$(Left$(t, Len(t) - 1))
            .TextRange.Text = t & ChrW(&H2026)
        Loop
    End With
    s.Width = maxW: s.Height = h: s.Top = tp
End Sub

Private Function TextW(ByVal s As Shape) As Single
    With s.TextFrame2
        .AutoSize = msoAutoSizeShapeToFitText
        TextW = s.Width
        .AutoSize = msoAutoSizeNone
    End With
End Function

' Setzt das erste Zeichen eines Shape-Textes in die Symbolschrift
Public Sub IconFirstChar(ByVal s As Shape, Optional ByVal size As Single = 0)
    With s.TextFrame2.TextRange.Characters(1, 1).Font
        .Name = ICON_FONT
        If size > 0 Then .Size = size
    End With
End Sub

' Schaltflaeche; kind: primary | secondary | ghost | tab | tabActive | chip | danger | link
Public Function Btn(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, _
                    ByVal macro As String, Optional ByVal kind As String = "secondary", Optional ByVal icon As String = "", _
                    Optional ByVal h As Single = BTN_HT, Optional ByVal minW As Single = 0, Optional ByVal fixedW As Single = 0) As Shape
    Dim s As Shape, txt As String, w As Single, fg As Long
    If Len(nm) = 0 Then mBtnSeq = mBtnSeq + 1: nm = "btn" & mBtnSeq
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, 120, h)
    s.Name = nm
    s.Placement = xlFreeFloating
    On Error Resume Next
    s.Adjustments.Item(1) = 0.22
    On Error GoTo 0
    If Len(icon) > 0 Then txt = icon & "  " & caption Else txt = caption
    Select Case kind
        Case "primary"
            s.Fill.ForeColor.RGB = CAccentFill(): s.Line.Visible = msoFalse: fg = COnAccent()
        Case "danger"
            s.Fill.ForeColor.RGB = CWhite(): s.Line.ForeColor.RGB = CDanger(): s.Line.Weight = 0.75: fg = CDanger()
        Case "tabActive"
            s.Fill.ForeColor.RGB = CSurface2(): s.Line.Visible = msoFalse: fg = CText()
        Case "tab", "ghost"
            s.Fill.ForeColor.RGB = CWhite(): s.Line.Visible = msoFalse: fg = CMuted()
        Case "chip"
            s.Fill.ForeColor.RGB = CAccentSoft(): s.Line.Visible = msoFalse: fg = Blend(CAccent(), RGB(0, 0, 0), 0.25)
        Case "link"
            s.Fill.Visible = msoFalse: s.Line.Visible = msoFalse: fg = CAccent()
        Case Else
            s.Fill.ForeColor.RGB = CWhite(): s.Line.ForeColor.RGB = CBorderStrong(): s.Line.Weight = 0.75: fg = CText()
    End Select
    If kind = "primary" Or kind = "secondary" Or kind = "danger" Then
        With s.Shadow
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0.9
            .OffsetX = 0: .OffsetY = 1
            .Blur = 2
            .Size = 100
        End With
    End If
    With s.TextFrame2
        .MarginLeft = 9: .MarginRight = 9: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        With .TextRange
            .Text = txt
            .Font.Name = IIf(kind = "tabActive" Or kind = "primary", UI_FONT_SEMI, UI_FONT)
            .Font.Size = IIf(kind = "chip", 9, 9.5)
            .Font.Fill.ForeColor.RGB = fg
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
        If Len(icon) > 0 Then
            With .TextRange.Characters(1, 1).Font
                .Name = ICON_FONT
                .Size = IIf(kind = "chip", 8.5, 10)
            End With
        End If
        If fixedW > 0 Then
            s.Width = fixedW
        Else
            ' Textbreite je Beschriftung/Art nur einmal je Sitzung messen (AutoSize ist langsam)
            w = CachedWidth(kind & "|" & txt)
            If w <= 0 Then
                .AutoSize = msoAutoSizeShapeToFitText
                w = s.Width
                .AutoSize = msoAutoSizeNone
                StoreWidth kind & "|" & txt, w
            End If
            If w < minW Then w = minW
            s.Width = w + 2
        End If
    End With
    s.Height = h
    s.Top = y
    s.Left = x
    If Len(macro) > 0 Then s.OnAction = MacroRef(macro)
    s.Locked = True
    Set Btn = s
End Function

Private Function CachedWidth(ByVal key As String) As Single
    On Error Resume Next
    If mWCache Is Nothing Then Exit Function
    CachedWidth = mWCache(key)
End Function

Private Sub StoreWidth(ByVal key As String, ByVal w As Single)
    On Error Resume Next
    If mWCache Is Nothing Then Set mWCache = New Collection
    If mWCache.Count > 2000 Then Set mWCache = New Collection
    mWCache.Add w, key
End Sub

'--- Fliesslayout mit Zeilenumbruch -----------------------------------------------------------
Public Sub FlowBegin(ByVal ws As Worksheet, ByVal row As Long, ByVal leftX As Single, ByVal rightX As Single)
    Set fWs = ws
    fRow = row
    fLeft = leftX: fRight = rightX
    fX = leftX
    fY = ws.Rows(row).Top + (ws.Rows(row).Height - BTN_HT) / 2
    If fY < ws.Rows(row).Top Then fY = ws.Rows(row).Top + 2
    fLines = 1
End Sub

Public Function FlowBtn(ByVal nm As String, ByVal caption As String, ByVal macro As String, Optional ByVal kind As String = "secondary", _
                        Optional ByVal icon As String = "", Optional ByVal gapAfter As Single = 6) As Shape
    Dim s As Shape
    Set s = Btn(fWs, nm, caption, fX, fY, macro, kind, icon)
    If fX + s.Width > fRight And fX > fLeft Then
        fX = fLeft
        fY = fY + BTN_HT + 6
        fLines = fLines + 1
        s.Left = fX: s.Top = fY
    End If
    fX = fX + s.Width + gapAfter
    Set FlowBtn = s
End Function

Public Function FlowText(ByVal nm As String, ByVal txt As String, Optional ByVal color As Long = -1, Optional ByVal size As Single = 9.5, _
                         Optional ByVal bold As Boolean = False, Optional ByVal gapAfter As Single = 10) As Shape
    Dim s As Shape
    Set s = TextShape(fWs, nm, txt, fX, fY, 300, BTN_HT, size, color, bold, 1, True)
    If fX + s.Width > fRight And fX > fLeft Then
        fX = fLeft
        fY = fY + BTN_HT + 6
        fLines = fLines + 1
        s.Left = fX: s.Top = fY
    End If
    fX = fX + s.Width + gapAfter
    Set FlowText = s
End Function

Public Sub FlowGap(ByVal w As Single)
    fX = fX + w
End Sub

Public Function FlowX() As Single
    FlowX = fX
End Function

Public Sub FlowEnd()
    Dim needH As Single
    needH = fLines * (BTN_HT + 6) + 2
    If fWs.Rows(fRow).RowHeight < needH Then fWs.Rows(fRow).RowHeight = needH
End Sub

'--- Seitenkopf ----------------------------------------------------------------------------
Public Sub DrawTitle(ByVal ws As Worksheet, ByVal icon As String, ByVal title As String, Optional ByVal subtitle As String = "")
    Dim x As Single, y As Single, s As Shape
    x = ContentLeft(ws)
    y = ws.Rows(R_TITLE).Top
    If Len(icon) > 0 Then
        Set s = TextShape(ws, "t_icon", icon, x, y, 28, 34, 17, CAccent(), False, 1)
        s.TextFrame2.TextRange.Font.Name = ICON_FONT
        x = x + 30
    End If
    Set s = TextShape(ws, "t_title", title, x, y, 400, 34, 18, CText(), True, 1, True)
    DrawStatus ws
    If Len(subtitle) > 0 Then
        ' Untertitel nur im freien Raum zwischen Titel und Status/Logo
        Dim subW As Single, lim As Single, nm As Variant, o As Shape
        lim = PageRight()
        For Each nm In Array("st_text", "st_dirty", "st_save", "img_logo")
            Set o = Nothing
            On Error Resume Next
            Set o = ws.Shapes(CStr(nm))
            On Error GoTo 0
            If Not o Is Nothing Then
                If o.Visible And o.Left < lim Then lim = o.Left
            End If
        Next
        subW = lim - (s.Left + s.Width + 14) - 12
        If subW > 80 Then
            ' immer einzeilig: zu lang -> an einer Wortgrenze mit "…" kürzen (nie halb abgeschnitten)
            Dim sub_ As Shape, txt As String, p As Long
            txt = subtitle
            Set sub_ = TextShape(ws, "t_sub", txt, s.Left + s.Width + 14, y + 4, subW, 30, 9.5, CMuted(), False, 1, True)
            Do While sub_.Width > subW
                p = InStrRev(RTrim$(Left$(txt, Len(txt) - 1)), " ")
                If p < 8 Then sub_.Delete: Set sub_ = Nothing: Exit Do
                txt = RTrim$(Left$(txt, p - 1))
                Do While Len(txt) > 0 And InStr(".,;:-" & ChrW(&HB7) & ChrW(&H2013), Right$(txt, 1)) > 0
                    txt = RTrim$(Left$(txt, Len(txt) - 1))
                Loop
                txt = txt & ChrW(&H2026)
                sub_.TextFrame2.AutoSize = msoAutoSizeNone
                sub_.TextFrame2.TextRange.Text = txt
                sub_.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                sub_.TextFrame2.AutoSize = msoAutoSizeNone
                sub_.Height = 30: sub_.Top = y + 4
                txt = Left$(txt, Len(txt) - 1)
            Loop
        End If
    End If
End Sub

Public Sub DrawStatus(ByVal ws As Worksheet)
    Dim txt As String, s As Shape, rightX As Single, y As Single, b As Shape, i As Long, t As Shape, titleRight As Single, w As Single
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "st_" Then ws.Shapes(i).Delete
    Next
    rightX = PageRight()
    y = ws.Rows(R_TITLE).Top
    txt = mBackup.StatusText()
    ' nie über den Seitentitel: Platz rechts davon ermitteln (bei schmalen Fenstern bzw. Zoom)
    titleRight = ContentLeft(ws) + 200
    Set t = Nothing
    Set t = ws.Shapes("t_title")
    If Not t Is Nothing Then titleRight = t.Left + t.Width
    On Error GoTo 0
    If mApp.gDataDirty And Not mApp.gReadOnly Then
        Set b = Btn(ws, "st_save", "Speichern", rightX - 90, y + 5, "mCmd.Cmd_Save", "primary", ChrW(&HE74E))
        b.Left = Application.Max(rightX - b.Width, titleRight + 16)
        rightX = b.Left - 8
        If rightX - 150 >= titleRight + 16 Then
            Set s = TextShape(ws, "st_dirty", ChrW(&H25CF) & " Ungespeichert", rightX - 110, y, 110, 34, 8.5, CWarn(), True, 3)
            rightX = rightX - 116
        End If
    End If
    ' Statuszeile: höchstens so breit wie der freie Platz, bricht sonst auf zwei Zeilen um
    w = Application.Min(330, rightX - (titleRight + 16))
    If w >= 120 Then
        Set s = TextShape(ws, "st_text", txt, rightX - w, y, w, 34, IIf(w < 220, 7.5, 8.5), CMuted(), False, 3)
    End If
    ' Logo (Hilfe, Technik) nur, wenn rechts daneben genug Platz bleibt
    PlaceTitleLogo ws, titleRight, IIf(w >= 120, rightX - w, rightX)
End Sub

' MidPoint-Logo zwischen Seitentitel und Status (nur Seiten mit Logo, nur wenn es Platz hat)
Private Sub PlaceTitleLogo(ByVal ws As Worksheet, ByVal leftLimit As Single, ByVal rightLimit As Single)
    Dim s As Shape
    On Error Resume Next
    Set s = ws.Shapes("img_logo")
    On Error GoTo 0
    If s Is Nothing Then Exit Sub
    s.Visible = msoFalse
    If ws.Name <> SH_HELP And ws.Name <> SH_TECH Then Exit Sub
    If rightLimit - leftLimit < s.Width + 48 Then Exit Sub
    s.Left = rightLimit - s.Width - 24
    s.Top = ws.Rows(R_TITLE).Top + (34 - s.Height) / 2
    s.Visible = msoTrue
End Sub

' Aktualisiert nur das Statusband der sichtbaren Seite
Public Sub RefreshStatus()
    Dim ws As Worksheet
    On Error Resume Next
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    Set ws = ActiveSheet
    If Not mProtect.IsUiSheet(ws.Name) Then Exit Sub
    mApp.BeginWork
    ws.Unprotect Password:=""
    DrawStatus ws
    mProtect.ProtectUi ws
    mApp.EndWork
End Sub

' Suchfeld: verbundene, entsperrte Zelle mit Kontur, Lupe und Schaltflaeche "Suchen"
Public Function SearchBox(ByVal ws As Worksheet, ByVal row As Long, ByVal value As String, ByVal macro As String, _
                          Optional ByVal minW As Single = 220) As Range
    Dim c As Long, w As Single, rng As Range, s As Shape, y As Single
    c = C_CONTENT
    w = 0
    Do
        w = w + ws.Columns(c).Width
        If w >= minW Then Exit Do
        c = c + 1
    Loop While c < 60
    Set rng = ws.Range(ws.Cells(row, C_CONTENT), ws.Cells(row, c))
    rng.Merge
    With rng
        .Value = value
        .NumberFormat = "@"
        .Locked = False
        .Interior.Color = CWhite()
        .IndentLevel = 2
        .HorizontalAlignment = xlLeft
        .Font.Size = 10
        .Font.Color = CText()
        .Borders.LineStyle = xlContinuous
        .Borders.Color = CBorderStrong()
        .Borders.Weight = xlThin
    End With
    ws.Rows(row).RowHeight = Application.Max(ws.Rows(row).RowHeight, 30)
    y = ws.Rows(row).Top
    Set s = TextShape(ws, "srch_icon", ChrW(&HE721), rng.Left + 5, y, 16, ws.Rows(row).Height, 10, CMuted(), False, 1)
    s.TextFrame2.TextRange.Font.Name = ICON_FONT
    s.OnAction = MacroRef(macro)
    Set SearchBox = rng
End Function

' Leerzustand mit Erklaerung und optionaler Aktion
Public Sub EmptyState(ByVal ws As Worksheet, ByVal row As Long, ByVal title As String, ByVal text As String, _
                      Optional ByVal btnCaption As String = "", Optional ByVal macro As String = "", Optional ByVal icon As String = "")
    Dim x As Single, y As Single, s As Shape
    x = ContentLeft(ws) + 10
    y = ws.Rows(row).Top + 16
    If Len(icon) > 0 Then
        Set s = TextShape(ws, "es_icon", icon, x, y, 40, 40, 24, CMuted(), False, 1)
        s.TextFrame2.TextRange.Font.Name = ICON_FONT
        y = y + 44
    End If
    TextShape ws, "es_title", title, x, y, 560, 24, 13, CText(), True, 1
    y = y + 26
    TextShape ws, "es_text", text, x, y, 560, 40, 10, CMuted(), False, 1
    y = y + 46
    If Len(btnCaption) > 0 Then Btn ws, "es_btn", btnCaption, x, y, macro, "primary", ChrW(&HE710)
End Sub

'--- Kontextmenues (temporaer, werden nicht dauerhaft gespeichert) ----------------------------
' items: Array("Beschriftung|Makro|Parameter|aktiv(1/0)|Gruppe(1/0)", ...)
Public Sub Popup(ByVal items As Variant)
    Dim cb As CommandBar, it As Variant, p As Variant, c As CommandBarButton
    On Error Resume Next
    Application.CommandBars("DBK_Menu").Delete
    On Error GoTo 0
    Set cb = Application.CommandBars.Add(Name:="DBK_Menu", Position:=msoBarPopup, MenuBar:=False, Temporary:=True)
    For Each it In items
        p = Split(CStr(it), "|")
        Set c = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
        c.caption = p(0)
        If UBound(p) >= 1 Then If Len(p(1)) > 0 Then c.OnAction = MacroRef(CStr(p(1)))
        If UBound(p) >= 2 Then c.Parameter = CStr(p(2))
        If UBound(p) >= 3 Then c.Enabled = (p(3) <> "0")
        If UBound(p) >= 4 Then c.BeginGroup = (p(4) = "1")
        c.Style = msoButtonCaption
    Next
    cb.ShowPopup
End Sub

Public Function MenuParam() As String
    On Error Resume Next
    MenuParam = Application.CommandBars.ActionControl.Parameter
End Function

Public Function CallerName() As String
    On Error Resume Next
    CallerName = CStr(Application.Caller)
End Function

' Zellbereich mit duennen Linien (Tabellenraster)
Public Sub GridBorders(ByVal rng As Range, Optional ByVal color As Long = -1)
    If color = -1 Then color = CBorder()
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = color
    End With
End Sub
