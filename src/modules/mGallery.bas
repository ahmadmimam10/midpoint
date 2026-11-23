Attribute VB_Name = "mGallery"
'==============================================================================
' mGallery - "Welches Diagramm passt?": Vorschaubilder aller Darstellungen mit
' Beispieldaten. Die Bilder entstehen offline mit Excel selbst (Diagramm bzw.
' Formen in einer unsichtbaren Hilfsmappe, Export als GIF in den Temp-Ordner)
' und werden je Programmversion und Farbstil zwischengespeichert.
'==============================================================================
Option Explicit

' Wofür eignet sich die Darstellung? (Galerie, eine Zeile)
Public Function TypeUse(ByVal typ As String) As String
    Select Case typ
        Case "Kachel": TypeUse = "Eine Kennzahl auf einen Blick: Wie viele? Wie hoch ist die Summe?"
        Case "Balken": TypeUse = "Gruppen vergleichen, auch mit langen Namen (waagerecht)."
        Case "Säulen": TypeUse = "Wenige Gruppen vergleichen (senkrecht)."
        Case "Zeitverlauf": TypeUse = "Entwicklung über Monate, Quartale oder Jahre."
        Case "Fläche": TypeUse = "Entwicklung über die Zeit mit Betonung der Menge."
        Case "Anteil": TypeUse = "Anteile am Ganzen als Ring, bis etwa 7 Gruppen."
        Case "Kreis": TypeUse = "Anteile am Ganzen als Kreis, bis etwa 7 Gruppen."
        Case "Netz": TypeUse = "Profil über mehrere Merkmale (mindestens 3 Gruppen)."
        Case "Trichter": TypeUse = "Stufen eines Ablaufs, die kleiner werden."
        Case "Baumdiagramm": TypeUse = "Viele Gruppen als Flächen: Größenverhältnisse erkennen."
        Case "Pareto": TypeUse = "Welche wenigen Gruppen machen den Großteil aus (80/20)?"
        Case "Rangliste": TypeUse = "Reihenfolge mit Namen und Wert, gut lesbar."
        Case "Sankey": TypeUse = "Wie verteilen sich Einträge von einer Gruppierung auf eine andere?"
        Case "Tabelle": TypeUse = "Genaue Werte je Gruppe als Liste."
    End Select
End Function

' Pfad des Vorschaubildes (legt bei Bedarf alle Bilder an); leer, wenn das nicht möglich war
Public Function PreviewPath(ByVal typ As String) As String
    Dim folder As String
    folder = PreviewFolder()
    If Len(Dir$(folder & "\" & FileKey(typ) & ".gif")) = 0 Then BuildAll folder
    If Len(Dir$(folder & "\" & FileKey(typ) & ".gif")) > 0 Then PreviewPath = folder & "\" & FileKey(typ) & ".gif"
End Function

Private Function PreviewFolder() As String
    Dim dirPath As String, stl As String
    stl = mDocStyle.DefaultStyle()
    If stl = "Akzentfarbe" Then stl = stl & "_" & CStr(CAccentRaw())
    dirPath = Environ$("TEMP") & "\MidPoint_Diagramme_" & Replace(PROG_VERSION, ".", "_") & "_" & FileKey(stl)
    If Len(Dir$(dirPath, vbDirectory)) = 0 Then
        On Error Resume Next
        MkDir dirPath
        On Error GoTo 0
    End If
    PreviewFolder = dirPath
End Function

Private Function FileKey(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then r = r & ch Else r = r & "_"
    Next
    FileKey = r
End Function

' Alle Vorschaubilder in einer unsichtbaren Hilfsmappe zeichnen und exportieren
Private Sub BuildAll(ByVal folder As String)
    Dim wb As Workbook, ws As Worksheet, typ As Variant, co As ChartObject, res As cMap, pal As String, i As Long
    Dim prevUpd As Boolean, prevEv As Boolean, prevAlerts As Boolean, actWb As Workbook, actWs As Object
    Dim cats As Variant, vals As Variant, mon As Variant, mvals As Variant
    prevUpd = Application.ScreenUpdating: prevEv = Application.EnableEvents: prevAlerts = Application.DisplayAlerts
    Set actWb = ActiveWorkbook: Set actWs = ActiveSheet
    On Error GoTo done
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    pal = mDocStyle.DefaultStyle()
    cats = Array("Nord", "Süd", "West", "Ost", "Mitte", "Rand")
    vals = Array(42, 33, 26, 18, 12, 7)
    mon = Array("Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug")
    mvals = Array(12, 15, 13, 19, 22, 20, 26, 29)
    For i = 0 To 5
        ws.Cells(i + 1, 1).Value = cats(i): ws.Cells(i + 1, 2).Value = vals(i)
    Next
    For i = 0 To 7
        ws.Cells(i + 1, 5).Value = mon(i): ws.Cells(i + 1, 6).Value = mvals(i)
    Next
    Set res = New cMap
    res.SetVal "measure", "ANZAHL"
    res.SetVal "n", 6
    For Each typ In mDash.ElemTypes()
        Select Case CStr(typ)
            Case "Kachel", "Rangliste", "Tabelle", "Sankey"
                Set co = ws.ChartObjects.Add(300, 10, 240, 150)
                co.Chart.ChartArea.Format.Line.Visible = msoFalse
                DrawShapePreview co.Chart, CStr(typ), pal, cats, vals
            Case "Zeitverlauf", "Fläche"
                Set co = mDash.NewChart(ws, CStr(typ), 300, 10, 240, 150, ws.Range("E1:E8"), ws.Range("F1:F8"), res, pal)
            Case Else
                Set co = mDash.NewChart(ws, CStr(typ), 300, 10, 240, 150, ws.Range("A1:A6"), ws.Range("B1:B6"), res, pal)
        End Select
        Tidy co.Chart, CStr(typ)
        co.Chart.Export folder & "\" & FileKey(CStr(typ)) & ".gif", "GIF"
        co.Delete
        ws.Range("C1:C10").ClearContents      ' Pareto schreibt kumulierte Werte daneben
    Next
done:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error Resume Next
    If Not actWb Is Nothing Then actWb.Activate
    If Not actWs Is Nothing Then actWs.Activate
    Application.DisplayAlerts = prevAlerts
    Application.EnableEvents = prevEv
    Application.ScreenUpdating = prevUpd
End Sub

' Vorschau ruhig halten: ohne Achsentitel und Legende, kleine Schrift
Private Sub Tidy(ByVal c As Chart, ByVal typ As String)
    On Error Resume Next
    c.HasTitle = False
    c.HasLegend = False
    c.ChartArea.Format.Fill.Visible = msoTrue
    c.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    c.ChartArea.Format.Line.Visible = msoFalse
    c.ChartArea.Font.Size = 7
    c.ChartArea.Font.Name = UI_FONT
End Sub

' Kachel, Rangliste, Tabelle und Sankey bestehen im Dashboard aus Formen: hier vereinfacht nachgebildet
Private Sub DrawShapePreview(ByVal c As Chart, ByVal typ As String, ByVal pal As String, ByVal cats As Variant, ByVal vals As Variant)
    Dim s As Shape, i As Long, y As Single, col As Long, ff As FreeformBuilder
    Dim ya As Variant, yb As Variant
    col = mDocStyle.DocInk(mDash.PaletteColor(pal, 0))
    Select Case typ
        Case "Kachel"
            AddText c, "Offene Aufgaben", 18, 28, 200, 16, 9, RGB(55, 65, 81), True
            AddText c, "128", 18, 48, 200, 46, 30, col, True
            AddText c, "Anzahl Datensätze", 18, 98, 200, 14, 8, RGB(107, 114, 128), False
        Case "Rangliste"
            For i = 0 To 4
                y = 14 + i * 26
                AddText c, CStr(cats(i)), 14, y, 120, 12, 8, RGB(55, 65, 81), False
                AddText c, CStr(vals(i)), 170, y, 52, 12, 8, RGB(55, 65, 81), True, True
                Set s = c.Shapes.AddShape(msoShapeRoundedRectangle, 14, y + 14, 208, 6)
                s.Line.Visible = msoFalse: s.Fill.ForeColor.RGB = RGB(243, 244, 246)
                Set s = c.Shapes.AddShape(msoShapeRoundedRectangle, 14, y + 14, 208 * vals(i) / vals(0), 6)
                s.Line.Visible = msoFalse: s.Fill.ForeColor.RGB = mDash.PaletteColor(pal, IIf(pal = "Akzentfarbe", 0, i))
            Next
        Case "Tabelle"
            For i = 0 To 5
                y = 12 + i * 20
                AddText c, CStr(cats(i)), 16, y, 140, 14, 8.5, RGB(31, 41, 55), False
                AddText c, CStr(vals(i)), 160, y, 62, 14, 8.5, RGB(31, 41, 55), True, True
                Set s = c.Shapes.AddLine(16, y + 17, 222, y + 17)
                s.Line.ForeColor.RGB = RGB(229, 231, 235): s.Line.Weight = 0.5
            Next
        Case "Sankey"
            ' drei Gruppen links, zwei rechts
            ya = Array(14, 62, 102): yb = Array(20, 88)
            Dim fa As Variant, fb As Variant, ft As Variant, offA(0 To 2) As Single, offB(0 To 1) As Single
            fa = Array(0, 0, 1, 1, 2): fb = Array(0, 1, 0, 1, 1): ft = Array(26, 16, 14, 18, 22)
            For i = 0 To 4
                Set ff = c.Shapes.BuildFreeform(msoEditingAuto, 70, ya(fa(i)) + offA(fa(i)))
                ff.AddNodes msoSegmentCurve, msoEditingCorner, 120, ya(fa(i)) + offA(fa(i)), 120, yb(fb(i)) + offB(fb(i)), 168, yb(fb(i)) + offB(fb(i))
                ff.AddNodes msoSegmentLine, msoEditingAuto, 168, yb(fb(i)) + offB(fb(i)) + ft(i)
                ff.AddNodes msoSegmentCurve, msoEditingCorner, 120, yb(fb(i)) + offB(fb(i)) + ft(i), 120, ya(fa(i)) + offA(fa(i)) + ft(i), 70, ya(fa(i)) + offA(fa(i)) + ft(i)
                ff.AddNodes msoSegmentLine, msoEditingAuto, 70, ya(fa(i)) + offA(fa(i))
                Set s = ff.ConvertToShape
                s.Line.Visible = msoFalse
                s.Fill.ForeColor.RGB = mDash.PaletteColor(pal, fa(i))
                s.Fill.Transparency = IIf(IsLightColor(mDash.PaletteColor(pal, fa(i))), 0.25, 0.5)
                offA(fa(i)) = offA(fa(i)) + ft(i): offB(fb(i)) = offB(fb(i)) + ft(i)
            Next
            For i = 0 To 2
                Set s = c.Shapes.AddShape(msoShapeRectangle, 62, ya(i), 8, offA(i))
                s.Line.Visible = msoFalse: s.Fill.ForeColor.RGB = mDash.PaletteColor(pal, i)
                AddText c, Choose(i + 1, "Offen", "In Arbeit", "Erledigt"), 4, ya(i) + offA(i) / 2 - 6, 56, 12, 7.5, RGB(55, 65, 81), False, True
            Next
            For i = 0 To 1
                Set s = c.Shapes.AddShape(msoShapeRectangle, 168, yb(i), 8, offB(i))
                s.Line.Visible = msoFalse: s.Fill.ForeColor.RGB = RGB(100, 116, 139)
                AddText c, Choose(i + 1, "Hoch", "Normal"), 180, yb(i) + offB(i) / 2 - 6, 56, 12, 7.5, RGB(55, 65, 81), False
            Next
    End Select
End Sub

Private Sub AddText(ByVal c As Chart, ByVal txt As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                    ByVal size As Single, ByVal col As Long, ByVal bold As Boolean, Optional ByVal alignRight As Boolean = False)
    Dim s As Shape
    Set s = c.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, w, h)
    s.Line.Visible = msoFalse
    s.Fill.Visible = msoFalse
    With s.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        .TextRange.Text = txt
        .TextRange.Font.Size = size
        .TextRange.Font.Name = UI_FONT
        .TextRange.Font.Bold = IIf(bold, msoTrue, msoFalse)
        .TextRange.Font.Fill.ForeColor.RGB = col
        .TextRange.ParagraphFormat.Alignment = IIf(alignRight, msoAlignRight, msoAlignLeft)
    End With
End Sub
