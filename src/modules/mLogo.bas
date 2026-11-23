Attribute VB_Name = "mLogo"
'==============================================================================
' mLogo - Programmlogo (MidPoint) nur in der Hauptdatei: Navigation oben links,
' Hinweisblatt und Technische Dokumentation. Erzeugte Dateien (Erfassungs-
' dateien, Exporte) erhalten bewusst kein Logo.
' Die Bilder liegen als Base64-Text auf dem ausgeblendeten Blatt "_Ressourcen"
' (beim Build eingebettet) und werden bei Bedarf in den Temp-Ordner geschrieben.
' Für die Navigation rendert Excel das Symbol auf der Navigationsfarbe (GIF),
' weil Formular-Bilder kein PNG mit Transparenz kennen. Alles offline.
'==============================================================================
Option Explicit

Private mNavPic As Object        ' StdPicture des Navigationssymbols
Private mNavPicFor As Long       ' ... für diese Hintergrundfarbe

Public Function HasLogo() As Boolean
    HasLogo = mApp.SheetExists(SH_RES)
End Function

' Ressource als Datei bereitstellen (Temp-Ordner); "" wenn nicht vorhanden
Public Function ResFile(ByVal nm As String) As String
    Dim ws As Worksheet, r As Long, c As Long, b64 As String, path As String, bytes() As Byte, fn As Integer, dirPath As String
    On Error GoTo failed
    If Not HasLogo() Then Exit Function
    dirPath = Environ$("TEMP") & "\MidPoint_" & Replace(PROG_VERSION, ".", "_")
    If Len(Dir$(dirPath, vbDirectory)) = 0 Then MkDir dirPath
    path = dirPath & "\" & nm
    If Len(Dir$(path)) > 0 Then ResFile = path: Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_RES)
    For r = 1 To 50
        If CStr(ws.Cells(r, 1).Value) = nm Then
            c = 2
            Do While Len(CStr(ws.Cells(r, c).Value)) > 0
                b64 = b64 & CStr(ws.Cells(r, c).Value)
                c = c + 1
            Loop
            Exit For
        End If
    Next
    If Len(b64) = 0 Then Exit Function
    bytes = Base64Decode(b64)
    fn = FreeFile
    Open path For Binary Access Write As #fn
    Put #fn, , bytes
    Close #fn
    ResFile = path
    Exit Function
failed:
    On Error Resume Next
    If fn > 0 Then Close #fn
    ResFile = ""
End Function

Private Function Base64Decode(ByVal s As String) As Byte()
    Dim doc As Object, el As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    Set el = doc.createElement("b64")
    el.DataType = "bin.base64"
    el.Text = s
    Base64Decode = el.nodeTypedValue
End Function

' Bild des Symbols für die Navigation (weiß auf dunkler, farbig auf heller Navigation)
Public Function NavPicture(ByVal bg As Long) As Object
    Dim p As String
    On Error GoTo failed
    If Not mNavPic Is Nothing And mNavPicFor = bg Then Set NavPicture = mNavPic: Exit Function
    p = RenderOnColor(IIf(IsLightColor(bg), "logo_symbol.png", "logo_symbol_weiss.png"), bg, 48, 46, "nav_" & Hex$(bg) & ".gif")
    If Len(p) = 0 Then Exit Function
    Set mNavPic = LoadPicture(p)
    mNavPicFor = bg
    Set NavPicture = mNavPic
    Exit Function
failed:
    Set NavPicture = Nothing
End Function

' PNG mit Transparenz auf einfarbigem Grund als GIF ausgeben (über ein Excel-Diagramm)
Private Function RenderOnColor(ByVal resName As String, ByVal bg As Long, ByVal wPt As Single, ByVal hPt As Single, ByVal outName As String) As String
    Dim src As String, outPath As String, ws As Worksheet, co As ChartObject, prevUpd As Boolean
    On Error GoTo failed
    src = ResFile(resName)
    If Len(src) = 0 Then Exit Function
    outPath = Left$(src, InStrRev(src, "\")) & outName
    If Len(Dir$(outPath)) > 0 Then RenderOnColor = outPath: Exit Function
    prevUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets(SH_RES)
    Set co = ws.ChartObjects.Add(0, 0, wPt, hPt)
    With co.Chart
        .ChartArea.Format.Fill.Visible = msoTrue
        .ChartArea.Format.Fill.Solid
        .ChartArea.Format.Fill.ForeColor.RGB = bg
        .ChartArea.Format.Line.Visible = msoFalse
        .Shapes.AddPicture src, msoFalse, msoTrue, 0, 0, wPt, hPt
        .Export outPath, "GIF"
    End With
    co.Delete
    Application.ScreenUpdating = prevUpd
    RenderOnColor = outPath
    Exit Function
failed:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    Application.ScreenUpdating = prevUpd
    RenderOnColor = ""
End Function

' Logo als Form auf einem Blatt (Hinweisblatt, Technische Dokumentation); Name beginnt mit "img_"
Public Sub PlaceSheetLogo(ByVal ws As Worksheet, ByVal x As Single, ByVal y As Single, ByVal h As Single, Optional ByVal resName As String = "logo_horizontal.png")
    Dim p As String, s As Shape, wasProt As Boolean
    On Error GoTo failed
    p = ResFile(resName)
    If Len(p) = 0 Then Exit Sub
    wasProt = ws.ProtectContents
    If wasProt Then ws.Unprotect Password:=""
    On Error Resume Next
    ws.Shapes("img_logo").Delete
    On Error GoTo failed
    Set s = ws.Shapes.AddPicture(p, msoFalse, msoTrue, x, y, -1, -1)
    s.LockAspectRatio = msoTrue
    s.Height = h
    s.Name = "img_logo"
    s.Placement = xlFreeFloating
    s.AlternativeText = "MidPoint"
    If wasProt Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True
    Exit Sub
failed:
End Sub
