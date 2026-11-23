Attribute VB_Name = "mDocStyle"
'==============================================================================
' mDocStyle - Gestaltung erzeugter Dateien (Erfassungsdateien, Excel-Export):
' Farben für Kopf und Akzente, Schriftarten und Tabellenstile. Zurückhaltend:
' eine Akzentfarbe, helle Kopfzeile, feine Linien.
'==============================================================================
Option Explicit

Private mAppFont As String           ' zwischengespeicherte Schrift der Oberfläche

'--- Dokumentstil (Einrichten): Vorlagen aus Navigation, Akzent, Farbstil und Schrift ----------
Public Function DesignNames() As Variant
    DesignNames = Array("Klassisch Blau", "Petrol & Sand", "Wald", "Schiefer", "Bordeaux elegant", "Erde warm", _
                        "Nordisch hell", "Flieder hell", "Salbei hell", "Kontrast klar", "Mainz")
End Function

' Array(Navigationsfarbe, Akzentfarbe, Farbstil, Schrift, Beschreibung)
Public Function DesignSpec(ByVal nm As String) As Variant
    Select Case nm
        Case "Klassisch Blau": DesignSpec = Array("Nachtblau", "Blau", "Ozean", "Segoe UI", "Sachlich und vertraut, für Verwaltung und Berichte.")
        Case "Petrol & Sand": DesignSpec = Array("Petrol", "Petrol", "Salbei & Sand (pastell)", "Calibri", "Ruhiges Petrol mit warmen Sandtönen.")
        Case "Wald": DesignSpec = Array("Waldgrün", "Grün", "Wald", "Aptos", "Gedeckte Grüntöne, modern und klar.")
        Case "Schiefer": DesignSpec = Array("Schiefer", "Graphit", "Schiefer", "Segoe UI", "Neutral und zurückhaltend: Inhalte stehen im Vordergrund.")
        Case "Bordeaux elegant": DesignSpec = Array("Bordeaux", "Weinrot", "Bordeaux", "Garamond", "Weinrot mit klassischer Serifenschrift, edel für Präsentationen.")
        Case "Erde warm": DesignSpec = Array("Mokka", "Terrakotta", "Erde", "Georgia", "Terrakotta, Ocker und Oliv mit Serifenschrift, warm und natürlich.")
        Case "Nordisch hell": DesignSpec = Array("Nebel", "Himmel hell", "Nordisch (pastell)", "Aptos", "Helles Graublau und Pastell, leicht und luftig.")
        Case "Flieder hell": DesignSpec = Array("Flieder", "Flieder hell", "Flieder (pastell)", "Calibri", "Sanfte Flieder- und Rosétöne.")
        Case "Salbei hell": DesignSpec = Array("Salbei", "Mint", "Salbei & Sand (pastell)", "Cambria", "Salbeigrün und Sand mit ruhiger Serifenschrift.")
        Case "Kontrast klar": DesignSpec = Array("Anthrazit", "Blau", "Kontrast (barrierearm)", "Verdana", "Hoher Kontrast, gut unterscheidbare Farben, sehr gut lesbare Schrift.")
        ' Design „Mainz“ (angelehnt an das Erscheinungsbild der Stadt Mainz): kräftiges Gelb (RGB 235/189/0), Navigation in der 20-%-Abstufung,
        ' Schrift Calibri
        Case "Mainz": DesignSpec = Array("Mainz hell", "Mainz Gelb", "Mainz", "Calibri", "Gelb und Hellgrün, Schrift Calibri: im Stil der Stadt Mainz.")
        Case Else: DesignSpec = Empty
    End Select
End Function

' Vorlage übernehmen: setzt alle vier Teile des Dokumentstils
Public Sub ApplyDesign(ByVal nm As String)
    Dim d As Variant
    d = DesignSpec(nm)
    If IsEmpty(d) Then Exit Sub
    SetDocStyle nm, CStr(d(0)), CStr(d(1)), CStr(d(2)), CStr(d(3))
End Sub

' Dokumentstil speichern (Vorlagenname oder "Eigene Auswahl")
Public Sub SetDocStyle(ByVal designName As String, ByVal navName As String, ByVal accName As String, ByVal styleName As String, ByVal fontName As String)
    mCfg.SysSet "Dokumentstil", designName
    mCfg.SysSet "Navigationsfarbe", CStr(mCmd.NavValue(navName))
    mCfg.SysSet "Akzentfarbe", CStr(mCmd.AccentValue(accName))
    mCfg.SysSet "Farbstil", IIf(IsStyleName(styleName), styleName, "Akzentfarbe")
    mCfg.SysSet "Schrift", DocFont(fontName)
    mAppFont = ""
End Sub

Public Function CurrentDesign() As String
    CurrentDesign = mCfg.SysGet("Dokumentstil", "")
End Function

' Standard-Farbstil der Datei (Dashboards, Erfassungsdateien, Export, solange dort nichts anderes gewählt ist)
Public Function DefaultStyle() As String
    DefaultStyle = mCfg.SysGet("Farbstil", "Akzentfarbe")
    If Not IsStyleName(DefaultStyle) Then DefaultStyle = "Akzentfarbe"
End Function

' Standard-Schrift der Datei (Oberfläche, Dashboards, erzeugte Dateien)
Public Function DefaultFont() As String
    DefaultFont = AppFont(False)
End Function

Public Function AppFont(ByVal semi As Boolean) As String
    If Len(mAppFont) = 0 Then
        On Error Resume Next
        mAppFont = DocFont(mCfg.SysGet("Schrift", "Segoe UI"))
        On Error GoTo 0
        If Len(mAppFont) = 0 Then mAppFont = "Segoe UI"
    End If
    If semi And mAppFont = "Segoe UI" Then AppFont = "Segoe UI Semibold" Else AppFont = mAppFont
End Function

Public Sub ResetFontCache()
    mAppFont = ""
End Sub

'--- Farbstile (gemeinsam für Dashboards, Erfassungsdateien und Excel-Export) -------------
' Elf ruhige Stile: kräftig-gedämpfte, pastellige und Mainz (Gelb und Hellgrün). Je Stil acht Farben; die erste ist
' die Leitfarbe (Kopfzeile, erste Diagrammreihe). Keine grellen Töne.
Public Function StyleNames() As Variant
    StyleNames = Array("Akzentfarbe", "Ozean", "Wald", "Schiefer", "Bordeaux", "Erde", "Nordisch (pastell)", "Flieder (pastell)", "Salbei & Sand (pastell)", "Kontrast (barrierearm)", "Mainz")
End Function

Public Function IsStyleName(ByVal nm As String) As Boolean
    Dim s As Variant
    For Each s In StyleNames()
        If CStr(s) = nm Then IsStyleName = True: Exit Function
    Next
End Function

Public Function StyleHelp(ByVal nm As String) As String
    Select Case nm
        Case "Akzentfarbe": StyleHelp = "Abstufungen der Akzentfarbe dieser Datei, einheitlich und ruhig."
        Case "Ozean": StyleHelp = "Nachtblau, Blau und Petrol: sachlich, gut für Berichte."
        Case "Wald": StyleHelp = "Gedeckte Grüntöne: ruhig, naturnah."
        Case "Schiefer": StyleHelp = "Graphit, Schiefer und Taupe: neutral, sehr dezent."
        Case "Bordeaux": StyleHelp = "Weinrot bis Rosé, gedämpft: edel, für Präsentationen."
        Case "Erde": StyleHelp = "Terrakotta, Ocker und Oliv: warm und natürlich."
        Case "Nordisch (pastell)": StyleHelp = "Helles Graublau, Salbei und Sand: zurückhaltend; mit dunkler Schrift."
        Case "Flieder (pastell)": StyleHelp = "Flieder, Rosé und Lavendel: sanft; mit dunkler Schrift."
        Case "Salbei & Sand (pastell)": StyleHelp = "Salbeigrün und Sandtöne: freundlich-ruhig; mit dunkler Schrift."
        Case "Kontrast (barrierearm)": StyleHelp = "Gut unterscheidbar, auch bei Farbsehschwäche, gedämpfte Töne."
        Case "Mainz": StyleHelp = "Gelb, Hellgrün, Blau und Orange: im Stil der Stadt Mainz."
    End Select
End Function

' Acht Farben eines Stils (VBA-RGB)
Public Function StyleColors(ByVal nm As String) As Variant
    Dim a As Long
    Select Case nm
        Case "Ozean": StyleColors = Array(RGB(23, 63, 95), RGB(32, 99, 155), RGB(60, 150, 160), RGB(110, 155, 195), RGB(16, 110, 120), RGB(160, 196, 222), RGB(70, 90, 125), RGB(130, 185, 185))
        Case "Wald": StyleColors = Array(RGB(30, 80, 60), RGB(50, 125, 90), RGB(110, 150, 95), RGB(150, 180, 120), RGB(70, 105, 80), RGB(190, 208, 168), RGB(95, 120, 65), RGB(135, 160, 140))
        Case "Schiefer": StyleColors = Array(RGB(45, 55, 72), RGB(71, 85, 105), RGB(100, 116, 139), RGB(148, 163, 184), RGB(120, 113, 108), RGB(196, 205, 216), RGB(30, 41, 59), RGB(165, 155, 145))
        Case "Bordeaux": StyleColors = Array(RGB(110, 32, 48), RGB(150, 55, 65), RGB(185, 100, 95), RGB(210, 150, 135), RGB(120, 75, 90), RGB(228, 192, 180), RGB(85, 45, 65), RGB(165, 120, 125))
        Case "Erde": StyleColors = Array(RGB(150, 80, 48), RGB(190, 135, 62), RGB(110, 112, 62), RGB(125, 85, 62), RGB(205, 170, 120), RGB(85, 95, 72), RGB(170, 112, 85), RGB(222, 200, 165))
        Case "Nordisch (pastell)": StyleColors = Array(RGB(176, 196, 218), RGB(178, 202, 186), RGB(222, 206, 172), RGB(210, 182, 182), RGB(178, 178, 205), RGB(198, 216, 226), RGB(226, 220, 205), RGB(160, 184, 172))
        Case "Flieder (pastell)": StyleColors = Array(RGB(200, 186, 230), RGB(230, 190, 210), RGB(182, 198, 235), RGB(214, 206, 238), RGB(240, 208, 216), RGB(172, 162, 210), RGB(224, 218, 244), RGB(200, 170, 200))
        Case "Salbei & Sand (pastell)": StyleColors = Array(RGB(176, 202, 180), RGB(222, 208, 172), RGB(196, 214, 194), RGB(200, 182, 150), RGB(150, 178, 158), RGB(234, 226, 198), RGB(186, 202, 166), RGB(210, 194, 170))
        Case "Kontrast (barrierearm)": StyleColors = Array(RGB(0, 95, 150), RGB(205, 135, 35), RGB(30, 135, 100), RGB(185, 100, 140), RGB(90, 155, 195), RGB(175, 85, 35), RGB(195, 175, 65), RGB(70, 70, 70))
        ' Mainz: Gelb und Hellgrün, dann weitere kräftige Farben
        Case "Mainz": StyleColors = Array(RGB(235, 189, 0), RGB(169, 208, 142), RGB(27, 121, 179), RGB(243, 148, 0), RGB(0, 119, 140), RGB(128, 163, 18), RGB(115, 115, 115), RGB(226, 218, 192))
        Case Else
            a = CAccent()
            StyleColors = Array(a, Blend(a, RGB(255, 255, 255), 0.35), Blend(a, RGB(0, 0, 0), 0.3), Blend(a, RGB(255, 255, 255), 0.6), _
                                Blend(a, RGB(0, 0, 0), 0.5), Blend(a, RGB(255, 255, 255), 0.2), Blend(a, RGB(255, 255, 255), 0.75), Blend(a, RGB(0, 0, 0), 0.15))
    End Select
End Function

' Auswahl für Dokumente = Farbstile (frühere Einzelfarben bleiben gültig)
Public Function DocColorNames() As Variant
    DocColorNames = StyleNames()
End Function

' Leitfarbe zu einem Namen: Stil, frühere Einzelfarbe oder (ältere Dateien) Dashboard-Palette
Public Function DocColor(ByVal nm As String) As Long
    If IsStyleName(nm) Then DocColor = StyleColors(nm)(0): Exit Function
    Select Case nm
        Case "Nachtblau": DocColor = RGB(30, 41, 82)
        Case "Königsblau": DocColor = RGB(37, 99, 235)
        Case "Petrol": DocColor = RGB(15, 110, 120)
        Case "Smaragd": DocColor = RGB(5, 150, 105)
        Case "Waldgrün": DocColor = RGB(22, 101, 52)
        Case "Olive": DocColor = RGB(101, 117, 35)
        Case "Aubergine": DocColor = RGB(88, 28, 92)
        Case "Violett": DocColor = RGB(109, 40, 217)
        Case "Weinrot": DocColor = RGB(127, 29, 29)
        Case "Kirschrot": DocColor = RGB(190, 18, 60)
        Case "Terrakotta": DocColor = RGB(194, 65, 12)
        Case "Bernstein": DocColor = RGB(180, 120, 10)
        Case "Graphit": DocColor = RGB(55, 65, 81)
        Case "Schiefer": DocColor = RGB(71, 85, 105)
        Case "Flieder (hell)": DocColor = RGB(216, 204, 240)
        Case "Salbei (hell)": DocColor = RGB(204, 226, 210)
        Case "Himmel (hell)": DocColor = RGB(199, 222, 244)
        Case "Pfirsich (hell)": DocColor = RGB(252, 214, 190)
        Case "Sand (hell)": DocColor = RGB(236, 224, 196)
        Case "Rosé (hell)": DocColor = RGB(246, 208, 216)
        Case "", "Akzentfarbe": DocColor = CAccent()
        Case Else: DocColor = mDash.PaletteColor(nm, 0)
    End Select
End Function

' Schrift auf einer Fläche mit dieser Farbe: dunkel bei hellen Farben, sonst weiß
Public Function DocOnColor(ByVal c As Long) As Long
    If IsLightColor(c) Then DocOnColor = RGB(31, 41, 55) Else DocOnColor = RGB(255, 255, 255)
End Function

' Lesbare Textfarbe dieser Farbe auf Weiß (helle Farben abgedunkelt)
Public Function DocInk(ByVal c As Long) As Long
    If IsLightColor(c) Then DocInk = Blend(c, RGB(0, 0, 0), 0.55) Else DocInk = c
End Function

Public Function DocFonts() As Variant
    DocFonts = Array("Segoe UI", "Calibri", "Aptos", "Arial", "Verdana", "Garamond", "Georgia", "Cambria")
End Function

' Schriftart prüfen (unbekannt -> Segoe UI)
Public Function DocFont(ByVal nm As String) As String
    Dim f As Variant
    For Each f In DocFonts()
        If StrComp(CStr(f), nm, vbTextCompare) = 0 Then DocFont = CStr(f): Exit Function
    Next
    DocFont = "Segoe UI"
End Function

' Serifenschriften laufen etwas kleiner: Grundgröße leicht anheben
Public Function DocFontSize(ByVal nm As String, ByVal baseSize As Single) As Single
    Select Case DocFont(nm)
        Case "Garamond": DocFontSize = baseSize + 1
        Case "Georgia", "Cambria": DocFontSize = baseSize + 0.5
        Case Else: DocFontSize = baseSize
    End Select
End Function
