Attribute VB_Name = "mConst"
'==============================================================================
' mConst - Versionen, Blattnamen, Feldtypen, Gestaltungswerte
'==============================================================================
Option Explicit

Public Const APP_TITLE As String = "MidPoint"
Public Const PROG_VERSION As String = "1.8.2"
Public Const FORMAT_VERSION As Long = 3
Public Const PROG_DATE As String = "November 2026"
Public Const AUTHOR_NOTE As String = "Entwickler: Ahmed Imam mit Claude Code"

' --- Oberflaechenblaetter ---
Public Const SH_HINWEIS As String = "Hinweis"
Public Const SH_START As String = "Start"
Public Const SH_TABLE As String = "Tabelle"
Public Const SH_CAL As String = "Kalender"
Public Const SH_BOARD As String = "Board"
Public Const SH_GANTT As String = "Gantt"
Public Const SH_DASH As String = "Dashboard"
Public Const SH_SETUP As String = "Einrichten"
Public Const SH_HELP As String = "Hilfe"
Public Const SH_TECH As String = "Technik"
Public Const SH_RES As String = "_Ressourcen"          ' eingebettete Bilder (Logo), Base64
Public Const SH_SEARCH As String = "Suche"
Public Const SH_CAPTURE As String = "Erfassung"
Public Const SH_CAP As String = "_Erfassungen"
Public Const SH_RET As String = "_Ruecklauf"

' --- Daten- und Konfigurationsblaetter ---
Public Const SH_SYS As String = "_System"
Public Const SH_TAB As String = "_Tabellen"
Public Const SH_FLD As String = "_Felder"
Public Const SH_OPT As String = "_Optionen"
Public Const SH_REL As String = "_Zuordnungen"
Public Const SH_VIEW As String = "_Ansichten"
Public Const SH_DSB As String = "_Dashboards"
Public Const SH_ELM As String = "_Elemente"
Public Const SH_LOG As String = "_Protokoll"
Public Const SH_LST As String = "_Listen"
Public Const SH_TPL As String = "_Vorlagen"
Public Const SH_CODE As String = "_Code"
Public Const SH_IMP As String = "_Importbericht"
Public Const SH_CALC As String = "_Rechnen"
Public Const DATA_PREFIX As String = "D_"

' --- Systemspalten der Datentabellen ---
Public Const C_ID As String = "_ID"
Public Const C_CREATED As String = "_Erstellt"
Public Const C_CHANGED As String = "_Geaendert"
Public Const C_USER As String = "_Benutzer"
Public Const C_ARCH As String = "_Archiviert"
Public Const SYS_COLS As Long = 5

' --- Feldtypen (interne Codes) ---
Public Const FT_TEXT As String = "TEXT"
Public Const FT_LTEXT As String = "LTEXT"
Public Const FT_ZAHL As String = "ZAHL"
Public Const FT_GELD As String = "GELD"
Public Const FT_PROZENT As String = "PROZENT"
Public Const FT_DATUM As String = "DATUM"
Public Const FT_DATUMZEIT As String = "DATUMZEIT"
Public Const FT_JANEIN As String = "JANEIN"
Public Const FT_AUSWAHL As String = "AUSWAHL"
Public Const FT_MEHRFACH As String = "MEHRFACH"
Public Const FT_STATUS As String = "STATUS"
Public Const FT_LINK As String = "LINK"
Public Const FT_BACKLINK As String = "BACKLINK"
Public Const FT_FORMEL As String = "FORMEL"
Public Const FT_PFAD As String = "PFAD"
' Weitere Eigenschaftsarten (wie in Notion); sie bauen auf einem Basistyp auf
Public Const FT_KONTROLL As String = "KONTROLL"
Public Const FT_PERSON As String = "PERSON"
Public Const FT_DATEIEN As String = "DATEIEN"
Public Const FT_URL As String = "URL"
Public Const FT_EMAIL As String = "EMAIL"
Public Const FT_TELEFON As String = "TELEFON"
Public Const FT_ORT As String = "ORT"
Public Const FT_ROLLUP As String = "ROLLUP"
Public Const FT_ERST_UM As String = "ERSTELLT_UM"
Public Const FT_ERST_VON As String = "ERSTELLT_VON"
Public Const FT_AEND_UM As String = "GEAENDERT_UM"
Public Const FT_AEND_VON As String = "GEAENDERT_VON"
Public Const FT_BUTTON As String = "BUTTON"
Public Const FT_AUTOID As String = "AUTOID"

' --- Ansichtstypen ---
Public Const VT_TABLE As String = "Tabelle"
Public Const VT_CAL As String = "Kalender"
Public Const VT_BOARD As String = "Board"
Public Const VT_GANTT As String = "Gantt"

' --- Fehlerkennzeichen in berechneten Feldern ---
Public Const ERR_PREFIX As String = "#FEHLER:"

' --- Layout (Punkte) ---
Public Const NAV_WIDTH_PT As Double = 184
Public Const GUTTER_PT As Double = 16
Public Const PAGE_SIZE_DEFAULT As Long = 100

Public Const ICON_FONT As String = "Segoe MDL2 Assets"
' Schrift der Oberfläche: aus dem Dokumentstil (Einrichten), Standard Segoe UI
Public Function UI_FONT() As String: UI_FONT = mDocStyle.AppFont(False): End Function
Public Function UI_FONT_SEMI() As String: UI_FONT_SEMI = mDocStyle.AppFont(True): End Function

' --- Farben (VBA-RGB) ---
Public Function CText() As Long: CText = RGB(31, 41, 55): End Function
Public Function CMuted() As Long: CMuted = RGB(107, 114, 128): End Function
Public Function CBorder() As Long: CBorder = RGB(229, 231, 235): End Function
Public Function CBorderStrong() As Long: CBorderStrong = RGB(209, 213, 219): End Function
Public Function CSurface() As Long: CSurface = RGB(249, 250, 251): End Function
Public Function CSurface2() As Long: CSurface2 = RGB(243, 244, 246): End Function
Public Function CWhite() As Long: CWhite = RGB(255, 255, 255): End Function
Public Function CDanger() As Long: CDanger = RGB(220, 38, 38): End Function
Public Function CDangerSoft() As Long: CDangerSoft = RGB(254, 226, 226): End Function
Public Function CSuccess() As Long: CSuccess = RGB(22, 163, 74): End Function
Public Function CWarn() As Long: CWarn = RGB(217, 119, 6): End Function
Public Function CWarnSoft() As Long: CWarnSoft = RGB(254, 243, 199): End Function

' Akzent- und Navigationsfarbe sind einstellbar (Einrichten). Helle Pastellfarben sind
' erlaubt: dann wird Schrift darauf dunkel und Akzent-Schrift auf Weiss abgedunkelt.
Public Function CAccentRaw() As Long
    CAccentRaw = mCfg.SysLong("Akzentfarbe", RGB(37, 99, 235))
End Function
' Akzent fuer Schrift, Symbole und Linien auf hellem Grund (immer gut lesbar)
Public Function CAccent() As Long
    CAccent = CAccentRaw()
    If IsLightColor(CAccent) Then CAccent = Blend(CAccent, RGB(24, 24, 40), 0.58)
End Function
' Akzent fuer Flaechen (Hauptschaltflaechen, Markierungen)
Public Function CAccentFill() As Long
    CAccentFill = CAccentRaw()
End Function
' Schrift auf Akzentflaechen
Public Function COnAccent() As Long
    If IsLightColor(CAccentRaw()) Then COnAccent = CText() Else COnAccent = RGB(255, 255, 255)
End Function
Public Function CAccentSoft() As Long
    If IsLightColor(CAccentRaw()) Then
        CAccentSoft = Blend(CAccentRaw(), RGB(255, 255, 255), 0.5)
    Else
        CAccentSoft = Blend(CAccentRaw(), RGB(255, 255, 255), 0.86)
    End If
End Function
Public Function CNavBg() As Long
    CNavBg = mCfg.SysLong("Navigationsfarbe", RGB(30, 41, 59))
End Function
Public Function NavIsLight() As Boolean
    NavIsLight = IsLightColor(CNavBg())
End Function
' Hervorgehobene Schrift in der Navigation (Titel, Hover): weiss bzw. bei hellen Farben fast schwarz
Public Function CNavStrong() As Long
    If NavIsLight() Then CNavStrong = RGB(17, 24, 39) Else CNavStrong = RGB(255, 255, 255)
End Function
Public Function CNavText() As Long
    If NavIsLight() Then CNavText = Blend(CNavBg(), RGB(17, 24, 39), 0.82) Else CNavText = Blend(CNavBg(), RGB(255, 255, 255), 0.82)
End Function
Public Function CNavMuted() As Long
    If NavIsLight() Then CNavMuted = Blend(CNavBg(), RGB(17, 24, 39), 0.55) Else CNavMuted = Blend(CNavBg(), RGB(255, 255, 255), 0.55)
End Function

' Relative Helligkeit (sRGB, vereinfacht) ueber 0,6 = helle Farbe, dunkle Schrift noetig
Public Function IsLightColor(ByVal c As Long) As Boolean
    Dim r As Long, g As Long, b As Long
    r = c And &HFF&: g = (c \ &H100&) And &HFF&: b = (c \ &H10000) And &HFF&
    IsLightColor = ((0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 > 0.6)
End Function

Public Function Blend(ByVal c1 As Long, ByVal c2 As Long, ByVal t As Double) As Long
    Dim r1 As Long, g1 As Long, b1 As Long, r2 As Long, g2 As Long, b2 As Long
    r1 = c1 And &HFF&: g1 = (c1 \ &H100&) And &HFF&: b1 = (c1 \ &H10000) And &HFF&
    r2 = c2 And &HFF&: g2 = (c2 \ &H100&) And &HFF&: b2 = (c2 \ &H10000) And &HFF&
    Blend = RGB(r1 + (r2 - r1) * t, g1 + (g2 - g1) * t, b1 + (b2 - b1) * t)
End Function

' --- Optionsfarben (Hintergrund / Schrift), Notion-aehnlich ruhig ---
Public Function PaletteKeys() As Variant
    PaletteKeys = Array("grau", "braun", "orange", "gelb", "gruen", "blau", "lila", "rosa", "rot", "tuerkis")
End Function

Public Function PaletteName(ByVal key As String) As String
    Select Case LCase$(key)
        Case "grau": PaletteName = "Grau"
        Case "braun": PaletteName = "Braun"
        Case "orange": PaletteName = "Orange"
        Case "gelb": PaletteName = "Gelb"
        Case "gruen": PaletteName = "Grün"
        Case "blau": PaletteName = "Blau"
        Case "lila": PaletteName = "Lila"
        Case "rosa": PaletteName = "Rosa"
        Case "rot": PaletteName = "Rot"
        Case "tuerkis": PaletteName = "Türkis"
        Case Else: PaletteName = "Grau"
    End Select
End Function

Public Function PaletteKeyFromName(ByVal nm As String) As String
    Dim k As Variant
    For Each k In PaletteKeys()
        If LCase$(PaletteName(CStr(k))) = LCase$(Trim$(nm)) Or LCase$(CStr(k)) = LCase$(Trim$(nm)) Then
            PaletteKeyFromName = CStr(k): Exit Function
        End If
    Next
    PaletteKeyFromName = "grau"
End Function

Public Function PaletteBg(ByVal key As String) As Long
    Select Case LCase$(key)
        Case "braun": PaletteBg = RGB(238, 226, 218)
        Case "orange": PaletteBg = RGB(251, 226, 204)
        Case "gelb": PaletteBg = RGB(253, 239, 200)
        Case "gruen": PaletteBg = RGB(219, 237, 221)
        Case "blau": PaletteBg = RGB(214, 230, 243)
        Case "lila": PaletteBg = RGB(233, 224, 241)
        Case "rosa": PaletteBg = RGB(246, 224, 234)
        Case "rot": PaletteBg = RGB(253, 222, 218)
        Case "tuerkis": PaletteBg = RGB(208, 239, 235)
        Case Else: PaletteBg = RGB(237, 237, 235)
    End Select
End Function

Public Function PaletteFg(ByVal key As String) As Long
    Select Case LCase$(key)
        Case "braun": PaletteFg = RGB(98, 68, 52)
        Case "orange": PaletteFg = RGB(140, 70, 20)
        Case "gelb": PaletteFg = RGB(113, 84, 6)
        Case "gruen": PaletteFg = RGB(28, 94, 52)
        Case "blau": PaletteFg = RGB(25, 76, 122)
        Case "lila": PaletteFg = RGB(92, 52, 128)
        Case "rosa": PaletteFg = RGB(128, 40, 84)
        Case "rot": PaletteFg = RGB(150, 38, 28)
        Case "tuerkis": PaletteFg = RGB(18, 94, 88)
        Case Else: PaletteFg = RGB(64, 64, 62)
    End Select
End Function

' --- Feldtyp-Bezeichnungen (Reihenfolge wie in Notion, danach Ergaenzungen) ---
Public Function FieldTypeCodes() As Variant
    FieldTypeCodes = Array(FT_TEXT, FT_ZAHL, FT_AUSWAHL, FT_MEHRFACH, FT_STATUS, FT_DATUM, FT_PERSON, FT_DATEIEN, FT_KONTROLL, _
                           FT_URL, FT_EMAIL, FT_TELEFON, FT_FORMEL, FT_LINK, FT_ROLLUP, FT_ERST_UM, FT_ERST_VON, FT_AEND_UM, _
                           FT_AEND_VON, FT_BUTTON, FT_ORT, FT_AUTOID, FT_LTEXT, FT_GELD, FT_PROZENT, FT_DATUMZEIT, FT_JANEIN, FT_PFAD)
End Function

Public Function FieldTypeName(ByVal code As String) As String
    Select Case code
        Case FT_TEXT: FieldTypeName = "Text"
        Case FT_LTEXT: FieldTypeName = "Langer Text / Notiz"
        Case FT_ZAHL: FieldTypeName = "Nummer"
        Case FT_GELD: FieldTypeName = "Geldbetrag (€)"
        Case FT_PROZENT: FieldTypeName = "Prozent"
        Case FT_DATUM: FieldTypeName = "Datum"
        Case FT_DATUMZEIT: FieldTypeName = "Datum mit Uhrzeit"
        Case FT_JANEIN: FieldTypeName = "Ja/Nein (auch leer)"
        Case FT_AUSWAHL: FieldTypeName = "Auswählen"
        Case FT_MEHRFACH: FieldTypeName = "Mehrfachauswahl"
        Case FT_STATUS: FieldTypeName = "Status"
        Case FT_LINK: FieldTypeName = "Verknüpfung"
        Case FT_BACKLINK: FieldTypeName = "Rückverknüpfung"
        Case FT_FORMEL: FieldTypeName = "Formel"
        Case FT_PFAD: FieldTypeName = "Lokaler Pfad"
        Case FT_KONTROLL: FieldTypeName = "Kontrollkästchen"
        Case FT_PERSON: FieldTypeName = "Person"
        Case FT_DATEIEN: FieldTypeName = "Dateien und Medien"
        Case FT_URL: FieldTypeName = "URL"
        Case FT_EMAIL: FieldTypeName = "E-Mail"
        Case FT_TELEFON: FieldTypeName = "Telefon"
        Case FT_ORT: FieldTypeName = "Ort"
        Case FT_ROLLUP: FieldTypeName = "Rollup"
        Case FT_ERST_UM: FieldTypeName = "Erstellt um"
        Case FT_ERST_VON: FieldTypeName = "Erstellt von"
        Case FT_AEND_UM: FieldTypeName = "Letzte Änderung um"
        Case FT_AEND_VON: FieldTypeName = "Zuletzt bearbeitet von"
        Case FT_BUTTON: FieldTypeName = "Button"
        Case FT_AUTOID: FieldTypeName = "ID"
        Case Else: FieldTypeName = code
    End Select
End Function

Public Function FieldTypeHelp(ByVal code As String) As String
    Select Case code
        Case FT_TEXT: FieldTypeHelp = "Kurzer Text, z. B. Name oder Titel"
        Case FT_LTEXT: FieldTypeHelp = "Mehrzeilige Notiz"
        Case FT_ZAHL: FieldTypeHelp = "Ganzzahl oder Dezimalzahl, optional mit Grenzen"
        Case FT_GELD: FieldTypeHelp = "Betrag in Euro, centgenau"
        Case FT_PROZENT: FieldTypeHelp = "Prozentwert (Eingabe 25 = 25 %)"
        Case FT_DATUM: FieldTypeHelp = "Datum mit Kalenderauswahl"
        Case FT_DATUMZEIT: FieldTypeHelp = "Datum und Uhrzeit"
        Case FT_JANEIN: FieldTypeHelp = "Ja, Nein oder keine Angabe"
        Case FT_AUSWAHL: FieldTypeHelp = "Genau eine Option aus einer farbigen Liste"
        Case FT_MEHRFACH: FieldTypeHelp = "Mehrere Optionen aus einer farbigen Liste"
        Case FT_STATUS: FieldTypeHelp = "Geordnete Zustände, z. B. Offen, In Arbeit, Erledigt"
        Case FT_LINK: FieldTypeHelp = "Einträge einer anderen Tabelle zuordnen (1:n oder n:m)"
        Case FT_FORMEL: FieldTypeHelp = "Excel-Formel mit Spalten in [ ], z. B. =[Preis]*[Menge]"
        Case FT_PFAD: FieldTypeHelp = "Ein lokaler Datei- oder Ordnerpfad"
        Case FT_KONTROLL: FieldTypeHelp = "Häkchen an/aus; Doppelklick in der Tabelle schaltet um"
        Case FT_PERSON: FieldTypeHelp = "Eine oder mehrere Personen aus einer gepflegten Liste"
        Case FT_DATEIEN: FieldTypeHelp = "Mehrere lokale Dateien (Verweise, keine Kopien)"
        Case FT_URL: FieldTypeHelp = "Webadresse; Öffnen nur nach bewusster Aktion"
        Case FT_EMAIL: FieldTypeHelp = "E-Mail-Adresse mit Formatprüfung"
        Case FT_TELEFON: FieldTypeHelp = "Telefonnummer mit Formatprüfung"
        Case FT_ORT: FieldTypeHelp = "Adresse oder Ort als Text (offline, ohne Karte)"
        Case FT_ROLLUP: FieldTypeHelp = "Anzahl, Summe, Min, Max, Durchschnitt verknüpfter Einträge"
        Case FT_ERST_UM: FieldTypeHelp = "Wird beim Anlegen automatisch gesetzt"
        Case FT_ERST_VON: FieldTypeHelp = "Windows-Kennung beim Anlegen (automatisch)"
        Case FT_AEND_UM: FieldTypeHelp = "Wird bei jeder Änderung automatisch gesetzt"
        Case FT_AEND_VON: FieldTypeHelp = "Windows-Kennung der letzten Änderung (automatisch)"
        Case FT_BUTTON: FieldTypeHelp = "Schaltfläche, die Felder setzt (z. B. Status auf Erledigt, Datum auf heute)"
        Case FT_AUTOID: FieldTypeHelp = "Fortlaufende Nummer mit optionalem Präfix, z. B. AUF-12"
    End Select
End Function

Public Function FieldTypeFromName(ByVal nm As String) As String
    Dim c As Variant
    For Each c In FieldTypeCodes()
        If FieldTypeName(CStr(c)) = nm Then FieldTypeFromName = CStr(c): Exit Function
    Next
    If nm = FieldTypeName(FT_BACKLINK) Then FieldTypeFromName = FT_BACKLINK
End Function

' Basistyp einer Eigenschaftsart (darauf beruhen Speicherung, Pruefung, Filter, Sortierung)
Public Function BaseTypeOf(ByVal code As String) As String
    Select Case code
        Case FT_KONTROLL: BaseTypeOf = FT_JANEIN
        Case FT_PERSON: BaseTypeOf = FT_MEHRFACH
        Case FT_DATEIEN: BaseTypeOf = FT_PFAD
        Case FT_URL, FT_EMAIL, FT_TELEFON, FT_ORT, FT_ERST_VON, FT_AEND_VON, FT_AUTOID, FT_BUTTON: BaseTypeOf = FT_TEXT
        Case FT_ERST_UM, FT_AEND_UM: BaseTypeOf = FT_DATUMZEIT
        Case FT_ROLLUP: BaseTypeOf = FT_FORMEL
        Case Else: BaseTypeOf = code
    End Select
End Function

' Automatisch gepflegte Arten (nicht manuell aenderbar)
Public Function IsAutoKind(ByVal code As String) As Boolean
    Select Case code
        Case FT_ERST_UM, FT_ERST_VON, FT_AEND_UM, FT_AEND_VON, FT_AUTOID, FT_BUTTON: IsAutoKind = True
    End Select
End Function

' Symbol (Segoe MDL2 Assets) je Feldtyp
Public Function FieldTypeIcon(ByVal code As String) As String
    Select Case code
        Case FT_TEXT: FieldTypeIcon = ChrW(&HE8D2)
        Case FT_LTEXT: FieldTypeIcon = ChrW(&HE8A5)
        Case FT_ZAHL: FieldTypeIcon = ChrW(&HE8EF)
        Case FT_GELD: FieldTypeIcon = ChrW(&HE8C7)
        Case FT_PROZENT: FieldTypeIcon = ChrW(&HE9D2)
        Case FT_DATUM, FT_DATUMZEIT: FieldTypeIcon = ChrW(&HE787)
        Case FT_JANEIN: FieldTypeIcon = ChrW(&HE73E)
        Case FT_AUSWAHL: FieldTypeIcon = ChrW(&HE70D)
        Case FT_MEHRFACH: FieldTypeIcon = ChrW(&HE762)
        Case FT_STATUS: FieldTypeIcon = ChrW(&HE91F)
        Case FT_LINK, FT_BACKLINK: FieldTypeIcon = ChrW(&HE8A7)
        Case FT_FORMEL: FieldTypeIcon = ChrW(&HE8EF)
        Case FT_PFAD: FieldTypeIcon = ChrW(&HE8B7)
        Case FT_KONTROLL: FieldTypeIcon = ChrW(&HE73A)
        Case FT_PERSON: FieldTypeIcon = ChrW(&HE716)
        Case FT_DATEIEN: FieldTypeIcon = ChrW(&HE723)
        Case FT_URL: FieldTypeIcon = ChrW(&HE71B)
        Case FT_EMAIL: FieldTypeIcon = ChrW(&HE715)
        Case FT_TELEFON: FieldTypeIcon = ChrW(&HE717)
        Case FT_ORT: FieldTypeIcon = ChrW(&HE707)
        Case FT_ROLLUP: FieldTypeIcon = ChrW(&HE721)
        Case FT_ERST_UM, FT_AEND_UM: FieldTypeIcon = ChrW(&HE823)
        Case FT_ERST_VON: FieldTypeIcon = ChrW(&HE77B)
        Case FT_AEND_VON: FieldTypeIcon = ChrW(&HE70F)
        Case FT_BUTTON: FieldTypeIcon = ChrW(&HE768)
        Case FT_AUTOID: FieldTypeIcon = ChrW(&HE8EC)
        Case Else: FieldTypeIcon = ChrW(&HE8D2)
    End Select
End Function

' Auswahl an Tabellen-Symbolen (Segoe MDL2 Assets)
Public Function TableIconCodes() As Variant
    TableIconCodes = Array(&HE8FD&, &HE9D5&, &HE8F1&, &HE77B&, &HE8B7&, &HE787&, &HE7C3&, &HE821&, _
                           &HE913&, &HE80F&, &HE7BE&, &HE734&, &HE8D4&, &HE719&, &HE82D&, &HEC92&)
End Function

Public Function IconChar(ByVal code As Variant) As String
    If IsNumeric(code) And Len(CStr(code)) > 0 Then
        IconChar = ChrW(CLng(code))
    Else
        IconChar = ChrW(&HE8FD)
    End If
End Function

