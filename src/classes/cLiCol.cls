VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLiCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cLiCol - geplantes Zielfeld der Listenuebernahme.
' Herkunft (Mode): "wert" eine Spalte | "kreuz" Kreuze ueber mehrere Spalten |
' "liste" mehrere Werte je Zelle | "gruppe" Gruppenueberschrift | "blatt" Blattname |
' "monat"/"monatswert" Zeitraumspalten in Eintraege umgewandelt | "code" 2ask-Spalte |
' "codemulti" 2ask-Mehrfachauswahl ueber mehrere Spalten | "formular" Formularangabe.
'==============================================================================
Option Explicit

Public Key As String               ' stabiler Schluessel (Aenderungen des Nutzers haengen daran)
Public Name As String
Public FType As String
Public Include As Boolean
Public Mode As String
Public Src As Collection           ' Quellspalten (Index der vereinheitlichten Spalten bzw. 2ask-Spalte)
Public SrcNames As Collection      ' Optionsname je Quellspalte (Kreuze, 2ask-Mehrfachauswahl)
Public SrcCol As Long              ' Formular: Spalte des Werts
Public Sep As String               ' Trennzeichen bei "liste", z. B. ",;/"
Public Opts As Collection          ' Optionsnamen in Reihenfolge
Public ValMap As cMap              ' Rohwert -> Zielwert ("" = leer; bei Mehrfachauswahl mit ; getrennt)
Public CodeMap As cMap             ' 2ask: Code (bzw. Code@Spalte) -> Optionsname
Public BlankFalse As Boolean       ' Ja/Nein: leere Zelle = Nein (Kreuzspalte)
Public Decimals As Long
Public Hint As String
Public Sample As String
Public Unsure As Boolean
Public IsFormula As Boolean
Public IsHidden As Boolean
Public Year_ As Long               ' Zeitraumspalten: Jahr (0 = unbekannt)
Public CanBeLabel As Boolean
Public AllUnique As Boolean        ' alle Werte verschieden (Kandidat fuer Bezeichnungsfeld/Schluessel)
Public QKey As String              ' 2ask: Fragenummer (Q4)
Public Tag As String               ' 2ask: Zusatzkennung (z. B. bip:Q8)
Public Help As String              ' Hilfetext des Felds (vollstaendige Ueberschrift, Stichtag, Uebersetzung)
Public DistRatio As Double         ' Anteil verschiedener Werte (1 = alle verschieden)
Public FillRatio As Double         ' Anteil ausgefuellter Zellen (1 = keine leere)
Public Ditto As Boolean            ' "dito" / "siehe oben" = Wert der Zeile darueber

Private Sub Class_Initialize()
    Set Src = New Collection
    Set SrcNames = New Collection
    Set Opts = New Collection
    Set ValMap = New cMap
    Set CodeMap = New cMap
    Include = True
    Decimals = -1
End Sub

Public Function FirstSrc() As Long
    If Src.Count > 0 Then FirstSrc = CLng(Src(1))
End Function

Public Function IsOptionType() As Boolean
    IsOptionType = (FType = FT_AUSWAHL Or FType = FT_MEHRFACH Or FType = FT_STATUS)
End Function
