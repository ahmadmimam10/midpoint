VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cField - Felddefinition (gelesen aus _Felder)
'==============================================================================
Option Explicit

Public ID As String
Public TableID As String
Public Name As String
Public FType As String          ' Basistyp (Speicherung/Pruefung)
Public Kind As String           ' gespeicherte Eigenschaftsart (z. B. KONTROLL, PERSON, URL)
Public Help As String
Public Required As Boolean
Public DefaultVal As String
Public Order As Double
Public Visible As Boolean
Public Group As String
Public Unique As Boolean
Public MaxLen As Long
Public Decimals As Long
Public MinVal As Variant
Public MaxVal As Variant
Public AllowEmpty As Boolean
Public TargetTable As String
Public Multi As Boolean
Public PairField As String
Public Formula As String
Public FormulaDesc As String
Public ResultType As String
Public NotBefore As String
Public Status As String
Public CfgRow As Long

Public Function IsActive() As Boolean
    IsActive = (Status <> "ausgeblendet")
End Function

Public Function IsLinkType() As Boolean
    IsLinkType = (FType = FT_LINK Or FType = FT_BACKLINK)
End Function

Public Function IsOptionType() As Boolean
    IsOptionType = (FType = FT_AUSWAHL Or FType = FT_STATUS Or FType = FT_MEHRFACH)
End Function

Public Function IsSingleOption() As Boolean
    IsSingleOption = (FType = FT_AUSWAHL Or FType = FT_STATUS)
End Function

' Effektiver Werttyp (bei berechneten Feldern der Ergebnistyp)
Public Function ValueType() As String
    If FType = FT_FORMEL Then
        If Len(ResultType) > 0 Then ValueType = ResultType Else ValueType = FT_ZAHL
    Else
        ValueType = FType
    End If
End Function

Public Function IsNumeric_() As Boolean
    Select Case ValueType()
        Case FT_ZAHL, FT_GELD, FT_PROZENT: IsNumeric_ = True
    End Select
End Function

Public Function IsDate_() As Boolean
    Select Case ValueType()
        Case FT_DATUM, FT_DATUMZEIT: IsDate_ = True
    End Select
End Function

Public Function IsTextLike() As Boolean
    Select Case ValueType()
        Case FT_TEXT, FT_LTEXT, FT_PFAD: IsTextLike = True
    End Select
End Function

' Direkt in Zellen/Formular bearbeitbar
Public Function IsEditable() As Boolean
    IsEditable = Not (FType = FT_FORMEL) And Not IsAutoKind(Kind)
End Function

Public Function IsAuto() As Boolean
    IsAuto = IsAutoKind(Kind) Or FType = FT_FORMEL
End Function

Public Function Icon() As String
    If FType = FT_FORMEL And Left$(Formula, 6) = "ROLLUP" Then Icon = FieldTypeIcon(FT_ROLLUP) Else Icon = FieldTypeIcon(Kind)
End Function

' Hat eine eigene Datenspalte im Datenblatt
Public Function HasColumn() As Boolean
    HasColumn = True
End Function

Public Function TypeName_() As String
    If FType = FT_FORMEL And Left$(Formula, 6) = "ROLLUP" Then TypeName_ = FieldTypeName(FT_ROLLUP) Else TypeName_ = FieldTypeName(Kind)
End Function

Public Function MinBound() As Variant
    MinBound = MinVal
End Function
