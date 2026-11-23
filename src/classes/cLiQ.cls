VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLiQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cLiQ - Rueckfrage der Listenuebernahme mit Vorschlag (Default) und Antwort.
'==============================================================================
Option Explicit

Public ID As String
Public Text As String
Public Choices As Variant          ' Array der Antworttexte
Public Default_ As Long            ' Index des Vorschlags (0-basiert)
Public Answer As Long              ' gewaehlte Antwort (0-basiert)
Public Topic As String             ' kurze Ueberschrift, z. B. Spaltenname
