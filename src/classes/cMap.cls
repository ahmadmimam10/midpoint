VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cMap - Schluessel/Wert-Zuordnung ohne externe Bibliothek (Collection-basiert).
' Schluessel sind nicht gross-/kleinschreibungssensitiv. Reihenfolge der
' Einfuegung bleibt fuer Keys/Items erhalten.
'==============================================================================
Option Explicit

Private mVals As Collection
Private mKeys As Collection

Private Sub Class_Initialize()
    Set mVals = New Collection
    Set mKeys = New Collection
End Sub

Private Function K(ByVal key As String) As String
    K = "k" & key
End Function

Public Function Exists(ByVal key As String) As Boolean
    Dim v As Variant
    On Error GoTo nope
    v = IsObject(mVals(K(key)))
    Exists = True
    Exit Function
nope:
    Exists = False
End Function

Public Sub SetVal(ByVal key As String, ByVal value As Variant)
    If Exists(key) Then
        mVals.Remove K(key)
        mVals.Add value, K(key)
    Else
        mVals.Add value, K(key)
        mKeys.Add key, K(key)
    End If
End Sub

' Wie SetVal, aber fuer Schluessel, die meist neu sind (z. B. Werte einer Importzeile): erst einfuegen,
' nur bei vorhandenem Schluessel ersetzen. Ein neuer Schluessel kostet so keine Fehlerbehandlung;
' fuer haeufige Aktualisierungen vorhandener Schluessel ist SetVal schneller.
Public Sub SetNew(ByVal key As String, ByVal value As Variant)
    Dim kk As String, isNew As Boolean
    kk = K(key)
    On Error Resume Next
    mVals.Add value, kk
    isNew = (Err.Number = 0)
    On Error GoTo 0
    If isNew Then
        mKeys.Add key, kk
    Else
        mVals.Remove kk
        mVals.Add value, kk
    End If
End Sub

Public Function GetVal(ByVal key As String, Optional ByVal dflt As Variant) As Variant
    On Error GoTo nope
    If IsObject(mVals(K(key))) Then
        Set GetVal = mVals(K(key))
    Else
        GetVal = mVals(K(key))
    End If
    Exit Function
nope:
    If IsMissing(dflt) Then
        GetVal = Empty
    ElseIf IsObject(dflt) Then
        Set GetVal = dflt
    Else
        GetVal = dflt
    End If
End Function

Public Sub Remove(ByVal key As String)
    If Exists(key) Then
        mVals.Remove K(key)
        mKeys.Remove K(key)
    End If
End Sub

Public Function Count() As Long
    Count = mVals.Count
End Function

Public Function Keys() As Variant
    Dim a() As Variant, i As Long
    If mKeys.Count = 0 Then
        Keys = Array()
        Exit Function
    End If
    ReDim a(0 To mKeys.Count - 1)
    For i = 1 To mKeys.Count
        a(i - 1) = mKeys(i)
    Next
    Keys = a
End Function

Public Function Items() As Variant
    Dim a() As Variant, i As Long
    If mVals.Count = 0 Then
        Items = Array()
        Exit Function
    End If
    ReDim a(0 To mVals.Count - 1)
    For i = 1 To mKeys.Count
        If IsObject(mVals(K(mKeys(i)))) Then
            Set a(i - 1) = mVals(K(mKeys(i)))
        Else
            a(i - 1) = mVals(K(mKeys(i)))
        End If
    Next
    Items = a
End Function

Public Sub Clear()
    Set mVals = New Collection
    Set mKeys = New Collection
End Sub

' Zaehler erhoehen (fuer Gruppierungen). Ein Nachschlagen statt drei: ein neuer Schluessel kostet
' nur eine Fehlerbehandlung (wichtig beim Zaehlen tausender verschiedener Werte).
Public Sub AddNum(ByVal key As String, ByVal n As Double)
    Dim kk As String, cur As Variant, has As Boolean
    kk = K(key)
    On Error Resume Next
    cur = mVals(kk)
    has = (Err.Number = 0)
    On Error GoTo 0
    If has Then
        mVals.Remove kk
        mVals.Add CDbl(cur) + n, kk
    Else
        mVals.Add n, kk
        mKeys.Add key, kk
    End If
End Sub
