Attribute VB_Name = "mUtil"
'==============================================================================
' mUtil - Allgemeine Hilfen: Zeichenketten, Serialisierung, deutsche Zahlen-/
' Datumsdarstellung und -erkennung, Sortierung, Dateien, UTF-8, CRC32.
'==============================================================================
Option Explicit

Private mCrcTable(0 To 255) As Long
Private mCrcReady As Boolean

'--- Leere Werte ---------------------------------------------------------------
Public Function IsBlank(ByVal v As Variant) As Boolean
    If IsObject(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then IsBlank = True: Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then IsBlank = (Len(Trim$(v)) = 0)
End Function

' Zellsicherer Text: ein fuehrendes Apostroph wuerde von Excel als Praefix verschluckt
Public Function CellText(ByVal v As Variant) As Variant
    If VarType(v) = vbString Then
        ' Text, der mit ' = + - @ beginnt, bleibt Text (wird nie als Formel ausgefuehrt)
        If Len(v) > 0 And InStr("'=+-@", Left$(v, 1)) > 0 Then CellText = "'" & v Else CellText = v
    Else
        CellText = v
    End If
End Function

' CellText fuer einen ganzen 2D-Bereich (Kopieren von Blattinhalten)
Public Function CellTexts(ByVal v As Variant) As Variant
    Dim r As Long, c As Long
    If Not IsArray(v) Then CellTexts = CellText(v): Exit Function
    For r = LBound(v, 1) To UBound(v, 1)
        For c = LBound(v, 2) To UBound(v, 2)
            If VarType(v(r, c)) = vbString Then v(r, c) = CellText(v(r, c))
        Next
    Next
    CellTexts = v
End Function

Public Function NzS(ByVal v As Variant) As String
    If IsBlank(v) Or IsError(v) Then NzS = "" Else NzS = CStr(v)
End Function

Public Function IsErrVal(ByVal v As Variant) As Boolean
    If IsError(v) Then IsErrVal = True: Exit Function
    If VarType(v) = vbString Then IsErrVal = (Left$(v, Len(ERR_PREFIX)) = ERR_PREFIX)
End Function

'--- Maskierung fuer gespeicherte Einstellungen ----------------------------------
Public Function Esc(ByVal s As String) As String
    s = Replace(s, "%", "%25")
    s = Replace(s, "|", "%7C")
    s = Replace(s, ";", "%3B")
    s = Replace(s, "=", "%3D")
    s = Replace(s, "~", "%7E")
    s = Replace(s, ",", "%2C")
    s = Replace(s, ":", "%3A")
    s = Replace(s, vbCr, "%0D")
    s = Replace(s, vbLf, "%0A")
    Esc = s
End Function

Public Function Unesc(ByVal s As String) As String
    s = Replace(s, "%0A", vbLf)
    s = Replace(s, "%0D", vbCr)
    s = Replace(s, "%3A", ":")
    s = Replace(s, "%2C", ",")
    s = Replace(s, "%7E", "~")
    s = Replace(s, "%3D", "=")
    s = Replace(s, "%3B", ";")
    s = Replace(s, "%7C", "|")
    s = Replace(s, "%25", "%")
    Unesc = s
End Function

' Schluessel=Wert;Schluessel=Wert
Public Function KVGet(ByVal kv As String, ByVal key As String, Optional ByVal dflt As String = "") As String
    Dim parts As Variant, i As Long, p As Long
    KVGet = dflt
    If Len(kv) = 0 Then Exit Function
    parts = Split(kv, ";")
    For i = LBound(parts) To UBound(parts)
        p = InStr(parts(i), "=")
        If p > 0 Then
            If LCase$(Left$(parts(i), p - 1)) = LCase$(key) Then
                KVGet = Unesc(Mid$(parts(i), p + 1))
                Exit Function
            End If
        End If
    Next
End Function

Public Function KVSet(ByVal kv As String, ByVal key As String, ByVal value As String) As String
    Dim parts As Variant, i As Long, p As Long, res As String, found As Boolean
    If Len(kv) > 0 Then
        parts = Split(kv, ";")
        For i = LBound(parts) To UBound(parts)
            p = InStr(parts(i), "=")
            If p > 0 Then
                If LCase$(Left$(parts(i), p - 1)) = LCase$(key) Then
                    If Len(value) > 0 Then res = res & IIf(Len(res) > 0, ";", "") & key & "=" & Esc(value)
                    found = True
                Else
                    res = res & IIf(Len(res) > 0, ";", "") & parts(i)
                End If
            End If
        Next
    End If
    If Not found And Len(value) > 0 Then res = res & IIf(Len(res) > 0, ";", "") & key & "=" & Esc(value)
    KVSet = res
End Function

'--- Listen -------------------------------------------------------------------
Public Function ListHas(ByVal lst As String, ByVal item As String, Optional ByVal sep As String = ";") As Boolean
    If Len(item) = 0 Then Exit Function
    ListHas = (InStr(1, sep & lst & sep, sep & item & sep, vbTextCompare) > 0)
End Function

Public Function ListAdd(ByVal lst As String, ByVal item As String, Optional ByVal sep As String = ";") As String
    If Len(item) = 0 Or ListHas(lst, item, sep) Then
        ListAdd = lst
    ElseIf Len(lst) = 0 Then
        ListAdd = item
    Else
        ListAdd = lst & sep & item
    End If
End Function

Public Function ListRemove(ByVal lst As String, ByVal item As String, Optional ByVal sep As String = ";") As String
    Dim a As Variant, i As Long, res As String
    If Len(lst) = 0 Then Exit Function
    a = Split(lst, sep)
    For i = LBound(a) To UBound(a)
        If StrComp(a(i), item, vbTextCompare) <> 0 And Len(a(i)) > 0 Then
            res = res & IIf(Len(res) > 0, sep, "") & a(i)
        End If
    Next
    ListRemove = res
End Function

Public Function ListCount(ByVal lst As String, Optional ByVal sep As String = ";") As Long
    If Len(lst) = 0 Then Exit Function
    ListCount = UBound(Split(lst, sep)) + 1
End Function

Public Function ArrLen(ByVal a As Variant) As Long
    On Error GoTo nope
    If Not IsArray(a) Then Exit Function
    ArrLen = UBound(a) - LBound(a) + 1
    Exit Function
nope:
    ArrLen = 0
End Function

'--- Deutsche Darstellung -----------------------------------------------------
Public Function FmtInt(ByVal n As Double) As String
    Dim s As String, res As String, i As Long, neg As Boolean
    neg = (n < 0)
    s = Format$(Abs(Fix(n)), "0")
    For i = Len(s) To 1 Step -1
        res = Mid$(s, i, 1) & res
        If (Len(s) - i + 1) Mod 3 = 0 And i > 1 Then res = "." & res
    Next
    If neg Then res = "-" & res
    FmtInt = res
End Function

' Zahl mit festen Nachkommastellen, deutsch (1.234,56)
Public Function FmtNum(ByVal n As Double, Optional ByVal decimals As Long = 2, Optional ByVal trimZeros As Boolean = False) As String
    Dim r As Double, ip As Double, fp As Double, fs As String, neg As Boolean, s As String
    neg = (n < 0)
    r = Round(Abs(n), decimals)
    ip = Fix(r)
    s = FmtInt(ip)
    If decimals > 0 Then
        fp = Round((r - ip) * 10 ^ decimals, 0)
        If fp >= 10 ^ decimals Then
            ip = ip + 1: fp = 0: s = FmtInt(ip)
        End If
        fs = Format$(fp, String$(decimals, "0"))
        If trimZeros Then
            Do While Len(fs) > 0 And Right$(fs, 1) = "0"
                fs = Left$(fs, Len(fs) - 1)
            Loop
        End If
        If Len(fs) > 0 Then s = s & "," & fs
    End If
    If neg And (ip <> 0 Or fp <> 0) Then s = "-" & s
    FmtNum = s
End Function

Public Function FmtDate(ByVal d As Variant) As String
    If Not IsDate(d) And Not IsNumeric(d) Then FmtDate = NzS(d): Exit Function
    Dim dd As Date
    dd = CDate(d)
    FmtDate = Format$(Day(dd), "00") & "." & Format$(Month(dd), "00") & "." & Format$(Year(dd), "0000")
End Function

Public Function FmtDateTime(ByVal d As Variant, Optional ByVal withSeconds As Boolean = False) As String
    Dim dd As Date
    If Not IsDate(d) And Not IsNumeric(d) Then FmtDateTime = NzS(d): Exit Function
    dd = CDate(d)
    FmtDateTime = FmtDate(dd) & " " & Format$(Hour(dd), "00") & ":" & Format$(Minute(dd), "00")
    If withSeconds Then FmtDateTime = FmtDateTime & ":" & Format$(Second(dd), "00")
End Function

Public Function FmtStamp(ByVal d As Variant) As String
    If IsBlank(d) Then FmtStamp = "": Exit Function
    FmtStamp = FmtDate(d) & " " & Format$(Hour(CDate(d)), "00") & ":" & Format$(Minute(CDate(d)), "00")
End Function

Public Function MonthName_DE(ByVal m As Long) As String
    MonthName_DE = Choose(m, "Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", _
                          "September", "Oktober", "November", "Dezember")
End Function

Public Function WeekdayShort_DE(ByVal wd As Long) As String
    ' wd: 1 = Montag
    WeekdayShort_DE = Choose(wd, "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So")
End Function

Public Function WeekdayLong_DE(ByVal d As Date) As String
    WeekdayLong_DE = Choose(Weekday(d, vbMonday), "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag")
End Function

Public Function DateLong_DE(ByVal d As Date) As String
    DateLong_DE = WeekdayLong_DE(d) & ", " & Day(d) & ". " & MonthName_DE(Month(d)) & " " & Year(d)
End Function

Public Function IsoWeek(ByVal d As Date) As Long
    Dim t As Date
    t = DateSerial(Year(d - Weekday(d - 1) + 4), 1, 3)
    IsoWeek = Int((d - t + Weekday(t) + 5) / 7)
End Function

'--- Deutsche Eingaben erkennen -----------------------------------------------
Public Function ParseNumberDE(ByVal s As String, ByRef result As Double) As Boolean
    Dim t As String, i As Long, ch As String, nDot As Long, nCom As Long, lastDot As Long, lastCom As Long
    t = Trim$(s)
    t = Replace(t, " ", "")
    t = Replace(t, ChrW(160), "")
    t = Replace(t, "€", "")
    t = Replace(t, "%", "")
    If Len(t) = 0 Then Exit Function
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        Select Case ch
            Case "0" To "9"
            Case "."
                nDot = nDot + 1: lastDot = i
            Case ","
                nCom = nCom + 1: lastCom = i
            Case "-", "+"
                If i <> 1 Then Exit Function
            Case Else
                Exit Function
        End Select
    Next
    If nCom > 1 Then Exit Function
    If nCom = 1 Then
        If nDot > 0 And lastDot > lastCom Then Exit Function
        t = Replace(t, ".", "")
        t = Replace(t, ",", ".")
    ElseIf nDot > 1 Then
        t = Replace(t, ".", "")
    ElseIf nDot = 1 Then
        ' "1.234" = Tausendertrennung; "1.5" = Dezimalpunkt
        If Len(t) - lastDot = 3 And lastDot > 1 And Mid$(t, 1, 1) <> "0" Then
            t = Replace(t, ".", "")
        End If
    End If
    If t = "-" Or t = "+" Or t = "." Then Exit Function
    On Error GoTo nope
    result = Val(t)
    ParseNumberDE = True
    Exit Function
nope:
    ParseNumberDE = False
End Function

' Erkennt TT.MM.JJJJ, TT.MM.JJ, TT.MM., JJJJ-MM-TT, jeweils optional mit hh:mm[:ss]
Public Function ParseDateDE(ByVal s As String, ByRef result As Date, ByRef hasTime As Boolean) As Boolean
    Dim t As String, dpart As String, tpart As String, p As Long
    Dim a As Variant, d As Long, m As Long, y As Long, hh As Long, mi As Long, ss As Long
    t = Trim$(s)
    hasTime = False
    If Len(t) = 0 Then Exit Function
    t = Replace(t, "T", " ")
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    p = InStr(t, " ")
    If p > 0 Then
        dpart = Left$(t, p - 1): tpart = Trim$(Mid$(t, p + 1))
        If Right$(dpart, 1) = "," Then dpart = Left$(dpart, Len(dpart) - 1)
    Else
        dpart = t
        If InStr(t, ":") > 0 Then Exit Function
    End If
    If InStr(dpart, "-") > 0 Then
        a = Split(dpart, "-")
        If UBound(a) <> 2 Then Exit Function
        If Not (IsAllDigits(a(0)) And IsAllDigits(a(1)) And IsAllDigits(a(2))) Then Exit Function
        y = CLng(a(0)): m = CLng(a(1)): d = CLng(a(2))
    Else
        a = Split(dpart, ".")
        If UBound(a) < 1 Or UBound(a) > 2 Then Exit Function
        If Not (IsAllDigits(a(0)) And IsAllDigits(a(1))) Then Exit Function
        d = CLng(a(0)): m = CLng(a(1))
        If UBound(a) = 2 Then
            If Len(a(2)) = 0 Then
                y = Year(Date)
            ElseIf IsAllDigits(a(2)) Then
                y = CLng(a(2))
                If Len(a(2)) <= 2 Then y = 2000 + y
            Else
                Exit Function
            End If
        Else
            y = Year(Date)
        End If
    End If
    If y < 1900 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Then Exit Function
    If d > Day(DateSerial(y, m + 1, 0)) Then Exit Function
    If Len(tpart) > 0 Then
        a = Split(tpart, ":")
        If UBound(a) < 1 Or UBound(a) > 2 Then Exit Function
        If Not (IsAllDigits(a(0)) And IsAllDigits(a(1))) Then Exit Function
        hh = CLng(a(0)): mi = CLng(a(1))
        If UBound(a) = 2 Then
            If Not IsAllDigits(a(2)) Then Exit Function
            ss = CLng(a(2))
        End If
        If hh > 23 Or mi > 59 Or ss > 59 Then Exit Function
        hasTime = True
    End If
    result = DateSerial(y, m, d) + TimeSerial(hh, mi, ss)
    ParseDateDE = True
End Function

Public Function IsAllDigits(ByVal s As Variant) As Boolean
    Dim i As Long, sv As String
    sv = CStr(s)
    If Len(sv) = 0 Then Exit Function
    For i = 1 To Len(sv)
        If Mid$(sv, i, 1) < "0" Or Mid$(sv, i, 1) > "9" Then Exit Function
    Next
    IsAllDigits = True
End Function

' Zahl invariant (Punkt) fuer gespeicherte Konfiguration
Public Function NumInv(ByVal n As Double) As String
    NumInv = Trim$(Str$(n))
    If Left$(NumInv, 1) = "." Then NumInv = "0" & NumInv
    If Left$(NumInv, 2) = "-." Then NumInv = "-0" & Mid$(NumInv, 2)
End Function

Public Function InvNum(ByVal s As String) As Double
    InvNum = Val(Trim$(s))
End Function

Public Function DateInv(ByVal d As Date) As String
    DateInv = Format$(Year(d), "0000") & "-" & Format$(Month(d), "00") & "-" & Format$(Day(d), "00")
    If d - Int(d) > 0 Then DateInv = DateInv & " " & Format$(Hour(d), "00") & ":" & Format$(Minute(d), "00") & ":" & Format$(Second(d), "00")
End Function

Public Function InvDate(ByVal s As String) As Variant
    Dim d As Date, ht As Boolean
    If ParseDateDE(s, d, ht) Then InvDate = d Else InvDate = Empty
End Function

'--- Sortierung (stabil, Indizes) ---------------------------------------------
' keys: Variant-Array (1..n) je Index; dirs: 1 aufsteigend, -1 absteigend
' idx: Long-Array mit zu sortierenden Positionen
Public Sub MergeSortIdx(ByRef idx() As Long, ByRef keys() As Variant, ByVal nKeys As Long, ByRef dirs() As Long)
    Dim n As Long, tmp() As Long
    n = UBound(idx) - LBound(idx) + 1
    If n < 2 Then Exit Sub
    ReDim tmp(LBound(idx) To UBound(idx))
    MSort idx, tmp, LBound(idx), UBound(idx), keys, nKeys, dirs
End Sub

Private Sub MSort(ByRef a() As Long, ByRef tmp() As Long, ByVal lo As Long, ByVal hi As Long, ByRef keys() As Variant, ByVal nKeys As Long, ByRef dirs() As Long)
    Dim mid_ As Long, i As Long, j As Long, k As Long
    If hi - lo < 1 Then Exit Sub
    If hi - lo < 12 Then
        ' Einfuegesortierung fuer kleine Bereiche
        Dim x As Long
        For i = lo + 1 To hi
            x = a(i): j = i - 1
            Do While j >= lo
                If CmpKeys(a(j), x, keys, nKeys, dirs) <= 0 Then Exit Do
                a(j + 1) = a(j): j = j - 1
            Loop
            a(j + 1) = x
        Next
        Exit Sub
    End If
    mid_ = (lo + hi) \ 2
    MSort a, tmp, lo, mid_, keys, nKeys, dirs
    MSort a, tmp, mid_ + 1, hi, keys, nKeys, dirs
    i = lo: j = mid_ + 1: k = lo
    Do While i <= mid_ And j <= hi
        If CmpKeys(a(i), a(j), keys, nKeys, dirs) <= 0 Then
            tmp(k) = a(i): i = i + 1
        Else
            tmp(k) = a(j): j = j + 1
        End If
        k = k + 1
    Loop
    Do While i <= mid_: tmp(k) = a(i): i = i + 1: k = k + 1: Loop
    Do While j <= hi: tmp(k) = a(j): j = j + 1: k = k + 1: Loop
    For k = lo To hi: a(k) = tmp(k): Next
End Sub

Private Function CmpKeys(ByVal r1 As Long, ByVal r2 As Long, ByRef keys() As Variant, ByVal nKeys As Long, ByRef dirs() As Long) As Long
    Dim kk As Long, c As Long, ea As Boolean, eb As Boolean
    For kk = 1 To nKeys
        ea = IsBlank(keys(kk)(r1)): eb = IsBlank(keys(kk)(r2))
        If ea Or eb Then
            ' leere Werte stehen unabhaengig von der Richtung am Ende
            If ea And Not eb Then CmpKeys = 1: Exit Function
            If eb And Not ea Then CmpKeys = -1: Exit Function
        Else
            c = CmpVal(keys(kk)(r1), keys(kk)(r2))
            If c <> 0 Then CmpKeys = c * dirs(kk): Exit Function
        End If
    Next
End Function

' Leere Werte immer am Ende (unabhaengig von der Richtung wird das ueber dirs korrigiert)
Public Function CmpVal(ByVal a As Variant, ByVal b As Variant) As Long
    Dim ea As Boolean, eb As Boolean
    ea = IsBlank(a): eb = IsBlank(b)
    If ea And eb Then Exit Function
    If ea Then CmpVal = 1: Exit Function
    If eb Then CmpVal = -1: Exit Function
    If IsError(a) Or IsError(b) Then Exit Function
    If VarType(a) = vbString Or VarType(b) = vbString Then
        CmpVal = StrComp(CStr(a), CStr(b), vbTextCompare)
    Else
        If a < b Then
            CmpVal = -1
        ElseIf a > b Then
            CmpVal = 1
        End If
    End If
End Function

'--- Dateien ------------------------------------------------------------------
Public Function FileExists(ByVal p As String) As Boolean
    On Error GoTo nope
    If Len(p) = 0 Then Exit Function
    FileExists = ((GetAttr(p) And vbDirectory) = 0)
    Exit Function
nope:
End Function

Public Function DirExists(ByVal p As String) As Boolean
    On Error GoTo nope
    If Len(p) = 0 Then Exit Function
    DirExists = ((GetAttr(p) And vbDirectory) = vbDirectory)
    Exit Function
nope:
End Function

Public Function EnsureDir(ByVal p As String) As Boolean
    On Error GoTo nope
    If DirExists(p) Then EnsureDir = True: Exit Function
    MkDir p
    EnsureDir = DirExists(p)
    Exit Function
nope:
    EnsureDir = False
End Function

Public Function SafeFileName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next
    s = Trim$(s)
    If Len(s) = 0 Then s = "Export"
    If Len(s) > 80 Then s = Left$(s, 80)
    SafeFileName = s
End Function

Public Function StampForFile(Optional ByVal d As Date = 0) As String
    If d = 0 Then d = Now
    StampForFile = Format$(Year(d), "0000") & "-" & Format$(Month(d), "00") & "-" & Format$(Day(d), "00") & "_" & _
                   Format$(Hour(d), "00") & Format$(Minute(d), "00") & Format$(Second(d), "00")
End Function

Public Function UniquePath(ByVal folder As String, ByVal baseName As String, ByVal ext As String) As String
    Dim p As String, i As Long
    p = folder & "\" & baseName & ext
    Do While FileExists(p)
        i = i + 1
        p = folder & "\" & baseName & " (" & i & ")" & ext
    Loop
    UniquePath = p
End Function

Public Function UserId() As String
    UserId = Environ$("USERNAME")
    If Len(UserId) = 0 Then UserId = Application.UserName
End Function

'--- UTF-8 ----------------------------------------------------------------------
Public Function Utf8Bytes(ByVal s As String) As Byte()
    Dim b() As Byte, n As Long, i As Long, c As Long, c2 As Long
    If Len(s) = 0 Then
        Utf8Bytes = b
        Exit Function
    End If
    ReDim b(0 To Len(s) * 4)
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c >= &HD800& And c <= &HDBFF& And i < Len(s) Then
            c2 = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
            If c2 >= &HDC00& And c2 <= &HDFFF& Then
                c = &H10000 + (c - &HD800&) * &H400& + (c2 - &HDC00&)
                i = i + 1
            End If
        End If
        If c < &H80& Then
            b(n) = c: n = n + 1
        ElseIf c < &H800& Then
            b(n) = &HC0& Or (c \ &H40&): b(n + 1) = &H80& Or (c And &H3F&): n = n + 2
        ElseIf c < &H10000 Then
            b(n) = &HE0& Or (c \ &H1000&): b(n + 1) = &H80& Or ((c \ &H40&) And &H3F&)
            b(n + 2) = &H80& Or (c And &H3F&): n = n + 3
        Else
            b(n) = &HF0& Or (c \ &H40000): b(n + 1) = &H80& Or ((c \ &H1000&) And &H3F&)
            b(n + 2) = &H80& Or ((c \ &H40&) And &H3F&): b(n + 3) = &H80& Or (c And &H3F&): n = n + 4
        End If
    Next
    ReDim Preserve b(0 To n - 1)
    Utf8Bytes = b
End Function

' Dekodiert UTF-8; bei ungueltigen Folgen ok=False
Public Function Utf8Decode(ByRef b() As Byte, ByVal startPos As Long, ByRef ok As Boolean) As String
    Dim i As Long, n As Long, c As Long, need As Long, out() As Integer, o As Long, cp As Long
    n = UBound(b)
    ReDim out(0 To (n - startPos + 1) * 2 + 1)
    ok = True
    i = startPos
    Do While i <= n
        c = b(i)
        If c < &H80& Then
            cp = c: need = 0
        ElseIf (c And &HE0&) = &HC0& Then
            cp = c And &H1F&: need = 1
        ElseIf (c And &HF0&) = &HE0& Then
            cp = c And &HF&: need = 2
        ElseIf (c And &HF8&) = &HF0& Then
            cp = c And &H7&: need = 3
        Else
            ok = False: Exit Function
        End If
        If i + need > n Then ok = False: Exit Function
        Do While need > 0
            i = i + 1
            If (b(i) And &HC0&) <> &H80& Then ok = False: Exit Function
            cp = cp * &H40& + (b(i) And &H3F&)
            need = need - 1
        Loop
        If cp >= &H10000 Then
            cp = cp - &H10000
            out(o) = ToInt(&HD800& + (cp \ &H400&)): o = o + 1
            out(o) = ToInt(&HDC00& + (cp And &H3FF&)): o = o + 1
        Else
            out(o) = ToInt(cp): o = o + 1
        End If
        i = i + 1
    Loop
    If o = 0 Then Utf8Decode = "": Exit Function
    Dim s As String
    s = String$(o, " ")
    For i = 0 To o - 1
        Mid$(s, i + 1, 1) = ChrW(out(i))
    Next
    Utf8Decode = s
End Function

Private Function ToInt(ByVal v As Long) As Integer
    If v > 32767 Then ToInt = CInt(v - 65536) Else ToInt = CInt(v)
End Function

Public Function Ansi1252Decode(ByRef b() As Byte, ByVal startPos As Long) As String
    Dim i As Long, s As String
    s = String$(UBound(b) - startPos + 1, " ")
    For i = startPos To UBound(b)
        Mid$(s, i - startPos + 1, 1) = Cp1252Char(b(i))
    Next
    Ansi1252Decode = s
End Function

Private Function Cp1252Char(ByVal c As Byte) As String
    Select Case c
        Case &H80: Cp1252Char = ChrW(&H20AC)
        Case &H84: Cp1252Char = ChrW(&H201E)
        Case &H93: Cp1252Char = ChrW(&H201C)
        Case &H94: Cp1252Char = ChrW(&H201D)
        Case &H96: Cp1252Char = ChrW(&H2013)
        Case &H97: Cp1252Char = ChrW(&H2014)
        Case Else: Cp1252Char = ChrW(c)
    End Select
End Function

Public Function ReadFileBytes(ByVal p As String) As Byte()
    Dim f As Integer, b() As Byte, n As Long
    f = FreeFile
    Open p For Binary Access Read Shared As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, 1, b
    End If
    Close #f
    ReadFileBytes = b
End Function

Public Sub WriteFileBytes(ByVal p As String, ByRef b() As Byte)
    Dim f As Integer
    If FileExists(p) Then Kill p
    f = FreeFile
    Open p For Binary Access Write As #f
    If ArrLen(b) > 0 Then Put #f, 1, b
    Close #f
End Sub

' Schreibt Text als UTF-8 (mit BOM, wenn gewuenscht)
Public Sub WriteUtf8File(ByVal p As String, ByVal text As String, Optional ByVal withBom As Boolean = True)
    Dim body() As Byte, allB() As Byte, i As Long, n As Long
    body = Utf8Bytes(text)
    n = ArrLen(body)
    If withBom Then
        ReDim allB(0 To n + 2)
        allB(0) = &HEF: allB(1) = &HBB: allB(2) = &HBF
        For i = 0 To n - 1: allB(i + 3) = body(i): Next
        WriteFileBytes p, allB
    Else
        WriteFileBytes p, body
    End If
End Sub

'--- CRC32 (fuer ZIP) -------------------------------------------------------------
Public Function Crc32(ByRef b() As Byte) As Long
    Dim i As Long, c As Long, n As Long
    If Not mCrcReady Then InitCrc
    c = &HFFFFFFFF
    n = ArrLen(b)
    For i = 0 To n - 1
        c = mCrcTable((c Xor b(i)) And &HFF&) Xor ShrL8(c)
    Next
    Crc32 = c Xor &HFFFFFFFF
End Function

Private Sub InitCrc()
    Dim i As Long, j As Long, c As Long
    For i = 0 To 255
        c = i
        For j = 1 To 8
            If (c And 1) Then
                c = ShrL1(c) Xor &HEDB88320
            Else
                c = ShrL1(c)
            End If
        Next
        mCrcTable(i) = c
    Next
    mCrcReady = True
End Sub

Private Function ShrL1(ByVal v As Long) As Long
    If v < 0 Then
        ShrL1 = ((v And &H7FFFFFFF) \ 2) Or &H40000000
    Else
        ShrL1 = v \ 2
    End If
End Function

Private Function ShrL8(ByVal v As Long) As Long
    If v < 0 Then
        ShrL8 = ((v And &H7FFFFFFF) \ &H100&) Or &H800000
    Else
        ShrL8 = v \ &H100&
    End If
End Function

'--- Textvergleich fuer Dubletten/Eindeutigkeit ---------------------------------
Public Function NormKey(ByVal v As Variant) As String
    If IsBlank(v) Then Exit Function
    If VarType(v) = vbString Then
        NormKey = LCase$(Trim$(v))
        Do While InStr(NormKey, "  ") > 0: NormKey = Replace(NormKey, "  ", " "): Loop
    ElseIf VarType(v) = vbDate Then
        NormKey = "d" & NumInv(CDbl(v))
    ElseIf VarType(v) = vbBoolean Then
        NormKey = IIf(v, "b1", "b0")
    Else
        NormKey = "n" & NumInv(CDbl(v))
    End If
End Function

Public Function Truncate(ByVal s As String, ByVal n As Long) As String
    s = Replace(Replace(s, vbCr, " "), vbLf, " ")
    If Len(s) > n Then Truncate = Left$(s, n - 1) & ChrW(&H2026) Else Truncate = s
End Function

Public Function Plural(ByVal n As Double, ByVal one As String, ByVal many As String) As String
    If n = 1 Then Plural = "1 " & one Else Plural = FmtInt(n) & " " & many
End Function
