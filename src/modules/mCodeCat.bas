Attribute VB_Name = "mCodeCat"
'==============================================================================
' mCodeCat - Eingebetteter Codekatalog: Der vollstaendige Quellcode wird beim
' Build in das Blatt _Code geschrieben und ist ohne Zugriff auf das VBA-
' Projektobjektmodell lesbar und exportierbar (ZIP, HTML, Markdown, Text).
' ZIP wird ohne externe Komponenten erzeugt (unkomprimiert, mit CRC32).
'==============================================================================
Option Explicit

' Liste der Module: Array(Name, Art, Zeilen)
Public Function Modules() As Collection
    Dim ws As Worksheet, r As Long, lastR As Long, res As New Collection, cur As String, cnt As Long, kind As String
    Set ws = ThisWorkbook.Worksheets(SH_CODE)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value) <> cur Then
            If Len(cur) > 0 Then res.Add Array(cur, kind, cnt)
            cur = CStr(ws.Cells(r, 1).Value)
            kind = CStr(ws.Cells(r, 2).Value)
            cnt = 0
        End If
        cnt = cnt + 1
    Next
    If Len(cur) > 0 Then res.Add Array(cur, kind, cnt)
    Set Modules = res
End Function

Public Function ModuleText(ByVal nm As String) As String
    Dim ws As Worksheet, r As Long, lastR As Long, parts() As String, n As Long, v As Variant
    Set ws = ThisWorkbook.Worksheets(SH_CODE)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function
    v = ws.Range("A2:D" & lastR).Value
    ReDim parts(1 To lastR)
    For r = 1 To UBound(v, 1)
        If CStr(v(r, 1)) = nm Then
            n = n + 1
            parts(n) = NzS(v(r, 4))
        End If
    Next
    If n = 0 Then Exit Function
    ReDim Preserve parts(1 To n)
    ModuleText = Join(parts, vbCrLf)
End Function

Public Function CatalogVersion() As String
    CatalogVersion = mCfg.SysGet("CodeKatalog", "")
End Function

Public Function TotalLines() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_CODE)
    TotalLines = Application.Max(0, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1)
End Function

Private Function FileExt(ByVal kind As String) As String
    Select Case kind
        Case "Klasse": FileExt = ".cls"
        Case "Formular": FileExt = ".frm.vba"
        Case "Dokument": FileExt = ".cls"
        Case Else: FileExt = ".bas"
    End Select
End Function

'--- Anzeige ------------------------------------------------------------------------------
Public Sub ShowCode()
    Dim f As frmCode
    Set f = New frmCode
    f.Setup
    f.Show: mNav.FocusGrid
    Unload f
End Sub

'--- Exporte ------------------------------------------------------------------------------
Public Sub ExportAs(ByVal fmt As String)
    Dim path As String, m As Variant, s As String, parts As New Collection, html As String, toc As String, i As Long
    Dim hdr As String
    hdr = APP_TITLE & ": Quellcode, Programmversion " & PROG_VERSION & ", Formatversion " & FORMAT_VERSION & ", Katalog " & CatalogVersion()
    Select Case fmt
        Case "ZIP"
            path = mIO.AskExportPath(APP_TITLE & "_Quellcode_v" & PROG_VERSION, ".zip")
            If Len(path) = 0 Then Exit Sub
            WriteZip path
        Case "TXT"
            path = mIO.AskExportPath(APP_TITLE & "_Quellcode_v" & PROG_VERSION, ".txt")
            If Len(path) = 0 Then Exit Sub
            s = hdr & vbCrLf & String$(78, "=") & vbCrLf
            For Each m In Modules()
                s = s & vbCrLf & String$(78, "-") & vbCrLf & "' Datei: " & m(0) & FileExt(CStr(m(1))) & "  (" & m(1) & ", " & m(2) & " Zeilen)" & vbCrLf & String$(78, "-") & vbCrLf
                s = s & ModuleText(CStr(m(0))) & vbCrLf
            Next
            WriteUtf8File path, s, True
        Case "MD"
            path = mIO.AskExportPath(APP_TITLE & "_Quellcode_v" & PROG_VERSION, ".md")
            If Len(path) = 0 Then Exit Sub
            s = "# " & hdr & vbCrLf & vbCrLf & "## Inhalt" & vbCrLf & vbCrLf
            For Each m In Modules()
                s = s & "- " & m(0) & FileExt(CStr(m(1))) & " (" & m(1) & ", " & m(2) & " Zeilen)" & vbCrLf
            Next
            For Each m In Modules()
                s = s & vbCrLf & "## " & m(0) & FileExt(CStr(m(1))) & vbCrLf & vbCrLf & "```vb" & vbCrLf & ModuleText(CStr(m(0))) & vbCrLf & "```" & vbCrLf
            Next
            WriteUtf8File path, s, True
        Case "HTML"
            path = mIO.AskExportPath(APP_TITLE & "_Quellcode_v" & PROG_VERSION, ".html")
            If Len(path) = 0 Then Exit Sub
            For Each m In Modules()
                i = i + 1
                toc = toc & "<li><a href=""#m" & i & """>" & HtmlEsc(m(0) & FileExt(CStr(m(1)))) & "</a> <small>(" & m(1) & ", " & m(2) & " Zeilen)</small></li>"
            Next
            html = "<!DOCTYPE html><html lang=""de""><head><meta charset=""utf-8""><title>" & HtmlEsc(hdr) & "</title>" & _
                   "<style>body{font-family:Segoe UI,Arial,sans-serif;color:#1f2937;margin:32px;max-width:1100px}h1{font-size:22px}h2{font-size:17px;margin-top:32px;border-bottom:1px solid #e5e7eb}" & _
                   "pre{background:#f9fafb;border:1px solid #e5e7eb;padding:12px;overflow:auto;font-family:Consolas,monospace;font-size:12px;line-height:1.4}small{color:#6b7280}</style></head><body>" & _
                   "<h1>" & HtmlEsc(hdr) & "</h1><p><small>Erzeugt am " & FmtDateTime(Now) & " aus dem eingebetteten Codekatalog. " & HtmlEsc(AUTHOR_NOTE) & "</small></p><h2>Inhalt</h2><ol>" & toc & "</ol>"
            i = 0
            For Each m In Modules()
                i = i + 1
                html = html & "<h2 id=""m" & i & """>" & HtmlEsc(m(0) & FileExt(CStr(m(1)))) & "</h2><pre>" & HtmlEsc(ModuleText(CStr(m(0)))) & "</pre>"
            Next
            html = html & "</body></html>"
            WriteUtf8File path, html, False
    End Select
    mLog.Add "Codeexport", "", "", "", fmt
    mUI.Alert "Quellcode exportiert:" & vbLf & path, "Codeexport"
End Sub

Private Function HtmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HtmlEsc = Replace(s, """", "&quot;")
End Function

'--- ZIP (Methode "stored") -------------------------------------------------------------------
Private Sub WriteZip(ByVal path As String)
    Dim f As Integer, m As Variant, names As New Collection, datas As New Collection, i As Long, b() As Byte, nm() As Byte
    Dim offsets() As Long, crcs() As Long, sizes() As Long, n As Long, pos As Long, cdStart As Long, cdSize As Long, manifest As String
    Dim dosTime As Long, dosDate As Long, d As Date
    For Each m In Modules()
        names.Add "src/" & m(0) & FileExt(CStr(m(1)))
        datas.Add Utf8Bytes(ModuleText(CStr(m(0))) & vbCrLf)
    Next
    manifest = "Manifest " & APP_TITLE & vbCrLf & "Programmversion: " & PROG_VERSION & vbCrLf & "Formatversion: " & FORMAT_VERSION & vbCrLf & _
               "Codekatalog: " & CatalogVersion() & vbCrLf & "Erzeugt: " & FmtDateTime(Now) & vbCrLf & vbCrLf & "Datei;Bytes;CRC32" & vbCrLf
    For i = 1 To names.Count
        b = datas(i)
        manifest = manifest & names(i) & ";" & (ArrLen(b)) & ";" & Hex$(Crc32(b)) & vbCrLf
    Next
    manifest = manifest & vbCrLf & "Hinweis: Die Formulare werden zur Laufzeit aus Code aufgebaut und benötigen keine binären .frx-Ressourcen." & vbCrLf
    names.Add "MANIFEST.txt"
    datas.Add Utf8Bytes(manifest)
    n = names.Count
    ReDim offsets(1 To n): ReDim crcs(1 To n): ReDim sizes(1 To n)
    d = Now
    dosTime = (Hour(d) * 2048&) + (Minute(d) * 32&) + (Second(d) \ 2)
    dosDate = ((Year(d) - 1980) * 512&) + (Month(d) * 32&) + Day(d)
    If FileExists(path) Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    pos = 1
    For i = 1 To n
        b = datas(i)
        nm = Utf8Bytes(CStr(names(i)))
        crcs(i) = Crc32(b)
        sizes(i) = ArrLen(b)
        offsets(i) = pos - 1
        PutLong f, pos, &H4034B50
        PutInt f, pos, 20: PutInt f, pos, &H800: PutInt f, pos, 0
        PutInt f, pos, dosTime: PutInt f, pos, dosDate
        PutLong f, pos, crcs(i): PutLong f, pos, sizes(i): PutLong f, pos, sizes(i)
        PutInt f, pos, ArrLen(nm): PutInt f, pos, 0
        Put #f, pos, nm: pos = pos + ArrLen(nm)
        If sizes(i) > 0 Then Put #f, pos, b: pos = pos + sizes(i)
    Next
    cdStart = pos - 1
    For i = 1 To n
        nm = Utf8Bytes(CStr(names(i)))
        PutLong f, pos, &H2014B50
        PutInt f, pos, 20: PutInt f, pos, 20: PutInt f, pos, &H800: PutInt f, pos, 0
        PutInt f, pos, dosTime: PutInt f, pos, dosDate
        PutLong f, pos, crcs(i): PutLong f, pos, sizes(i): PutLong f, pos, sizes(i)
        PutInt f, pos, ArrLen(nm): PutInt f, pos, 0: PutInt f, pos, 0: PutInt f, pos, 0: PutInt f, pos, 0
        PutLong f, pos, 0: PutLong f, pos, offsets(i)
        Put #f, pos, nm: pos = pos + ArrLen(nm)
    Next
    cdSize = pos - 1 - cdStart
    PutLong f, pos, &H6054B50
    PutInt f, pos, 0: PutInt f, pos, 0: PutInt f, pos, n: PutInt f, pos, n
    PutLong f, pos, cdSize: PutLong f, pos, cdStart
    PutInt f, pos, 0
    Close #f
End Sub

Private Sub PutLong(ByVal f As Integer, ByRef pos As Long, ByVal v As Long)
    Put #f, pos, v
    pos = pos + 4
End Sub

Private Sub PutInt(ByVal f As Integer, ByRef pos As Long, ByVal v As Long)
    Dim i As Integer
    If v > 32767 Then i = CInt(v - 65536) Else i = CInt(v)
    Put #f, pos, i
    pos = pos + 2
End Sub
