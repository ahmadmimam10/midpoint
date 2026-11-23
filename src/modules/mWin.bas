Attribute VB_Name = "mWin"
'==============================================================================
' mWin - Windows-API fuer Fensterverankerung, DPI und Fokus (32/64 Bit, VBA7)
' Ausschliesslich lokale Systembibliotheken (user32, gdi32). Kein Netzwerk.
'==============================================================================
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetFocusApi Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetDpiForWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMetricsForDpi Lib "user32" (ByVal nIndex As Long, ByVal dpi As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#If Win64 Then
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_CLIPCHILDREN As Long = &H2000000
Private Const LOGPIXELSX As Long = 88
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10

Private Declare PtrSafe Function SendMessageA Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long
Private Const WM_SETREDRAW As Long = &HB
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4
Private Const RDW_ALLCHILDREN As Long = &H80
Private Const RDW_UPDATENOW As Long = &H100
Private Const RDW_FRAME As Long = &H400

Private Declare PtrSafe Function SystemParametersInfoA Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA As Long = 48

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private mDpi As Long

' Linke Maustaste gerade gedrueckt (z. B. waehrend eine Form gezogen wird)
Public Function MouseDown() As Boolean
    MouseDown = (GetAsyncKeyState(1) And &H8000) <> 0
End Function

Private Function WorkArea() As RECT
    Dim r As RECT
    SystemParametersInfoA SPI_GETWORKAREA, 0, r, 0
    WorkArea = r
End Function

Public Function WorkAreaHeightPt() As Single
    Dim r As RECT
    r = WorkArea()
    WorkAreaHeightPt = PxToPt(r.Bottom - r.Top)
    If WorkAreaHeightPt < 200 Then WorkAreaHeightPt = Application.UsableHeight
End Function

Public Function WorkAreaWidthPt() As Single
    Dim r As RECT
    r = WorkArea()
    WorkAreaWidthPt = PxToPt(r.Right - r.Left)
    If WorkAreaWidthPt < 300 Then WorkAreaWidthPt = Application.UsableWidth
End Function

'--- DPI des Bildschirms (96 = 100 %, 120 = 125 %) ---------------------------
Public Function ScreenDpi() As Long
    Dim hDC As LongPtr
    If mDpi = 0 Then
        hDC = GetDC(0)
        mDpi = GetDeviceCaps(hDC, LOGPIXELSX)
        ReleaseDC 0, hDC
        If mDpi <= 0 Then mDpi = 96
    End If
    ScreenDpi = mDpi
End Function

' Breite der senkrechten Bildlaufleiste in Punkten (Excels UsableWidth schließt sie ein)
Public Function VScrollWidthPt() As Double
    Dim d As Long, px As Long
    d = ScreenDpi()
    On Error Resume Next
    px = GetSystemMetricsForDpi(2, d)                    ' SM_CXVSCROLL
    On Error GoTo 0
    If px <= 0 Then px = CLng(17# * d / 96#)
    VScrollWidthPt = px * 72# / d
End Function

' DPI eines Fensters (Windows 10 und neuer: je Bildschirm); sonst die des Hauptbildschirms
Public Function WindowDpi(ByVal hWnd As LongPtr) As Long
    On Error GoTo fallback
    If hWnd <> 0 Then WindowDpi = GetDpiForWindow(hWnd)
    If WindowDpi >= 72 Then Exit Function
fallback:
    WindowDpi = ScreenDpi()
End Function

Public Function PtToPx(ByVal pt As Double) As Long
    PtToPx = CLng(pt * ScreenDpi() / 72#)
End Function

Public Function PxToPt(ByVal px As Long) As Double
    PxToPt = px * 72# / ScreenDpi()
End Function

'--- Fenster der Arbeitsmappe (EXCEL7 innerhalb XLDESK des aktiven XLMAIN) -----
Public Function WorkbookGridHwnd() As LongPtr
    Dim hMain As LongPtr, hDesk As LongPtr
    hMain = Application.hWnd
    hDesk = FindWindowExA(hMain, 0, "XLDESK", vbNullString)
    If hDesk = 0 Then Exit Function
    WorkbookGridHwnd = FindWindowExA(hDesk, 0, "EXCEL7", vbNullString)
End Function

Public Function FormHwnd(ByVal caption As String) As LongPtr
    FormHwnd = FindWindowA("ThunderDFrame", caption)
End Function

Public Function HwndAlive(ByVal h As LongPtr) As Boolean
    If h = 0 Then Exit Function
    HwndAlive = (IsWindow(h) <> 0)
End Function

Public Function ParentOf(ByVal h As LongPtr) As LongPtr
    ParentOf = GetParent(h)
End Function

' Macht ein modelloses Formular zum rahmenlosen Kindfenster von hParent.
Public Sub DockAsChild(ByVal hForm As LongPtr, ByVal hParent As LongPtr)
    Dim st As LongPtr
    st = WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN
    SetWindowLongPtr hForm, GWL_STYLE, st
    SetWindowLongPtr hForm, GWL_EXSTYLE, 0
    SetParent hForm, hParent
    DrawMenuBar hForm
End Sub

Public Sub PlaceWindow(ByVal h As LongPtr, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal hgt As Long)
    MoveWindow h, x, y, w, hgt, 1
End Sub

Public Sub BringToTop(ByVal h As LongPtr)
    SetWindowPos h, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
End Sub

Public Function ClientHeightPx(ByVal h As LongPtr) As Long
    Dim r As RECT
    GetClientRect h, r
    ClientHeightPx = r.Bottom - r.Top
End Function

Public Function ClientWidthPx(ByVal h As LongPtr) As Long
    Dim r As RECT
    GetClientRect h, r
    ClientWidthPx = r.Right - r.Left
End Function

' Zeichnen eines Fensters aussetzen (Neuaufbau ohne Flackern) und danach in einem Zug zeichnen
Public Sub SetRedraw(ByVal h As LongPtr, ByVal enable As Boolean)
    If h = 0 Then Exit Sub
    SendMessageA h, WM_SETREDRAW, IIf(enable, 1, 0), 0
    If enable Then RedrawWindow h, 0, 0, RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN Or RDW_UPDATENOW Or RDW_FRAME
End Sub

' Fenster samt Kindfenstern sofort vollständig neu zeichnen (z. B. nach neuem Hintergrundbild)
Public Sub Repaint(ByVal h As LongPtr)
    If h = 0 Then Exit Sub
    RedrawWindow h, 0, 0, RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN Or RDW_UPDATENOW Or RDW_FRAME
End Sub

' Nur aktivieren (kein SetFocus: das umginge die Fokusverwaltung der Formularfelder)
Public Sub ActivateWindow(ByVal h As LongPtr)
    If h = 0 Then Exit Sub
    SetForegroundWindow h
End Sub

Public Sub FocusWindow(ByVal h As LongPtr)
    If h <> 0 Then SetFocusApi h
End Sub

' Excel-Hauptfenster in den Vordergrund (nach modalen Dialogen), damit der naechste Klick sofort wirkt
Public Sub ForegroundExcel()
    On Error Resume Next
    ReleaseCapture
    SetForegroundWindow Application.hWnd
End Sub

