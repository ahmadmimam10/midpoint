Attribute VB_Name = "mGdi"
'==============================================================================
' mGdi - Zeichnet den Navigationshintergrund mit offener, weich gerundeter
' Aktivmarkierung (GDI+ aus Windows, 6-fache Abtastung nur der Streifen mit
' Rundungen, hochwertig verkleinert).
' Ergebnis ist ein IPictureDisp fuer UserForm.Picture. Farben/Groessen/DPI
' sind Parameter; nichts ist fest eingebaut.
'==============================================================================
Option Explicit

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type PICTDESC
    cbSizeofstruct As Long
    picType As Long
    hImage As LongPtr
    xExt As LongPtr
End Type

Private Type GUIDT
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, inputbuf As GdiplusStartupInput, ByVal outputbuf As LongPtr) As Long
Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal w As Long, ByVal h As Long, ByVal stride As Long, ByVal pf As Long, ByVal scan0 As LongPtr, bmp As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal img As LongPtr, g As LongPtr) As Long
Private Declare PtrSafe Function GdipSetSmoothingMode Lib "gdiplus" (ByVal g As LongPtr, ByVal m As Long) As Long
Private Declare PtrSafe Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal g As LongPtr, ByVal m As Long) As Long
Private Declare PtrSafe Function GdipSetInterpolationMode Lib "gdiplus" (ByVal g As LongPtr, ByVal m As Long) As Long
Private Declare PtrSafe Function GdipGraphicsClear Lib "gdiplus" (ByVal g As LongPtr, ByVal argb As Long) As Long
Private Declare PtrSafe Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, brush As LongPtr) As Long
Private Declare PtrSafe Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As LongPtr) As Long
Private Declare PtrSafe Function GdipCreatePath Lib "gdiplus" (ByVal fillMode As Long, path As LongPtr) As Long
Private Declare PtrSafe Function GdipDeletePath Lib "gdiplus" (ByVal path As LongPtr) As Long
Private Declare PtrSafe Function GdipAddPathArc Lib "gdiplus" (ByVal path As LongPtr, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare PtrSafe Function GdipAddPathLine Lib "gdiplus" (ByVal path As LongPtr, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare PtrSafe Function GdipClosePathFigure Lib "gdiplus" (ByVal path As LongPtr) As Long
Private Declare PtrSafe Function GdipFillPath Lib "gdiplus" (ByVal g As LongPtr, ByVal brush As LongPtr, ByVal path As LongPtr) As Long
Private Declare PtrSafe Function GdipFillRectangle Lib "gdiplus" (ByVal g As LongPtr, ByVal brush As LongPtr, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single) As Long
Private Declare PtrSafe Function GdipDrawImageRectI Lib "gdiplus" (ByVal g As LongPtr, ByVal img As LongPtr, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Long
Private Declare PtrSafe Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal g As LongPtr, ByVal img As LongPtr, ByVal dx As Long, ByVal dy As Long, ByVal dw As Long, ByVal dh As Long, _
    ByVal sx As Long, ByVal sy As Long, ByVal sw As Long, ByVal sh As Long, ByVal srcUnit As Long, ByVal attr As LongPtr, ByVal cb As LongPtr, ByVal cbData As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateImageAttributes Lib "gdiplus" (attr As LongPtr) As Long
Private Declare PtrSafe Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal attr As LongPtr, ByVal wrap As Long, ByVal argb As Long, ByVal clamp As Long) As Long
Private Declare PtrSafe Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal attr As LongPtr) As Long
Private Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal g As LongPtr) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal img As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bmp As LongPtr, hbm As LongPtr, ByVal background As Long) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (desc As PICTDESC, riid As GUIDT, ByVal fOwn As Long, ipic As IPicture) As Long

Private Const PF32ARGB As Long = &H26200A
Private Const SS As Long = 6   ' Abtastfaktor

' Fein gezeichneten Streifen verkleinert einsetzen. Randmodus "gespiegelt": Der Filter
' liest an Ober- und Unterkante sonst Pixel ausserhalb des Bildes (durchsichtig) mit -
' auf hellen Farben als feine helle Linie ueber/unter der Markierung sichtbar.
Private Sub DrawStrip(ByVal g As LongPtr, ByVal img As LongPtr, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim attr As LongPtr
    If GdipCreateImageAttributes(attr) <> 0 Or attr = 0 Then
        GdipDrawImageRectI g, img, x, y, w, h
        Exit Sub
    End If
    GdipSetImageAttributesWrapMode attr, 3, 0, 0
    GdipDrawImageRectRectI g, img, x, y, w, h, 0, 0, w * SS, h * SS, 2, attr, 0, 0
    GdipDisposeImageAttributes attr
End Sub

' VBA-Farbe (BGR) -> GDI+ ARGB (deckend)
Private Function Argb(ByVal vbColor As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = vbColor And &HFF&
    g = (vbColor \ &H100&) And &HFF&
    b = (vbColor \ &H10000) And &HFF&
    Argb = &HFF000000 Or (r * &H10000) Or (g * &H100&) Or b
End Function

' Zeichnet Hintergrund (bg) und eine Aktivform (act) ab y=topPx mit Hoehe hPx.
' Linke Seite rund, rechts offen mit konkaven Uebergaengen (Radius rcPx).
' topPx < 0 bedeutet: keine Aktivmarkierung.
Public Function NavPicture(ByVal wPx As Long, ByVal hPx As Long, ByVal bg As Long, ByVal act As Long, _
                           ByVal topPx As Long, ByVal itemHPx As Long, ByVal insetPx As Long, ByVal rcPx As Long, _
                           Optional ByVal tgX As Long = -1, Optional ByVal tgY As Long = 0, Optional ByVal tgW As Long = 0, _
                           Optional ByVal tgH As Long = 0, Optional ByVal tgOn As Boolean = False, Optional ByVal tgColor As Long = 0) As IPictureDisp
    Dim si As GdiplusStartupInput, token As LongPtr, brT As LongPtr, d As Single
    Dim big As LongPtr, gBig As LongPtr, small As LongPtr, gSmall As LongPtr
    Dim brBg As LongPtr, brAct As LongPtr, path As LongPtr, hbm As LongPtr
    Dim W As Single, T As Single, H As Single, L As Single, r As Single, rc As Single, visW As Long

    Dim y0 As Long, sh As Long, offC As Long
    If wPx < 4 Or hPx < 4 Then Exit Function
    si.GdiplusVersion = 1
    If GdiplusStartup(token, si, 0) <> 0 Then Exit Function
    ' Bild rechts um einige Pixel breiter als sichtbar: das Verkleinern mischt die letzte
    ' Spalte mit "nichts" (graue Kante) - diese Spalte liegt so ausserhalb des Fensters.
    ' Die Kurven richten sich weiter nach der sichtbaren Breite (visW).
    visW = wPx
    wPx = wPx + 6
    offC = Blend(bg, IIf(IsLightColor(bg), RGB(0, 0, 0), RGB(255, 255, 255)), IIf(IsLightColor(bg), 0.14, 0.3))

    ' Zielbild in Originalgroesse; nur die Streifen mit Rundungen werden mehrfach
    ' abgetastet gezeichnet (statt des ganzen Bildes) - deutlich schneller
    GdipCreateBitmapFromScan0 wPx, hPx, 0, PF32ARGB, 0, small
    GdipGetImageGraphicsContext small, gSmall
    GdipSetInterpolationMode gSmall, 7
    GdipSetPixelOffsetMode gSmall, 4
    GdipGraphicsClear gSmall, Argb(bg)

    If topPx >= 0 And itemHPx > 0 Then
        y0 = topPx - rcPx - 2
        sh = itemHPx + 2 * rcPx + 4
        GdipCreateBitmapFromScan0 wPx * SS, sh * SS, 0, PF32ARGB, 0, big
        GdipGetImageGraphicsContext big, gBig
        GdipSetSmoothingMode gBig, 4
        GdipSetPixelOffsetMode gBig, 4
        GdipGraphicsClear gBig, Argb(bg)
        W = visW * SS                ' sichtbarer rechter Rand: Kurve endet senkrecht genau dort
        T = (topPx - y0) * SS
        H = itemHPx * SS
        L = insetPx * SS
        r = H / 2!
        rc = rcPx * SS
        GdipCreateSolidFill Argb(act), brAct
        GdipCreatePath 0, path
        ' oberer konkaver Uebergang
        GdipAddPathArc path, W - 2 * rc, T - 2 * rc, 2 * rc, 2 * rc, 0!, 90!
        ' links oben rund
        GdipAddPathArc path, L, T, 2 * r, 2 * r, 270!, -90!
        ' links unten rund
        GdipAddPathArc path, L, T + H - 2 * r, 2 * r, 2 * r, 180!, -90!
        ' unterer konkaver Uebergang
        GdipAddPathArc path, W - 2 * rc, T + H, 2 * rc, 2 * rc, 270!, 90!
        ' rechts davon (unsichtbar) bis zum Bildrand fuellen: keine Mischkante am sichtbaren Rand
        GdipAddPathLine path, W, T + H + rc, wPx * SS + SS, T + H + rc
        GdipAddPathLine path, wPx * SS + SS, T + H + rc, wPx * SS + SS, T - rc
        GdipClosePathFigure path
        GdipFillPath gBig, brAct, path
        GdipDeletePath path
        GdipDeleteBrush brAct
        ' hochwertig verkleinert an die Stelle des Streifens
        DrawStrip gSmall, big, 0, y0, wPx, sh
        GdipDeleteGraphics gBig
        GdipDisposeImage big
    End If

    ' Schiebeschalter (Pille mit rundem Knopf); tgX < 0 = keiner
    If tgX >= 0 And tgW > tgH And tgH > 4 Then
        ' eigener kleiner Streifen nur um den Schalter (ueberdeckt keine Aktivmarkierung)
        y0 = tgY - 2
        sh = tgH + 4
        GdipCreateBitmapFromScan0 (tgW + 4) * SS, sh * SS, 0, PF32ARGB, 0, big
        GdipGetImageGraphicsContext big, gBig
        GdipSetSmoothingMode gBig, 4
        GdipSetPixelOffsetMode gBig, 4
        GdipGraphicsClear gBig, Argb(bg)
        d = tgH * SS
        GdipCreateSolidFill Argb(IIf(tgOn, tgColor, offC)), brT
        GdipCreatePath 0, path
        GdipAddPathArc path, 2 * SS, 2 * SS, d, d, 90!, 180!
        GdipAddPathArc path, (tgW + 2) * SS - d, 2 * SS, d, d, 270!, 180!
        GdipClosePathFigure path
        GdipFillPath gBig, brT, path
        GdipDeletePath path
        GdipDeleteBrush brT
        GdipCreateSolidFill Argb(RGB(255, 255, 255)), brT
        GdipCreatePath 0, path
        If tgOn Then
            GdipAddPathArc path, (tgW + 2) * SS - d + 2 * SS, 4 * SS, d - 4 * SS, d - 4 * SS, 0!, 360!
        Else
            GdipAddPathArc path, 4 * SS, 4 * SS, d - 4 * SS, d - 4 * SS, 0!, 360!
        End If
        GdipFillPath gBig, brT, path
        GdipDeletePath path
        GdipDeleteBrush brT
        DrawStrip gSmall, big, tgX - 2, y0, tgW + 4, sh
        GdipDeleteGraphics gBig
        GdipDisposeImage big
    End If

    GdipCreateHBITMAPFromBitmap small, hbm, Argb(bg)

    GdipDeleteGraphics gSmall
    GdipDisposeImage small
    GdiplusShutdown token

    If hbm <> 0 Then Set NavPicture = HbmToPicture(hbm)
End Function

Private Function HbmToPicture(ByVal hbm As LongPtr) As IPictureDisp
    Dim pd As PICTDESC, iid As GUIDT, pic As IPicture
    pd.cbSizeofstruct = LenB(pd)
    pd.picType = 1
    pd.hImage = hbm
    ' IID_IPicture {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    iid.Data1 = &H7BF80980
    iid.Data2 = &HBF32
    iid.Data3 = &H101A
    iid.Data4(0) = &H8B: iid.Data4(1) = &HBB: iid.Data4(2) = &H0: iid.Data4(3) = &HAA
    iid.Data4(4) = &H0: iid.Data4(5) = &H30: iid.Data4(6) = &HC: iid.Data4(7) = &HAB
    If OleCreatePictureIndirect(pd, iid, 1, pic) = 0 Then Set HbmToPicture = pic
End Function
