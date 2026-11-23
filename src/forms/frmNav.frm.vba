'==============================================================================
' frmNav - Feste linke Navigation, als Kindfenster im Arbeitsmappenfenster
' verankert (bleibt beim Scrollen stehen, ohne Fensterfixierung). Die aktive
' Markierung wird per GDI+ als glatte offene Form gezeichnet.
' Seitenwechsel bauen die Eintraege nicht neu auf (nur Markierung umsetzen);
' neu aufgebaut wird nur, wenn sich Eintraege, Titel, Farben oder Groesse aendern.
'==============================================================================
Option Explicit

Private mItems As Collection        ' Array(key, icon, text, kind)
Private mEvts As Collection
Private mTops() As Single
Private mKeys() As String
Private mN As Long
Private mWpt As Single, mHpt As Single, mWpx As Long, mHpx As Long, mHFullPx As Long, mHFullPt As Single
Private mIH As Single
Private mActiveKey As String
Private mHover As Long
Private mContentH As Single
Private mToggleY As Single
Private mSecH As Single
Private mOff As Single              ' eigener Bildlauf (Punkte)
Private mUserScrolled As Boolean    ' mit den Pfeilen selbst geblättert (aktiven Eintrag dann nicht nachführen)
Private mBottomY As Single, mMainEnd As Single, mNeed As Single, mLvl As Long
Private mLastActive As String
Private mBuilding As Boolean        ' Neuaufbau läuft (Zeichnen ist bis zum Ende ausgesetzt)
Private mToggleOn As Boolean
Private mSig As String              ' Stand der Eintraege/Farben beim letzten Aufbau
Private mTxH As Single, mIcH As Single, mSizedFor As Long, mBrandH As Single, mSubH As Single

Private Const TOP0 As Single = 74
Private Const TX_SIZE As Single = 9.75
Private Const IC_SIZE As Single = 11

Private Sub UserForm_Initialize()
    Me.BackColor = CNavBg()
    Me.Font.Name = UI_FONT
    Set mEvts = New Collection
    mIH = 30
End Sub

Public Sub Teardown()
    Dim e As cEvt
    If Not mEvts Is Nothing Then
        For Each e In mEvts
            e.Unbind
        Next
    End If
    Set mEvts = New Collection
End Sub

' Groesse in Punkten so, wie das Formular sie selbst sieht (nach dem Setzen der Fenstergroesse):
' MSForms rechnet bei manchen Skalierungen (z. B. 175 %) nicht genau mit der Bildschirm-DPI
' hPx = sichtbare Hoehe fuer die Eintraege, hFullPx = Fensterhoehe (bis unter die Bildlaufleiste)
Public Sub Relayout(ByVal wPx As Long, ByVal hPx As Long, ByVal hFullPx As Long)
    mWpx = wPx: mHpx = hPx: mHFullPx = hFullPx
    MeasureSize
    BuildItems
End Sub

Private Sub MeasureSize()
    Dim w As Single, h As Single
    On Error Resume Next
    w = Me.InsideWidth: h = Me.InsideHeight
    On Error GoTo 0
    If w < 40 Or h < 40 Then
        w = mWin.PxToPt(mWpx): h = mWin.PxToPt(mHpx)
    ElseIf mHFullPx > 0 Then
        h = h * mHpx / mHFullPx
    End If
    mWpt = w: mHpt = h
    mHFullPt = Me.InsideHeight
End Sub

' Aktualisieren nach Seitenwechsel/Aenderung: neu aufbauen nur bei geaendertem Stand,
' sonst nur die Aktivmarkierung umsetzen (schnell, ohne Flackern)
Public Sub Update(Optional ByVal force As Boolean = False)
    If mWpt = 0 Then Exit Sub
    If force Or mNav.NavSignature() <> mSig Then
        BuildItems
    Else
        SetActive mNav.ActiveNavKey()
    End If
End Sub

Private Sub SetActive(ByVal key As String)
    Dim i As Long, oldI As Long, newI As Long
    If key = mActiveKey Then Exit Sub
    For i = 1 To mN
        If mKeys(i) = mActiveKey And Len(mActiveKey) > 0 Then oldI = i
        If mKeys(i) = key And Len(key) > 0 Then newI = i
    Next
    ' mit Bildlauf: aktiven Eintrag ggf. sichtbar machen (dann Neuaufbau)
    If newI > 0 And mContentH > mHpt + 2 Then
        If mTops(newI) < 20 Or mTops(newI) + mIH > mHpt - 20 Then
            mActiveKey = key
            BuildItems
            Exit Sub
        End If
    End If
    If oldI > 0 Then StyleItem oldI, False
    If newI > 0 Then StyleItem newI, True
    mActiveKey = key
    mLastActive = key
    mHover = 0
    RenderPicture
End Sub

Private Sub StyleItem(ByVal i As Long, ByVal active As Boolean)
    With Me.Controls("ic" & i)
        .ForeColor = IIf(active, CAccent(), CNavText())
    End With
    With Me.Controls("tx" & i)
        .Font.Name = IIf(active, UI_FONT_SEMI, UI_FONT)
        .Font.Size = TX_SIZE
        .ForeColor = IIf(active, CText(), CNavText())
    End With
End Sub

' Hoehenaenderung ohne Neuaufbau (sicher auch waehrend eines Maus-Ereignisses eines
' Eintrags): nur die untere Gruppe und den Schalter verschieben. False = Neuaufbau noetig.
Public Function TryShift(ByVal hPx As Long, ByVal hFullPx As Long) As Boolean
    Dim delta As Single, c As MSForms.Control, k As Long, hPt As Single
    hPt = Me.InsideHeight
    If hFullPx > 0 Then hPt = hPt * hPx / hFullPx
    If hPt < 40 Then hPt = mWin.PxToPt(hPx)
    If mWpt = 0 Or mOff > 0 Or mContentH > mHpt + 2 Then Exit Function
    If mNeed > hPt Then Exit Function
    delta = hPt - mHpt
    If mBottomY + delta < mMainEnd + 12 Then Exit Function
    For Each c In Me.Controls
        If c.Top >= mBottomY - 1 Then c.Top = c.Top + delta
    Next
    For k = 1 To mN
        If mTops(k) >= mBottomY - 1 Then mTops(k) = mTops(k) + delta
    Next
    If mToggleY >= 0 Then mToggleY = mToggleY + delta
    mBottomY = mBottomY + delta
    mContentH = mContentH + delta
    mHpt = hPt: mHpx = hPx
    RenderPicture
    TryShift = True
End Function

' Tatsaechliche Zeilenhoehe einer Schrift in einem Label (fuer exakte senkrechte Mitte)
Private Function LineHeight(ByVal fontName As String, ByVal size As Single, ByVal sample As String) As Single
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", "probe")
    lbl.Font.Name = fontName
    lbl.Font.Size = size
    lbl.WordWrap = False
    lbl.caption = sample
    lbl.AutoSize = True
    LineHeight = lbl.Height
    Me.Controls.Remove "probe"
End Function

Public Sub BuildItems()
    Dim it As Variant, i As Long, y As Single, bottomStart As Long, nBottom As Long, yBottom As Single
    Dim lbl As MSForms.Label, e As cEvt, kind As String, need As Single, hwnd As LongPtr
    If mWpt = 0 Then Exit Sub
    hwnd = mNav.NavHwnd()
    mWin.SetRedraw hwnd, False
    mBuilding = True
    On Error GoTo done
    Teardown
    Me.Controls.Clear
    Me.BackColor = CNavBg()
    Set mItems = mNav.NavItems()
    mActiveKey = mNav.ActiveNavKey()
    mSig = mNav.NavSignature()
    mN = 0
    mHover = 0
    ReDim mTops(1 To mItems.Count + 1)
    ReDim mKeys(1 To mItems.Count + 1)
    If mSizedFor <> mWin.ScreenDpi() Then
        mTxH = LineHeight(UI_FONT, TX_SIZE, "Hg")
        mIcH = LineHeight(ICON_FONT, IC_SIZE, ChrW(&HE80F))
        mBrandH = LineHeight(UI_FONT, 12, "Hg")
        mSubH = LineHeight(UI_FONT, 8.5, "Hg")
        mSizedFor = mWin.ScreenDpi()
    End If

    ' Markenbereich: MidPoint-Symbol links, daneben Titel und Untertitel der Datei
    Dim bx As Single, pic As Object, img As MSForms.Image
    bx = 20
    Set pic = mLogo.NavPicture(CNavBg())
    If Not pic Is Nothing Then
        Set img = Me.Controls.Add("Forms.Image.1", "brandlogo")
        img.Left = 16: img.Top = 18: img.Width = 30: img.Height = 29
        img.BorderStyle = fmBorderStyleNone
        img.BackStyle = fmBackStyleTransparent
        img.PictureSizeMode = fmPictureSizeModeZoom
        Set img.Picture = pic
        bx = 54
    End If
    Set lbl = mUI.AddLabel(Me, "brand", mCfg.SysGet("AppName", APP_TITLE), bx, 16, mWpt - bx - 10, mBrandH, True, 12, CNavStrong())
    lbl.WordWrap = False
    Set lbl = mUI.AddLabel(Me, "brand2", mCfg.SysGet("Untertitel", "Offline-Datenbank in Excel"), bx, 37, mWpt - bx - 10, mSubH, False, 8.5, CNavMuted())
    lbl.WordWrap = False
    ' Titel jederzeit bearbeitbar: Klick auf den Markenbereich
    Set lbl = mUI.AddLabel(Me, "brandhit", "", 8, 10, mWpt - 8, 46)
    lbl.BackStyle = fmBackStyleTransparent
    lbl.ControlTipText = "Klicken, um Titel und Untertitel dieser Datei zu ändern"
    Set e = New cEvt
    e.Bind Me, lbl, "TITLE"
    mEvts.Add e

    ' Platzbedarf pruefen: bei wenig Hoehe (kleines Fenster, Menueband ausgeklappt)
    ' stufenweise verdichten, erst danach Bildlauf
    Dim nItem As Long, nSec As Long, lvl As Long, yTop As Single
    For Each it In mItems
        Select Case it(3)
            Case "item": nItem = nItem + 1
            Case "section": nSec = nSec + 1
        End Select
    Next
    For lvl = 0 To 2
        mIH = Choose(lvl + 1, 30, 26, 22)
        mSecH = Choose(lvl + 1, 22, 19, 16)
        yTop = Choose(lvl + 1, TOP0, 66, 52)
        need = yTop + nItem * mIH + nSec * mSecH + 12 + 46
        If need <= mHpt Then Exit For
    Next
    If lvl > 2 Then lvl = 2
    If lvl = 2 Then Me.Controls("brand2").Visible = False
    ' Symbol und Text senkrecht aufeinander ausrichten: ohne Untertitel steht der Titel mittig
    ' neben dem Symbol, mit Untertitel die Mitte beider Zeilen
    If Not img Is Nothing Then
        If lvl = 2 Then
            img.Top = 14
            Me.Controls("brand").Top = img.Top + img.Height / 2 - mBrandH / 2
        Else
            Me.Controls("brand").Top = 16
            Me.Controls("brand2").Top = 16 + mBrandH + 1
            img.Top = (16 + (16 + mBrandH + 1 + mSubH)) / 2 - img.Height / 2
        End If
    End If

    ' Hauptgruppe
    y = yTop
    i = 0
    For Each it In mItems
        i = i + 1
        kind = it(3)
        If kind = "bottom" Then bottomStart = i + 1: Exit For
        If kind = "section" Then
            Set lbl = mUI.AddLabel(Me, "sec" & i, CStr(it(2)), 22, y + mSecH - 16, mWpt - 40, 14, True, 7.5, CNavMuted())
            y = y + mSecH
        Else
            AddItem it, y
            y = y + mIH
        End If
    Next
    ' Untere Gruppe (unten verankert, wenn Platz)
    nBottom = mItems.Count - bottomStart + 1
    yBottom = mHpt - IIf(lvl = 0, 58, 46) - nBottom * mIH
    If yBottom < y + 12 Then yBottom = y + 12
    mBottomY = yBottom: mMainEnd = y: mNeed = need: mLvl = lvl
    i = 0
    For Each it In mItems
        i = i + 1
        If i >= bottomStart Then
            AddItem it, yBottom
            yBottom = yBottom + mIH
        End If
    Next
    ' Modusschalter
    AddModeRow yBottom + 6
    mContentH = yBottom + 40
    ' Zu wenig Hoehe: eigener Bildlauf mit Pfeilen (die MSForms-Bildlaufleiste wird im
    ' angedockten Fenster nicht gezeichnet); der aktive Eintrag bleibt sichtbar
    Me.ScrollBars = fmScrollBarsNone
    Me.ScrollTop = 0
    If mContentH > mHpt + 2 Then
        Dim maxOff As Single, k As Long, c As MSForms.Control
        maxOff = mContentH - mHpt + 8
        ' aktiven Eintrag sichtbar halten (auch nach Größenänderung), außer man blättert selbst
        If mActiveKey <> mLastActive Then mUserScrolled = False
        If Not mUserScrolled Then
            For k = 1 To mN
                If mKeys(k) = mActiveKey Then
                    If mTops(k) - 24 < mOff Then mOff = mTops(k) - 24
                    If mTops(k) + mIH + 26 > mOff + mHpt Then mOff = mTops(k) + mIH + 26 - mHpt
                End If
            Next
        End If
        If mOff > maxOff Then mOff = maxOff
        If mOff < 0 Then mOff = 0
        If mOff > 0 Then
            For Each c In Me.Controls
                c.Top = c.Top - mOff
            Next
            For k = 1 To mN
                mTops(k) = mTops(k) - mOff
            Next
            If mToggleY >= 0 Then mToggleY = mToggleY - mOff
        End If
        ' Was nur halb neben einem Pfeilstreifen stünde, ausblenden (kein halber Text). Einträge
        ' als Ganzes (Symbol, Text, Klickfläche), damit nichts Sichtbares unanklickbar wird.
        For Each c In Me.Controls
            Select Case True
                Case c.Name Like "ic#*", c.Name Like "tx#*", c.Name Like "hit#*", c.Name = "modehit"
                Case UnderArrow(c.Top, c.Height, maxOff)
                    c.Visible = False
                    If c.Name = "mode" Then
                        mToggleY = -1
                        If Not mApp.gReadOnly Then Me.Controls("modehit").Visible = False
                    End If
            End Select
        Next
        For k = 1 To mN
            Set c = Me.Controls("tx" & k)
            If UnderArrow(c.Top, c.Height, maxOff) Then
                c.Visible = False
                Me.Controls("ic" & k).Visible = False
                Me.Controls("hit" & k).Visible = False
            End If
        Next
        If mOff > 0 Then AddScrollArrow "UP", ChrW(&HE70E), 0
        If mOff < maxOff Then AddScrollArrow "DOWN", ChrW(&HE70D), mHpt - 18
    Else
        mOff = 0
    End If
    mLastActive = mActiveKey
    RenderPicture
done:
    mBuilding = False
    mWin.SetRedraw hwnd, True
End Sub

' Liegt ein Steuerelement teilweise unter einem Pfeilstreifen (oben bzw. unten, je 18 pt)?
Private Function UnderArrow(ByVal t As Single, ByVal h As Single, ByVal maxOff As Single) As Boolean
    UnderArrow = (mOff > 0 And t < 17) Or (mOff < maxOff And t + h > mHpt - 17)
End Function

Private Sub AddScrollArrow(ByVal key As String, ByVal icon As String, ByVal y As Single)
    Dim lbl As MSForms.Label, e As cEvt
    Set lbl = mUI.AddLabel(Me, "sc" & key, icon, 0, y, mWpt, 18, False, 9, CNavStrong())
    lbl.Font.Name = ICON_FONT
    lbl.TextAlign = fmTextAlignCenter
    lbl.BackStyle = fmBackStyleOpaque
    lbl.BackColor = Blend(CNavBg(), CNavStrong(), 0.12)
    lbl.ControlTipText = IIf(key = "UP", "Nach oben blättern", "Weitere Einträge: nach unten blättern")
    lbl.ZOrder 0
    Set e = New cEvt
    e.Bind Me, lbl, key
    mEvts.Add e
End Sub

' Symbol und Text exakt auf die senkrechte Mitte der Markierung (Eintragsmitte) setzen
Private Sub AddItem(ByVal it As Variant, ByVal y As Single)
    Dim ic As MSForms.Label, tx As MSForms.Label, hit As MSForms.Label, e As cEvt, active As Boolean, cy As Single
    mN = mN + 1
    mTops(mN) = y
    mKeys(mN) = CStr(it(0))
    active = (CStr(it(0)) = mActiveKey And Len(mActiveKey) > 0)
    cy = y + mIH / 2
    Set ic = mUI.AddLabel(Me, "ic" & mN, CStr(it(1)), 22, cy - mIcH / 2, 20, mIcH, False, IC_SIZE, IIf(active, CAccent(), CNavText()))
    ic.Font.Name = ICON_FONT
    ic.WordWrap = False
    ' Segoe UI: Grossbuchstaben sitzen etwas unter der Zeilenmitte -> minimal hoeher setzen
    Set tx = mUI.AddLabel(Me, "tx" & mN, CStr(it(2)), 48, cy - mTxH / 2 - 0.5, mWpt - 58, mTxH, active, TX_SIZE, IIf(active, CText(), CNavText()))
    tx.WordWrap = False
    ' Langer Text (fett, hohe Skalierung): Schrift minimal verkleinern statt rechts abschneiden
    Dim fs As Single
    fs = TX_SIZE
    tx.AutoSize = True
    Do While tx.Width > mWpt - 52 And fs > TX_SIZE - 2
        fs = fs - 0.5
        tx.Font.Size = fs
    Loop
    tx.AutoSize = False
    tx.Width = mWpt - 50: tx.Height = mTxH
    tx.Top = cy - mTxH / 2 - 0.5
    Set hit = mUI.AddLabel(Me, "hit" & mN, "", 8, y, mWpt - 8, mIH)
    hit.BackStyle = fmBackStyleTransparent
    hit.ControlTipText = Replace(CStr(it(2)), "&", "&&")    ' sonst zeigt der Hinweis "Hilfe_Sicherung"
    hit.MousePointer = fmMousePointerArrow
    Set e = New cEvt
    e.Bind Me, hit, CStr(mN)
    mEvts.Add e
End Sub

' Schiebeschalter "Automatisch speichern" (der Schalter selbst wird in RenderPicture gezeichnet)
Private Sub AddModeRow(ByVal y As Single)
    Dim lbl As MSForms.Label, hit As MSForms.Label, e As cEvt, txt As String
    mToggleY = -1
    If mApp.gReadOnly Then
        Set lbl = mUI.AddLabel(Me, "mode", ChrW(&H25CF) & "  Lesemodus", 20, y + 8, mWpt - 28, 16, False, 8.5, CNavMuted())
        lbl.WordWrap = False
        Exit Sub
    End If
    mToggleY = y + 9
    mToggleOn = mBackup.AutoSaveOn()
    txt = "Automatisch speichern"
    Set lbl = mUI.AddLabel(Me, "mode", txt, 56, y + 8, mWpt - 64, 16, False, 8.5, IIf(mToggleOn, CNavStrong(), CNavMuted()))
    lbl.WordWrap = False
    Set hit = mUI.AddLabel(Me, "modehit", "", 8, y, mWpt - 8, 34)
    hit.BackStyle = fmBackStyleTransparent
    hit.ControlTipText = IIf(mToggleOn, "Ein: Jede Änderung wird sofort in der Datei gespeichert. Klicken zum Ausschalten.", _
                                         "Aus: Gespeichert wird mit „Speichern“ bzw. Strg+S. Klicken zum Einschalten.")
    hit.MousePointer = fmMousePointerArrow
    Set e = New cEvt
    e.Bind Me, hit, "MODE"
    mEvts.Add e
End Sub

Private Sub RenderPicture()
    Dim i As Long, topPt As Single, scale_ As Double, pic As IPictureDisp, tgX As Long, tgY As Long, picW As Long, picH As Long
    On Error GoTo nope
    ' Das Bild wird in Punkten des Formulars gezeichnet: Ein Bitmap-Bild erhält seine Größe
    ' über die Bildschirm-DPI; das Formular rechnet sie in seine Punkte um. So sitzen Markierung
    ' und Kurven bei jeder Skalierung genau unter dem Eintrag (statt vergrößert abgeschnitten).
    scale_ = mWin.ScreenDpi() / 72#
    picW = CLng(mWpt * scale_): picH = CLng(mHpt * scale_)
    tgX = -1
    If mToggleY >= 0 Then
        tgX = CLng(20 * scale_)
        tgY = CLng((mToggleY - Me.ScrollTop) * scale_)
    End If
    For i = 1 To mN
        If mKeys(i) = mActiveKey And Len(mActiveKey) > 0 Then
            topPt = mTops(i) - Me.ScrollTop
            Exit For
        End If
    Next
    If i > mN Then
        Set pic = mGdi.NavPicture(picW, picH, CNavBg(), CWhite(), -1, 0, 0, 0, _
                                  tgX, tgY, CLng(28 * scale_), CLng(15 * scale_), mToggleOn, CSuccess())
    Else
        ' konkave Uebergaenge als glatter Viertelkreis; Markierung genau so hoch wie der Eintrag
        ' minus je 1 pt oben/unten, damit Mitte von Markierung und Beschriftung uebereinstimmen
        Set pic = mGdi.NavPicture(picW, picH, CNavBg(), CWhite(), CLng((topPt + 1) * scale_), CLng((mIH - 2) * scale_), _
                                  CLng(10 * scale_), CLng(18 * scale_), tgX, tgY, CLng(28 * scale_), CLng(15 * scale_), mToggleOn, CSuccess())
    End If
    If Not pic Is Nothing Then
        Set Me.Picture = pic
        Me.PictureAlignment = fmPictureAlignmentTopLeft
        Me.PictureSizeMode = fmPictureSizeModeClip
        ' sofort ganz neu zeichnen: sonst zeigt Excel Teile eines älteren Bildes (Markierung und
        ' Schalter wirkten dann verrutscht). Beim Neuaufbau erledigt das SetRedraw am Ende.
        If Not mBuilding Then mWin.Repaint mNav.NavHwnd()
    End If
    Exit Sub
nope:
End Sub

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    Dim i As Long
    If key = "MODE" Or key = "TITLE" Then
        If evt = "Click" Then mNav.NavClick key
        Exit Sub
    End If
    If key = "UP" Or key = "DOWN" Then
        If evt = "Click" Then
            mOff = mOff + IIf(key = "UP", -1, 1) * (mHpt * 0.6)
            mUserScrolled = True
            BuildItems
        End If
        Exit Sub
    End If
    i = CLng(Val(key))
    If i < 1 Or i > mN Then Exit Sub
    Select Case evt
        Case "Click"
            mNav.NavClick mKeys(i)
            mNav.FocusGrid
        Case "Move"
            ' Menueband ein-/ausgeklappt? Hier nur verschieben, nicht neu aufbauen
            mNav.CheckNavSizeThrottled True
            If mHover <> i Then
                If mHover > 0 And mHover <= mN Then
                    If mKeys(mHover) <> mActiveKey Then Me.Controls("tx" & mHover).ForeColor = CNavText()
                End If
                If mKeys(i) <> mActiveKey Then Me.Controls("tx" & i).ForeColor = CNavStrong()
                mHover = i
            End If
    End Select
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If mHover > 0 And mHover <= mN Then
        If mKeys(mHover) <> mActiveKey Then Me.Controls("tx" & mHover).ForeColor = CNavText()
    End If
    mHover = 0
    ' Menueband ein-/ausgeklappt? (Excel meldet das nicht) - Hoehe beim Ueberfahren pruefen
    mNav.CheckNavSizeThrottled
End Sub

Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    RenderPicture
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

Private Sub UserForm_Terminate()
    Teardown
End Sub
