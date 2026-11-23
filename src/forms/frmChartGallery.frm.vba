'==============================================================================
' frmChartGallery - "Welches Diagramm passt?": alle Darstellungen als Kacheln
' mit Vorschaubild (Beispieldaten, nicht Ihre Daten), Name und Einsatzzweck.
' Klick wählt, Doppelklick übernimmt.
'==============================================================================
Option Explicit

Public Result As String
Public OK As Boolean
Private mTypes As Variant
Private mSel As Long
Private mEvts As Collection
Private WithEvents btnOk As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton

Private Const TILE_W As Single = 176
Private Const TILE_H As Single = 172
Private Const GAP As Single = 10

Public Sub Setup(ByVal current As String)
    Dim i As Long, nCol As Long, x As Single, y As Single, w As Single, y0 As Single, rows_ As Long
    Set mEvts = New Collection
    mTypes = mDash.ElemTypes()
    mUI.StyleForm Me
    nCol = 5
    If PAD * 2 + nCol * TILE_W + (nCol - 1) * GAP + 12 > mUI.MaxFormWidth() Then nCol = 4
    If PAD * 2 + nCol * TILE_W + (nCol - 1) * GAP + 12 > mUI.MaxFormWidth() Then nCol = 3
    Me.Width = PAD * 2 + nCol * TILE_W + (nCol - 1) * GAP + (Me.Width - Me.InsideWidth)
    Me.caption = "Welches Diagramm passt?"
    w = Me.InsideWidth - PAD * 2
    mUI.AddLabel Me, "t", "Welches Diagramm passt zu meiner Frage?", PAD, PAD - 2, w, 20, True, 12
    mUI.AddLabel Me, "t2", "Die Bilder zeigen Beispieldaten. So sieht die jeweilige Darstellung aus. Klick wählt, Doppelklick übernimmt.", _
                 PAD, PAD + 20, w, 16, False, 8.5, CMuted()
    y0 = PAD + 44
    mSel = 0
    For i = 0 To UBound(mTypes)
        x = PAD + (i Mod nCol) * (TILE_W + GAP)
        y = y0 + (i \ nCol) * (TILE_H + GAP)
        AddTile i, x, y
        If CStr(mTypes(i)) = current Then mSel = i
    Next
    rows_ = -Int(-(UBound(mTypes) + 1) / nCol)
    y = y0 + rows_ * (TILE_H + GAP) + 4
    Set btnOk = mUI.AddButton(Me, "ok", "Diese Darstellung wählen", PAD, y, 190, "primary")
    Set btnCancel = mUI.AddButton(Me, "cancel", "Abbrechen", Me.InsideWidth - PAD - 104, y, 104)
    btnCancel.Cancel = True
    Me.Height = y + BTN_H + PAD + (Me.Height - Me.InsideHeight)
    If Me.Height > mUI.MaxFormHeight() Then
        Me.Height = mUI.MaxFormHeight()
        Me.ScrollBars = fmScrollBarsVertical
        Me.ScrollHeight = y + BTN_H + PAD
        Me.ScrollTop = 0
    End If
    MarkSel
End Sub

Private Sub AddTile(ByVal i As Long, ByVal x As Single, ByVal y As Single)
    Dim bg As MSForms.Label, lbl As MSForms.Label, img As MSForms.Image, typ As String, pth As String
    typ = CStr(mTypes(i))
    Set bg = mUI.AddLabel(Me, "bg" & i, "", x, y, TILE_W, TILE_H)
    bg.BackStyle = fmBackStyleOpaque
    bg.BackColor = CWhite()
    bg.BorderStyle = fmBorderStyleSingle
    bg.BorderColor = CBorder()
    Bind bg, i
    Set img = Me.Controls.Add("Forms.Image.1", "im" & i)
    img.Left = x + 6: img.Top = y + 6: img.Width = TILE_W - 12: img.Height = 98
    img.BorderStyle = fmBorderStyleNone
    img.BackStyle = fmBackStyleTransparent
    img.PictureSizeMode = fmPictureSizeModeZoom
    pth = mGallery.PreviewPath(typ)
    If Len(pth) > 0 Then
        On Error Resume Next
        Set img.Picture = LoadPicture(pth)
        On Error GoTo 0
    End If
    Bind img, i
    Set lbl = mUI.AddLabel(Me, "nm" & i, typ, x + 10, y + 106, TILE_W - 20, 16, True, 10)
    Bind lbl, i
    Set lbl = mUI.AddLabel(Me, "hp" & i, mGallery.TypeUse(typ), x + 10, y + 123, TILE_W - 20, 44, False, 8, CMuted())
    lbl.WordWrap = True
    Bind lbl, i
End Sub

Private Sub Bind(ByVal c As Object, ByVal i As Long)
    Dim e As cEvt
    Set e = New cEvt
    e.Bind Me, c, CStr(i)
    mEvts.Add e
End Sub

Private Sub MarkSel()
    Dim i As Long
    For i = 0 To UBound(mTypes)
        With Me.Controls("bg" & i)
            .BackColor = IIf(i = mSel, CAccentSoft(), CWhite())
            .BorderColor = IIf(i = mSel, CAccent(), CBorder())
        End With
    Next
End Sub

Public Sub HandleEvent(ByVal key As String, ByVal evt As String, Optional ByVal arg As Variant)
    Dim i As Long
    i = CLng(Val(key))
    If i < 0 Or i > UBound(mTypes) Then Exit Sub
    If evt = "Click" Then
        mSel = i
        MarkSel
    ElseIf evt = "DblClick" Then
        mSel = i
        btnOk_Click
    End If
End Sub

Private Sub btnOk_Click()
    Result = CStr(mTypes(mSel))
    OK = True
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    OK = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True: btnCancel_Click
End Sub

Private Sub UserForm_Terminate()
    Dim e As cEvt
    If Not mEvts Is Nothing Then
        For Each e In mEvts: e.Unbind: Next
    End If
End Sub
