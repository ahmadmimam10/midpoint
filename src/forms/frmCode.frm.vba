'==============================================================================
' frmCode - Lesbare Codeansicht des eingebetteten Codekatalogs mit Modulliste.
'==============================================================================
Option Explicit

Private WithEvents lst As MSForms.ListBox
Private WithEvents btnClose As MSForms.CommandButton
Private txt As MSForms.TextBox
Private mNames() As String

Public Sub Setup()
    Dim m As Variant, n As Long, w As Single, h As Single
    mUI.StyleForm Me
    w = mUI.MaxFormWidth()
    If w > 1100 Then w = 1100
    h = mUI.MaxFormHeight()
    If h > 720 Then h = 720
    Me.Width = w: Me.Height = h
    Me.caption = "Quellcode: " & APP_TITLE & " " & PROG_VERSION
    mUI.AddLabel Me, "t", "Quellcode (" & FmtInt(mCodeCat.TotalLines) & " Zeilen, Katalog " & mCodeCat.CatalogVersion() & ")", PAD, PAD - 4, 600, 20, True, 12
    Set lst = mUI.AddList(Me, "lst", PAD, PAD + 22, 220, Me.InsideHeight - PAD * 2 - 60)
    ReDim mNames(0 To mCodeCat.Modules().Count)
    For Each m In mCodeCat.Modules()
        lst.AddItem m(0) & "  (" & m(2) & ")"
        mNames(n) = m(0)
        n = n + 1
    Next
    Set txt = mUI.AddText(Me, "code", PAD + 230, PAD + 22, Me.InsideWidth - PAD * 2 - 230, Me.InsideHeight - PAD * 2 - 60, True)
    txt.Font.Name = "Consolas"
    txt.Font.Size = 9
    txt.Locked = True
    txt.ScrollBars = fmScrollBarsBoth
    txt.WordWrap = False
    Set btnClose = mUI.AddButton(Me, "close", "Schließen", Me.InsideWidth - PAD - 110, Me.InsideHeight - PAD - BTN_H, 110, "primary")
    btnClose.Cancel = True
    If lst.ListCount > 0 Then lst.ListIndex = 0
End Sub

Private Sub lst_Change()
    If lst.ListIndex >= 0 Then
        txt.Text = mCodeCat.ModuleText(mNames(lst.ListIndex))
        txt.SelStart = 0
    End If
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub
