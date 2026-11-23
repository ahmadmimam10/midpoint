VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEvt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cEvt - Ereignisweiterleitung fuer zur Laufzeit erzeugte Formular-Steuerelemente.
' Der Besitzer (UserForm) muss Public Sub HandleEvent(key, evt, arg) besitzen.
' Referenzen werden mit Unbind gezielt freigegeben (keine haengenden Senken).
'==============================================================================
Option Explicit

Public key As String
Private mOwner As Object
Private WithEvents lbl As MSForms.Label
Private WithEvents btn As MSForms.CommandButton
Private WithEvents txt As MSForms.TextBox
Private WithEvents cbo As MSForms.ComboBox
Private WithEvents lst As MSForms.ListBox
Private WithEvents chk As MSForms.CheckBox
Private WithEvents opt As MSForms.OptionButton
Private WithEvents img As MSForms.Image

Public Sub Bind(ByVal owner As Object, ByVal ctl As Object, ByVal k As String)
    Set mOwner = owner
    key = k
    Select Case TypeName(ctl)
        Case "Label": Set lbl = ctl
        Case "CommandButton": Set btn = ctl
        Case "TextBox": Set txt = ctl
        Case "ComboBox": Set cbo = ctl
        Case "ListBox": Set lst = ctl
        Case "CheckBox": Set chk = ctl
        Case "OptionButton": Set opt = ctl
        Case "Image": Set img = ctl
    End Select
End Sub

Public Sub Unbind()
    Set lbl = Nothing: Set btn = Nothing: Set txt = Nothing: Set cbo = Nothing
    Set lst = Nothing: Set chk = Nothing: Set opt = Nothing: Set img = Nothing
    Set mOwner = Nothing
End Sub

Private Sub Fire(ByVal evt As String, Optional ByVal arg As Variant)
    If mOwner Is Nothing Then Exit Sub
    mOwner.HandleEvent key, evt, arg
End Sub

Private Sub lbl_Click(): Fire "Click": End Sub
Private Sub lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Fire "Move": End Sub
Private Sub lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Fire "DblClick": End Sub
Private Sub btn_Click(): Fire "Click": End Sub
Private Sub txt_Change(): Fire "Change": End Sub
Private Sub txt_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Fire "KeyUp", CLng(KeyCode): End Sub
Private Sub txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not txt.MultiLine Then Fire "Enter"
    If KeyCode = vbKeyDown Then Fire "Down"
End Sub
Private Sub txt_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Fire "MouseUp": End Sub
Private Sub txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Fire "DblClick": End Sub
Private Sub cbo_Change(): Fire "Change": End Sub
Private Sub cbo_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Fire "KeyUp", CLng(KeyCode): End Sub
Private Sub lst_Click(): Fire "Click": End Sub
Private Sub lst_Change(): Fire "Change": End Sub
Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Fire "DblClick": End Sub
Private Sub lst_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Fire "KeyUp", CLng(KeyCode): End Sub
Private Sub chk_Click(): Fire "Click": End Sub
Private Sub chk_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Fire "KeyUp", CLng(KeyCode): End Sub
Private Sub opt_Click(): Fire "Click": End Sub
Private Sub img_Click(): Fire "Click": End Sub
Private Sub img_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Fire "DblClick": End Sub
