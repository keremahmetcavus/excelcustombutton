VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "hover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents btnimg As MSForms.Image
Attribute btnimg.VB_VarHelpID = -1
Public WithEvents btnimg_hover As MSForms.Image
Attribute btnimg_hover.VB_VarHelpID = -1
Public WithEvents btnform As MSForms.UserForm
Attribute btnform.VB_VarHelpID = -1
Public button_parent_index
Public button_index
Public button_count
Private Sub btnimg_hover_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    button_index = btnimg_hover.Name
     userform_index.Controls(button_index).Move userform_index.Controls(button_index).Left + 2, userform_index.Controls(button_index).Top + 2
End Sub
Private Sub btnimg_hover_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     button_index = btnimg_hover.Name
     button_parent_index = Left(btnimg_hover.Name, Len(btnimg_hover.Name) - 6)
     userform_index.Controls(button_index).Move userform_index.Controls(button_parent_index).Left, userform_index.Controls(button_parent_index).Top
End Sub
Sub btnimg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo Message
    button_index = btnimg.Name
    userform_index.Controls(button_index).Visible = False
    userform_index.Controls(button_index & "_hover").Visible = True
    Exit Sub
Message:
    userform_index.Controls(button_index).Visible = True
    MsgBox "Lütfen " & button_index & " objenize " & button_index & "_hover isminde Image Control objesi ekleyiniz."

End Sub
Sub btnform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    button_count = ((col_button.Count) - 1) / 2
    
    If button_count = 0 Then
    MsgBox "Lütfen Image Tipinde ve button(i) isminde Control objesi ekleyiniz."
    Unload userform_index
    Exit Sub
    End If
    
    For button_index = 1 To button_count
        userform_index.Controls("button" & button_index).Visible = True
        userform_index.Controls("button" & button_index & "_hover").Visible = False
    Next
End Sub

