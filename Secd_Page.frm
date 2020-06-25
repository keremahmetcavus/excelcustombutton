VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Secd_Page 
   Caption         =   "Hafta Verisi"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5955
   OleObjectBlob   =   "Secd_Page.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Secd_Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub button2_hover_Click()
    Unload Me
    Main_Page.Show
End Sub
Sub UserForm_Activate()
    Set userform_index = Me
    Call Functions.Button_Hover
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Debug.Print KeyAscii
If KeyAscii >= 48 And KeyAscii <= 57 Then
    Debug.Print "number"
Else
    Debug.Print "other"
    KeyAscii = 0
End If
End Sub


