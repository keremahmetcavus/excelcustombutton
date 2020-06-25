VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Main_Page 
   Caption         =   "Anasayfa"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5955
   OleObjectBlob   =   "Main_Page.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Main_Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Activate()
    Set userform_index = Me
    Call Functions.Button_Hover
End Sub

Private Sub button3_hover_Click()
    Unload Me
    Secd_Page.Show
End Sub
