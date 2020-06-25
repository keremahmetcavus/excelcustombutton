Attribute VB_Name = "Functions"
Option Explicit
Public col_button As Collection
Public userform_index As UserForm
Public temp_ctl
Sub Run()
Main_Page.Show
End Sub
Function Button_Hover()
           Set col_button = New Collection
            For Each temp_ctl In userform_index.Controls
                If Left(temp_ctl.Name, 6) = "button" And TypeName(temp_ctl) = "Image" Then
                    col_button.Add getHover(temp_ctl)
                End If
            Next
            col_button.Add getHover(userform_index)
End Function
Private Function getHover(temp_ctl) '' this func set hover as image or userform.
    Dim temp_hover As New hover
    If Not TypeName(temp_ctl) = "Image" Then
    Set temp_hover.btnform = temp_ctl
    Else
        If TypeName(temp_ctl) = "Image" And Right(temp_ctl.Name, 5) = "hover" Then
         Set temp_hover.btnimg_hover = temp_ctl
        ElseIf TypeName(temp_ctl) = "Image" Then
         Set temp_hover.btnimg = temp_ctl
        End If
    End If
    Set getHover = temp_hover
End Function
